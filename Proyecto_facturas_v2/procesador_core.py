# procesador_core.py
import os
import re
import shutil
from datetime import datetime
import fitz  # PyMuPDF
import pandas as pd
from config import CARPETA_A, CARPETA_B, CARPETA_C, ARCHIVO_PROVEEDORES, MESES_ESPANOL, REPORTE_MAESTRO_HTML
from pdf_rules import REGLAS_EXTRACCION
from proveedores import cargar_excel_proveedores, buscar_proveedor
from downloader import descargar_facturas_outlook
from reporter import generar_reporte_html, obtener_facturas_procesadas
from proveedor_rules import es_factura_por_remitente, analizar_pdf_ecosan
import oc_manager

# ─── OCR FALLBACK ────────────────────────────────────────────────────────────
def extraer_texto_con_ocr(ruta_pdf: str, max_paginas: int = 3) -> str:
    """Fallback OCR para PDFs escaneados o imagen usando RapidOCR + PyMuPDF (limitado a max_paginas)."""
    print("   📷 Activando OCR (RapidOCR + PyMuPDF)...")
    texto_extraido = []
    try:
        from rapidocr_onnxruntime import RapidOCR
        engine = RapidOCR()
        doc = fitz.open(ruta_pdf)
        num_paginas = min(len(doc), max_paginas)
        for page_num in range(num_paginas):
            page = doc[page_num]
            pix = page.get_pixmap(dpi=200)
            img_bytes = pix.tobytes("png")
            result, _ = engine(img_bytes)
            if result:
                texto_pagina = "\n".join([line[1] for line in result])
                texto_extraido.append(texto_pagina)
        doc.close()
        res = "\n".join(texto_extraido).strip()
        if res:
            print(f"   ✅ RapidOCR exitoso: {len(res)} caracteres extraídos")
            return res
    except Exception as e:
        print(f"   ⚠️ RapidOCR falló: {e}")

    # Fallback secundario pytesseract
    try:
        from pdf2image import convert_from_path
        import pytesseract
        paginas = convert_from_path(ruta_pdf, dpi=200)
        texto = "\n".join([pytesseract.image_to_string(p, lang="spa") for p in paginas])
        return texto.strip()
    except Exception as e:
        print(f"   ⚠️ Error OCR secundario: {e}")
        return ""
# ─────────────────────────────────────────────────────────────────────────────

def es_factura_valida(texto_lower, remitente=""):
    """Valida si el texto corresponde a una factura/NC/ND válida de AFIP."""
    # 1. Remitente de confianza
    if es_factura_por_remitente(remitente):
        return True
        
    # Documentos prohibidos que no son facturas
    documentos_no_factura = [
        "orden de compra", "orden de compras", "vista de deuda",
        "resumen de cuenta", "estado de cuenta", "cuenta corriente de clientes",
        "cuenta corriente", "presupuesto", "nota de pedido", "para cotizar",
        "cotización", "cotizacion", "anulada", "duplicado de recibo",
        "recibo de pago", "certificado de cobertura"
    ]
    
    # Frases secundarias comunes en facturas reales que se limpian antes del chequeo
    texto_seguro = (
        texto_lower.replace("presupuesto económico", "")
        .replace("presupuesto economico", "")
        .replace("orden de compra:", "")
        .replace("orden de compra :", "")
        .replace("orden de compra n", "")
        .replace("orden de compra-", "")
    )
    
    # Chequeo de CAE / CAEA / CAI AFIP
    has_cae_valido = False
    if re.search(r"\bC[.\s]*A[.\s]*E[.\s]*[A\s:\.\|N°º]*\n?\s*\|?\s*(\d{14,15})\b", texto_lower, re.IGNORECASE):
        has_cae_valido = True
    elif re.search(r"\b[678]\d{13}\b", texto_lower):
        has_cae_valido = True
    elif re.search(r"\bcae\b\s*[:\.]?\s*\d{14}", texto_lower, re.IGNORECASE | re.MULTILINE):
        has_cae_valido = True

    # Rechazar si contiene tipo de documento prohibido y NO posee CAE válido
    for doc_tipo in documentos_no_factura:
        if doc_tipo in texto_seguro and not has_cae_valido:
            return False

    # Si tiene un CAE AFIP válido, es un comprobante válido
    if has_cae_valido:
        return True

    # Fallback para comprobante AFIP (emergencia) si dice "Factura" y tiene número de comprobante AFIP
    if re.search(r"\bfactura\b", texto_lower) and re.search(r"\b\d{4,5}[-\s]\d{7,8}\b", texto_lower):
        if "vista de deuda" not in texto_lower and "cuenta corriente" not in texto_lower:
            return True

    return False

def extraer_datos_pdf(ruta_pdf, remitente=""):
    try:
        doc = fitz.open(ruta_pdf)
        num_paginas = len(doc)
        
        # Extraer texto de las primeras 3 páginas (donde radica la factura principal)
        texto_cabecera = "\n".join([doc[i].get_text("text", sort=True) for i in range(min(3, num_paginas))])
        texto_cabecera_lower = texto_cabecera.lower()

        texto_original = "\n".join([pagina.get_text("text", sort=True) for pagina in doc])
        texto_lower = texto_original.lower()
        doc.close()

        ocr_ejecutado = False
        es_valida = es_factura_valida(texto_cabecera_lower, remitente) or es_factura_valida(texto_lower, remitente)

        if not es_valida and (len(texto_cabecera_lower.strip()) < 50 or len(texto_lower.strip()) < 50):
            print(f"   ℹ️ Texto escaso o sin validación previa. Intentando OCR automático...")
            texto_ocr = extraer_texto_con_ocr(ruta_pdf)
            if texto_ocr.strip():
                texto_original = texto_original + "\n" + texto_ocr
                texto_lower = texto_original.lower()
                texto_cabecera = texto_cabecera + "\n" + texto_ocr
                texto_cabecera_lower = texto_cabecera.lower()
                ocr_ejecutado = True
                es_valida = es_factura_valida(texto_cabecera_lower, remitente) or es_factura_valida(texto_lower, remitente)
            elif not texto_lower.strip():
                return False, "PDF vacío o imagen (OCR también falló)"
            
        if not es_valida:
            if num_paginas > 3:
                return False, "Documento descartado (Tiene más de 3 páginas y no contiene CAE/Factura - Catálogo/Manual)"
            return False, "Documento descartado (No tiene CAE o contiene palabras prohibidas)"
            
        tipo = "FC"
        if re.search(REGLAS_EXTRACCION["rx_nota_credito"], texto_lower, re.IGNORECASE): 
            tipo = "NC"
        elif re.search(REGLAS_EXTRACCION["rx_nota_debito"], texto_lower, re.IGNORECASE): 
            tipo = "ND"
        
        nro = "SinNumero"
        
        # 1. Intentar el formato estándar XXXX-XXXXXXXX primero
        m_exacto = re.search(r"\b\d{4,5}\s*[-]\s*(\d{8})\b", texto_original)
        if m_exacto:
            nro = str(int(m_exacto.group(1)))
        else:
            # 1.5 Intentar el formato con letra en el medio: 0070A00371583
            m_letra = re.search(r"\b\d{4,5}[A-Za-z](\d{8})\b", texto_original)
            if m_letra:
                nro = str(int(m_letra.group(1)))
            else:
                # 2. Intentar buscar el texto "Nro:" 
                m_arca = re.search(r"(?:Comp\.?\s*Nro\.?|Nro\.?|Número)\s*[:\.]?\s*(\d{4,8})", texto_original, re.IGNORECASE)
                if m_arca:
                    nro = str(int(m_arca.group(1)))
                else:
                    m_comp = re.search(REGLAS_EXTRACCION["rx_numero_comp"], texto_original, re.IGNORECASE)
                    if m_comp: 
                        nro_crudo = m_comp.group(1).replace("-", "").replace(" ", "")
                        nro = str(int(nro_crudo[-8:]))
                    else:
                        m_guion = re.search(REGLAS_EXTRACCION["rx_numero_guion"], texto_original)
                        if m_guion: 
                            nro = str(int(m_guion.group(1)))
            
        razon_social = ""
        m_razon = re.search(REGLAS_EXTRACCION["rx_razon_social"], texto_original, re.IGNORECASE)
        if m_razon:
            rs_raw = m_razon.group(1).strip().upper()
            for basura in REGLAS_EXTRACCION["basura_afip"]:
                if basura in rs_raw: rs_raw = rs_raw.split(basura)[0]
            razon_social = rs_raw.strip()
            if REGLAS_EXTRACCION["empresa_propia"] in razon_social or "EUROSAT" in razon_social:
                razon_social = "" 
                
        if not razon_social:
            m_etiqueta = re.search(REGLAS_EXTRACCION.get("rx_razon_social_etiqueta", ""), texto_original, re.IGNORECASE)
            if m_etiqueta:
                rs_raw = m_etiqueta.group(1).strip().upper()
                for basura in REGLAS_EXTRACCION["basura_afip"]:
                    if basura in rs_raw: rs_raw = rs_raw.split(basura)[0]
                rs_candidate = rs_raw.strip()
                if rs_candidate and "TELCOM VENTURES" not in rs_candidate and "EUROSAT" not in rs_candidate:
                    razon_social = rs_candidate
                
        # ─── EXTRACCIÓN ROBUSTA DE CUIT DEL PROVEEDOR ─────────────────────────
        cuit_proveedor = ""
        cuit_propio_limpio = "30696428898"      # TELCOM VENTURES
        cuit_eurosat_limpio = "33676549639"     # EUROSAT
        cuits_clientes = {cuit_propio_limpio, cuit_eurosat_limpio}
        
        # 1. Buscar todos los CUITs de 11 dígitos en el texto
        todos_cuits_raw = re.findall(r"\b((?:20|23|24|27|30|33|34)[-\s]?\d{2}[-\s]?\d{6}[-\s]?\d{1})\b", texto_original)
        posibles_cuits = []
        for c in todos_cuits_raw:
            c_limpio = "".join(filter(str.isdigit, c))
            if c_limpio and c_limpio not in cuits_clientes and len(c_limpio) == 11:
                if c_limpio not in posibles_cuits:
                    posibles_cuits.append(c_limpio)

        if posibles_cuits:
            cuit_proveedor = posibles_cuits[0]

        # 2. Si no halló nada y no ejecutó OCR, forzar OCR y buscar de nuevo
        if not cuit_proveedor and not ocr_ejecutado:
            print(f"   ℹ️ CUIT del proveedor no hallado en texto digital. Forzando OCR completo...")
            texto_ocr = extraer_texto_con_ocr(ruta_pdf)
            if texto_ocr.strip():
                texto_original += "\n" + texto_ocr
                todos_cuits_raw = re.findall(r"\b((?:20|23|24|27|30|33|34)[-\s]?\d{2}[-\s]?\d{6}[-\s]?\d{1})\b", texto_original)
                for c in todos_cuits_raw:
                    c_limpio = "".join(filter(str.isdigit, c))
                    if c_limpio and c_limpio not in cuits_clientes and len(c_limpio) == 11:
                        cuit_proveedor = c_limpio
                        break 

        # --- EXTRACCIÓN DE FECHA DE EMISIÓN Y DETECCIÓN DE EUROSAT ---
        fecha_emision_str = ""
        m_fecha = re.search(REGLAS_EXTRACCION["rx_fecha_emision"], texto_original, re.IGNORECASE)
        if m_fecha:
            fecha_emision_str = m_fecha.group(1).replace("-", "/")

        es_eurosat = False
        # Detección por palabra clave EUROSAT o por su CUIT (con o sin guiones/espacios)
        if bool(re.search(r"\bEUROSAT\b", texto_original, re.IGNORECASE)) or bool(re.search(r"33[-\s]?67654963[-\s]?9", texto_original)):
            es_eurosat = True
                
        # --- EXTRACCIÓN DE OC Y SUBTOTAL ---
        oc_encontrada = ""
        m_oc = re.search(REGLAS_EXTRACCION["rx_oc"], texto_original, re.IGNORECASE)
        if m_oc:
            oc_encontrada = m_oc.group(1).strip()
            
        subtotal_encontrado = ""
        m_sub = re.search(REGLAS_EXTRACCION["rx_subtotal"], texto_original, re.IGNORECASE)
        if m_sub:
            subtotal_encontrado = m_sub.group(1).strip()
                
        es_tarjeta = bool(re.search(r"\b(visa\s*corpo|visacorpo|visa|tarjeta\s*de\s*credito|pago\s*con\s*tarjeta)\b", texto_original, re.IGNORECASE))

        return True, {
            "tipo": tipo, 
            "numero": nro, 
            "razon_social": razon_social, 
            "cuit": cuit_proveedor,
            "fecha_emision": fecha_emision_str,
            "es_eurosat": es_eurosat,
            "es_tarjeta": es_tarjeta,
            "oc_extraida": oc_encontrada,
            "subtotal": subtotal_encontrado
        }
    except Exception as e:
        return False, f"Error de lectura: {str(e)}"

def fusionar_pdfs_adjuntos(lista_rutas_pdf, ruta_salida):
    doc_salida = fitz.open()
    for ruta in lista_rutas_pdf:
        try:
            doc_parcial = fitz.open(ruta)
            if doc_parcial.is_pdf:
                doc_salida.insert_pdf(doc_parcial)
            else:
                pdfbytes = doc_parcial.convert_to_pdf()
                img_pdf = fitz.open("pdf", pdfbytes)
                doc_salida.insert_pdf(img_pdf)
                img_pdf.close()
            doc_parcial.close()
        except Exception as e:
            print(f"⚠️ No se pudo fusionar {ruta}: {e}")
    doc_salida.save(ruta_salida)
    doc_salida.close()

def obtener_ruta_unica(ruta):
    if not os.path.exists(ruta): return ruta
    nombre, ext = os.path.splitext(ruta)
    c = 1
    while os.path.exists(f"{nombre} ({c}){ext}"): c += 1
    return f"{nombre} ({c}){ext}"

def obtener_ultima_fecha_procesada(ruta_html):
    if not os.path.exists(ruta_html):
        return None
    try:
        # Leemos todas las tablas que contenga el HTML
        tablas = pd.read_html(ruta_html)
        for df_existente in tablas:
            # Buscamos la columna de fecha (puede llamarse 'Fecha')
            col_fecha = next((c for c in df_existente.columns if 'fecha' in str(c).lower()), None)
            if col_fecha and not df_existente.empty:
                fechas = pd.to_datetime(df_existente[col_fecha], format='%d/%m/%Y', errors='coerce')
                # Si fallara el formato con hora, probamos formato mixto
                if fechas.isna().all():
                    fechas = pd.to_datetime(df_existente[col_fecha], format='%d/%m/%Y %H:%M:%S', errors='coerce')
                
                ultima = fechas.max()
                if pd.notna(ultima):
                    return ultima.strftime('%d/%m/%Y')
    except Exception as e:
        print(f"⚠️ Error leyendo fecha del reporte: {e}")
    return None

def iniciar_proceso_interactivo(ejecutar_descarga=False, usar_filtro_fecha=True, fecha_inicio="01/07/2026", fecha_fin="", filtro_correo=""):
    ruta_html_reporte = REPORTE_MAESTRO_HTML
    
    ultima_fecha_historica = obtener_ultima_fecha_procesada(ruta_html_reporte)
    fecha_a_utilizar = ultima_fecha_historica if ultima_fecha_historica else fecha_inicio
    
    print(f"📅 Rango de búsqueda optimizado desde: {fecha_a_utilizar}")

    mapa_remitentes = {}
    if ejecutar_descarga:
        mapa_remitentes = descargar_facturas_outlook(fecha_inicio=fecha_a_utilizar, filtro_proveedor=filtro_correo)
        print("-" * 50)

    print("📥 Cargando bases de datos de proveedores y sectores...")
    base_prov = cargar_excel_proveedores(ARCHIVO_PROVEEDORES)
    
    facturas_procesadas_historico = obtener_facturas_procesadas(ruta_html_reporte)
    print(f"🛡️ Filtro activo: Se encontraron {len(facturas_procesadas_historico)} facturas previas registradas en el historial.")

    if not os.path.exists(CARPETA_A):
        print("❌ La Carpeta A no existe.")
        return

    archivos = [f for f in os.listdir(CARPETA_A) if f.lower().endswith(".pdf")]
    if not archivos:
        print("📂 No hay PDFs en la Carpeta A para procesar.")
        return

    os.makedirs(CARPETA_B, exist_ok=True)
    os.makedirs(CARPETA_C, exist_ok=True)

    f_ini = datetime.strptime(fecha_a_utilizar, "%d/%m/%Y").date() if usar_filtro_fecha and fecha_a_utilizar else None
    f_fin = datetime.strptime(fecha_fin, "%d/%m/%Y").date() if usar_filtro_fecha and fecha_fin.strip() else datetime.today().date()

    mes_actual = int(datetime.now().month)
    nombre_mes_defecto = MESES_ESPANOL.get(mes_actual, "MES")
    
    datos_para_html = []
    procesados_contador = 0
    duplicados_omitidos = 0
    processed_files = set()

    for archivo in archivos:
        if archivo in processed_files: 
            continue
            
        ruta_completa = os.path.join(CARPETA_A, archivo)
        
        remitente_correo = mapa_remitentes.get(archivo, "")

        match_marino = re.search(r'(?:VOLQUETES.*?MARINO|MARINO).*?(\d+)', archivo, re.IGNORECASE)
        archivos_a_fusionar = [ruta_completa]
        if match_marino:
            nro_detectado = match_marino.group(1)
            for otro_arch in archivos:
                if otro_arch != archivo and nro_detectado in otro_arch:
                    archivos_a_fusionar.append(os.path.join(CARPETA_A, otro_arch))
                    processed_files.add(otro_arch)

        es_valido, datos = extraer_datos_pdf(ruta_completa, remitente_correo)
        
        if not es_valido:
            if "vacío" in datos or "imagen" in datos:
                print(f"\n⚠️ ATENCIÓN: {archivo} parece ser un PDF escaneado o imagen (no se pudo leer el texto).")
                opc = input("¿Deseas procesarlo manualmente como factura? (s/n): ").strip().lower()
                if opc == 's':
                    es_valido = True
                    datos = {"tipo": "FC", "numero": "SinNumero", "razon_social": "PROVEEDOR_DESCONOCIDO", "cuit": ""}
                    # Intentar rescatar el número desde el nombre del archivo
                    m_exacto = re.search(r"\b\d{4,5}\s*[-_]\s*(\d{8})\b", archivo)
                    if m_exacto:
                        datos["numero"] = str(int(m_exacto.group(1)))
                    else:
                        m_arca = re.search(r"(\d{4,8})\.pdf$", archivo, re.IGNORECASE)
                        if m_arca:
                            datos["numero"] = str(int(m_arca.group(1)))
            
            if not es_valido:
                for r_fusa in archivos_a_fusionar:
                    destino_b = obtener_ruta_unica(os.path.join(CARPETA_B, os.path.basename(r_fusa)))
                    shutil.move(r_fusa, destino_b)
                print(f"[-] {archivo} -> NO ES FACTURA. Movido automáticamente a Carpeta B.")
                continue
            
        nombre_final, sector, es_negra = buscar_proveedor(datos.get("cuit", ""), datos["razon_social"], base_prov)
        
        if es_negra:
            for r_fusa in archivos_a_fusionar:
                destino_b = obtener_ruta_unica(os.path.join(CARPETA_B, os.path.basename(r_fusa)))
                shutil.move(r_fusa, destino_b)
            print(f"[!] {archivo} -> LISTA NEGRA. Movido a Carpeta B.")
            continue

        if "ECOSAN" in nombre_final.upper() or "facturacion@ecosan.com.ar" in remitente_correo:
            _, locacion_ecosan = analizar_pdf_ecosan(ruta_completa)
            if not sector or sector == "GENERAL":
                sector = locacion_ecosan

        # --- 🛡️ VALIDACIÓN ANTICIPADA DE DUPLICADOS (CUIT + NÚMERO) ---
        cuit_str = str(datos.get('cuit', '')).strip()
        nro_str = str(datos.get('numero', '')).strip()
        identidad_factura = (cuit_str, nro_str) if cuit_str else (nombre_final.strip().upper(), nro_str)
        
        if identidad_factura in facturas_procesadas_historico or nro_str == "SinNumero":
            for r_fusa in archivos_a_fusionar:
                destino_duplicado = obtener_ruta_unica(os.path.join(CARPETA_B, os.path.basename(r_fusa)))
                shutil.move(r_fusa, destino_duplicado)
            print(f"[⏭️ DUPLICADO OMITIDO] {archivo} (Comprobante {nro_str}) -> Ya procesado anteriormente. Movido a Carpeta B.")
            duplicados_omitidos += 1
            continue
            
        # --- CÁLCULO DINÁMICO DEL MES SEGÚN FECHA DE EMISIÓN ---
        nombre_mes = nombre_mes_defecto
        if datos.get("fecha_emision"):
            try:
                dt_emision = datetime.strptime(datos["fecha_emision"], "%d/%m/%Y")
                nro_mes = dt_emision.month
                nombre_mes = MESES_ESPANOL.get(nro_mes, nombre_mes_defecto)
            except Exception:
                pass

        # --- PREFIJO DE EMPRESA (EUROSAT SI CORRESPONDE) ---
        prefijo_empresa = "(EUROSAT) " if datos.get("es_eurosat") else ""

        sec_str = f" {sector}" if sector else ""
        nombre_sugerido = f"{prefijo_empresa}({nombre_mes}) {nombre_final} {datos['tipo']} {datos['numero']}{sec_str}.pdf"
        
        print("\n" + "="*60)
        print(f"📄 Archivo principal: {archivo}")
        if len(archivos_a_fusionar) > 1:
            print(f"   📎 Se fusionarán {len(archivos_a_fusionar)-1} adjuntos en un único PDF.")
        print(f"   ↳ Proveedor detectado : {nombre_final}")
        print(f"   ↳ CUIT                  : {datos.get('cuit', 'No detectado')}")
        print(f"   ↳ Fecha Emisión         : {datos.get('fecha_emision', 'No detectada')}")
        print(f"   ↳ Tipo y Número         : {datos['tipo']} - {datos['numero']}")
        print(f"   ↳ Sector                : {sector if sector else 'GENERAL'}")
        print(f"   ✨ Nombre Propuesto     : {nombre_sugerido}")
        print("="*60)
        
        opcion = input("¿Qué deseas hacer?\n  [Enter] ➔ Aprobar y mover a 'fc a subir'\n  [e]     ➔ Editar nombre manualmente\n  [n]     ➔ Rechazar (Mover a 'no es factura')\nTu opción: ").strip().lower()
        
        if opcion == 'n':
            for r_fusa in archivos_a_fusionar:
                destino_b = obtener_ruta_unica(os.path.join(CARPETA_B, os.path.basename(r_fusa)))
                shutil.move(r_fusa, destino_b)
            print("❌ Rechazado. Movido a Carpeta B.")
            continue
            
        elif opcion == 'e':
            nuevo_input = input(f"Escribe el nuevo nombre (sin .pdf) o presiona Enter para usar '{nombre_sugerido}': ").strip()
            if nuevo_input:
                if not nuevo_input.lower().endswith('.pdf'):
                    nombre_sugerido = f"{nuevo_input}.pdf"
                else:
                    nombre_sugerido = nuevo_input
                    
        # --- INTEGRACIÓN DE ÓRDENES DE COMPRA (OC) ---
        print("\n[DEBUG] Iniciando búsqueda de OCs...")
        cuit_busq = datos.get("cuit", "")
        ocs_disponibles = oc_manager.leer_oc_disponibles(cuit_busq, nombre_final)
        print(f"[DEBUG] OCs devueltas por el manager: {len(ocs_disponibles)}")
        
        oc_elegida = ""
        oc_extraida = datos.get("oc_extraida", "")
        
        if oc_extraida:
            print(f"📌 Se detectó automáticamente la leyenda OC {oc_extraida} en el PDF.")
            # Verificar si la OC extraída está en las disponibles
            if any(o['nro_oc'].lower() == oc_extraida.lower() for o in ocs_disponibles):
                print(f"✅ La OC {oc_extraida} está disponible en la base de datos.")
                oc_elegida = oc_extraida
            else:
                # Opcional: Permitir fusionar igual aunque no esté en disponibles?
                # Vamos a forzar la búsqueda de todas formas para no trabar el proceso
                oc_elegida = oc_extraida
        
        if ocs_disponibles and not oc_elegida:
            print(f"\n🛒 OCs disponibles para {nombre_final}:")
            for i, oc in enumerate(ocs_disponibles, 1):
                print(f"  {i}) OC {oc['nro_oc']} | {oc['fecha']} | {oc['articulo']} | Tipo: {oc['tipo']}")
                
            oc_elegida = input("Escribe el Nro de OC a vincular (o presiona Enter para omitir): ").strip()
            
        if oc_elegida:
            # Buscar en la lista (si está) o usar la extraída directamente
            oc_encontrada = next((o for o in ocs_disponibles if o['nro_oc'].lower() == oc_elegida.lower()), None)
            nro_oc_final = oc_encontrada['nro_oc'] if oc_encontrada else oc_elegida
            
            ruta_pdf_oc = oc_manager.buscar_pdf_oc(nro_oc_final)
            if ruta_pdf_oc:
                print(f"✅ Se adjuntará el PDF de la OC: {os.path.basename(ruta_pdf_oc)}")
                archivos_a_fusionar.append(ruta_pdf_oc)
                
                # Agregar sufijo al nombre
                base_n, ext_n = os.path.splitext(nombre_sugerido)
                nombre_sugerido = f"{base_n} -OC {nro_oc_final}{ext_n}"
                
                # Registrar consumo
                if oc_encontrada:
                    oc_manager.registrar_consumo_oc(nro_oc_final, nombre_sugerido)
            else:
                print(f"⚠️ No se encontró el archivo PDF para la OC {nro_oc_final} en la red. Se omitirá la fusión de la OC.")
                    
        nombre_limpio = re.sub(r'[\\/*?:"<>|]', "", nombre_sugerido)
        ruta_destino_c = obtener_ruta_unica(os.path.join(CARPETA_C, nombre_limpio))
        
        sector_final = sector if sector else "GENERAL"
        carpeta_autorizacion = r"\\10.10.10.210\AyF_Trabajoadistancia\AUTORIZACION DE FACTURAS"
        carpeta_sector = os.path.join(carpeta_autorizacion, sector_final)
        os.makedirs(carpeta_sector, exist_ok=True)
        ruta_destino_sector = obtener_ruta_unica(os.path.join(carpeta_sector, nombre_limpio))
        
        if len(archivos_a_fusionar) > 1:
            fusionar_pdfs_adjuntos(archivos_a_fusionar, ruta_destino_c)
            # Copiar el fusionado a Autorizacion
            shutil.copy2(ruta_destino_c, ruta_destino_sector)
            for r_fusa in archivos_a_fusionar:
                if os.path.exists(r_fusa) and r_fusa != ruta_completa: 
                    # No borrar la OC original, solo borrar adjuntos del correo
                    if CARPETA_A in r_fusa:
                        os.remove(r_fusa)
            if os.path.exists(ruta_completa):
                os.remove(ruta_completa)
            print(f"✅ ¡Fusionados y guardados en 'fc a subir' y '{sector_final}'!")
        else:
            shutil.copy2(ruta_completa, ruta_destino_c)
            shutil.move(ruta_completa, ruta_destino_sector)
            print(f"✅ ¡Aprobado y movido a 'fc a subir' y '{sector_final}'!")
        
        datos_para_html.append({
            "Fecha": datetime.now().strftime("%d/%m/%Y"),
            "Tipo": datos['tipo'],
            "Nro Comprobante": datos['numero'],
            "Razón Social": nombre_final,
            "Sector": sector_final,
            "OC": oc_elegida if oc_elegida else "N/A",
            "Subtotal": datos.get("subtotal", "")
        })
        
        procesados_contador += 1

    print(f"\n🎯 Proceso finalizado. Aprobadas: {procesados_contador} | Duplicados omitidos: {duplicados_omitidos}")

    # Genera el reporte acumulativo de la carpeta 'fc a subir'
    generar_reporte_html()