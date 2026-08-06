# procesador_core.py
import os
import re
import shutil
from datetime import datetime
import fitz  # PyMuPDF
import pandas as pd
from config import CARPETA_A, CARPETA_B, CARPETA_C, ARCHIVO_PROVEEDORES, MESES_ESPANOL
from pdf_rules import REGLAS_EXTRACCION
from proveedores import cargar_excel_proveedores, buscar_proveedor
from downloader import descargar_facturas_outlook
from reporter import generar_reporte_html, obtener_facturas_procesadas
from proveedor_rules import es_factura_por_remitente, analizar_pdf_ecosan
import oc_manager

def es_factura_valida(texto_lower, remitente=""):
    if es_factura_por_remitente(remitente):
        return True
        
    texto_seguro = texto_lower.replace("presupuesto económico", "").replace("presupuesto economico", "")
    
    for prohibida in REGLAS_EXTRACCION["palabras_prohibidas"]:
        if prohibida.lower() in texto_seguro: 
            return False
            
    texto_sin_espacios_raros = re.sub(r'[\s\n\r\t:\-\._\|]+', '', texto_lower)
    if "cae" in texto_sin_espacios_raros:
        return True
    if re.search(r'\b8\d{13,14}\b', texto_lower):
        return True
    return False

def extraer_datos_pdf(ruta_pdf, remitente=""):
    try:
        doc = fitz.open(ruta_pdf)
        texto_original = "\n".join([pagina.get_text("text", sort=True) for pagina in doc])
        texto_lower = texto_original.lower()
        doc.close()
        if not texto_lower.strip():
            return False, "PDF vacío o imagen"
            
        if not es_factura_valida(texto_lower, remitente):
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
            if REGLAS_EXTRACCION["empresa_propia"] in razon_social:
                razon_social = "" 
                
        cuit_proveedor = ""
        todos_cuits = re.findall(REGLAS_EXTRACCION["rx_cuit"], texto_original, re.IGNORECASE)
        cuit_propio_limpio = "30696428898"
        for c in todos_cuits:
            c_limpio = "".join(filter(str.isdigit, c))
            if c_limpio and c_limpio != cuit_propio_limpio:
                cuit_proveedor = c_limpio
                break 

        # --- EXTRACCIÓN DE FECHA DE EMISIÓN Y DETECCIÓN DE EUROSAT ---
        fecha_emision_str = ""
        m_fecha = re.search(REGLAS_EXTRACCION["rx_fecha_emision"], texto_original, re.IGNORECASE)
        if m_fecha:
            fecha_emision_str = m_fecha.group(1).replace("-", "/")

        es_eurosat = False
        cuit_eurosat_limpio = REGLAS_EXTRACCION.get("cuit_eurosat", "33676549639")
        nombre_eurosat = REGLAS_EXTRACCION.get("nombre_eurosat", "EUROSAT")
        if nombre_eurosat in texto_original.upper() or cuit_eurosat_limpio in texto_original:
            es_eurosat = True
                
        return True, {
            "tipo": tipo, 
            "numero": nro, 
            "razon_social": razon_social, 
            "cuit": cuit_proveedor,
            "fecha_emision": fecha_emision_str,
            "es_eurosat": es_eurosat
        }
    except Exception as e:
        return False, f"Error de lectura: {str(e)}"

def fusionar_pdfs_adjuntos(lista_rutas_pdf, ruta_salida):
    doc_salida = fitz.open()
    for ruta in lista_rutas_pdf:
        try:
            doc_parcial = fitz.open(ruta)
            doc_salida.insert_pdf(doc_parcial)
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
    ruta_html_reporte = r"\\10.10.10.210\AyF_Trabajoadistancia\Compras\Reporte_Maestro.html"
    
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
        
        if ocs_disponibles:
            print(f"\n🛒 OCs disponibles para {nombre_final}:")
            for i, oc in enumerate(ocs_disponibles, 1):
                print(f"  {i}) OC {oc['nro_oc']} | {oc['fecha']} | {oc['articulo']} | Tipo: {oc['tipo']}")
                
            oc_elegida = input("Escribe el Nro de OC a vincular (o presiona Enter para omitir): ").strip()
            
            if oc_elegida:
                oc_encontrada = next((o for o in ocs_disponibles if o['nro_oc'].lower() == oc_elegida.lower()), None)
                if oc_encontrada:
                    ruta_pdf_oc = oc_manager.buscar_pdf_oc(oc_encontrada['nro_oc'])
                    if ruta_pdf_oc:
                        print(f"✅ Se adjuntará el PDF de la OC: {os.path.basename(ruta_pdf_oc)}")
                        archivos_a_fusionar.append(ruta_pdf_oc)
                        
                        # Agregar sufijo al nombre
                        base_n, ext_n = os.path.splitext(nombre_sugerido)
                        nombre_sugerido = f"{base_n} -OC {oc_encontrada['nro_oc']}{ext_n}"
                        
                        # Registrar consumo
                        oc_manager.registrar_consumo_oc(oc_encontrada['nro_oc'], nombre_sugerido)
                    else:
                        print(f"⚠️ No se encontró el archivo PDF para la OC {oc_encontrada['nro_oc']} en la red. Se omitirá la fusión de la OC.")
                else:
                    print("⚠️ Nro de OC no coincide con la lista. Se omitirá.")
                    
        nombre_limpio = re.sub(r'[\\/*?:"<>|]', "", nombre_sugerido)
        ruta_destino_c = obtener_ruta_unica(os.path.join(CARPETA_C, nombre_limpio))
        
        if len(archivos_a_fusionar) > 1:
            fusionar_pdfs_adjuntos(archivos_a_fusionar, ruta_destino_c)
            for r_fusa in archivos_a_fusionar:
                if os.path.exists(r_fusa): os.remove(r_fusa)
            print(f"✅ ¡Fusionados y guardados como un único PDF en Carpeta C!")
        else:
            shutil.move(ruta_completa, ruta_destino_c)
            print(f"✅ ¡Aprobado y movido a Carpeta C!")
        
        datos_para_html.append({
            "Fecha": datetime.now().strftime("%d/%m/%Y"),
            "Tipo": datos['tipo'],
            "Nro Comprobante": datos['numero'],
            "Razon Social": nombre_final,
            "Sector": sector if sector else "GENERAL",
            "OC": "N/A"
        })
        
        procesados_contador += 1

    print(f"\n🎯 Proceso finalizado. Aprobadas: {procesados_contador} | Duplicados omitidos: {duplicados_omitidos}")

    # Genera el reporte acumulativo de la carpeta 'fc a subir'
    generar_reporte_html()