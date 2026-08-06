# prueba_fase1.py
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

# =======================================================================
# ⚙️ PANEL DE CONTROL
# =======================================================================
EJECUTAR_DESCARGA = False 
USAR_FILTRO_FECHA = True
FECHA_INICIO = "01/07/2026"  # Fecha por defecto si el reporte está vacío
FECHA_FIN = ""               
FILTRO_CORREO = ""           
# =======================================================================

def es_factura_valida(texto_lower, remitente=""):
    if es_factura_por_remitente(remitente):
        return True
        
    for prohibida in REGLAS_EXTRACCION["palabras_prohibidas"]:
        if prohibida.lower() in texto_lower: 
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
        
        if not es_factura_valida(texto_lower, remitente):
            return False, "Documento descartado (No tiene CAE o contiene palabras prohibidas)"
            
        # --- EXTRACCIÓN DE TIPO Y NÚMERO DE COMPROBANTE ---
        tipo = "FC"
        if re.search(REGLAS_EXTRACCION["rx_nota_credito"], texto_lower, re.IGNORECASE): 
            tipo = "NC"
        elif re.search(REGLAS_EXTRACCION["rx_nota_debito"], texto_lower, re.IGNORECASE): 
            tipo = "ND"
        
        nro = "SinNumero"
        
        # Búsqueda optimizada estándar ARCA (Comp. Nro:, Nro:, etc.)
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
                
        return True, {"tipo": tipo, "numero": nro, "razon_social": razon_social, "cuit": cuit_proveedor}
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
    """Busca en el reporte HTML la fecha más reciente para continuar desde ahí de forma inteligente."""
    if not os.path.exists(ruta_html):
        return None
    try:
        df_existente = pd.read_html(ruta_html)[0]
        if 'Fecha' in df_existente.columns and not df_existente.empty:
            fechas = pd.to_datetime(df_existente['Fecha'], format='%d/%m/%Y', errors='coerce')
            ultima = fechas.max()
            if pd.notna(ultima):
                return ultima.strftime('%d/%m/%Y')
    except Exception:
        pass
    return None

def iniciar_proceso_interactivo():
    ruta_html_reporte = r"\\10.10.10.210\AyF_Trabajoadistancia\Compras\Reporte_Maestro.html"
    
    # Auto-detección del punto de partida inteligente
    ultima_fecha_historica = obtener_ultima_fecha_procesada(ruta_html_reporte)
    fecha_a_utilizar = ultima_fecha_historica if ultima_fecha_historica else FECHA_INICIO
    
    print(f"📅 Rango de búsqueda optimizado desde: {fecha_a_utilizar}")

    mapa_remitentes = {}
    if EJECUTAR_DESCARGA:
        mapa_remitentes = descargar_facturas_outlook(fecha_inicio=fecha_a_utilizar, filtro_proveedor=FILTRO_CORREO)
        print("-" * 50)

    print("📥 Cargando bases de datos de proveedores y sectores...")
    base_prov = cargar_excel_proveedores(ARCHIVO_PROVEEDORES)
    
    facturas_procesadas_historico = obtener_facturas_procesadas(ruta_html_reporte)
    print(f"🛡️ Filtro activo: Se encontraron {len(facturas_procesadas_historico)} facturas previas registradas en el historial.")

    print(f"\n🔍 Analizando archivos en la Carpeta A: {CARPETA_A}\n")
    if not os.path.exists(CARPETA_A):
        print("❌ La Carpeta A no existe.")
        return

    archivos = [f for f in os.listdir(CARPETA_A) if f.lower().endswith(".pdf")]
    if not archivos:
        print("📂 No hay PDFs en la Carpeta A para procesar.")
        return

    os.makedirs(CARPETA_B, exist_ok=True)
    os.makedirs(CARPETA_C, exist_ok=True)

    f_ini = datetime.strptime(fecha_a_utilizar, "%d/%m/%Y").date() if USAR_FILTRO_FECHA and fecha_a_utilizar else None
    f_fin = datetime.strptime(FECHA_FIN, "%d/%m/%Y").date() if USAR_FILTRO_FECHA and FECHA_FIN.strip() else datetime.today().date()

    mes_actual = int(datetime.now().month)
    nombre_mes = MESES_ESPANOL.get(mes_actual, "MES")
    
    datos_para_html = []
    procesados_contador = 0
    duplicados_omitidos = 0
    processed_files = set()

    for archivo in archivos:
        if archivo in processed_files: 
            continue
            
        ruta_completa = os.path.join(CARPETA_A, archivo)
        
        remitente_correo = mapa_remitentes.get(archivo, "")

        # Excepción de agrupación de múltiples PDFs (ej: Transportes Marino)
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

        # Regla específica Ecosan
        if "ECOSAN" in nombre_final.upper() or "facturacion@ecosan.com.ar" in remitente_correo:
            _, locacion_ecosan = analizar_pdf_ecosan(ruta_completa)
            if not sector or sector == "GENERAL":
                sector = locacion_ecosan

        # --- 🛡️ FILTRO DE DUPLICADOS ---
        identidad_factura = (nombre_final.strip().upper(), str(datos['numero']).strip())
        if identidad_factura in facturas_procesadas_historico:
            for r_fusa in archivos_a_fusionar:
                destino_duplicado = obtener_ruta_unica(os.path.join(CARPETA_B, os.path.basename(r_fusa)))
                shutil.move(r_fusa, destino_duplicado)
            print(f"[⏭️ DUPLICADO] {archivo} -> Ya procesada anteriormente. Movida a Carpeta B.")
            duplicados_omitidos += 1
            continue
            
        sec_str = f" {sector}" if sector else ""
        nombre_sugerido = f"({nombre_mes}) {nombre_final} {datos['tipo']} {datos['numero']}{sec_str}.pdf"
        
        # --- 🛑 PARADA INTERACTIVA ---
        print("\n" + "="*60)
        print(f"📄 Archivo principal: {archivo}")
        if len(archivos_a_fusionar) > 1:
            print(f"   📎 Se fusionarán {len(archivos_a_fusionar)-1} adjuntos en un único PDF.")
        print(f"   ↳ Proveedor detectado : {nombre_final}")
        print(f"   ↳ CUIT                  : {datos.get('cuit', 'No detectado')}")
        print(f"   ↳ Tipo y Número         : {datos['tipo']} - {datos['numero']}")
        print(f"   ↳ Sector                : {sector if sector else 'GENERAL'}")
        print(f"   ✨ Nombre Propuesto     : {nombre_sugerido}")
        print("="*60)
        
        opcion = input("¿Qué deseas hacer?\n  [Enter] ➔ Aprobar y mover a 'fc a subir'\n  [e]     ➔ Editar nombre manualmente\n  [n]     ➔ Rechazar (Mover a 'no es factura')\nTu opción: ").strip().lower()
        
        if opcion == 'n':
            for r_fusa in archivos_a_fusionar:
                destino_b = obtener_ruta_unica(os.path.join(CARPETA_B, os.path.basename(r_fusa)))
                shutil.move(r_fusa, destino_b)
            print(f"❌ Archivo rechazado y movido a Carpeta B.")
            continue
            
        elif opcion == 'e':
            nuevo_input = input("Escribe el nombre final del archivo (sin .pdf): ").strip()
            if nuevo_input:
                nombre_sugerido = f"{nuevo_input}.pdf"
                
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

    if datos_para_html:
        generar_reporte_html(datos_para_html, ruta_html_reporte)

if __name__ == "__main__":
    iniciar_proceso_interactivo()