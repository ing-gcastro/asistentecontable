import os
import re
import sys
import pandas as pd
from datetime import datetime
from config import CARPETA_A, ARCHIVO_PROVEEDORES, REPORTE_MAESTRO_HTML
from proveedores import cargar_excel_proveedores, buscar_proveedor
from procesador_core import extraer_datos_pdf
from reporter import obtener_facturas_procesadas

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

# 📍 El Excel ahora se ubica directamente en la carpeta de entrada de la red
EXCEL_VALIDACION = os.path.join(CARPETA_A, "Validacion_Lote.xlsx")

def cerrar_excel_si_abierto():
    """Fuerza el cierre de Excel en Windows si el usuario lo dejó abierto para evitar bloqueos."""
    try:
        os.system("taskkill /f /im excel.exe >nul 2>&1")
    except Exception:
        pass

def generar_excel_lote():
    print("Iniciando Preparación de Lote (Paso 1)...")
    cerrar_excel_si_abierto()
    
    if not os.path.exists(CARPETA_A):
        print(f"❌ La carpeta de entrada no existe: {CARPETA_A}")
        return

    # Omitimos el propio Excel de validación si existe en la carpeta
    archivos = [f for f in os.listdir(CARPETA_A) if f.lower().endswith(".pdf")]
    if not archivos:
        print("📂 No hay facturas pendientes en la bandeja de entrada para procesar.")
        return
        
    print("📥 Cargando bases de proveedores, sectores e historial...")
    base_prov = cargar_excel_proveedores(ARCHIVO_PROVEEDORES)
    facturas_procesadas_historico = obtener_facturas_procesadas(REPORTE_MAESTRO_HTML)
    print(f"🛡️ Historial cargado: {len(facturas_procesadas_historico)} facturas registradas previamente.")

    datos_excel = []
    
    for archivo in archivos:
        ruta_completa = os.path.join(CARPETA_A, archivo)
        print(f"  Analizando: {archivo}")
        
        exito, datos = extraer_datos_pdf(ruta_completa, remitente="")
        
        # ─── CASO 1: NO ES FACTURA O ES CATÁLOGO (>3 PÁGS) ────────────────────────
        if not exito:
            print(f"  ⚠️ Documento no válido/descartado. Marcando para RECHAZAR.")
            datos_excel.append({
                "Archivo_Original": archivo,
                "Empresa": "TELCOM VENTURES",
                "Razón_Social": "NO FACTURA",
                "Sector": "GENERAL",
                "Tipo": "N/A",
                "Nro_Comprobante": "N/A",
                "Subtotal": "",
                "OC_Detectada": "",
                "Adjuntos_Extra": "",
                "Nombre_Propuesto": "NO FACTURA",
                "Estado_Historial": "NO ES FACTURA / RECHAZAR",
                "Acción": "RECHAZAR"
            })
            continue

        # ─── CASO 2: DOCUMENTO VÁLIDO -> BUSCAR PROVEEDOR ─────────────────────────
        nombre_final, sector, es_lista_negra = buscar_proveedor(datos["cuit"], datos["razon_social"], base_prov)
        
        tipo_str = datos["tipo"] if datos["tipo"] else "FC"
        nro_str = datos["numero"] if datos["numero"] else "S_N"
        
        mes_actual = int(datetime.now().month)
        meses = {1:"ENERO", 2:"FEBRERO", 3:"MARZO", 4:"ABRIL", 5:"MAYO", 6:"JUNIO", 
                 7:"JULIO", 8:"AGOSTO", 9:"SEPTIEMBRE", 10:"OCTUBRE", 11:"NOVIEMBRE", 12:"DICIEMBRE"}
        nombre_mes = meses.get(mes_actual, "MES")
        
        sector_str = sector if sector else "GENERAL"
        
        oc_extraida_limpia = str(datos.get("oc_extraida", "")).strip()
        if oc_extraida_limpia.endswith('.0'): oc_extraida_limpia = oc_extraida_limpia[:-2]
        
        es_eurosat = datos.get("es_eurosat", False)
        empresa_str = "EUROSAT" if es_eurosat else "TELCOM VENTURES"
        prefijo_empresa = "(EUROSAT) " if es_eurosat else ""
        
        nombre_sugerido = f"{prefijo_empresa}({nombre_mes}) {nombre_final} {tipo_str} {nro_str} {sector_str}"
        if oc_extraida_limpia:
            nombre_sugerido += f" -OC {oc_extraida_limpia}"

        # ─── CASO 3: DUPLICADAS ───────────────────────────────────────────────────
        razon_limpia = re.sub(r'^\([^)]+\)\s*', '', nombre_final).strip().upper()
        
        es_duplicada = False
        for (r_hist, n_hist) in facturas_procesadas_historico:
            if n_hist == str(nro_str).strip() and (r_hist in razon_limpia or razon_limpia in r_hist):
                es_duplicada = True
                break
                
        if es_duplicada:
            print(f"  🛑 DUPLICADA DETECTADA: {nombre_final} {nro_str} ya fue procesada anteriormente.")
            datos_excel.append({
                "Archivo_Original": archivo,
                "Empresa": empresa_str,
                "Razón_Social": nombre_final,
                "Sector": sector_str,
                "Tipo": tipo_str,
                "Nro_Comprobante": str(nro_str),
                "Subtotal": datos.get("subtotal", ""),
                "OC_Detectada": oc_extraida_limpia,
                "Adjuntos_Extra": "",
                "Nombre_Propuesto": f"{nombre_sugerido} [DUPLICADA]",
                "Estado_Historial": "YA PROCESADA (DUPLICADA)",
                "Acción": "RECHAZAR"
            })
        else:
            accion_defecto = "RECHAZAR" if es_lista_negra else "APROBAR"
            datos_excel.append({
                "Archivo_Original": archivo,
                "Empresa": empresa_str,
                "Razón_Social": nombre_final,
                "Sector": sector_str,
                "Tipo": tipo_str,
                "Nro_Comprobante": str(nro_str),
                "Subtotal": datos.get("subtotal", ""),
                "OC_Detectada": oc_extraida_limpia,
                "Adjuntos_Extra": "",
                "Nombre_Propuesto": nombre_sugerido,
                "Estado_Historial": "NUEVA",
                "Acción": accion_defecto
            })
        
    df = pd.DataFrame(datos_excel)
    
    try:
        with pd.ExcelWriter(EXCEL_VALIDACION, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name="Lote_Pendiente")
            worksheet = writer.sheets['Lote_Pendiente']
            for col in worksheet.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = (max_length + 2)
                worksheet.column_dimensions[column].width = min(adjusted_width, 50)
                
        print(f"\n✅ ¡Lote generado con éxito en la carpeta de red!")
        print(f"👉 Excel generado en: '{EXCEL_VALIDACION}'")
        print(f"👉 Revisa las Acciones, edita Sector u OC si hace falta, guarda y cierra el Excel.")
        print(f"👉 Luego, ejecuta el Paso 2 (Aplicar Lote).")
    except Exception as e:
        print(f"❌ Error al guardar el Excel: {e}")

if __name__ == "__main__":
    generar_excel_lote()
    input("\nPresiona ENTER para salir...")
