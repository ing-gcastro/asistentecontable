import os
import shutil
import re
import sys
import pandas as pd
from datetime import datetime
from config import CARPETA_A, CARPETA_B, CARPETA_C, CARPETA_REVISION, REPORTE_MAESTRO_HTML
from procesador_core import fusionar_pdfs_adjuntos, obtener_ruta_unica
import oc_manager
import reporter

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

# 📍 El Excel ahora se ubica directamente en la carpeta de entrada de la red
EXCEL_VALIDACION = os.path.join(CARPETA_A, "Validacion_Lote.xlsx")
CARPETA_AUTORIZACION = r"\\10.10.10.210\AyF_Trabajoadistancia\AUTORIZACION DE FACTURAS"

def cerrar_excel_si_abierto():
    """Cierra Excel para desbloquear el archivo en la red."""
    try:
        os.system("taskkill /f /im excel.exe >nul 2>&1")
    except Exception:
        pass

def limpiar_cadena_oc(val):
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return "" if s.lower() == 'nan' else s

def procesar_excel_lote():
    print("Iniciando Procesamiento de Lote (Paso 2)...")
    cerrar_excel_si_abierto()
    
    if not os.path.exists(EXCEL_VALIDACION):
        print(f"❌ No se encontró el archivo {EXCEL_VALIDACION}. Ejecuta primero el Paso 1.")
        return
        
    try:
        # Cargar con dtype=str para que los números de OC o Comprobante no se conviertan en floats (.0)
        df = pd.read_excel(EXCEL_VALIDACION, engine='openpyxl', dtype=str)
    except Exception as e:
        print(f"❌ Error al leer el Excel (Asegúrate de haberlo cerrado): {e}")
        return
        
    if df.empty:
        print("📂 El Excel está vacío.")
        return
        
    os.makedirs(CARPETA_B, exist_ok=True)
    os.makedirs(CARPETA_C, exist_ok=True)
    os.makedirs(CARPETA_REVISION, exist_ok=True)
    
    datos_para_html = []
    aprobadas = 0
    rechazadas = 0
    revision = 0
    
    for index, row in df.iterrows():
        archivo_orig = str(row.get("Archivo_Original", "")).strip()
        accion = str(row.get("Acción", "")).strip().upper()
        
        if not archivo_orig or archivo_orig == 'nan':
            continue
            
        ruta_completa = os.path.join(CARPETA_A, archivo_orig)
        
        if not os.path.exists(ruta_completa):
            print(f"⚠️ Archivo no encontrado en Carpeta A: {archivo_orig} (Omitiendo)")
            continue
            
        # Normalizar acción (quitar acentos y espacios extra)
        accion_norm = accion.replace("Ó", "O").replace("Í", "I").replace("Á", "A").replace("É", "E").replace("Ú", "U")
        
        if "RECHAZ" in accion_norm or "NO" in accion_norm or "DESCART" in accion_norm:
            destino_b = obtener_ruta_unica(os.path.join(CARPETA_B, archivo_orig))
            shutil.move(ruta_completa, destino_b)
            rechazadas += 1
            print(f"❌ RECHAZADA: {archivo_orig}")
            
        elif "REVIS" in accion_norm:
            destino_rev = obtener_ruta_unica(os.path.join(CARPETA_REVISION, archivo_orig))
            shutil.move(ruta_completa, destino_rev)
            revision += 1
            print(f"⏸️  A REVISIÓN: {archivo_orig}")
            
        elif "APROB" in accion_norm or "OK" in accion_norm or "SI" in accion_norm:
            # Extraer campos corregidos manualmente del Excel
            nombre_propuesto = str(row.get("Nombre_Propuesto", "")).strip()
            sector_editado = str(row.get("Sector", "")).strip().upper()
            if not sector_editado or sector_editado == 'NAN': sector_editado = "GENERAL"
            
            oc_detectada = limpiar_cadena_oc(row.get("OC_Detectada"))
            subtotal = str(row.get("Subtotal", "")).strip()
            adjuntos_extra = str(row.get("Adjuntos_Extra", "")).strip()
            
            # ─── RE-CONSTRUIR DINÁMICAMENTE EL NOMBRE Y EL SECTOR SI SE MODIFICÓ LA OC O EL SECTOR ───
            # Eliminar la marca [DUPLICADA] si el usuario forzó APROBAR
            nombre_propuesto = nombre_propuesto.replace(" [DUPLICADA]", "")
            
            # Ajustar extensión
            if not nombre_propuesto.lower().endswith('.pdf'):
                nombre_propuesto += ".pdf"
                
            # Si el sector fue editado en la celda del Excel, asegurar que el nombre lo refleje
            # Reemplazar la palabra de sector si estaba al final del nombre base
            base_sin_ext = os.path.splitext(nombre_propuesto)[0]

            archivos_a_fusionar = [ruta_completa]
            
            # 1. Fusión de OC (usando el número de OC limpio sin .0)
            if oc_detectada:
                ruta_oc = oc_manager.buscar_pdf_oc(oc_detectada)
                if ruta_oc:
                    print(f"  📎 Adjuntando PDF de OC {oc_detectada} a {archivo_orig}")
                    archivos_a_fusionar.append(ruta_oc)
                    # Asegurar que el sufijo -OC quede en el nombre final
                    if f"-OC {oc_detectada}" not in nombre_propuesto:
                        base_n, ext_n = os.path.splitext(nombre_propuesto)
                        nombre_propuesto = f"{base_n} -OC {oc_detectada}{ext_n}"
                else:
                    print(f"  ⚠️ OC {oc_detectada} no encontrada. Se omite fusión.")
                    
            # 2. Fusión de Adjuntos Extra
            if adjuntos_extra and adjuntos_extra != 'nan':
                rutas_extra = [r.strip() for r in adjuntos_extra.split(",")]
                for r_extra in rutas_extra:
                    r_extra = r_extra.strip('"\'')
                    if os.path.exists(r_extra):
                        print(f"  📎 Adjuntando archivo extra: {os.path.basename(r_extra)}")
                        archivos_a_fusionar.append(r_extra)
                    else:
                        print(f"  ⚠️ Archivo extra no encontrado: {r_extra}")
            
            empresa_val = "EUROSAT" if ("(EUROSAT)" in nombre_propuesto.upper() or str(row.get("Empresa", "")).strip().upper() == "EUROSAT") else "TELCOM VENTURES"
            if empresa_val == "EUROSAT" and not nombre_propuesto.upper().startswith("(EUROSAT)"):
                nombre_propuesto = f"(EUROSAT) {nombre_propuesto}"

            # ─── RUTEADOR DINÁMICO AL SECTOR EDITADO ──────────────────────────────────
            nombre_limpio = re.sub(r'[\\/*?:"<>|]', "", nombre_propuesto)
            ruta_destino_c = obtener_ruta_unica(os.path.join(CARPETA_C, nombre_limpio))
            
            # Usar la carpeta del sector editado en el Excel
            carpeta_sector = os.path.join(CARPETA_AUTORIZACION, sector_editado)
            os.makedirs(carpeta_sector, exist_ok=True)
            ruta_destino_sector = obtener_ruta_unica(os.path.join(carpeta_sector, nombre_limpio))
            
            if len(archivos_a_fusionar) > 1:
                fusionar_pdfs_adjuntos(archivos_a_fusionar, ruta_destino_c)
                shutil.copy2(ruta_destino_c, ruta_destino_sector)
                if os.path.exists(ruta_completa):
                    os.remove(ruta_completa)
            else:
                shutil.copy2(ruta_completa, ruta_destino_c)
                shutil.move(ruta_completa, ruta_destino_sector)
                
            print(f"✅ APROBADA: {nombre_limpio} -> Enviada a sector {sector_editado}")
            
            # Registrar en HTML
            datos_para_html.append({
                "Fecha": datetime.now().strftime("%d/%m/%Y"),
                "Empresa": empresa_val,
                "Tipo": str(row.get("Tipo", "")),
                "Nro Comprobante": str(row.get("Nro_Comprobante", "")),
                "Razon Social": str(row.get("Razón_Social", "")),
                "Sector": sector_editado,
                "OC": oc_detectada if oc_detectada else "N/A",
                "Subtotal": subtotal if subtotal != 'nan' else ""
            })
            
            aprobadas += 1
            
        else:
            print(f"⚠️ Acción desconocida para {archivo_orig}: {accion}")
            
    # Resumen final
    print("=" * 60)
    print(f"🎯 Procesamiento finalizado. Aprobadas: {aprobadas} | Rechazadas: {rechazadas} | A Revisión: {revision}")
    
    if aprobadas > 0 and datos_para_html:
        print("🔄 Actualizando reportes HTML...")
        df_nuevos = pd.DataFrame(datos_para_html)
        
        try:
            ruta_html_maestro = REPORTE_MAESTRO_HTML
            df_historico = pd.DataFrame()
            if os.path.exists(ruta_html_maestro):
                tablas = pd.read_html(ruta_html_maestro)
                if tablas:
                    df_historico = tablas[0]
            
            if not df_historico.empty:
                df_tracked = pd.concat([df_historico, df_nuevos], ignore_index=True)
            else:
                df_tracked = df_nuevos
                
            df_tracked.drop_duplicates(subset=["Razón Social", "Nro Comprobante"], keep="last", inplace=True)
            
            reporter._escribir_html(df_tracked, ruta_html_maestro, "Reporte Histórico de Facturas Aprobadas")
            
            if os.path.exists(CARPETA_AUTORIZACION):
                sectores_unicos = df_tracked['Sector'].dropna().unique()
                for sec in sectores_unicos:
                    if not sec or sec == "GENERAL": continue
                    df_sec = df_tracked[df_tracked['Sector'] == sec]
                    if df_sec.empty: continue
                    dir_sector = os.path.join(CARPETA_AUTORIZACION, sec)
                    os.makedirs(dir_sector, exist_ok=True)
                    ruta_sec = os.path.join(dir_sector, "Reporte.html")
                    reporter._escribir_html(df_sec, ruta_sec, f"Reporte de Facturas - {sec}")
            print("✅ Reportes actualizados exitosamente.")
        except Exception as e:
            print(f"⚠️ Error actualizando reportes HTML: {e}")
            
    # Archivar el Excel procesado
    if aprobadas > 0 or rechazadas > 0 or revision > 0:
        excel_bkp = os.path.join(CARPETA_A, "Validacion_Lote_Procesado.xlsx")
        excel_bkp_unico = obtener_ruta_unica(excel_bkp)
        try:
            shutil.move(EXCEL_VALIDACION, excel_bkp_unico)
            print(f"📁 Excel original archivado como {excel_bkp_unico}.")
        except Exception as e:
            print(f"⚠️ No se pudo archivar el Excel: {e}")

if __name__ == "__main__":
    procesar_excel_lote()
    input("\nPresiona ENTER para salir...")
