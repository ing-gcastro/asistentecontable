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

def mover_archivo_seguro(origen, destino):
    """Mueve un archivo evitando bloqueos de permisos en Windows (ej: Acrobat o Vista Previa)."""
    if not os.path.exists(origen):
        return False, "Archivo no existe"
    destino_final = obtener_ruta_unica(destino)
    try:
        shutil.move(origen, destino_final)
        return True, destino_final
    except PermissionError:
        # Si shutil.move falla por bloqueo de lectura/borrado, intentar copiar al destino
        try:
            shutil.copy2(origen, destino_final)
            try:
                os.remove(origen)
            except Exception:
                print(f"  ⚠️ El archivo de origen quedó bloqueado por otro programa (ej: Acrobat/Vista Previa) y no se pudo borrar de entrada: {os.path.basename(origen)}")
            return True, destino_final
        except Exception as e:
            print(f"  ❌ Error de permisos al copiar {os.path.basename(origen)} (Asegúrate de no tenerlo abierto en un visor de PDF): {e}")
            return False, str(e)
    except Exception as e:
        print(f"  ❌ Error al mover {os.path.basename(origen)}: {e}")
        return False, str(e)

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
    
    aprobados_en_ejecucion = set()
    
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
            exito_m, _ = mover_archivo_seguro(ruta_completa, os.path.join(CARPETA_B, archivo_orig))
            if exito_m:
                rechazadas += 1
                print(f"❌ RECHAZADA: {archivo_orig}")
            
        elif "REVIS" in accion_norm:
            exito_m, _ = mover_archivo_seguro(ruta_completa, os.path.join(CARPETA_REVISION, archivo_orig))
            if exito_m:
                revision += 1
                print(f"⏸️  A REVISIÓN: {archivo_orig}")
            
        elif "APROB" in accion_norm or "OK" in accion_norm or "SI" in accion_norm:
            nro_comp = str(row.get("Nro_Comprobante", "")).strip()
            razon_social = str(row.get("Razón_Social", row.get("Razon_Social", ""))).strip()
            ident_prov = razon_social.upper()

            clave_proc = (ident_prov, nro_comp)
            if nro_comp not in ["SinNumero", "N/A", "S_N", "", "nan"]:
                if clave_proc in aprobados_en_ejecucion:
                    print(f"🛑 DUPLICADO OMITIDO EN PROCESAMIENTO: {archivo_orig} (Comp. {nro_comp}) ya fue procesado en este lote. Moviendo a 'no es factura' para evitar legajos duplicados.")
                    mover_archivo_seguro(ruta_completa, os.path.join(CARPETA_B, archivo_orig))
                    rechazadas += 1
                    continue
                aprobados_en_ejecucion.add(clave_proc)

            # Extraer campos corregidos manualmente del Excel
            nombre_propuesto = str(row.get("Nombre_Propuesto", "")).strip()
            sector_editado = str(row.get("Sector", "")).strip().upper()
            if not sector_editado or sector_editado == 'NAN': sector_editado = "GENERAL"
            
            oc_detectada = limpiar_cadena_oc(row.get("OC_Detectada"))
            subtotal = str(row.get("Subtotal", "")).strip()
            adjuntos_extra = str(row.get("Adjuntos_Extra", "")).strip()
            
            # ─── RE-CONSTRUIR DINÁMICAMENTE EL NOMBRE Y EL SECTOR SI SE MODIFICÓ LA OC O EL SECTOR ───
            # Eliminar la marca [DUPLICADA] si el usuario forzó APROBAR
            nombre_propuesto = nombre_propuesto.replace(" [DUPLICADA HISTORIAL]", "").replace(" [DUPLICADA EN LOTE]", "").replace(" [DUPLICADA]", "")
            
            # Ajustar extensión
            if not nombre_propuesto.lower().endswith('.pdf'):
                nombre_propuesto += ".pdf"
                
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
                    r_extra = r_extra.strip().strip('"\'`').strip()
                    if os.path.exists(r_extra):
                        print(f"  📎 Adjuntando archivo extra: {os.path.basename(r_extra)}")
                        archivos_a_fusionar.append(r_extra)
                    else:
                        print(f"  ⚠️ Archivo extra no encontrado: {r_extra}")
            
            valor_tarjeta = str(row.get("Tarjeta", "")).strip().upper()
            es_tarjeta = valor_tarjeta in ["SI", "S", "VISA", "VISA CORPO", "TARJETA", "TRUE", "1"]
            
            if es_tarjeta and not nombre_propuesto.upper().startswith("(VISA CORPO)"):
                nombre_propuesto = f"(VISA CORPO) {nombre_propuesto}"
            elif valor_tarjeta in ["NO", "N", "FALSE", "0"] and nombre_propuesto.upper().startswith("(VISA CORPO)"):
                nombre_propuesto = re.sub(r'^\(VISA CORPO\)\s*', '', nombre_propuesto, flags=re.IGNORECASE).strip()

            empresa_val = "EUROSAT" if ("(EUROSAT)" in nombre_propuesto.upper() or str(row.get("Empresa", "")).strip().upper() == "EUROSAT") else "TELCOM VENTURES"
            if empresa_val == "EUROSAT" and "(EUROSAT)" not in nombre_propuesto.upper():
                if nombre_propuesto.upper().startswith("(VISA CORPO)"):
                    nombre_propuesto = f"(VISA CORPO) (EUROSAT) {nombre_propuesto[12:].strip()}"
                else:
                    nombre_propuesto = f"(EUROSAT) {nombre_propuesto}"

            # ─── RUTEADOR DINÁMICO AL SECTOR EDITADO O SUBDIRECTORIO AUTORIZADA ─────
            nombre_limpio = re.sub(r'[\\/*?:"<>|]', "", nombre_propuesto)
            ruta_destino_c = obtener_ruta_unica(os.path.join(CARPETA_C, nombre_limpio))
            
            valor_autorizar = str(row.get("Autorizar_Directo", row.get("Autorizar", ""))).strip().upper()
            es_autorizada_directo = valor_autorizar in ["SI", "S", "AUTORIZADA", "AUTORIZADO", "TRUE", "1"]
            
            # Usar la carpeta del sector o subcarpeta AUTORIZADA si corresponde
            if es_autorizada_directo:
                carpeta_sector = os.path.join(CARPETA_AUTORIZACION, sector_editado, "AUTORIZADA")
                estado_registro = "AUTORIZADA"
            else:
                carpeta_sector = os.path.join(CARPETA_AUTORIZACION, sector_editado)
                estado_registro = "PENDIENTE"

            os.makedirs(carpeta_sector, exist_ok=True)
            ruta_destino_sector = obtener_ruta_unica(os.path.join(carpeta_sector, nombre_limpio))
            
            if len(archivos_a_fusionar) > 1:
                fusionar_pdfs_adjuntos(archivos_a_fusionar, ruta_destino_c)
                shutil.copy2(ruta_destino_c, ruta_destino_sector)
                if os.path.exists(ruta_completa):
                    try:
                        os.remove(ruta_completa)
                    except Exception:
                        pass
            else:
                shutil.copy2(ruta_completa, ruta_destino_c)
                mover_archivo_seguro(ruta_completa, ruta_destino_sector)
                
            msg_aut = " [AUTORIZADA DIRECTO]" if es_autorizada_directo else ""
            print(f"✅ APROBADA{msg_aut}: {nombre_limpio} -> Enviada a sector {sector_editado}")
            
            # Registrar en HTML
            datos_para_html.append({
                "Fecha": datetime.now().strftime("%d/%m/%Y"),
                "Empresa": empresa_val,
                "Tipo": str(row.get("Tipo", "")),
                "Nro Comprobante": str(row.get("Nro_Comprobante", "")),
                "Razón Social": str(row.get("Razón_Social", row.get("Razon_Social", ""))),
                "Sector": sector_editado,
                "Estado": estado_registro,
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
                if 'Razon Social' in df_historico.columns:
                    if 'Razón Social' in df_historico.columns:
                        df_historico['Razón Social'] = df_historico['Razón Social'].combine_first(df_historico['Razon Social'])
                        df_historico.drop(columns=['Razon Social'], inplace=True)
                    else:
                        df_historico.rename(columns={'Razon Social': 'Razón Social'}, inplace=True)
                df_tracked = pd.concat([df_historico, df_nuevos], ignore_index=True)
            else:
                df_tracked = df_nuevos
                
            df_tracked.drop_duplicates(subset=["Razón Social", "Nro Comprobante"], keep="last", inplace=True)
            
            reporter._escribir_html(df_tracked, ruta_html_maestro, "Reporte Histórico de Facturas Aprobadas")
            
            # Actualizar Consolidado Maestro Excel
            try:
                from config import REPORTE_MAESTRO_EXCEL
                if os.path.exists(REPORTE_MAESTRO_EXCEL):
                    try:
                        df_excel_hist = pd.read_excel(REPORTE_MAESTRO_EXCEL, dtype=str)
                        if 'Razon Social' in df_excel_hist.columns:
                            if 'Razón Social' in df_excel_hist.columns:
                                df_excel_hist['Razón Social'] = df_excel_hist['Razón Social'].combine_first(df_excel_hist['Razon Social'])
                                df_excel_hist.drop(columns=['Razon Social'], inplace=True)
                            else:
                                df_excel_hist.rename(columns={'Razon Social': 'Razón Social'}, inplace=True)
                        df_excel_tracked = pd.concat([df_excel_hist, df_nuevos], ignore_index=True)
                    except Exception:
                        df_excel_tracked = df_nuevos
                else:
                    df_excel_tracked = df_nuevos
                
                df_excel_tracked.drop_duplicates(subset=["Razón Social", "Nro Comprobante"], keep="last", inplace=True)
                df_excel_tracked.to_excel(REPORTE_MAESTRO_EXCEL, index=False)
                print(f"📊 Consolidado Excel actualizado exitosamente en: {REPORTE_MAESTRO_EXCEL}")
            except Exception as e_ex:
                print(f"⚠️ Advertencia actualizando Reporte_Maestro.xlsx: {e_ex}")

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
        exito_bkp, ruta_bkp = mover_archivo_seguro(EXCEL_VALIDACION, excel_bkp)
        if exito_bkp:
            print(f"📁 Excel original archivado como {os.path.basename(ruta_bkp)}.")
        else:
            print(f"⚠️ No se pudo archivar el Excel de validación: {ruta_bkp}")

if __name__ == "__main__":
    procesar_excel_lote()
    input("\nPresiona ENTER para salir...")
