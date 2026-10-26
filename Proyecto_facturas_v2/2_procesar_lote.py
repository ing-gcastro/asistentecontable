# 2_procesar_lote.py — PASO 2: aplica las decisiones del Excel de validación.
#   APROBAR  -> copia a \Compras y mueve a AUTORIZACION DE FACTURAS\<sector>[\AUTORIZADA]
#   RECHAZAR -> Carpeta B ("no es factura")
#   REVISION -> Carpeta de revisión
import os
import re
import shutil
import sys
import unicodedata
from datetime import datetime

import pandas as pd

from config import (CARPETA_A, CARPETA_AUTORIZACION, CARPETA_B, CARPETA_C, CARPETA_REVISION,
                    LOTE_DATOS_PATH, REPORTE_MAESTRO_EXCEL)
from estado import cargar_json, esperar_archivo_libre, registrar_facturas
from procesador_core import fusionar_pdfs_adjuntos, obtener_ruta_unica, parsear_monto
import notificador_sectores
import oc_manager
import reporter

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

EXCEL_VALIDACION = os.path.join(CARPETA_A, "Validacion_Lote.xlsx")
NUMEROS_INVALIDOS = {"SinNumero", "N/A", "S_N", "", "nan"}
VALORES_SI = {"SI", "S", "TRUE", "1", "VISA", "VISA CORPO", "TARJETA", "AUTORIZADA", "AUTORIZADO"}


def limpiar_cadena_oc(val):
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return "" if s.lower() == 'nan' else s


def _texto(row, *claves):
    for c in claves:
        v = row.get(c)
        if v is not None and not pd.isna(v) and str(v).strip().lower() != "nan":
            return str(v).strip()
    return ""


def clasificar_accion(valor):
    """Devuelve APROBAR / RECHAZAR / REVISION / None a partir del texto de la columna Acción."""
    v = unicodedata.normalize("NFKD", str(valor or "")).encode("ascii", "ignore").decode().strip().upper()
    if not v or v == "NAN":
        return None
    if v.startswith("REVIS"):
        return "REVISION"
    if v.startswith(("RECHAZ", "DESCART")) or v in ("NO", "N"):
        return "RECHAZAR"
    if v.startswith("APROB") or v in ("OK", "SI", "S"):
        return "APROBAR"
    return None


def mover_archivo_seguro(origen, destino):
    """Mueve un archivo evitando bloqueos de permisos en Windows (ej: Acrobat o Vista Previa)."""
    if not os.path.exists(origen):
        return False, "Archivo no existe"
    destino_final = obtener_ruta_unica(destino)
    try:
        shutil.move(origen, destino_final)
        return True, destino_final
    except PermissionError:
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


def _ajustar_prefijos(nombre_propuesto, row):
    valor_tarjeta = _texto(row, "Tarjeta").upper()
    if valor_tarjeta in VALORES_SI and not nombre_propuesto.upper().startswith("(VISA CORPO)"):
        nombre_propuesto = f"(VISA CORPO) {nombre_propuesto}"
    elif valor_tarjeta in ("NO", "N", "FALSE", "0") and nombre_propuesto.upper().startswith("(VISA CORPO)"):
        nombre_propuesto = re.sub(r'^\(VISA CORPO\)\s*', '', nombre_propuesto, flags=re.IGNORECASE).strip()

    empresa = "EUROSAT" if ("(EUROSAT)" in nombre_propuesto.upper() or _texto(row, "Empresa").upper() == "EUROSAT") else "TELCOM VENTURES"
    if empresa == "EUROSAT" and "(EUROSAT)" not in nombre_propuesto.upper():
        if nombre_propuesto.upper().startswith("(VISA CORPO)"):
            nombre_propuesto = f"(VISA CORPO) (EUROSAT) {nombre_propuesto[12:].strip()}"
        else:
            nombre_propuesto = f"(EUROSAT) {nombre_propuesto}"
    return nombre_propuesto, empresa


def ajustar_sector_en_nombre(nombre_propuesto, sector):
    """
    Si el usuario cambió el Sector en el Excel, el nombre del archivo debe reflejarlo
    ("... FC 123 GENERAL -OC 5.pdf" -> "... FC 123 ADM -OC 5.pdf").
    """
    sectores = sorted(set(reporter.SECTORES_VALIDOS + ["GENERAL"]), key=len, reverse=True)
    patron = r"(\s)(" + "|".join(re.escape(s) for s in sectores) + r")(?=(\s+-OC\s+\S+)?\.pdf$)"
    if re.search(patron, nombre_propuesto, flags=re.IGNORECASE):
        return re.sub(patron, lambda m: m.group(1) + sector, nombre_propuesto, count=1, flags=re.IGNORECASE)
    return nombre_propuesto


def _actualizar_excel_maestro(datos_para_html):
    """Consolidado Reporte_Maestro.xlsx (mismas columnas que el HTML)."""
    df_nuevos = pd.DataFrame(datos_para_html)
    try:
        df_hist = pd.read_excel(REPORTE_MAESTRO_EXCEL, dtype=str) if os.path.exists(REPORTE_MAESTRO_EXCEL) else pd.DataFrame()
        if 'Razon Social' in df_hist.columns:
            if 'Razón Social' in df_hist.columns:
                df_hist['Razón Social'] = df_hist['Razón Social'].combine_first(df_hist['Razon Social'])
                df_hist.drop(columns=['Razon Social'], inplace=True)
            else:
                df_hist.rename(columns={'Razon Social': 'Razón Social'}, inplace=True)
        if 'Fecha' in df_hist.columns and 'Fecha Carga' not in df_hist.columns:
            df_hist.rename(columns={'Fecha': 'Fecha Carga'}, inplace=True)
        df = pd.concat([df_hist, df_nuevos], ignore_index=True)
        df.drop_duplicates(subset=["Razón Social", "Nro Comprobante"], keep="last", inplace=True)
        if not esperar_archivo_libre(REPORTE_MAESTRO_EXCEL):
            print("⚠️ Reporte_Maestro.xlsx sigue abierto: no se actualizó (se completa en la próxima corrida).")
            return
        df.to_excel(REPORTE_MAESTRO_EXCEL, index=False)
        print(f"📊 Consolidado Excel actualizado exitosamente en: {REPORTE_MAESTRO_EXCEL}")
    except Exception as e:
        print(f"⚠️ Advertencia actualizando Reporte_Maestro.xlsx: {e}")


def procesar_excel_lote():
    print("Iniciando Procesamiento de Lote (Paso 2)...")

    if not os.path.exists(EXCEL_VALIDACION):
        print(f"❌ No se encontró el archivo {EXCEL_VALIDACION}. Ejecuta primero el Paso 1.")
        return

    if not esperar_archivo_libre(EXCEL_VALIDACION):
        print("❌ Cancelado: el Excel de validación sigue abierto.")
        return

    try:
        # dtype=str para que los números de OC o Comprobante no se conviertan en floats (.0)
        df = pd.read_excel(EXCEL_VALIDACION, engine='openpyxl', dtype=str)
    except Exception as e:
        print(f"❌ Error al leer el Excel (Asegúrate de haberlo cerrado): {e}")
        return

    if df.empty:
        print("📂 El Excel está vacío.")
        return

    for carpeta in (CARPETA_B, CARPETA_C, CARPETA_REVISION):
        os.makedirs(carpeta, exist_ok=True)

    datos_lote = cargar_json(LOTE_DATOS_PATH, {})
    datos_para_html = []
    registros = []
    aprobadas = rechazadas = revision = 0
    aprobados_en_ejecucion = set()

    for _, row in df.iterrows():
        archivo_orig = _texto(row, "Archivo_Original")
        if not archivo_orig:
            continue

        ruta_completa = os.path.join(CARPETA_A, archivo_orig)
        if not os.path.exists(ruta_completa):
            print(f"⚠️ Archivo no encontrado en Carpeta A: {archivo_orig} (Omitiendo)")
            continue

        accion_raw = _texto(row, "Acción", "Accion")
        accion = clasificar_accion(accion_raw)

        if accion == "RECHAZAR":
            exito_m, _ = mover_archivo_seguro(ruta_completa, os.path.join(CARPETA_B, archivo_orig))
            if exito_m:
                rechazadas += 1
                print(f"❌ RECHAZADA: {archivo_orig}")
            continue

        if accion == "REVISION":
            exito_m, _ = mover_archivo_seguro(ruta_completa, os.path.join(CARPETA_REVISION, archivo_orig))
            if exito_m:
                revision += 1
                print(f"⏸️  A REVISIÓN: {archivo_orig}")
            continue

        if accion != "APROBAR":
            print(f"⚠️ Acción desconocida para {archivo_orig}: '{accion_raw}' (se deja en la Carpeta A)")
            continue

        # ─── APROBAR ────────────────────────────────────────────────────────
        nro_comp = _texto(row, "Nro_Comprobante")
        razon_social = _texto(row, "Razón_Social", "Razon_Social")

        clave_proc = (razon_social.upper(), nro_comp)
        if nro_comp not in NUMEROS_INVALIDOS:
            if clave_proc in aprobados_en_ejecucion:
                print(f"🛑 DUPLICADO OMITIDO EN PROCESAMIENTO: {archivo_orig} (Comp. {nro_comp}) ya fue procesado en este lote. Moviendo a 'no es factura' para evitar legajos duplicados.")
                mover_archivo_seguro(ruta_completa, os.path.join(CARPETA_B, archivo_orig))
                rechazadas += 1
                continue
            aprobados_en_ejecucion.add(clave_proc)

        nombre_propuesto = _texto(row, "Nombre_Propuesto") or os.path.splitext(archivo_orig)[0]
        sector_editado = (_texto(row, "Sector") or "GENERAL").upper()
        oc_detectada = limpiar_cadena_oc(row.get("OC_Detectada"))
        adjuntos_extra = _texto(row, "Adjuntos_Extra")

        # Quitar la marca [DUPLICADA] si el usuario forzó APROBAR
        nombre_propuesto = re.sub(r"\s*\[DUPLICADA[^\]]*\]", "", nombre_propuesto).strip()
        if not nombre_propuesto.lower().endswith('.pdf'):
            nombre_propuesto += ".pdf"

        archivos_a_fusionar = [ruta_completa]

        # 1. Fusión de OC
        if oc_detectada:
            ruta_oc = oc_manager.buscar_pdf_oc(oc_detectada)
            if ruta_oc:
                print(f"  📎 Adjuntando PDF de OC {oc_detectada} a {archivo_orig}")
                archivos_a_fusionar.append(ruta_oc)
                if f"-OC {oc_detectada}" not in nombre_propuesto:
                    base_n, ext_n = os.path.splitext(nombre_propuesto)
                    nombre_propuesto = f"{base_n} -OC {oc_detectada}{ext_n}"
            else:
                print(f"  ⚠️ OC {oc_detectada} no encontrada. Se omite fusión.")

        # 2. Fusión de Adjuntos Extra (rutas separadas por coma)
        for r_extra in [r.strip().strip('"\'`').strip() for r in adjuntos_extra.split(",") if r.strip()]:
            if os.path.exists(r_extra):
                print(f"  📎 Adjuntando archivo extra: {os.path.basename(r_extra)}")
                archivos_a_fusionar.append(r_extra)
            else:
                print(f"  ⚠️ Archivo extra no encontrado: {r_extra}")

        nombre_propuesto, empresa_val = _ajustar_prefijos(nombre_propuesto, row)
        nombre_propuesto = ajustar_sector_en_nombre(nombre_propuesto, sector_editado)

        # ─── RUTEO: \Compras + sector (o subcarpeta AUTORIZADA) ─────────────
        nombre_limpio = re.sub(r'[\\/*?:"<>|]', "", nombre_propuesto)
        ruta_destino_c = obtener_ruta_unica(os.path.join(CARPETA_C, nombre_limpio))

        es_autorizada_directo = _texto(row, "Autorizar_Directo", "Autorizar").upper() in VALORES_SI
        if es_autorizada_directo:
            carpeta_sector = os.path.join(CARPETA_AUTORIZACION, sector_editado, "AUTORIZADA")
            estado_registro = "AUTORIZADA"
        else:
            carpeta_sector = os.path.join(CARPETA_AUTORIZACION, sector_editado)
            estado_registro = "PENDIENTE"

        try:
            os.makedirs(carpeta_sector, exist_ok=True)
            ruta_destino_sector = obtener_ruta_unica(os.path.join(carpeta_sector, nombre_limpio))

            if len(archivos_a_fusionar) > 1:
                fusionar_pdfs_adjuntos(archivos_a_fusionar, ruta_destino_c)
                shutil.copy2(ruta_destino_c, ruta_destino_sector)
                try:
                    os.remove(ruta_completa)
                except Exception:
                    print(f"  ⚠️ No se pudo borrar el original de la Carpeta A (¿abierto?): {archivo_orig}")
            else:
                shutil.copy2(ruta_completa, ruta_destino_c)
                ok, detalle = mover_archivo_seguro(ruta_completa, ruta_destino_sector)
                if not ok:
                    raise OSError(detalle)
        except Exception as e:
            print(f"❌ No se pudo archivar {archivo_orig}: {e}. Queda en la Carpeta A para reintentar.")
            continue

        nombre_final = os.path.basename(ruta_destino_c)
        msg_aut = " [AUTORIZADA DIRECTO]" if es_autorizada_directo else ""
        print(f"✅ APROBADA{msg_aut}: {nombre_final} -> Enviada a sector {sector_editado}")

        # ─── DATOS PARA REPORTES Y REGISTRO (BUG-04 / MEJ-02) ───────────────
        extra = datos_lote.get(archivo_orig, {})
        fecha_emision = _texto(row, "Fecha_Emision") or extra.get("fecha_emision") or ""
        total = parsear_monto(_texto(row, "Total"))
        if total is None:
            total = extra.get("total")
        subtotal = _texto(row, "Subtotal")

        datos_para_html.append({
            "Fecha Emisión": fecha_emision,
            "Fecha Carga": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "Empresa": empresa_val,
            "Tipo": _texto(row, "Tipo"),
            "Nro Comprobante": nro_comp,
            "Razón Social": razon_social,
            "Sector": sector_editado,
            "Estado": estado_registro,
            "Etapa Contable": "ENTRADA",
            "OC": oc_detectada or "N/A",
            "Subtotal": subtotal,
            "Total": reporter._formatear_monto(total),
            "Archivo": nombre_final,
            "Ruta Completa": ruta_destino_c,
        })

        registro = {
            "archivo": nombre_final,
            "razon_social": razon_social,
            "cuit": _texto(row, "CUIT") or extra.get("cuit"),
            "tipo": _texto(row, "Tipo"),
            "numero": nro_comp,
            "fecha_emision": fecha_emision,
            "sector": sector_editado,
            "estado": estado_registro,
            "empresa": empresa_val,
            "oc": oc_detectada,
            "subtotal": subtotal,
            "total": total,
            "items": extra.get("items") or [],
            "ruta_sector": ruta_destino_sector,
        }
        for col_excel, clave in (("Neto", "neto"), ("IVA_21", "iva_21"), ("IVA_10_5", "iva_105"),
                                 ("IVA_27", "iva_27"), ("Otros_Impuestos", "otros_impuestos")):
            valor = parsear_monto(_texto(row, col_excel))
            registro[clave] = valor if valor is not None else extra.get(clave)
        registros.append(registro)
        aprobadas += 1

    print("=" * 60)
    print(f"🎯 Procesamiento finalizado. Aprobadas: {aprobadas} | Rechazadas: {rechazadas} | A Revisión: {revision}")

    if registros:
        try:
            registrar_facturas(registros)
        except Exception as e:
            print(f"⚠️ No se pudo actualizar registro_facturas.json: {e}")

    if datos_para_html:
        print("🔄 Actualizando reportes...")
        try:
            reporter.agregar_al_reporte_maestro(datos_para_html)
            _actualizar_excel_maestro(datos_para_html)
            reporter.generar_reportes_sector()
            reporter.generar_reporte_detallado()
            print("✅ Reportes actualizados exitosamente.")
        except Exception as e:
            print(f"⚠️ Error actualizando reportes: {e}")

        # Aviso por mail a cada sector (apagado por defecto: ver config NOTIFICACION_MODO)
        try:
            notificador_sectores.notificar_lote([r for r in registros if r.get("estado") == "PENDIENTE"])
        except Exception as e:
            print(f"⚠️ Notificación a sectores no enviada: {e}")

    # Archivar el Excel procesado
    if aprobadas or rechazadas or revision:
        excel_bkp = os.path.join(CARPETA_A, "Validacion_Lote_Procesado.xlsx")
        exito_bkp, ruta_bkp = mover_archivo_seguro(EXCEL_VALIDACION, excel_bkp)
        if exito_bkp:
            print(f"📁 Excel original archivado como {os.path.basename(ruta_bkp)}.")
        else:
            print(f"⚠️ No se pudo archivar el Excel de validación: {ruta_bkp}")


if __name__ == "__main__":
    from estado import iniciar_log_auditoria
    iniciar_log_auditoria("paso2_procesar_lote")
    procesar_excel_lote()
    input("\nPresiona ENTER para salir...")
