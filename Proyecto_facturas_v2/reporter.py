# reporter.py — Reportes HTML/Excel del ciclo de compras.
#
#   Reporte_Maestro.html   : histórico de facturas aprobadas + etapa contable + estado.
#   <sector>\Reporte.html  : facturas físicamente presentes en cada sector de AUTORIZACION.
#   Reportesector.html     : consolidado de todos los sectores.
#   Reporte_Detallado.html : desglose de impuestos, ítems y evolución de precios (MEJ-02).
#
# Uso directo:  python reporter.py            (regenera todo)
#               python reporter.py --detallado (solo el detallado, sin escanear la red)
import html
import io
import json
import os
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

from config import (CACHE_DATOS_PDF, CARPETA_AUTORIZACION, CARPETA_C, CARPETAS_ETAPAS,
                    REPORTE_DETALLADO_HTML, REPORTE_MAESTRO_HTML, REPORTE_MAX_LECTURAS_PDF,
                    REPORTE_SECTORES_HTML)
from estado import cargar_json, cargar_registro, guardar_json

HTML_SALIDA = REPORTE_MAESTRO_HTML  # alias histórico

SECTORES_VALIDOS = [
    "RRHH",
    "REDES Y SOPORTE",
    "OPERACIONES",
    "PROGRAMACION",
    "ADM",
    "PRESIDENCIA",
    "INTERNET",
]

COLS_MAESTRO = [
    "Fecha Emisión", "Fecha Carga", "Empresa", "Razón Social", "Tipo", "Nro Comprobante",
    "Subtotal", "Total", "OC", "Información Adicional", "Sector", "Estado",
    "Etapa Contable", "Archivo", "Ruta Completa",
]
COLS_SECTOR = COLS_MAESTRO[:12]
_COLS_TEXTO = COLS_MAESTRO + ["Fecha", "Razon Social"]
_CARPETAS_EXCLUIDAS_SECTOR = {"ARCHIVOS EN REVISION", "REPORTES"}


# ─── LECTURA DE REPORTES EXISTENTES ──────────────────────────────────────────
def _limpiar_numero(valor):
    s = "" if valor is None else str(valor).strip()
    if s.lower() in ("nan", "none"):
        return ""
    return s[:-2] if s.endswith(".0") else s


def leer_tabla_html(ruta_html):
    """
    Lee la tabla de un reporte generado por este módulo, SIN corromper datos:
    - todo como texto (pandas convertía "47336" en 47336.0 y rompía la detección de duplicados),
    - sin separador de miles (evita que "10.000,00" se interprete como número),
    - descarta la fila de filtros del encabezado (que volvía como una fila vacía).
    Devuelve un DataFrame vacío si el archivo no existe o no se puede leer.
    """
    if not ruta_html or not os.path.exists(ruta_html):
        return pd.DataFrame()
    try:
        with open(ruta_html, "r", encoding="utf-8", errors="ignore") as f:
            contenido = f.read()
        tablas = pd.read_html(io.StringIO(contenido), thousands=None, keep_default_na=False,
                              converters={c: str for c in _COLS_TEXTO})
    except Exception as e:
        print(f"⚠️ No se pudo leer el reporte {os.path.basename(ruta_html)}: {e}")
        return pd.DataFrame()
    if not tablas:
        return pd.DataFrame()

    df = tablas[0]
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [c[0] for c in df.columns]
    df.columns = [str(c).strip() for c in df.columns]
    df = df.astype(str).replace({"nan": "", "NaN": "", "None": ""})
    df = df[(df != "").any(axis=1)].copy()

    if "Razon Social" in df.columns:
        if "Razón Social" in df.columns:
            df["Razón Social"] = df["Razón Social"].where(df["Razón Social"] != "", df["Razon Social"])
            df.drop(columns=["Razon Social"], inplace=True)
        else:
            df.rename(columns={"Razon Social": "Razón Social"}, inplace=True)
    # Reportes anteriores a v3: "Fecha" era la fecha de carga (fecha de modificación del archivo)
    if "Fecha" in df.columns and "Fecha Carga" not in df.columns:
        df.rename(columns={"Fecha": "Fecha Carga"}, inplace=True)
    elif "Fecha" in df.columns:
        df.drop(columns=["Fecha"], inplace=True)
    for col in ("Nro Comprobante", "OC"):
        if col in df.columns:
            df[col] = df[col].map(_limpiar_numero)
    return df.reset_index(drop=True)


def _razon_sin_prefijos(razon):
    return re.sub(r'^(\([^)]+\)\s*)+', '', str(razon).strip().upper()).strip()


def obtener_facturas_procesadas(ruta_html):
    """Set de (RAZÓN SOCIAL, NRO) ya registradas en el reporte (para evitar duplicados)."""
    df = leer_tabla_html(ruta_html)
    procesadas = set()
    if df.empty or "Razón Social" not in df.columns or "Nro Comprobante" not in df.columns:
        return procesadas
    for razon_raw, nro in zip(df["Razón Social"], df["Nro Comprobante"]):
        razon = _razon_sin_prefijos(razon_raw)
        nro = _limpiar_numero(nro)
        if razon and nro and razon not in ("NAN", "DESCONOCIDO", "PROVEEDOR_DESCONOCIDO") and nro.upper() != "NAN":
            procesadas.add((razon, nro))
    return procesadas


# ─── DATOS DE CADA PDF (registro de Paso 2 + caché de lecturas) ──────────────
class LectorDatosFactura:
    """
    Resuelve fecha de emisión e importes de un PDF archivado (BUG-04), en este orden:
      1) registro_facturas.json (lo escribe Paso 2 con los datos del Excel),
      2) cache_datos_pdf.json (lecturas anteriores),
      3) lectura liviana del PDF (con tope REPORTE_MAX_LECTURAS_PDF por corrida).
    """

    def __init__(self, max_lecturas=None):
        max_lecturas = REPORTE_MAX_LECTURAS_PDF if max_lecturas is None else max_lecturas
        self.registro = cargar_registro()
        self.cache = cargar_json(CACHE_DATOS_PDF, {})
        self.restantes = max_lecturas
        self.pendientes = 0
        self._cambios = False

    def datos(self, nombre_archivo, ruta=None):
        if nombre_archivo in self.registro:
            return self.registro[nombre_archivo]
        try:
            tam = os.path.getsize(ruta) if ruta and os.path.exists(ruta) else None
        except OSError:
            tam = None
        en_cache = self.cache.get(nombre_archivo)
        if en_cache and (tam is None or en_cache.get("tam") == tam):
            return en_cache
        if tam is None:
            return {}
        if self.restantes <= 0:
            self.pendientes += 1
            return {}
        from procesador_core import extraer_datos_livianos
        self.restantes -= 1
        leidos = extraer_datos_livianos(ruta) or {}
        leidos["tam"] = tam
        self.cache[nombre_archivo] = leidos
        self._cambios = True
        return leidos

    def guardar(self):
        if self._cambios:
            try:
                guardar_json(CACHE_DATOS_PDF, self.cache)
            except Exception as e:
                print(f"⚠️ No se pudo guardar la caché de datos de PDFs: {e}")
        if self.pendientes:
            print(f"ℹ️ Quedaron {self.pendientes} PDF(s) sin leer en esta corrida (tope {REPORTE_MAX_LECTURAS_PDF}); "
                  f"se completan en las próximas ejecuciones.")


def _formatear_monto(valor):
    if valor in (None, ""):
        return ""
    try:
        v = float(valor)
    except (TypeError, ValueError):
        return str(valor)
    return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


# ─── ESTADOS Y NOMBRES DE ARCHIVO ────────────────────────────────────────────
def _estado_por_ruta(rel_path):
    rel = rel_path.upper()
    if "AUTORIZAD" in rel:
        return "AUTORIZADA"
    if "REVIS" in rel:
        return "A REVISIÓN"
    return "PENDIENTE"


def obtener_estados_autorizacion():
    """Escanea físicamente AUTORIZACION DE FACTURAS para obtener el estado real de cada comprobante."""
    estados_por_archivo = {}
    estados_por_datos = {}
    if not os.path.exists(CARPETA_AUTORIZACION):
        return estados_por_archivo, estados_por_datos

    for root, _dirs, files in os.walk(CARPETA_AUTORIZACION):
        for f in files:
            if not f.lower().endswith(".pdf"):
                continue
            estado = _estado_por_ruta(os.path.relpath(root, CARPETA_AUTORIZACION))
            estados_por_archivo[f] = estado
            info = parsear_nombre_aprobado(f)
            rs = str(info["Razón Social"]).strip().upper()
            nro = _limpiar_numero(info["Nro Comprobante"])
            if rs and nro:
                estados_por_datos[(rs, nro)] = estado
    return estados_por_archivo, estados_por_datos


def parsear_nombre_aprobado(nombre_archivo):
    nombre_sin_ext = Path(nombre_archivo).stem.strip()

    razon_social = "DESCONOCIDO"
    tipo = "FC"
    numero = "SinNumero"
    oc = "N/A"
    sector = "GENERAL"
    empresa = "TELCOM VENTURES"
    info_adicional = ""

    if "EUROSAT" in nombre_sin_ext.upper():
        empresa = "EUROSAT"

    for s in SECTORES_VALIDOS:
        if re.search(r'\b' + re.escape(s) + r'\b', nombre_sin_ext, re.IGNORECASE):
            sector = s
            break

    m_tipo = re.search(r'\b(FC|NC|ND)\b', nombre_sin_ext, re.IGNORECASE)
    if m_tipo:
        tipo = m_tipo.group(1).upper()
        partes = re.split(r'\b' + tipo + r'\b', nombre_sin_ext, flags=re.IGNORECASE, maxsplit=1)

        izquierda = partes[0].strip()
        rs_limpia = re.sub(r'^(\([^)]+\)\s*)+', '', izquierda).strip()
        if rs_limpia:
            razon_social = rs_limpia

        if len(partes) > 1:
            tokens = partes[1].strip().split()
            if tokens:
                if tokens[0].upper() != 'OC':
                    numero = _limpiar_numero(tokens[0])
                    resto_tokens = tokens[1:]
                else:
                    resto_tokens = tokens

                resto_texto = " ".join(resto_tokens)

                parentesis_extra = re.findall(r'\(([^)]+)\)', resto_texto)
                if parentesis_extra:
                    info_adicional = " ".join(parentesis_extra)
                    resto_texto = re.sub(r'\([^)]+\)', '', resto_texto)

                m_oc = re.search(r'\bOC\s*[:\.]?\s*(\d+)', resto_texto, re.IGNORECASE)
                if m_oc:
                    oc = m_oc.group(1)
                    partes_oc = re.split(r'\bOC\s*[:\.]?\s*\d+', resto_texto, flags=re.IGNORECASE, maxsplit=1)
                    sucia = partes_oc[1].strip() if len(partes_oc) > 1 else ""
                else:
                    sucia = resto_texto.strip()
                sucia = sucia.strip(" -")
                if sector != "GENERAL":
                    sucia = re.sub(r'\b' + re.escape(sector) + r'\b', '', sucia, flags=re.IGNORECASE).strip(" -")
                if sucia:
                    info_adicional = (info_adicional + " " + sucia).strip()

    return {
        "Archivo Original": nombre_archivo,
        "Empresa": empresa,
        "Razón Social": razon_social,
        "Tipo": tipo,
        "Nro Comprobante": numero,
        "OC": oc,
        "Información Adicional": info_adicional,
        "Sector": sector,
    }


def _fecha_mtime(ruta):
    try:
        return datetime.fromtimestamp(os.path.getmtime(ruta)).strftime('%d/%m/%Y %H:%M:%S')
    except Exception:
        return datetime.now().strftime('%d/%m/%Y %H:%M:%S')


def _a_fecha(valor):
    return pd.to_datetime(pd.Series(valor, dtype="object"), dayfirst=True, errors="coerce", format="mixed")


def _ordenar_por_fecha(df):
    if df.empty:
        return df
    emision = _a_fecha(df.get("Fecha Emisión", pd.Series([""] * len(df))).tolist())
    carga = _a_fecha(df.get("Fecha Carga", pd.Series([""] * len(df))).tolist())
    orden = emision.fillna(carga)
    return df.assign(_orden=orden.values).sort_values("_orden", ascending=False, na_position="last").drop(columns="_orden")


def _completar_con_datos(df, lector):
    """Rellena Fecha Emisión y Total vacíos desde registro / caché / PDF."""
    for idx, row in df.iterrows():
        if row.get("Fecha Emisión") and row.get("Total"):
            continue
        nombre = row.get("Archivo", "")
        if not nombre:
            continue
        d = lector.datos(nombre, row.get("Ruta Completa", "") or None)
        if not row.get("Fecha Emisión") and d.get("fecha_emision"):
            df.at[idx, "Fecha Emisión"] = d["fecha_emision"]
        if not row.get("Total") and d.get("total") not in (None, ""):
            df.at[idx, "Total"] = _formatear_monto(d["total"])
    return df


def _fusionar(df_lista):
    """
    Une varias fuentes de filas por (Razón Social, Nro Comprobante): para cada columna
    se queda con el último valor NO vacío (así el histórico conserva Subtotal/Fecha
    Emisión aunque el escaneo de la carpeta no los tenga).
    """
    partes = [d for d in df_lista if d is not None and not d.empty]
    if not partes:
        return pd.DataFrame(columns=COLS_MAESTRO)
    df = pd.concat(partes, ignore_index=True)
    for col in COLS_MAESTRO:
        if col not in df.columns:
            df[col] = ""
    df = df.fillna("").astype(str)
    df["Razón Social"] = df["Razón Social"].replace("", "DESCONOCIDO")
    df["Nro Comprobante"] = df["Nro Comprobante"].map(_limpiar_numero).replace("", "SinNumero")
    df["_k_rs"] = df["Razón Social"].str.strip().str.upper()
    df["_k_nro"] = df["Nro Comprobante"].str.strip()
    df = df.replace("", pd.NA)
    df = df.groupby(["_k_rs", "_k_nro"], sort=False, as_index=False).last()
    return df.drop(columns=["_k_rs", "_k_nro"]).fillna("")


# ─── REPORTE MAESTRO ─────────────────────────────────────────────────────────
def _filas_desde_datos_nuevos(datos_nuevos):
    filas = []
    for d in datos_nuevos or []:
        fila = {c: "" for c in COLS_MAESTRO}
        fila.update({k: ("" if v is None else str(v)) for k, v in d.items() if k in COLS_MAESTRO})
        filas.append(fila)
    return pd.DataFrame(filas)


def _escanear_compras():
    """MEJ-04: solo los PDF sueltos de la raíz de \\Compras (sin subcarpetas)."""
    filas = []
    ruta_fc = Path(CARPETA_C)
    if not ruta_fc.exists():
        return pd.DataFrame()
    for archivo in ruta_fc.glob("*.pdf"):
        info = parsear_nombre_aprobado(archivo.name)
        filas.append({
            "Fecha Carga": _fecha_mtime(str(archivo)),
            "Empresa": info["Empresa"],
            "Razón Social": info["Razón Social"],
            "Tipo": info["Tipo"],
            "Nro Comprobante": info["Nro Comprobante"],
            "OC": info["OC"],
            "Información Adicional": info["Información Adicional"],
            "Sector": info["Sector"],
            "Archivo": archivo.name,
            "Ruta Completa": str(archivo),
        })
    return pd.DataFrame(filas)


def _buscar_etapas(rastreados):
    """Busca "gemelos" del archivo en las carpetas contables para saber su etapa."""
    etapas = {}
    for carpeta, recursiva, etapa, nivel in CARPETAS_ETAPAS:
        ruta_base = Path(carpeta)
        if not ruta_base.exists():
            continue
        archivos_pdf = ruta_base.rglob("*.pdf") if recursiva else ruta_base.glob("*.pdf")
        for archivo in archivos_pdf:
            info = parsear_nombre_aprobado(archivo.name)
            clave = (str(info["Razón Social"]).strip().upper(), _limpiar_numero(info["Nro Comprobante"]))
            if clave in rastreados and (clave not in etapas or nivel > etapas[clave][0]):
                etapas[clave] = (nivel, etapa, str(archivo))
    return etapas


def generar_reporte_html(datos_nuevos=None, ruta_html=None, con_sectores=True):
    """
    Regenera el Reporte Maestro (y por defecto los de sectores).
    datos_nuevos: filas recién aprobadas por Paso 2 (claves = COLS_MAESTRO).
    """
    ruta_html = ruta_html or REPORTE_MAESTRO_HTML
    lector = LectorDatosFactura()
    estados_por_archivo, estados_por_datos = obtener_estados_autorizacion()

    df_tracked = _fusionar([leer_tabla_html(ruta_html), _escanear_compras(), _filas_desde_datos_nuevos(datos_nuevos)])
    if df_tracked.empty:
        print("⚠️ No hay datos en el histórico ni en \\Compras para mostrar.")
        if con_sectores:
            generar_reportes_sector(lector)
        lector.guardar()
        return

    rastreados = set(zip(df_tracked["Razón Social"].str.strip().str.upper(),
                         df_tracked["Nro Comprobante"].str.strip()))
    etapas = _buscar_etapas(rastreados)

    for idx, row in df_tracked.iterrows():
        clave = (str(row["Razón Social"]).strip().upper(), str(row["Nro Comprobante"]).strip())
        if clave in etapas:
            _nivel, etapa, ruta = etapas[clave]
            df_tracked.at[idx, "Etapa Contable"] = etapa
            df_tracked.at[idx, "Ruta Completa"] = ruta
        else:
            df_tracked.at[idx, "Etapa Contable"] = "ENTRADA"

        nombre_arch = str(row.get("Archivo", "")).strip()
        estado = estados_por_archivo.get(nombre_arch) or estados_por_datos.get(clave) or "PENDIENTE"
        df_tracked.at[idx, "Estado"] = estado

    df_tracked = _completar_con_datos(df_tracked, lector)
    df_final = _ordenar_por_fecha(df_tracked)

    _escribir_html(df_final, ruta_html, "Reporte Histórico de Facturas Aprobadas")
    if con_sectores:
        generar_reportes_sector(lector)
    lector.guardar()
    return df_final


def generar_reportes_sector(lector=None):
    """<sector>\\Reporte.html y Reportesector.html, 100% según la ubicación física de los PDF."""
    if not os.path.exists(CARPETA_AUTORIZACION):
        return
    propio = lector is None
    lector = lector or LectorDatosFactura()
    todos = []

    for sec in sorted(os.listdir(CARPETA_AUTORIZACION)):
        dir_sector = os.path.join(CARPETA_AUTORIZACION, sec)
        if not os.path.isdir(dir_sector) or sec.upper() in _CARPETAS_EXCLUIDAS_SECTOR:
            continue
        registros = []
        for root, _dirs, files in os.walk(dir_sector):
            for f in files:
                if not f.lower().endswith(".pdf"):
                    continue
                ruta_pdf = os.path.join(root, f)
                info = parsear_nombre_aprobado(f)
                d = lector.datos(f, ruta_pdf)
                registros.append({
                    "Fecha Emisión": d.get("fecha_emision", ""),
                    "Fecha Carga": _fecha_mtime(ruta_pdf),
                    "Empresa": info["Empresa"],
                    "Razón Social": info["Razón Social"],
                    "Tipo": info["Tipo"],
                    "Nro Comprobante": _limpiar_numero(info["Nro Comprobante"]),
                    "Subtotal": d.get("subtotal", "") or "",
                    "Total": _formatear_monto(d.get("total")),
                    "OC": info["OC"],
                    "Información Adicional": info["Información Adicional"],
                    "Sector": sec,
                    "Estado": _estado_por_ruta(os.path.relpath(root, dir_sector)),
                })
        df_sec = pd.DataFrame(registros)
        if df_sec.empty:
            continue
        df_sec = _ordenar_por_fecha(df_sec.drop_duplicates(subset=["Razón Social", "Nro Comprobante"], keep="last"))
        _escribir_html(df_sec, os.path.join(dir_sector, "Reporte.html"), f"Reporte de Facturas - {sec}", es_sector=True)
        todos.extend(registros)

    df_consolidado = pd.DataFrame(todos)
    if not df_consolidado.empty:
        df_consolidado = _ordenar_por_fecha(
            df_consolidado.drop_duplicates(subset=["Razón Social", "Nro Comprobante"], keep="last"))
        _escribir_html(df_consolidado, REPORTE_SECTORES_HTML, "Reporte Consolidado por Sector", es_sector=True)
    if propio:
        lector.guardar()


# ─── ESCRITURA HTML ──────────────────────────────────────────────────────────
def obtener_data_order_fecha(fecha_val):
    """Atributo data-order (AAAA-MM-DD hh:mm:ss) para que DataTables ordene bien las fechas dd/mm/aaaa."""
    s = "" if fecha_val is None else str(fecha_val).strip()
    if not s or s.lower() == 'nan':
        return ""
    try:
        partes = s.split()
        hora_p = partes[1] if len(partes) > 1 else "00:00:00"
        sep = "/" if "/" in partes[0] else "-"
        fecha_p = partes[0].split(sep)
        if len(fecha_p) == 3:
            if len(fecha_p[0]) == 4:
                anio, mes, dia = fecha_p
            else:
                dia, mes, anio = fecha_p
            return f'data-order="{anio}-{mes.zfill(2)}-{dia.zfill(2)} {hora_p}"'
    except Exception:
        pass
    return ""


def _badge_estado(valor):
    v = str(valor).strip().upper()
    if "AUTORIZ" in v:
        return '<span class="badge bg-success">AUTORIZADA</span>'
    if "REVIS" in v:
        return '<span class="badge bg-warning text-dark">A REVISIÓN</span>'
    if v == "N/A":
        return '<span class="badge bg-danger">N/A</span>'
    return '<span class="badge bg-secondary">PENDIENTE</span>'


_BADGES_ETAPA = {
    "PAGADA": "bg-success", "PARA PAGAR": "bg-primary", "CONTABILIZADA": "bg-info text-dark",
    "ENTRADA": "bg-secondary",
}


def _celda(col, valor):
    e = html.escape("" if valor is None else str(valor))
    if col in ("Fecha Emisión", "Fecha Carga"):
        return f"<td {obtener_data_order_fecha(valor)}>{e}</td>"
    if col == "Empresa":
        return f'<td><span class="badge bg-secondary">{e or "TELCOM VENTURES"}</span></td>'
    if col == "Razón Social":
        return f'<td class="fw-bold">{e}</td>'
    if col in ("Subtotal", "Total"):
        return f'<td class="text-end fw-bold text-success">{e}</td>'
    if col == "Sector":
        return f'<td><span class="badge bg-info text-dark">{e or "GENERAL"}</span></td>'
    if col == "Estado":
        return f"<td>{_badge_estado(valor)}</td>"
    if col == "Etapa Contable":
        v = str(valor).strip().upper()
        return f'<td><span class="badge {_BADGES_ETAPA.get(v, "bg-light text-dark")}">{html.escape(v or "OTRA")}</span></td>'
    if col == "Ruta Completa":
        return f'<td style="font-family: monospace; font-size: 0.75rem; word-break: break-all;">{e}</td>'
    return f"<td>{e}</td>"


_PLANTILLA_TABLA = """<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
    <meta http-equiv="Pragma" content="no-cache">
    <meta http-equiv="Expires" content="0">
    <title>__TITULO__</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdn.datatables.net/1.13.6/css/dataTables.bootstrap5.min.css" rel="stylesheet">
    <style>
        body { font-family: Arial, sans-serif; background-color: #f4f7f6; margin: 0; padding: 10px; }
        .container-fluid { max-width: 98%; margin: auto; background: white; padding: 25px; border-radius: 8px; box-shadow: 0px 0px 10px rgba(0,0,0,0.1); }
        th, td { padding: 10px; text-align: left; font-size: 0.85rem; }
        thead tr:first-child th { background-color: #0d6efd; color: white; }
        thead tr:nth-child(2) th { background-color: #f8f9fa; padding: 5px; }
        thead input { width: 100%; padding: 4px; box-sizing: border-box; font-size: 0.75rem; border: 1px solid #ced4da; border-radius: 4px; }
        .footer { margin-top: 20px; font-size: 12px; color: #777; text-align: right; }
    </style>
</head>
<body>
    <div class="container-fluid">
        <h2 class="text-primary fw-bold">📋 __TITULO__</h2>
        <p class="text-muted">Última actualización: <strong>__TIMESTAMP__</strong> | Total de registros: <strong>__TOTAL__</strong></p>
        <hr>
        <div class="table-responsive">
            <table class="table table-striped table-bordered table-hover align-middle" id="tabla_reporte">
                <thead>
                    <tr>__HEADERS__</tr>
                    <tr>__FILTROS__</tr>
                </thead>
                <tbody>
__FILAS__
                </tbody>
            </table>
        </div>
        <div class="footer">Sistema de Gestión Documental - __TITULO__</div>
    </div>
    <script src="https://code.jquery.com/jquery-3.7.0.min.js"></script>
    <script src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
    <script src="https://cdn.datatables.net/1.13.6/js/dataTables.bootstrap5.min.js"></script>
    <script>
        $(document).ready(function() {
            var table = $('#tabla_reporte').DataTable({
                "language": { "url": "https://cdn.datatables.net/plug-ins/1.13.6/i18n/es-ES.json" },
                "orderCellsTop": true,
                "pageLength": 25,
                "order": [[0, "desc"]]
            });
            $('#tabla_reporte thead tr:eq(1) input').on('keyup change clear', function () {
                var index = $(this).parent().index();
                if (table.column(index).search() !== this.value) {
                    table.column(index).search(this.value).draw();
                }
            });
        });
    </script>
</body>
</html>
"""


def _escribir_html(df_final, ruta_html, titulo, es_sector=False):
    columnas = COLS_SECTOR if es_sector else COLS_MAESTRO
    headers = "".join(f"<th>{html.escape(c)}</th>" for c in columnas)
    filtros = "".join(f'<th><input type="text" placeholder="Filtrar {html.escape(c)}" /></th>' for c in columnas)

    filas = []
    for _, row in df_final.iterrows():
        celdas = "".join(_celda(c, row.get(c, "")) for c in columnas)
        filas.append(f"                    <tr>{celdas}</tr>")

    contenido = (_PLANTILLA_TABLA
                 .replace("__TITULO__", html.escape(titulo))
                 .replace("__TIMESTAMP__", datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
                 .replace("__TOTAL__", str(len(df_final)))
                 .replace("__HEADERS__", headers)
                 .replace("__FILTROS__", filtros)
                 .replace("__FILAS__", "\n".join(filas)))

    try:
        tmp = ruta_html + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            f.write(contenido)
        os.replace(tmp, ruta_html)
        print(f"Reporte actualizado con exito en: {ruta_html} (Total registros: {len(df_final)})")
    except Exception as e:
        print(f"⚠️ No se pudo escribir {ruta_html}: {e}")


# ─── REPORTE MAESTRO RÁPIDO (Paso 2) ─────────────────────────────────────────
def agregar_al_reporte_maestro(datos_nuevos, ruta_html=None):
    """
    Agrega las facturas recién aprobadas al Reporte Maestro SIN escanear las carpetas
    contables (rápido). La etapa contable se recalcula en la regeneración completa
    (ejecutar_reportes.bat / monitor de Compras).
    """
    ruta_html = ruta_html or REPORTE_MAESTRO_HTML
    df_nuevos = _filas_desde_datos_nuevos(datos_nuevos)
    df = _fusionar([leer_tabla_html(ruta_html), df_nuevos])
    for col, defecto in (("Estado", "PENDIENTE"), ("Etapa Contable", "ENTRADA")):
        df[col] = df[col].replace("", defecto)
    df = _ordenar_por_fecha(df)
    _escribir_html(df, ruta_html, "Reporte Histórico de Facturas Aprobadas")
    return df


# ─── REPORTE DETALLADO (MEJ-02) ──────────────────────────────────────────────
def _filas_detalladas(lector):
    filas = []
    fuentes = list(lector.registro.items()) + [(k, v) for k, v in lector.cache.items() if k not in lector.registro]
    for nombre, d in fuentes:
        if not isinstance(d, dict):
            continue
        info = parsear_nombre_aprobado(nombre)
        filas.append({
            "fecha": d.get("fecha_emision", ""),
            "razon_social": d.get("razon_social") or info["Razón Social"],
            "cuit": d.get("cuit", ""),
            "tipo": d.get("tipo") or info["Tipo"],
            "numero": _limpiar_numero(d.get("numero") or info["Nro Comprobante"]),
            "sector": d.get("sector") or info["Sector"],
            "neto": d.get("neto"),
            "iva_21": d.get("iva_21"),
            "iva_105": d.get("iva_105"),
            "iva_27": d.get("iva_27"),
            "otros": d.get("otros_impuestos"),
            "total": d.get("total"),
            "items": d.get("items") or [],
            "archivo": nombre,
        })
    return filas


_PLANTILLA_DETALLADO = """<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Reporte Detallado de Facturas</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdn.datatables.net/1.13.6/css/dataTables.bootstrap5.min.css" rel="stylesheet">
    <style>
        body { background-color: #f4f7f6; padding: 10px; font-family: Arial, sans-serif; }
        .caja { max-width: 98%; margin: auto; background: white; padding: 25px; border-radius: 8px; box-shadow: 0 0 10px rgba(0,0,0,0.1); }
        th, td { font-size: 0.82rem; }
        thead th { background-color: #0d6efd; color: white; }
        .num { text-align: right; white-space: nowrap; }
        #grafico-wrap { position: relative; height: 340px; }
        .items { font-size: 0.75rem; color: #555; }
    </style>
</head>
<body>
<div class="caja">
    <h2 class="text-primary fw-bold">📊 Reporte Detallado de Facturas</h2>
    <p class="text-muted">Última actualización: <strong>__TIMESTAMP__</strong> | Comprobantes: <strong id="total-reg"></strong>
       | Desglose de impuestos, ítems y evolución de precios por proveedor.</p>
    <hr>
    <div class="row g-3 align-items-end mb-3">
        <div class="col-md-6">
            <label class="form-label fw-bold" for="sel-prov">Proveedor (clic en una fila también lo selecciona)</label>
            <select id="sel-prov" class="form-select"><option value="">— Todos —</option></select>
        </div>
        <div class="col-md-6 text-muted small" id="resumen-prov"></div>
    </div>
    <div id="grafico-wrap" class="mb-4"><canvas id="grafico"></canvas></div>
    <div class="table-responsive">
        <table id="tabla" class="table table-striped table-bordered table-hover align-middle w-100">
            <thead><tr>
                <th>Fecha Emisión</th><th>Razón Social</th><th>CUIT</th><th>Tipo</th><th>Nro</th><th>Sector</th>
                <th>Neto</th><th>IVA 21%</th><th>IVA 10,5%</th><th>IVA 27%</th><th>Otros imp.</th><th>Total</th>
                <th>Ítems (descripción · cant. · precio unit.)</th>
            </tr></thead>
        </table>
    </div>
    <div class="text-end text-muted small mt-3">Los importes e ítems se extraen automáticamente del PDF/QR: verificar contra el original ante dudas.</div>
</div>
<script id="datos" type="application/json">__DATOS__</script>
<script src="https://code.jquery.com/jquery-3.7.0.min.js"></script>
<script src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
<script src="https://cdn.datatables.net/1.13.6/js/dataTables.bootstrap5.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<script>
(function () {
    var datos = JSON.parse(document.getElementById('datos').textContent);
    function iso(f) { if (!f) return ''; var p = f.split('/'); return p.length === 3 ? p[2] + '-' + p[1].padStart(2, '0') + '-' + p[0].padStart(2, '0') : f; }
    function money(v) { return (v === null || v === undefined || v === '') ? '' : Number(v).toLocaleString('es-AR', {minimumFractionDigits: 2, maximumFractionDigits: 2}); }
    function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return {'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;'}[c]; }); }
    document.getElementById('total-reg').textContent = datos.length;

    var provs = Array.from(new Set(datos.map(function (d) { return d.razon_social; }))).sort();
    var sel = document.getElementById('sel-prov');
    provs.forEach(function (p) { var o = document.createElement('option'); o.value = p; o.textContent = p; sel.appendChild(o); });

    var tabla = $('#tabla').DataTable({
        data: datos,
        pageLength: 25,
        order: [[0, 'desc']],
        language: { url: 'https://cdn.datatables.net/plug-ins/1.13.6/i18n/es-ES.json' },
        columns: [
            { data: 'fecha', render: function (v, t) { return t === 'sort' ? iso(v) : esc(v); } },
            { data: 'razon_social', render: esc, className: 'fw-bold' },
            { data: 'cuit', render: esc }, { data: 'tipo', render: esc }, { data: 'numero', render: esc },
            { data: 'sector', render: esc },
            { data: 'neto', render: money, className: 'num' }, { data: 'iva_21', render: money, className: 'num' },
            { data: 'iva_105', render: money, className: 'num' }, { data: 'iva_27', render: money, className: 'num' },
            { data: 'otros', render: money, className: 'num' },
            { data: 'total', render: money, className: 'num fw-bold text-success' },
            { data: 'items', orderable: false, render: function (items) {
                return '<div class="items">' + (items || []).slice(0, 5).map(function (i) {
                    return esc(i.descripcion) + ' · ' + money(i.cantidad) + ' · $' + money(i.precio_unitario);
                }).join('<br>') + ((items || []).length > 5 ? '<br>…' : '') + '</div>';
            } }
        ]
    });

    var chart = null;
    var colores = ['#0d6efd', '#dc3545', '#198754', '#fd7e14', '#6f42c1', '#20c997'];
    function dibujar(prov) {
        var filas = datos.filter(function (d) { return d.razon_social === prov && d.fecha; })
                         .sort(function (a, b) { return iso(a.fecha) < iso(b.fecha) ? -1 : 1; });
        var series = [{ label: 'Total del comprobante', data: filas.filter(function (d) { return d.total != null; })
                        .map(function (d) { return { x: iso(d.fecha), y: d.total }; }) }];
        var porItem = {};
        filas.forEach(function (d) { (d.items || []).forEach(function (i) {
            var k = (i.descripcion || '').toUpperCase().slice(0, 40);
            (porItem[k] = porItem[k] || []).push({ x: iso(d.fecha), y: i.precio_unitario });
        }); });
        Object.keys(porItem).sort(function (a, b) { return porItem[b].length - porItem[a].length; }).slice(0, 5)
            .forEach(function (k) { series.push({ label: 'Precio unit.: ' + k, data: porItem[k] }); });
        series.forEach(function (s, i) { s.borderColor = colores[i % colores.length]; s.backgroundColor = s.borderColor; s.tension = 0.2; });

        var totales = filas.map(function (d) { return d.total; }).filter(function (v) { return v != null; });
        document.getElementById('resumen-prov').textContent = filas.length ? (filas.length + ' comprobantes con fecha · último total $' +
            money(totales[totales.length - 1]) + (totales.length > 1 ? ' · variación vs. primero: ' +
            (((totales[totales.length - 1] / totales[0]) - 1) * 100).toFixed(1) + '%' : '')) : 'Sin comprobantes con fecha para graficar.';

        if (chart) chart.destroy();
        chart = new Chart(document.getElementById('grafico'), {
            type: 'line',
            data: { datasets: series },
            options: { maintainAspectRatio: false, parsing: true,
                       scales: { x: { type: 'category', labels: Array.from(new Set(series.flatMap(function (s) { return s.data.map(function (p) { return p.x; }); }))).sort() },
                                 y: { ticks: { callback: function (v) { return '$' + Number(v).toLocaleString('es-AR'); } } } },
                       plugins: { title: { display: true, text: 'Evolución de precios — ' + prov } } }
        });
    }
    function seleccionar(prov) {
        sel.value = prov;
        tabla.column(1).search(prov ? '^' + $.fn.dataTable.util.escapeRegex(prov) + '$' : '', true, false).draw();
        if (prov) { document.getElementById('grafico-wrap').hidden = false; dibujar(prov); }
        else { document.getElementById('grafico-wrap').hidden = true; document.getElementById('resumen-prov').textContent = ''; }
    }
    sel.addEventListener('change', function () { seleccionar(sel.value); });
    $('#tabla tbody').on('click', 'tr', function () { var d = tabla.row(this).data(); if (d) seleccionar(d.razon_social); });
    seleccionar('');
})();
</script>
</body>
</html>
"""


def generar_reporte_detallado(ruta_html=None, lector=None):
    """MEJ-02: genera Reporte_Detallado.html (coexiste con Reporte_Maestro.html)."""
    ruta_html = ruta_html or REPORTE_DETALLADO_HTML
    lector = lector or LectorDatosFactura()
    filas = _filas_detalladas(lector)
    datos_json = json.dumps(filas, ensure_ascii=False, default=str).replace("</", "<\\/")
    contenido = (_PLANTILLA_DETALLADO
                 .replace("__TIMESTAMP__", datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
                 .replace("__DATOS__", datos_json))
    try:
        tmp = ruta_html + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            f.write(contenido)
        os.replace(tmp, ruta_html)
        print(f"Reporte detallado actualizado en: {ruta_html} ({len(filas)} comprobantes)")
    except Exception as e:
        print(f"⚠️ No se pudo escribir {ruta_html}: {e}")


def generar_todos_los_reportes():
    generar_reporte_html()
    generar_reporte_detallado()


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if "--detallado" in sys.argv:
        generar_reporte_detallado()
    else:
        generar_todos_los_reportes()
