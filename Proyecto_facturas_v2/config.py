# config.py — Fuente única de verdad para rutas y constantes
# ⚠️  TODAS las rutas del proyecto se definen aquí y SOLO aquí.
#     Las rutas locales se calculan solas a partir del usuario de Windows y de
#     la ubicación del repo. Si en otra PC algo cambia (servidor, carpetas),
#     NO edites este archivo: copiá config_local.example.py como config_local.py
#     y sobrescribí ahí solo lo que haga falta (config_local.py no se sube a git).
import os

# ─── BASE DE RUTAS LOCALES ───────────────────────────────────────────────────
_USUARIO    = os.environ.get("USERNAME") or os.path.basename(os.path.expanduser("~"))
_HOME       = os.path.expanduser("~")
_BASE_REPO  = os.path.dirname(os.path.abspath(__file__))          # ...\Proyecto_facturas_v2
_RAIZ_REPO  = os.path.dirname(_BASE_REPO)                         # ...\asistentecontable
_BASE_LOCAL = os.path.join(_HOME, "ROBOTS")
_BASE_DATA  = os.path.join(_BASE_LOCAL, "data")

# ─── RUTAS DE RED (UNC) ──────────────────────────────────────────────────────
_SERVIDOR   = r"\\10.10.10.210"
_AYF        = rf"{_SERVIDOR}\AyF_Trabajoadistancia"

# ─── CARPETAS OPERATIVAS (LOCALES / DESKTOP) ──────────────────────────────────
_BASE_DESKTOP_PROCESO = os.path.join(_HOME, "Desktop", "PROCESO DE FC")
CARPETA_A           = os.path.join(_BASE_DESKTOP_PROCESO, "entrada de facturas")
CARPETA_B           = os.path.join(_BASE_DESKTOP_PROCESO, "no es factura")
CARPETA_REVISION    = os.path.join(_BASE_DESKTOP_PROCESO, "archivos en revision")

# ─── CARPETAS OPERATIVAS (RED UNC) ───────────────────────────────────────────
# MEJ-03: las facturas aprobadas van directo a la raíz de \Compras
# (antes: \Compras\fc a subir). El reporte escanea SOLO los PDF sueltos de esta
# carpeta (sin subcarpetas) — ver MEJ-04.
CARPETA_COMPRAS      = rf"{_AYF}\Compras"
CARPETA_C            = CARPETA_COMPRAS          # alias histórico, lo usan todos los módulos
CARPETA_AUTORIZACION = rf"{_AYF}\AUTORIZACION DE FACTURAS"
CARPETA_OC           = rf"{_SERVIDOR}\Compras\OC TELCOM"

# Carpetas donde el reporte busca "gemelos" para saber en qué etapa contable está
# cada factura: (ruta, recursiva, etapa, nivel). Mayor nivel = etapa más avanzada.
CARPETAS_ETAPAS = [
    (rf"{_AYF}\Facturas a pagar\2026",           True,  "PAGADA",        4),
    (rf"{_AYF}\Facturas a pagar\Periodo Actual", True,  "PARA PAGAR",    3),
    (rf"{_AYF}\Contabilidad",                    False, "CONTABILIZADA", 2),
    (CARPETA_COMPRAS,                            False, "ENTRADA",       1),
]

# ─── ARCHIVOS DE DATOS (EXCEL) ───────────────────────────────────────────────
ARCHIVO_PROVEEDORES = os.path.join(_BASE_DATA, "proveedores (2).xls")
ARCHIVO_SECTORES    = os.path.join(_BASE_DATA, "Sectores.xlsx")
OC_DB_PATH          = os.path.join(_BASE_DATA, "Comprobantes de Compras (Órdenes de Compras).xls")
CONSUMIDAS_PATH     = os.path.join(_BASE_DATA, "consumidas.xlsx")

# ─── ARCHIVOS DE ESTADO (JSON, locales, NO se versionan los nuevos) ──────────
ESTADO_EMAILS        = os.path.join(_BASE_REPO, "emails_procesados.json")
ESTADO_OC_PATH       = os.path.join(_BASE_REPO, "estado_oc.json")
# Remitente de cada PDF descargado de Outlook (lo escribe downloader, lo lee Paso 1)
REMITENTES_PATH      = os.path.join(_BASE_REPO, "remitentes_descargas.json")
# Datos de cada factura aprobada (fecha de emisión, importes, ítems). Lo escribe Paso 2.
REGISTRO_FACTURAS    = os.path.join(_BASE_REPO, "registro_facturas.json")
# Datos extraídos en Paso 1 (ítems, importes) para que Paso 2 los registre
LOTE_DATOS_PATH      = os.path.join(_BASE_REPO, "lote_datos.json")
# Caché de datos leídos de PDFs ya archivados (evita releerlos por la red)
CACHE_DATOS_PDF      = os.path.join(_BASE_REPO, "cache_datos_pdf.json")

# ─── REPORTES ────────────────────────────────────────────────────────────────
REPORTE_MAESTRO_HTML    = rf"{CARPETA_COMPRAS}\Reporte_Maestro.html"
REPORTE_MAESTRO_EXCEL   = rf"{CARPETA_COMPRAS}\Reporte_Maestro.xlsx"
REPORTE_DETALLADO_HTML  = rf"{CARPETA_COMPRAS}\Reporte_Detallado.html"   # MEJ-02
REPORTE_SECTORES_HTML   = rf"{CARPETA_AUTORIZACION}\Reportesector.html"
# BUG-04: para facturas viejas sin fecha de emisión registrada, el reporte lee el PDF
# (y guarda el resultado en caché). Tope de PDFs leídos por corrida, para que la
# primera ejecución por la red no tarde demasiado; el resto se completa en las siguientes.
REPORTE_MAX_LECTURAS_PDF = 300

# ─── MONITOR (MEJ-05) ────────────────────────────────────────────────────────
MONITOR_INTERVALO_SEG = 30     # cada cuánto revisa \Compras (PollingObserver, apto UNC)
MONITOR_ESPERA_SEG    = 20     # espera tras el último cambio antes de regenerar (agrupa ráfagas)

# ─── NOTIFICACIÓN A SECTORES (pendiente de SDD — ver docs/SDD.md §8) ─────────
# "desactivado": solo genera una vista previa HTML en logs\notificaciones (no toca Outlook)
# "borrador"   : crea el mail en Borradores de Outlook para revisarlo y enviarlo a mano
# "enviar"     : envía directamente (activar solo cuando el SDD esté aprobado)
NOTIFICACION_MODO = "desactivado"
# Destinatarios por sector (los define el SDD). Sector sin destinatarios = no se notifica.
NOTIFICACION_DESTINATARIOS = {
    "RRHH": [],
    "REDES Y SOPORTE": [],
    "OPERACIONES": [],
    "PROGRAMACION": [],
    "ADM": [],
    "PRESIDENCIA": [],
    "INTERNET": [],
    "GENERAL": [],
}
NOTIFICACION_CC = []
NOTIFICACION_ADJUNTAR_PDF = False      # True: adjunta cada PDF (ojo con el tamaño del mail)
NOTIFICACION_MAX_ADJUNTOS_MB = 15

# ─── LOGGING ─────────────────────────────────────────────────────────────────
LOGS_DIR = os.path.join(_RAIZ_REPO, "logs")

# ─── CONSTANTES DE NEGOCIO ───────────────────────────────────────────────────
MESES_ESPANOL = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
    9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE",
}

# ─── OVERRIDES LOCALES (opcional) ────────────────────────────────────────────
# config_local.py puede redefinir cualquier constante de arriba. Ej.:
#   CARPETA_A = r"D:\Facturas\entrada"
try:
    from config_local import *  # noqa: F401,F403
except ImportError:
    pass
