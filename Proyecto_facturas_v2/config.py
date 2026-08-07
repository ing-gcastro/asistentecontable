# config.py — Fuente única de verdad para rutas y constantes
# ⚠️  TODAS las rutas del proyecto se definen aquí y SOLO aquí.
#     Modificar este archivo al cambiar de equipo o usuario de Windows.
import os

# ─── BASE DE RUTAS LOCALES ───────────────────────────────────────────────────
_USUARIO    = r"gcastro"
_BASE_LOCAL = rf"C:\Users\{_USUARIO}\ROBOTS"
_BASE_REPO  = os.path.join(_BASE_LOCAL, "asistentecontable", "Proyecto_facturas_v2")
_BASE_DATA  = os.path.join(_BASE_LOCAL, "data")

# ─── RUTAS DE RED (UNC) ──────────────────────────────────────────────────────
_SERVIDOR   = r"\\10.10.10.210"

# ─── CARPETAS OPERATIVAS (RED) ───────────────────────────────────────────────
CARPETA_A           = rf"{_SERVIDOR}\Compras\entrada de facturas"
CARPETA_B           = rf"{_SERVIDOR}\Compras\no es factura"
CARPETA_C           = rf"{_SERVIDOR}\AyF_Trabajoadistancia\Compras\fc a subir"
CARPETA_OC          = rf"{_SERVIDOR}\Compras\OC TELCOM"

# ─── ARCHIVOS DE DATOS (EXCEL) ───────────────────────────────────────────────
ARCHIVO_PROVEEDORES = os.path.join(_BASE_DATA, "proveedores (2).xls")
ARCHIVO_SECTORES    = os.path.join(_BASE_DATA, "Sectores.xlsx")
OC_DB_PATH          = os.path.join(_BASE_DATA, "Comprobantes de Compras (Órdenes de Compras).xls")
CONSUMIDAS_PATH     = os.path.join(_BASE_DATA, "consumidas.xlsx")

# ─── ARCHIVOS DE ESTADO (JSON) ───────────────────────────────────────────────
ESTADO_EMAILS  = os.path.join(_BASE_REPO, "emails_procesados.json")
ESTADO_OC_PATH = os.path.join(_BASE_REPO, "estado_oc.json")

# ─── REPORTES ────────────────────────────────────────────────────────────────
REPORTE_MAESTRO_HTML = rf"{_SERVIDOR}\AyF_Trabajoadistancia\Compras\Reporte_Maestro.html"

# ─── LOGGING ─────────────────────────────────────────────────────────────────
LOGS_DIR = os.path.join(_BASE_REPO, "..", "logs")

# ─── CONSTANTES DE NEGOCIO ───────────────────────────────────────────────────
MESES_ESPANOL = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
    9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE",
}