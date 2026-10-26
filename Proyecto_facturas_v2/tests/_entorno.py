# _entorno.py — Arma un entorno aislado (carpetas temporales) ANTES de importar el proyecto.
# Escribe un config_local.py en una carpeta temporal y la pone primera en sys.path,
# así config.py redirige TODAS las rutas (red incluida) a carpetas de prueba.
import os
import sys
import tempfile

PROYECTO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BASE = tempfile.mkdtemp(prefix="asistente_test_")

RUTAS = {
    "CARPETA_A": os.path.join(BASE, "entrada"),
    "CARPETA_B": os.path.join(BASE, "no es factura"),
    "CARPETA_REVISION": os.path.join(BASE, "revision"),
    "CARPETA_COMPRAS": os.path.join(BASE, "Compras"),
    "CARPETA_AUTORIZACION": os.path.join(BASE, "AUTORIZACION DE FACTURAS"),
    "CARPETA_OC": os.path.join(BASE, "OC TELCOM"),
    "ARCHIVO_PROVEEDORES": os.path.join(BASE, "data", "no_existe.xls"),
    "ARCHIVO_SECTORES": os.path.join(BASE, "data", "no_existe.xlsx"),
    "OC_DB_PATH": os.path.join(BASE, "data", "no_existe_oc.xls"),
    "CONSUMIDAS_PATH": os.path.join(BASE, "data", "no_existe_cons.xlsx"),
    "ESTADO_EMAILS": os.path.join(BASE, "estado", "emails_procesados.json"),
    "ESTADO_OC_PATH": os.path.join(BASE, "estado", "estado_oc.json"),
    "REMITENTES_PATH": os.path.join(BASE, "estado", "remitentes_descargas.json"),
    "REGISTRO_FACTURAS": os.path.join(BASE, "estado", "registro_facturas.json"),
    "LOTE_DATOS_PATH": os.path.join(BASE, "estado", "lote_datos.json"),
    "CACHE_DATOS_PDF": os.path.join(BASE, "estado", "cache_datos_pdf.json"),
    "LOGS_DIR": os.path.join(BASE, "logs"),
}
_DERIVADAS = {
    "CARPETA_C": "CARPETA_COMPRAS",
}

for clave, ruta in RUTAS.items():
    if clave.startswith("CARPETA") or clave == "LOGS_DIR":
        os.makedirs(ruta, exist_ok=True)
os.makedirs(os.path.join(BASE, "estado"), exist_ok=True)

_dir_cfg = os.path.join(BASE, "_cfg")
os.makedirs(_dir_cfg, exist_ok=True)
with open(os.path.join(_dir_cfg, "config_local.py"), "w", encoding="utf-8") as f:
    f.write("import os\n")
    for clave, ruta in RUTAS.items():
        f.write(f"{clave} = {ruta!r}\n")
    for clave, origen in _DERIVADAS.items():
        f.write(f"{clave} = {origen}\n")
    f.write("REPORTE_MAESTRO_HTML = os.path.join(CARPETA_COMPRAS, 'Reporte_Maestro.html')\n")
    f.write("REPORTE_MAESTRO_EXCEL = os.path.join(CARPETA_COMPRAS, 'Reporte_Maestro.xlsx')\n")
    f.write("REPORTE_DETALLADO_HTML = os.path.join(CARPETA_COMPRAS, 'Reporte_Detallado.html')\n")
    f.write("REPORTE_SECTORES_HTML = os.path.join(CARPETA_AUTORIZACION, 'Reportesector.html')\n")
    f.write("CARPETAS_ETAPAS = [(os.path.join(%r, 'Facturas a pagar', '2026'), True, 'PAGADA', 4), "
            "(CARPETA_COMPRAS, False, 'ENTRADA', 1)]\n" % BASE)
    f.write("NOTIFICACION_MODO = 'desactivado'\n")
    f.write("NOTIFICACION_DESTINATARIOS = {'ADM': ['adm@example.com']}\n")

for p in (PROYECTO, _dir_cfg):
    if p in sys.path:
        sys.path.remove(p)
sys.path.insert(0, PROYECTO)
sys.path.insert(0, _dir_cfg)

import warnings  # noqa: E402
warnings.filterwarnings("ignore", category=FutureWarning)
warnings.filterwarnings("ignore", category=DeprecationWarning)
