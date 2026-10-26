# estado.py — Archivos de estado JSON y utilidades de archivos compartidas.
import json
import os
import time
from datetime import datetime

from config import REGISTRO_FACTURAS, REMITENTES_PATH


def cargar_json(ruta, defecto):
    if not os.path.exists(ruta):
        return defecto
    try:
        with open(ruta, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print(f"⚠️ No se pudo leer {os.path.basename(ruta)} ({e}). Se usa un estado vacío.")
        return defecto


def guardar_json(ruta, data):
    """Escritura atómica: si se corta a mitad de camino, el JSON anterior queda intacto."""
    os.makedirs(os.path.dirname(ruta) or ".", exist_ok=True)
    tmp = ruta + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, ruta)


# ─── REMITENTES (Outlook -> Paso 1) ──────────────────────────────────────────
def cargar_remitentes():
    """{nombre_pdf_en_carpeta_A: email_remitente}"""
    return cargar_json(REMITENTES_PATH, {})


def agregar_remitentes(mapa_nuevo):
    if not mapa_nuevo:
        return
    mapa = cargar_remitentes()
    mapa.update(mapa_nuevo)
    guardar_json(REMITENTES_PATH, mapa)


def podar_remitentes(archivos_presentes):
    """Elimina del mapa los PDFs que ya no están en la Carpeta A."""
    mapa = cargar_remitentes()
    presentes = set(archivos_presentes)
    podado = {k: v for k, v in mapa.items() if k in presentes}
    if len(podado) != len(mapa):
        guardar_json(REMITENTES_PATH, podado)
    return podado


# ─── REGISTRO DE FACTURAS APROBADAS (Paso 2 -> reportes) ─────────────────────
def cargar_registro():
    """{nombre_archivo_final.pdf: {fecha_emision, cuit, total, iva_21, ..., items}}"""
    return cargar_json(REGISTRO_FACTURAS, {})


def registrar_facturas(registros):
    """registros: lista de dicts con al menos la clave 'archivo'."""
    if not registros:
        return
    reg = cargar_registro()
    ahora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    for r in registros:
        nombre = r.get("archivo")
        if not nombre:
            continue
        entrada = dict(reg.get(nombre, {}))
        entrada.update({k: v for k, v in r.items() if v not in (None, "", "nan")})
        entrada.setdefault("fecha_proceso", ahora)
        reg[nombre] = entrada
    guardar_json(REGISTRO_FACTURAS, reg)


# ─── LOG DE AUDITORÍA ────────────────────────────────────────────────────────
class _Tee:
    """Duplica la salida de consola a un archivo (sin cambiar lo que ve el usuario)."""

    def __init__(self, consola, archivo):
        self._consola = consola
        self._archivo = archivo
        self.encoding = getattr(consola, "encoding", "utf-8")

    def write(self, texto):
        self._consola.write(texto)
        try:
            self._archivo.write(texto)
            self._archivo.flush()
        except Exception:
            pass
        return len(texto)

    def flush(self):
        self._consola.flush()

    def __getattr__(self, nombre):
        return getattr(self._consola, nombre)


def iniciar_log_auditoria(nombre_proceso):
    """
    Guarda todo lo que el proceso imprime en logs\\<proceso>_AAAA-MM.log (un archivo por mes),
    con usuario y fecha de cada corrida. Sirve para auditar qué se aprobó/rechazó y cuándo.
    """
    import sys
    from config import LOGS_DIR
    try:
        os.makedirs(LOGS_DIR, exist_ok=True)
        ruta = os.path.join(LOGS_DIR, f"{nombre_proceso}_{datetime.now():%Y-%m}.log")
        archivo = open(ruta, "a", encoding="utf-8")
        usuario = os.environ.get("USERNAME", "?")
        archivo.write(f"\n{'=' * 70}\n{datetime.now():%d/%m/%Y %H:%M:%S} | usuario: {usuario} | {nombre_proceso}\n{'=' * 70}\n")
        sys.stdout = _Tee(sys.stdout, archivo)
        return ruta
    except Exception as e:
        print(f"⚠️ No se pudo iniciar el log de auditoría: {e}")
        return None


# ─── ARCHIVOS BLOQUEADOS (Excel abierto) ─────────────────────────────────────
def archivo_bloqueado(ruta):
    if not os.path.exists(ruta):
        return False
    try:
        with open(ruta, "r+b"):
            return False
    except PermissionError:
        return True
    except OSError:
        return True


def esperar_archivo_libre(ruta, intentos_max=None):
    """
    Si el Excel está abierto, pide al usuario que lo cierre (en lugar de matar
    TODOS los Excel abiertos con taskkill, que hacía perder trabajo sin guardar).
    Devuelve True cuando el archivo queda libre, False si el usuario cancela.
    """
    intentos = 0
    while archivo_bloqueado(ruta):
        intentos += 1
        if intentos_max is not None and intentos > intentos_max:
            return False
        print(f"\n⚠️ El archivo está abierto en otro programa (¿Excel?):\n   {ruta}")
        resp = input("   Guardalo y cerralo, y presioná ENTER para reintentar (o escribí 'c' para cancelar): ").strip().lower()
        if resp == "c":
            return False
        time.sleep(0.5)
    return True
