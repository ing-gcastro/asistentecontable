# monitor_compras.py — Regenera los reportes cuando cambian los PDF de \Compras
# o de AUTORIZACION DE FACTURAS (MEJ-05).
#
# Por qué sondeo propio y no watchdog.Observer: en rutas de red UNC/SMB los eventos
# nativos de Windows no son confiables; la alternativa recomendada (PollingObserver)
# hace exactamente esto: comparar "fotos" del directorio cada N segundos.
# Hacerlo acá evita una dependencia extra y permite agrupar ráfagas de cambios.
#
# Uso: python monitor_compras.py            (corre hasta Ctrl+C)
#      python monitor_compras.py --una-vez  (una sola regeneración y sale)
import os
import sys
import time
from datetime import datetime

from config import (CARPETA_AUTORIZACION, CARPETA_COMPRAS, MONITOR_ESPERA_SEG,
                    MONITOR_INTERVALO_SEG)


def _foto_carpeta(ruta, recursiva):
    foto = {}
    if not os.path.isdir(ruta):
        return foto
    try:
        if recursiva:
            for root, _dirs, files in os.walk(ruta):
                for f in files:
                    if f.lower().endswith(".pdf"):
                        p = os.path.join(root, f)
                        try:
                            st = os.stat(p)
                            foto[p] = (st.st_size, int(st.st_mtime))
                        except OSError:
                            continue
        else:
            with os.scandir(ruta) as it:
                for e in it:
                    if e.is_file() and e.name.lower().endswith(".pdf"):
                        st = e.stat()
                        foto[e.path] = (st.st_size, int(st.st_mtime))
    except OSError as e:
        print(f"⚠️ No se pudo leer {ruta}: {e}")
    return foto


def foto_actual():
    """MEJ-04: \\Compras solo nivel raíz; AUTORIZACION recursivo (para ver lo autorizado)."""
    foto = _foto_carpeta(CARPETA_COMPRAS, recursiva=False)
    foto.update(_foto_carpeta(CARPETA_AUTORIZACION, recursiva=True))
    return foto


def regenerar():
    import reporter
    inicio = time.time()
    print(f"\n[{datetime.now():%d/%m/%Y %H:%M:%S}] 🔄 Regenerando reportes...")
    try:
        reporter.generar_todos_los_reportes()
        print(f"✅ Reportes al día ({time.time() - inicio:.0f} s).")
    except Exception as e:
        print(f"⚠️ Error regenerando reportes: {e}")


def vigilar(intervalo=MONITOR_INTERVALO_SEG, espera=MONITOR_ESPERA_SEG):
    print(f"👀 Vigilando:\n   {CARPETA_COMPRAS} (solo PDF sueltos)\n   {CARPETA_AUTORIZACION} (con subcarpetas)")
    print(f"   Revisión cada {intervalo} s. Cerrá esta ventana o Ctrl+C para detener.")
    anterior = foto_actual()
    regenerar()
    ultimo_cambio = None
    while True:
        time.sleep(intervalo)
        actual = foto_actual()
        if actual != anterior:
            nuevos = len(set(actual) - set(anterior))
            quitados = len(set(anterior) - set(actual))
            print(f"[{datetime.now():%H:%M:%S}] Cambios detectados (+{nuevos} / -{quitados}). Esperando que se estabilicen...")
            anterior = actual
            ultimo_cambio = time.time()
            continue
        if ultimo_cambio and time.time() - ultimo_cambio >= espera:
            regenerar()
            anterior = foto_actual()
            ultimo_cambio = None


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if "--una-vez" in sys.argv:
        regenerar()
    else:
        try:
            vigilar()
        except KeyboardInterrupt:
            print("\nMonitor detenido.")
