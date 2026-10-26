# verificar_entorno.py — Chequeo rápido de instalación y rutas (lo usa VERIFICAR.bat).
# Sale con código 1 si falta algo imprescindible.
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


def main():
    ok = True
    print(f"Python: {sys.executable} ({sys.version.split()[0]})")
    for modulo, uso in (("pymupdf", "PDF"), ("pandas", "Excel/reportes"), ("openpyxl", "Excel .xlsx"),
                        ("xlrd", "Excel .xls"), ("lxml", "leer historial HTML")):
        try:
            __import__(modulo)
            print(f"  OK     {modulo:<10} ({uso})")
        except ImportError:
            print(f"  FALTA  {modulo:<10} ({uso})  -> ejecutar INSTALAR.bat")
            ok = False

    import qr_afip
    print("  OK     lector QR AFIP" if qr_afip.hay_decodificador_qr()
          else "  AVISO  sin lector QR (opcional) -> ejecutar INSTALAR.bat")
    try:
        __import__("win32com.client")
        print("  OK     pywin32 (Outlook/Excel)")
    except ImportError:
        print("  AVISO  sin pywin32: no funcionan la descarga de Outlook ni los mails")

    import config
    print("\nRutas configuradas (config.py / config_local.py):")
    for nombre, ruta in (("Carpeta A (entrada)", config.CARPETA_A), ("Compras", config.CARPETA_COMPRAS),
                         ("Autorización", config.CARPETA_AUTORIZACION), ("OC TELCOM", config.CARPETA_OC),
                         ("Proveedores", config.ARCHIVO_PROVEEDORES), ("Sectores", config.ARCHIVO_SECTORES)):
        print(f"  {'OK   ' if os.path.exists(ruta) else 'FALTA'}  {nombre:<20} {ruta}")
    print(f"\nNotificación a sectores: modo '{config.NOTIFICACION_MODO}'")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
