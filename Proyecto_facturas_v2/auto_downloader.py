import os
import sys

# Agregar la ruta del proyecto para importar módulos si se ejecuta desde otra ubicación (ej. Tareas Programadas)
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from downloader import descargar_facturas_outlook

if __name__ == "__main__":
    print("Iniciando descarga en segundo plano...")
    # Descarga inicial desde el 04/08/2026. 
    # Gracias al nuevo sistema de tracking, ignorará automáticamente
    # cualquier correo que ya haya descargado hoy o en el futuro.
    descargar_facturas_outlook(fecha_inicio="04/08/2026")
    print("Finalizado.")
