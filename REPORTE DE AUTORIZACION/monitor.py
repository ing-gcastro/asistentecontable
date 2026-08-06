import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from indexar_pdfs import generar_reporte_html  # <--- Actualizado

CARPETA_A_MONITOREAR = r"\\10.10.10.210\AyF_Trabajoadistancia\AUTORIZACION DE FACTURAS"

class ManejadorCambios(FileSystemEventHandler):
    def on_any_event(self, event):
        # Evitamos reaccionar si cambia el propio archivo .html
        if event.src_path.endswith(".html"):
            return
            
        if event.src_path.endswith(".pdf"):
            print(f"\nDetectado cambio en: {event.src_path}. Actualizando reporte HTML...")
            generar_reporte_html()

if __name__ == "__main__":
    event_handler = ManejadorCambios()
    observer = Observer()
    observer.schedule(event_handler, path=CARPETA_A_MONITOREAR, recursive=True)
    
    print(f"Monitoreando cambios en '{CARPETA_A_MONITOREAR}'... (Presiona Ctrl+C para detener)")
    observer.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()