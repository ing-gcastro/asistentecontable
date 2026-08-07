import os
import sys
import pandas as pd
from config import REPORTE_MAESTRO_HTML

def main():
    if len(sys.argv) < 3:
        print("Uso: python eliminar_factura.py <Razon_Social> <Nro_Comprobante>")
        sys.exit(1)
        
    razon_social = sys.argv[1].strip()
    nro_comp = sys.argv[2].strip()
    
    print(f"Eliminando {razon_social} - {nro_comp} del historial...")
    
    if not os.path.exists(REPORTE_MAESTRO_HTML):
        print("El reporte maestro no existe.")
        sys.exit(1)
        
    try:
        # Leer el reporte actual
        df = pd.read_html(REPORTE_MAESTRO_HTML)[0]
        
        # Eliminar las filas que coincidan (conviertiendo a string para estar seguros)
        total_previo = len(df)
        df_nuevo = df[~((df['Razón Social'].astype(str) == razon_social) & 
                        (df['Nro Comprobante'].astype(str) == nro_comp))]
                        
        borrados = total_previo - len(df_nuevo)
        
        if borrados > 0:
            print(f"Se eliminaron {borrados} registros del historial.")
            
            # Ahora importamos reporter y usamos la funcion interna para regenerar todo
            import reporter
            
            # 1. Regenerar el Maestro
            reporter._escribir_html(df_nuevo, REPORTE_MAESTRO_HTML, "Reporte Histórico de Facturas Aprobadas")
            
            # 2. Regenerar los de sector
            base_sectores = r"\\10.10.10.210\AyF_Trabajoadistancia\AUTORIZACION DE FACTURAS"
            if os.path.exists(base_sectores):
                sectores_unicos = df_nuevo['Sector'].dropna().unique()
                for sec in sectores_unicos:
                    if not sec or sec == "GENERAL": continue
                    df_sec = df_nuevo[df_nuevo['Sector'] == sec]
                    if df_sec.empty: continue
                    dir_sector = os.path.join(base_sectores, sec)
                    os.makedirs(dir_sector, exist_ok=True)
                    ruta_sec = os.path.join(dir_sector, "Reporte.html")
                    reporter._escribir_html(df_sec, ruta_sec, f"Reporte de Facturas - {sec}")
            
            print("\n¡Borrado histórico exitoso! Los reportes se han actualizado.")
        else:
            print("No se encontró el registro en el historial.")
            
    except Exception as e:
        print(f"Error al eliminar: {e}")
        
    input("\nPresiona ENTER para salir...")

if __name__ == "__main__":
    main()
