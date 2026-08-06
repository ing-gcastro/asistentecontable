# proveedores.py
import pandas as pd
import os

def cargar_excel_proveedores(ruta_excel):
    """Carga el Excel usando índices exactos (Columna B y Columna F) y limpia formatos de texto."""
    if not os.path.exists(ruta_excel):
        print(f"⚠️ Advertencia: No se encontró el Excel en {ruta_excel}")
        return {}
        
    df = pd.read_excel(ruta_excel) 
    
    base = {}
    for _, fila in df.iterrows():
        try:
            # Columna B -> Índice 1 (Descripción / Razón Social)
            nombre = str(fila.iloc[1]).strip().upper()
            
            # Columna F -> Índice 5 (Número de Documento / CUIT)
            cuit_crudo = str(fila.iloc[5]).strip()
            
            # Limpieza profunda de texto para eliminar guiones, espacios, ".0" o marcas de texto de Excel
            cuit = "".join(filter(str.isdigit, cuit_crudo))
            
        except IndexError:
            continue 
            
        if cuit and cuit != '' and cuit != 'NAN':
            base[cuit] = {
                'nombre': nombre,
                'sector': str(fila['SECTOR']).strip() if 'SECTOR' in df.columns and str(fila['SECTOR']).strip() != 'nan' else "",
                'es_lista_negra': "LISTA NEGRA" in nombre
            }
    return base

def buscar_proveedor(cuit_pdf, razon_social_pdf, base_proveedores):
    """
    1. Búsqueda exacta por CUIT limpio.
    2. Plan B por Razón Social.
    3. Nombre de emergencia si todo falla.
    """
    cuit_limpio = "".join(filter(str.isdigit, str(cuit_pdf)))
    
    # 1. Búsqueda por CUIT
    if cuit_limpio and cuit_limpio in base_proveedores:
        datos = base_proveedores[cuit_limpio]
        return datos['nombre'], datos['sector'], datos['es_lista_negra']
        
    # 2. Búsqueda por Razón Social
    razon_limpia = str(razon_social_pdf).strip().upper()
    if razon_limpia:
        for cuit_key, datos in base_proveedores.items():
            nombre_oficial = datos['nombre']
            if nombre_oficial in razon_limpia or razon_limpia in nombre_oficial:
                return nombre_oficial, datos['sector'], datos['es_lista_negra']
                
    # 3. Emergencia
    nombre_emergencia = f"PROVEEDOR_CUIT_{cuit_limpio}" if cuit_limpio else "PROVEEDOR_DESCONOCIDO"
    return nombre_emergencia, "", False