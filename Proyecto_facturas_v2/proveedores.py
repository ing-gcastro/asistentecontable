# proveedores.py
import pandas as pd
import os
from config import ARCHIVO_SECTORES

def cargar_excel_proveedores(ruta_excel):
    """Carga el Excel principal de proveedores (CUITs)."""
    if not os.path.exists(ruta_excel):
        print(f"⚠️ Advertencia: No se encontró el Excel en {ruta_excel}")
        return {}
        
    df = pd.read_excel(ruta_excel) 
    base = {}
    for _, fila in df.iterrows():
        try:
            nombre = str(fila.iloc[1]).strip().upper()
            cuit_crudo = str(fila.iloc[5]).strip()
            cuit = "".join(filter(str.isdigit, cuit_crudo))
        except IndexError:
            continue 
            
        if cuit and cuit != '' and cuit != 'NAN':
            base[cuit] = {
                'nombre': nombre,
                'es_lista_negra': "LISTA NEGRA" in nombre
            }
    return base

def cargar_sectores():
    """Carga el archivo Sectores.xlsx y devuelve un diccionario {cuit: sector}."""
    if not os.path.exists(ARCHIVO_SECTORES):
        return {}
        
    df = pd.read_excel(ARCHIVO_SECTORES)
    df.columns = df.columns.str.strip()
    
    mapa_sectores = {}
    for _, fila in df.iterrows():
        try:
            # Buscamos la columna CUIT / Numero_Documento y la columna Sector
            cuit_val = str(fila.get('Numero_Documento', '')).strip()
            cuit = "".join(filter(str.isdigit, cuit_val))
            
            sector_val = str(fila.get('Sector', '')).strip()
            sector = sector_val.upper() if sector_val and sector_val != 'nan' else ""
            
            if cuit:
                mapa_sectores[cuit] = sector
        except Exception:
            continue
    return mapa_sectores

def buscar_proveedor(cuit_pdf, razon_social_pdf, base_proveedores):
    cuit_limpio = "".join(filter(str.isdigit, str(cuit_pdf)))
    
    # 1. Búsqueda por CUIT en proveedores
    if cuit_limpio and cuit_limpio in base_proveedores:
        datos = base_proveedores[cuit_limpio]
        # Buscamos el sector en el archivo independiente de sectores
        sectores_dict = cargar_sectores()
        sector = sectores_dict.get(cuit_limpio, "")
        return datos['nombre'], sector, datos['es_lista_negra']
        
    # 2. Búsqueda por Razón Social
    razon_limpia = str(razon_social_pdf).strip().upper()
    if razon_limpia:
        for cuit_key, datos in base_proveedores.items():
            nombre_oficial = datos['nombre']
            if nombre_oficial in razon_limpia or razon_limpia in nombre_oficial:
                sectores_dict = cargar_sectores()
                sector = sectores_dict.get(cuit_key, "")
                return nombre_oficial, sector, datos['es_lista_negra']
                
    # 3. Emergencia
    nombre_emergencia = f"PROVEEDOR_CUIT_{cuit_limpio}" if cuit_limpio else "PROVEEDOR_DESCONOCIDO"
    return nombre_emergencia, "", False