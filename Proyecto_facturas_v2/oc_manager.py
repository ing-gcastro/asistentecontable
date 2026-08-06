import os
import json
import pandas as pd
from config import CARPETA_B

# Base de datos de OC
OC_DB_PATH = r"C:\Users\gcastro\ROBOTS\data\Comprobantes de Compras (Órdenes de Compras).xls"
# Archivo de seguimiento
ESTADO_OC_PATH = r"C:\Users\gcastro\ROBOTS\asistentecontable\Proyecto_facturas_v2\estado_oc.json"

def cargar_estado_oc():
    if os.path.exists(ESTADO_OC_PATH):
        try:
            with open(ESTADO_OC_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print(f"⚠️ Error cargando estado_oc.json: {e}")
            return {}
    return {}

def guardar_estado_oc(estado):
    with open(ESTADO_OC_PATH, "w", encoding="utf-8") as f:
        json.dump(estado, f, indent=4, ensure_ascii=False)

def leer_oc_disponibles(proveedor_cuit, proveedor_nombre):
    """
    Busca todas las OCs disponibles para un proveedor.
    Filtra las OCs de tipo "Bien" (B) que ya hayan sido consumidas.
    """
    try:
        df = pd.read_excel(OC_DB_PATH)
    except Exception as e:
        print(f"⚠️ No se pudo leer el archivo de OCs: {e}")
        return []

    # Filtrar por proveedor (usamos una búsqueda flexible por si hay diferencias menores en el nombre)
    termino_busqueda = proveedor_nombre.split()[0] if proveedor_nombre else "NO_MATCH_XYZ"
    cuit_busq = str(proveedor_cuit) if proveedor_cuit else "NO_MATCH_XYZ"
    
    print(f"Buscando OC en BD para: '{termino_busqueda}' o CUIT '{cuit_busq}'")
    
    df_prov = df[
        df["Proveedor"].str.contains(termino_busqueda, case=False, na=False) |
        df["Proveedor"].str.contains(cuit_busq, case=False, na=False)
    ]
    
    # --- FILTRAR CONSUMIDAS MANUALMENTE ---
    try:
        df_cons = pd.read_excel(r"C:\Users\gcastro\ROBOTS\data\consumidas.xlsx")
        # Obtener lista de OCs consumidas (ignorando nulos y limpiando strings)
        ocs_consumidas_manual = [str(x).strip() for x in df_cons[df_cons['consumida'].notna()]['NoComprobante'].tolist()]
    except Exception as e:
        print(f"⚠️ No se pudo leer consumidas.xlsx: {e}")
        ocs_consumidas_manual = []
        
    # Filtrar el dataframe original
    df_prov = df_prov[~df_prov["NoComprobante"].astype(str).str.strip().isin(ocs_consumidas_manual)]
    
    print(f"OCs encontradas en BD para este proveedor: {len(df_prov)}")

    if df_prov.empty:
        return []

    estado_oc = cargar_estado_oc()
    ocs_disponibles = []

    for _, row in df_prov.iterrows():
        nro_oc = str(row.get("NoComprobante", "")).strip()
        if not nro_oc:
            continue
            
        # Al no existir la columna de Tipo de Consumo en la DB real, asumimos Servicio (S) temporalmente 
        # para que no desaparezca en el primer uso hasta que definamos cómo distinguirlo
        tipo = "S" 
        
        consumos = estado_oc.get(nro_oc, [])
        
        if tipo == "B" and len(consumos) > 0:
            continue
            
        ocs_disponibles.append({
            "nro_oc": nro_oc,
            "fecha": str(row.get("FechaComprobante", "")).split()[0], # Para quedarnos solo con la fecha YYYY-MM-DD
            "articulo": str(row.get("Articulo", "")),
            "cantidad": str(row.get("Cantidad", "")),
            "descripcion": str(row.get("Descripcion_Adicional", "")),
            "tipo": tipo
        })
        
    return ocs_disponibles

def registrar_consumo_oc(nro_oc, nombre_factura):
    estado_oc = cargar_estado_oc()
    
    if nro_oc not in estado_oc:
        estado_oc[nro_oc] = []
        
    if nombre_factura not in estado_oc[nro_oc]:
        estado_oc[nro_oc].append(nombre_factura)
        guardar_estado_oc(estado_oc)
        return True
    return False

def buscar_pdf_oc(nro_oc):
    carpeta_oc = r"\\10.10.10.210\Compras\OC TELCOM"
    if not os.path.exists(carpeta_oc):
        return None
        
    archivos = os.listdir(carpeta_oc)
    for f in archivos:
        if f.lower().endswith('.pdf'):
            # Buscar coincidencia exacta (12345.pdf) o contenida (12345 proveedor.pdf)
            if str(nro_oc) in f:
                return os.path.join(carpeta_oc, f)
    return None
