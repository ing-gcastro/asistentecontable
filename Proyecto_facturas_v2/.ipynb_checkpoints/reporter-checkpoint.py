import os
import re
import json
from pathlib import Path
from datetime import datetime
import pandas as pd
from config import CARPETA_C  # Ruta donde se guardará el reporte HTML de salida

HTML_SALIDA = os.path.join(CARPETA_C, "Reporte_FC_A_Subir.html")
CACHE_FILE = os.path.join(CARPETA_C, "cache_rutas_facturas.json")
RUTA_REPORTE_AUTORIZACION = r"\\10.10.10.210\AyF_Trabajoadistancia\AUTORIZACION DE FACTURAS\Reporte.html"

# Rutas donde se buscarán los archivos PDF
CARPETAS_BUSQUEDA = [
    r"\\10.10.10.210\AyF_Trabajoadistancia\Compras",
    r"\\10.10.10.210\AyF_Trabajoadistancia\Contabilidad",
    r"\\10.10.10.210\AyF_Trabajoadistancia\Facturas a pagar"
]

SECTORES_VALIDOS = [
    "RRHH", 
    "REDES Y SOPORTE", 
    "OPERACIONES", 
    "PROGRAMACION", 
    "ADM", 
    "PRESIDENCIA", 
    "INTERNET"
]

def obtener_facturas_procesadas(ruta_html):
    """Lee el reporte HTML existente para extraer las facturas ya procesadas y evitar duplicados en el core."""
    if not os.path.exists(ruta_html):
        return set()
    try:
        df_existente = pd.read_html(ruta_html)[0]
        procesadas = set()
        for _, row in df_existente.iterrows():
            razon_raw = str(row.get('Razón Social', row.get('Razon Social', ''))).strip().upper()
            razon = re.sub(r'^\([^)]+\)\s*', '', razon_raw).strip()
            nro = str(row.get('Nro Comprobante', '')).strip()
            if razon and nro and razon not in ["NAN", "DESCONOCIDO", "PROVEEDOR_DESCONOCIDO", ""] and nro != "NAN":
                procesadas.add((razon, nro))
        return procesadas
    except Exception:
        return set()

def cargar_cache():
    """Carga el caché local de rutas para evitar búsquedas repetitivas en la red."""
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def guardar_cache(cache):
    """Guarda el caché local actualizado."""
    try:
        with open(CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False, indent=4)
    except Exception as e:
        print(f"⚠️ Error al guardar el caché local: {e}")

def obtener_estados_autorizacion():
    """Lee el reporte de autorización externo para obtener los estados de las facturas."""
    if not os.path.exists(RUTA_REPORTE_AUTORIZACION):
        return {}, {}
    try:
        df_aut = pd.read_html(RUTA_REPORTE_AUTORIZACION)[0]
        estados_por_archivo = {}
        estados_por_datos = {}
        for _, row in df_aut.iterrows():
            archivo = str(row.get('Nombre completo del archivo', '')).strip()
            estado = str(row.get('Estado', 'PENDIENTE')).strip().upper()
            if archivo:
                estados_por_archivo[archivo] = estado
            
            rs = str(row.get('Razón Social', '')).strip().upper()
            tipo = str(row.get('Tipo', '')).strip().upper()
            nro = str(row.get('Número', '')).strip()
            if rs and nro:
                estados_por_datos[(rs, tipo, nro)] = estado
                
        return estados_por_archivo, estados_por_datos
    except Exception as e:
        print(f"⚠️ Nota al leer el reporte de autorización: {e}")
        return {}, {}

def parsear_nombre_aprobado(nombre_archivo):
    nombre_sin_ext = Path(nombre_archivo).stem.strip()
    
    razon_social = "DESCONOCIDO"
    tipo = "FC"
    numero = "SinNumero"
    oc = "N/A"
    sector = "GENERAL"
    empresa = "TELCOM VENTURES"
    info_adicional = ""

    if "(EUROSAT)" in nombre_sin_ext.upper() or "EUROSAT" in nombre_sin_ext.upper():
        empresa = "EUROSAT"

    for s in SECTORES_VALIDOS:
        if re.search(r'\b' + re.escape(s) + r'\b', nombre_sin_ext, re.IGNORECASE):
            sector = s
            break

    m_tipo = re.search(r'\b(FC|NC|ND)\b', nombre_sin_ext, re.IGNORECASE)
    if m_tipo:
        tipo = m_tipo.group(1).upper()
        partes = re.split(r'\b' + tipo + r'\b', nombre_sin_ext, flags=re.IGNORECASE, maxsplit=1)
        
        if len(partes) > 0:
            izquierda = partes[0].strip()
            rs_limpia = re.sub(r'^\([^)]+\)\s*', '', izquierda).strip()
            if rs_limpia:
                razon_social = rs_limpia

        if len(partes) > 1:
            derecha = partes[1].strip()
            tokens = derecha.split()
            if tokens:
                if tokens[0].upper() != 'OC':
                    numero = tokens[0]
                    resto_tokens = tokens[1:]
                else:
                    resto_tokens = tokens

                resto_texto = " ".join(resto_tokens)
                
                parentesis_extra = re.findall(r'\(([^)]+)\)', resto_texto)
                if parentesis_extra:
                    info_adicional = " ".join(parentesis_extra)
                    resto_texto = re.sub(r'\([^)]+\)', '', resto_texto)

                m_oc = re.search(r'\bOC\s*[:\.]?\s*(\d+)', resto_texto, re.IGNORECASE)
                if m_oc:
                    oc = m_oc.group(1)
                    partes_oc = re.split(r'\bOC\s*[:\.]?\s*\d+', resto_texto, flags=re.IGNORECASE, maxsplit=1)
                    if len(partes_oc) > 1:
                        sucia = partes_oc[1].strip()
                        if sector != "GENERAL":
                            sucia = re.sub(r'\b' + re.escape(sector) + r'\b', '', sucia, flags=re.IGNORECASE).strip()
                        if sucia:
                            info_adicional = (info_adicional + " " + sucia).strip()
                else:
                    sucia = resto_texto.strip()
                    if sector != "GENERAL":
                        sucia = re.sub(r'\b' + re.escape(sector) + r'\b', '', sucia, flags=re.IGNORECASE).strip()
                    if sucia:
                        info_adicional = (info_adicional + " " + sucia).strip()

    return {
        "Archivo Original": nombre_archivo,
        "Empresa": empresa,
        "Razón Social": razon_social,
        "Tipo": tipo,
        "Nro Comprobante": numero,
        "OC": oc,
        "Información Adicional": info_adicional,
        "Sector": sector
    }

def generar_reporte_html(datos_nuevos=None, ruta_html=HTML_SALIDA):
    estados_por_archivo, estados_por_datos = obtener_estados_autorizacion()
    cache_rutas = cargar_cache()
    
    datos_actuales = []
    archivos_procesados = set()
    
    ahora = datetime.now().timestamp()
    limite_antiguedad = 60 * 24 * 60 * 60  # 60 días sin cambios para considerar carpeta estable

    for carpeta in CARPETAS_BUSQUEDA:
        ruta_base = Path(carpeta)
        if not ruta_base.exists():
            continue
        
        for root, dirs, files in os.walk(ruta_base):
            root_path = Path(root)
            
            try:
                if root_path != ruta_base and (ahora - root_path.stat().st_mtime) > limite_antiguedad:
                    dirs.clear()
            except Exception:
                pass

            for file in files:
                if file.lower().endswith('.pdf'):
                    ruta_completa = str(root_path / file)
                    if ruta_completa in archivos_procesados:
                        continue
                    archivos_procesados.add(ruta_completa)
                    
                    cache_rutas[file] = ruta_completa

                    info = parsear_nombre_aprobado(file)
                    try:
                        fecha_mod = datetime.fromtimestamp(os.path.getmtime(ruta_completa)).strftime('%d/%m/%Y %H:%M:%S')
                    except Exception:
                        fecha_mod = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
                    
                    nombre_arch = info["Archivo Original"]
                    rs_key = str(info["Razón Social"]).strip().upper()
                    tipo_key = str(info["Tipo"]).strip().upper()
                    nro_key = str(info["Nro Comprobante"]).strip()
                    
                    estado = "PENDIENTE"
                    if nombre_arch in estados_por_archivo:
                        estado = estados_por_archivo[nombre_arch]
                    elif (rs_key, tipo_key, nro_key) in estados_por_datos:
                        estado = estados_por_datos[(rs_key, tipo_key, nro_key)]

                    datos_actuales.append({
                        "Fecha": fecha_mod,
                        "Empresa": info["Empresa"],
                        "Razón Social": info["Razón Social"],
                        "Tipo": info["Tipo"],
                        "Nro Comprobante": info["Nro Comprobante"],
                        "OC": info["OC"],
                        "Información Adicional": info["Información Adicional"],
                        "Sector": info["Sector"],
                        "Estado": estado,
                        "Archivo": info["Archivo Original"],
                        "Ruta Completa": ruta_completa
                    })

    guardar_cache(cache_rutas)

    df_nuevos = pd.DataFrame(datos_actuales)

    df_final = df_nuevos
    if os.path.exists(ruta_html):
        try:
            df_historico = pd.read_html(ruta_html)[0]
            for col in df_nuevos.columns:
                if col not in df_historico.columns:
                    df_historico[col] = ''
            df_final = pd.concat([df_historico, df_nuevos], ignore_index=True)
            df_final['Razón Social'] = df_final['Razón Social'].fillna('DESCONOCIDO')
            df_final['Nro Comprobante'] = df_final['Nro Comprobante'].fillna('SinNumero')
            df_final['Estado'] = df_final['Estado'].fillna('PENDIENTE')
            df_final['Ruta Completa'] = df_final['Ruta Completa'].fillna('')
            df_final.drop_duplicates(subset=["Razón Social", "Nro Comprobante"], keep="last", inplace=True)
        except Exception as e:
            print(f"⚠️ Nota al fusionar con el histórico anterior: {e}")

    for idx, row in df_final.iterrows():
        nombre_arch = str(row.get('Archivo', '')).strip()
        rs_key = str(row.get('Razón Social', '')).strip().upper()
        tipo_key = str(row.get('Tipo', '')).strip().upper()
        nro_key = str(row.get('Nro Comprobante', '')).strip()
        
        if nombre_arch in estados_por_archivo:
            df_final.at[idx, 'Estado'] = estados_por_archivo[nombre_arch]
        elif (rs_key, tipo_key, nro_key) in estados_por_datos:
            df_final.at[idx, 'Estado'] = estados_por_datos[(rs_key, tipo_key, nro_key)]

    if df_final.empty:
        print("⚠️ No hay datos para mostrar en el reporte.")
        return

    if "Fecha" in df_final.columns:
        df_final = df_final.sort_values(by="Fecha", ascending=False)

    timestamp_actual = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    
    html_content = f"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
    <meta http-equiv="Pragma" content="no-cache">
    <meta http-equiv="Expires" content="0">
    <title>Reporte Acumulado de Facturas Aprobadas</title>
    <!-- Bootstrap 5 CSS -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <!-- DataTables CSS -->
    <link href="https://cdn.datatables.net/1.13.6/css/dataTables.bootstrap5.min.css" rel="stylesheet">
    <style>
        body {{ font-family: Arial, sans-serif; background-color: #f4f7f6; margin: 0; padding: 10px; }}
        .container-fluid {{ max-width: 98%; margin: auto; background: white; padding: 25px; border-radius: 8px; box-shadow: 0px 0px 10px rgba(0,0,0,0.1); }}
        table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
        th, td {{ padding: 10px; text-align: left; border-bottom: 1px solid #ddd; font-size: 0.85rem; }}
        th {{ background-color: #0d6efd; color: white; }}
        tr:hover {{ background-color: #f1f1f1; }}
        .footer {{ margin-top: 20px; font-size: 12px; color: #777; text-align: right; }}
        thead input {{
            width: 100%;
            padding: 4px;
            box-sizing: border-box;
            font-size: 0.75rem;
            font-weight: normal;
            border: 1px solid #ced4da;
            border-radius: 4px;
        }}
        thead tr:nth-child(2) th {{
            background-color: #f8f9fa;
            padding: 5px;
        }}
    </style>
</head>
<body>
    <div class="container-fluid">
        <h2 class="text-primary fw-bold">📋 Reporte Histórico de Facturas Aprobadas</h2>
        <p class="text-muted">Última actualización: <strong>{timestamp_actual}</strong> | Total de registros acumulados: <strong>{len(df_final)}</strong></p>
        <hr>
        <div class="table-responsive">
            <table class="table table-striped table-bordered table-hover align-middle" id="tabla_reporte">
                <thead>
                    <tr style="text-align: right;">
                        <th>Fecha</th>
                        <th>Empresa</th>
                        <th>Razón Social</th>
                        <th>Tipo</th>
                        <th>Nro Comprobante</th>
                        <th>OC</th>
                        <th>Información Adicional</th>
                        <th>Sector</th>
                        <th>Estado</th>
                        <th>Archivo</th>
                        <th>Ruta Completa</th>
                    </tr>
                    <tr>
                        <th><input type="text" placeholder="Filtrar Fecha" /></th>
                        <th><input type="text" placeholder="Filtrar Empresa" /></th>
                        <th><input type="text" placeholder="Filtrar Razón Social" /></th>
                        <th><input type="text" placeholder="Filtrar Tipo" /></th>
                        <th><input type="text" placeholder="Filtrar Nro Comprobante" /></th>
                        <th><input type="text" placeholder="Filtrar OC" /></th>
                        <th><input type="text" placeholder="Filtrar Info Adicional" /></th>
                        <th><input type="text" placeholder="Filtrar Sector" /></th>
                        <th><input type="text" placeholder="Filtrar Estado" /></th>
                        <th><input type="text" placeholder="Filtrar Archivo" /></th>
                        <th><input type="text" placeholder="Filtrar Ruta Completa" /></th>
                    </tr>
                </thead>
                <tbody>
"""

    for _, row in df_final.iterrows():
        estado_val = str(row.get('Estado', 'PENDIENTE')).strip().upper()
        if "AUTORIZ" in estado_val:
            badge_estado = '<span class="badge bg-success">AUTORIZADA</span>'
        elif "REVIS" in estado_val:
            badge_estado = '<span class="badge bg-warning text-dark">A REVISIÓN</span>'
        else:
            badge_estado = '<span class="badge bg-secondary">PENDIENTE</span>'

        html_content += f"""
                    <tr>
                        <td>{row.get('Fecha', '')}</td>
                        <td><span class="badge bg-secondary">{row.get('Empresa', 'TELCOM VENTURES')}</span></td>
                        <td class="fw-bold">{row.get('Razón Social', '')}</td>
                        <td>{row.get('Tipo', '')}</td>
                        <td>{row.get('Nro Comprobante', '')}</td>
                        <td>{row.get('OC', 'N/A')}</td>
                        <td>{row.get('Información Adicional', '')}</td>
                        <td><span class="badge bg-info text-dark">{row.get('Sector', 'GENERAL')}</span></td>
                        <td>{badge_estado}</td>
                        <td>{row.get('Archivo', '')}</td>
                        <td style="font-family: monospace; font-size: 0.75rem; word-break: break-all;">{row.get('Ruta Completa', '')}</td>
                    </tr>
"""

    html_content += f"""
                </tbody>
            </table>
        </div>
        <div class="footer">Sistema de Gestión Documental - Histórico Acumulativo</div>
    </div>

    <!-- Scripts -->
    <script src="https://code.jquery.com/jquery-3.7.0.min.js"></script>
    <script src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
    <script src="https://cdn.datatables.net/1.13.6/js/dataTables.bootstrap5.min.js"></script>
    <script>
        $(document).ready(function() {{
            var table = $('#tabla_reporte').DataTable({{
                "language": {{
                    "url": "//cdn.datatables.net/plug-ins/1.13.6/i18n/es-ES.json"
                }},
                "pageLength": 25,
                "order": []
            }});

            $('#tabla_reporte thead tr:eq(1) input').on('keyup change clear', function () {{
                var index = $(this).parent().index();
                if (table.column(index).search() !== this.value) {{
                    table.column(index).search(this.value).draw();
                }}
            }});
        }});
    </script>
</body>
</html>
"""

    with open(ruta_html, "w", encoding="utf-8") as f:
        f.write(html_content)
        
    print(f"🌐 Reporte acumulativo actualizado con éxito en: {ruta_html} (Total registros: {len(df_final)})")