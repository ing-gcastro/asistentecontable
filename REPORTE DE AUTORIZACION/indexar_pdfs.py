import re
from pathlib import Path
import pandas as pd
from datetime import datetime

# 1. Define la ruta de tu carpeta principal y el archivo HTML de salida
CARPETA_PRINCIPAL = r"\\10.10.10.210\AyF_Trabajoadistancia\AUTORIZACION DE FACTURAS"
HTML_SALIDA = r"\\10.10.10.210\AyF_Trabajoadistancia\AUTORIZACION DE FACTURAS\Reporte.html"
HTML_SALIDA = r"\\10.10.10.210\AyF_Trabajoadistancia\AUTORIZACION DE FACTURAS\Reportesector.html"
TIPOS_VALIDOS = ["FC", "NC", "ND"]

def parsear_nombre_archivo(nombre_archivo):
    nombre_sin_ext = Path(nombre_archivo).stem.strip()
    
    tipo_encontrado = "NO IDENTIFICADO"
    numero = "NO IDENTIFICADO"
    razon_social = nombre_sin_ext

    try:
        # Identificar el Tipo (FC, NC, ND)
        for t in TIPOS_VALIDOS:
            if re.search(r'\b' + t + r'\b', nombre_sin_ext, re.IGNORECASE):
                tipo_encontrado = t
                break

        # Desglosar Razón Social y Número usando el Tipo como separador
        if tipo_encontrado != "NO IDENTIFICADO":
            partes = re.split(rf'\b{tipo_encontrado}\b', nombre_sin_ext, flags=re.IGNORECASE)
            if len(partes) > 1:
                parte_izq = partes[0].strip()
                razon_social_limpia = re.sub(r'^\([^)]+\)\s*', '', parte_izq).strip()
                if razon_social_limpia:
                    razon_social = razon_social_limpia

                parte_der = partes[1].strip()
                tokens = parte_der.split()
                if tokens:
                    numero = tokens[0]
    except Exception:
        razon_social = nombre_sin_ext

    return {
        "Nombre Completo": nombre_archivo,
        "Razón Social": razon_social,
        "Tipo": tipo_encontrado,
        "Número": numero
    }

def generar_reporte_html():
    ruta_base = Path(CARPETA_PRINCIPAL)
    
    if not ruta_base.exists():
        print(f"La ruta {CARPETA_PRINCIPAL} no existe.")
        return

    datos = []
    
    # Recorre recursivamente todas las subcarpetas buscando PDFs
    for archivo in ruta_base.rglob("*.pdf"):
        info = parsear_nombre_archivo(archivo.name)
        
        try:
            rel_path = archivo.relative_to(ruta_base)
            parts = rel_path.parts
            
            sector = parts[0] if len(parts) > 0 else "NO IDENTIFICADO"
            
            # Evaluar el estado según la carpeta en la que se encuentre
            estado = "PENDIENTE"
            for part in parts:
                p_upper = part.upper()
                if p_upper == "AUTORIZADA":
                    estado = "AUTORIZADA"
                    break
                elif p_upper in ["A REVISION", "A REVISIÓN"]:
                    estado = "A REVISION"
                    break
        except Exception:
            sector = "NO IDENTIFICADO"
            estado = "PENDIENTE"

        datos.append({
            "Nombre completo del archivo": info["Nombre Completo"],
            "Razón Social": info["Razón Social"],
            "Tipo": info["Tipo"],
            "Número": info["Número"],
            "Sector": sector,
            "Estado": estado,
            "Fecha del archivo": datetime.fromtimestamp(archivo.stat().st_mtime).strftime('%Y-%m-%d %H:%M:%S')
        })

    if not datos:
        df = pd.DataFrame(columns=["Nombre completo del archivo", "Razón Social", "Tipo", "Número", "Sector", "Estado", "Fecha del archivo"])
    else:
        df = pd.DataFrame(datos)

    # Convertir el DataFrame a una tabla HTML
    table_html = df.to_html(classes='table table-striped table-bordered table-hover align-middle', index=False, table_id='tabla_reporte')

    # Inyectar dinámicamente los inputs de búsqueda justo DEBAJO del encabezado (thead)
    cols = df.columns
    thead_filters = "<tr>" + "".join([f'<th><input type="text" placeholder="Filtrar {col}" /></th>' for col in cols]) + "</tr>"
    table_html = table_html.replace("</thead>", f"{thead_filters}\n</thead>")

    timestamp_actual = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    html_template = f"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <!-- Directivas anti-caché -->
    <meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
    <meta http-equiv="Pragma" content="no-cache">
    <meta http-equiv="Expires" content="0">
    <title>Reporte de Autorización de Facturas</title>
    <!-- Bootstrap 5 CSS -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <!-- DataTables CSS -->
    <link href="https://cdn.datatables.net/1.13.6/css/dataTables.bootstrap5.min.css" rel="stylesheet">
    <style>
        /* Estilo para los buscadores situados en la parte superior */
        thead input {{
            width: 100%;
            padding: 4px;
            box-sizing: border-box;
            font-size: 0.8rem;
            font-weight: normal;
            border: 1px solid #ced4da;
            border-radius: 4px;
        }}
        thead tr:nth-child(2) th {{
            background-color: #f8f9fa;
            padding: 6px;
        }}
    </style>
</head>
<body class="bg-light">
    <div class="container-fluid py-4">
        <div class="card shadow border-0 p-4">
            <div class="d-flex justify-content-between align-items-center mb-3 flex-wrap gap-2">
                <div>
                    <h2 class="text-primary fw-bold">📋 Reporte de Autorización de Facturas</h2>
                    <p class="text-muted mb-0">Última actualización del sistema: <strong>{timestamp_actual}</strong></p>
                </div>
                <div>
                    <span class="badge bg-success fs-6">Total registros: {len(datos)}</span>
                </div>
            </div>
            <hr>
            <div class="table-responsive">
                {table_html}
            </div>
        </div>
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

            // Aplicar el filtrado por columna leyendo los inputs de la segunda fila del thead
            $('#tabla_reporte thead tr:eq(1) input').on('keyup change clear', function () {{
                var index = $(this).parent().index();
                if (table.column(index).search() !== this.value) {{
                    table.column(index).search(this.value).draw();
                }}
            }});
        }});
    </script>
</body>
</html>"""

    try:
        with open(HTML_SALIDA, "w", encoding="utf-8") as f:
            f.write(html_template)
        print(f"¡Reporte HTML actualizado con éxito! Se procesaron {len(datos)} archivos.")
    except Exception as e:
        print(f"Error al guardar el reporte: {e}")

if __name__ == "__main__":
    generar_reporte_html()