# -*- coding: utf-8 -*-
import os
import re
import sys
import json
import time
try:
    import pymupdf as fitz  # PyMuPDF >= 1.24
except ImportError:  # versiones viejas
    import fitz
from concurrent.futures import ThreadPoolExecutor, as_completed

# Add path for Proyecto_facturas_v2 imports
_AQUI = os.path.dirname(os.path.abspath(__file__))
REPO_PATH = os.path.join(os.path.dirname(_AQUI), "Proyecto_facturas_v2")  # núcleo del repo (config, proveedores)
if REPO_PATH not in sys.path:
    sys.path.append(REPO_PATH)

import proveedores
import config

NETWORK_FOLDER = config.CARPETA_OC
OUTPUT_HTML = os.path.join(_AQUI, "reporte_oc_telcom.html")
OUTPUT_JSON = os.path.join(_AQUI, "datos_oc_telcom.json")
MIN_OC_NUMBER = 20174

# Load supplier database
PROV_BASE = {}
if os.path.exists(config.ARCHIVO_PROVEEDORES):
    try:
        PROV_BASE = proveedores.cargar_excel_proveedores(config.ARCHIVO_PROVEEDORES)
        print(f"[*] Base de Proveedores cargada: {len(PROV_BASE)} registros.", flush=True)
    except Exception as e:
        print(f"[!] Error al cargar base de proveedores: {e}", flush=True)

def clean_quantity(raw_cant):
    if not raw_cant:
        return '1'
    raw_cant = raw_cant.strip()
    
    # 1. Exact integer format like "1,00" or "240.00" -> "1" or "240"
    m = re.match(r'^(\d+)(?:[,.]0+)?$', raw_cant)
    if m:
        return m.group(1)
        
    # 2. Thousands separator format like "24.000" -> "24000" or "6.000,00" -> "6000"
    m2 = re.match(r'^(\d{1,3}(?:\.\d{3})+)(?:[,.]0+)?$', raw_cant)
    if m2:
        return m2.group(1).replace('.', '')
        
    # 3. Floating point format like "5.5" or "10.50"
    m3 = re.match(r'^(\d+[,.]\d+)', raw_cant)
    if m3:
        val = m3.group(1).replace(',', '.')
        try:
            fval = float(val)
            if fval.is_integer():
                return str(int(fval))
            return f"{fval:.2f}"
        except ValueError:
            pass

    # Fallback: extract first integer
    digits = re.findall(r'\d+', raw_cant)
    if digits:
        return digits[0]
        
    return raw_cant

def clean_article_name(raw_art):
    if not raw_art:
        return 'Sin especificar'
    # Remove leading article codes (e.g., "601050 - ARTICULO" or "5018 ARTICULO" -> "ARTICULO")
    clean = re.sub(r'^\d+\s*[-–:]?\s*', '', raw_art.strip())
    return clean.strip() if clean.strip() else raw_art.strip()

def parse_pdf_file(file_path):
    filename = os.path.basename(file_path)
    filename_lower = filename.lower()
    
    # Filter out test files (e.g. "1 prueba.pdf", "1 OTRA PRUEBA.pdf", "1.pdf", "2.pdf", "test.pdf")
    if 'prueba' in filename_lower or 'test' in filename_lower:
        return None
    if re.match(r'^\d{1,3}\.pdf$', filename_lower):
        return None
        
    # Pre-filter by filename digits if possible
    digits_in_name = re.findall(r'\d+', filename)
    if digits_in_name:
        max_digit = max(int(d) for d in digits_in_name)
        if max_digit < MIN_OC_NUMBER and max_digit > 100:
            return None

    try:
        doc = fitz.open(file_path)
        if len(doc) == 0:
            return None
        
        p1_text = doc[0].get_text('text')
        p1_text_lower = p1_text.lower()
        
        if not ('orden de compra' in p1_text_lower or 'ordenes de compra' in p1_text_lower or 'compras@antina' in p1_text_lower):
            return None
        
        # 1. NRO DE OC (below or next to ORDEN DE COMPRA)
        nro_oc_int = 0
        m_oc = re.search(r'(\d{5,6})\s*\n\s*Orden\s+de\s+Compra', p1_text, re.IGNORECASE)
        if not m_oc:
            m_oc = re.search(r'Orden\s+de\s+Compra\s*\n\s*(\d{5,6})', p1_text, re.IGNORECASE)
        if not m_oc:
            m_oc = re.search(r'(?:Ordenes\s+de\s+Compra\s*N[°\?]?|Orden\s+de\s+Compra)\s*\n?\s*(\d{5,6})', p1_text, re.IGNORECASE)
        if not m_oc:
            m_oc = re.search(r'Numero\s*\n\s*(\d{5,6})', p1_text, re.IGNORECASE)

        if m_oc:
            nro_oc_int = int(m_oc.group(1))
        elif digits_in_name:
            nro_oc_int = max(int(d) for d in digits_in_name if len(d) >= 5)

        # STRICT FILTER: OC >= 20174
        if nro_oc_int < MIN_OC_NUMBER:
            return None

        nro_oc = str(nro_oc_int)

        # 2. FECHA
        fecha = ''
        m = re.search(r'(\d{2}/\d{2}/\d{4})\s*\n\s*Fecha', p1_text, re.IGNORECASE)
        if not m:
            m = re.search(r'Fecha\s*de\s*Emisi[óo]n\s*\n?\s*(\d{2}/\d{2}/\d{4})', p1_text, re.IGNORECASE)
        if not m:
            m = re.search(r'FECHA\s*\n?\s*(\d{1,2}/\d{1,2}/\d{4})', p1_text, re.IGNORECASE)
        if m:
            fecha = m.group(1).strip()

        # 3. PROVEEDOR (Using CUIT Lookup in PROV_BASE)
        cuit_found = ''
        raw_cuits = re.findall(r'\b\d{2}-?\d{8}-?\d\b|\b\d{11}\b', p1_text)
        for c in raw_cuits:
            c_clean = ''.join(filter(str.isdigit, c))
            if c_clean != '30696428898' and len(c_clean) == 11:  # Exclude Antina CUIT
                cuit_found = c_clean
                break
                
        proveedor = ''
        if cuit_found and cuit_found in PROV_BASE:
            proveedor = PROV_BASE[cuit_found]['nombre']
        else:
            # Fallback to Proveedor text block
            m = re.search(r'Proveedor\s*\n\s*([^\n]+(?:\n[^\n]+)?)', p1_text, re.IGNORECASE)
            if m:
                raw_prov = m.group(1)
                lines = [l.strip() for l in raw_prov.splitlines() if l.strip()]
                filtered = []
                for l in lines:
                    if re.match(r'^(OTROS|C\.U\.I\.T\.|Teléfono|E-Mail|Condición|Comprador)', l, re.IGNORECASE):
                        break
                    filtered.append(l)
                proveedor = ' '.join(filtered)
                # Strip numeric code from supplier name if present
                proveedor = re.sub(r'^\(\d+\)\s*', '', proveedor).strip()

        if not proveedor:
            proveedor = "PROVEEDOR DESCONOCIDO"

        # 4. INFO ADICIONAL (between Fecha de Entrega and Facturas solo por...)
        info_adicional = ''
        m = re.search(
            r'(?:Fecha\s+de\s+Entrega|F\.\s+de\s+Entrega:)\s*\n?\s*(?:\d{2}/\d{2}/\d{4})?\s*\n+(.*?)\n+\s*(?:Facturas\s+solo\s+por|Recepcion\s+de\s+facturas|Co\.\s*Art[íi]culo|Articulo)',
            p1_text,
            re.DOTALL | re.IGNORECASE
        )
        if m:
            raw_info = m.group(1).strip()
            lines = [l.strip() for l in raw_info.splitlines() if l.strip()]
            clean_lines = [l for l in lines if not re.search(r'C\.U\.I\.T\.|Co\.|Lugar\s+de\s+Entrega|Depósito|Transporte', l, re.IGNORECASE)]
            info_adicional = ' '.join(clean_lines)
        else:
            m = re.search(r'(Solicitado\s+por[^\n]+)', p1_text, re.IGNORECASE)
            if m:
                info_adicional = m.group(1).strip()

        # 5. ITEMS EXTRACTION (Cantidad, Articulo sin codigo, Descripcion de abajo)
        items = []
        blocks = doc[0].get_text('blocks')
        blocks_text = [b[4].strip() for b in blocks if b[4].strip()]
        
        in_table = False
        current_item = None
        
        for text in blocks_text:
            text_strip = text.strip()
            
            # Start of table
            if 'Co. Art' in text_strip or 'Artculo' in text_strip or 'Articulo' in text_strip or 'Total $' in text_strip:
                if not in_table and ('Co. Art' in text_strip or 'Artculo' in text_strip or 'Articulo' in text_strip):
                    in_table = True
                    continue
            
            if not in_table:
                continue
                
            # End of table
            if 'Este comprobante' in text_strip or 'Subtotal' in text_strip or 'Ordenes de Compra' in text_strip or 'Total Cantidad' in text_strip or 'N° Copia' in text_strip:
                if current_item:
                    items.append(current_item)
                    current_item = None
                break
                
            lines = [l.strip() for l in text_strip.splitlines() if l.strip()]
            if not lines:
                continue
                
            # Detect item row by unit
            unit_indices = [idx for idx, l in enumerate(lines) if l.upper() in ['UNIDAD', 'UNI', 'PCE', 'HS', 'KG', 'MTS', 'METROS', 'JGO', 'UNIDADES']]
            
            if unit_indices:
                if current_item:
                    items.append(current_item)
                    
                u_idx = unit_indices[0]
                raw_cant = lines[u_idx - 1] if u_idx > 0 else '1'
                cant_clean = clean_quantity(raw_cant)
                
                if u_idx > 1:
                    raw_art_name = ' '.join(lines[1:u_idx-1])
                else:
                    raw_art_name = lines[1] if len(lines) > 1 else lines[0]
                    
                art_clean = clean_article_name(raw_art_name)
                    
                current_item = {
                    'cantidad': cant_clean,
                    'articulo': art_clean,
                    'descripcion_debajo': ''
                }
            else:
                if current_item:
                    if current_item['descripcion_debajo']:
                        current_item['descripcion_debajo'] += '\n' + text_strip
                    else:
                        current_item['descripcion_debajo'] = text_strip
                        
        if current_item:
            items.append(current_item)

        if not items:
            items.append({
                'cantidad': '1',
                'articulo': 'Sin especificar',
                'descripcion_debajo': ''
            })

        records = []
        for item in items:
            records.append({
                'nro_oc': nro_oc,
                'fecha': fecha,
                'proveedor': proveedor,
                'cuit': cuit_found,
                'info_adicional': info_adicional,
                'cantidad': item['cantidad'],
                'articulo': item['articulo'],
                'descripcion_debajo': item['descripcion_debajo'],
                'archivo': filename
            })
            
        return records

    except Exception as e:
        return None

def generate_html_report(data, output_path):
    json_data = json.dumps(data, ensure_ascii=False)
    
    html_content = f"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Reporte de Órdenes de Compra - Antina TELCOM (OC >= 20174)</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <style>
        :root {{
            --bg-primary: #0f172a;
            --bg-card: #1e293b;
            --bg-card-hover: #334155;
            --accent-primary: #0ea5e9;
            --accent-secondary: #6366f1;
            --accent-green: #10b981;
            --text-primary: #f8fafc;
            --text-secondary: #94a3b8;
            --border-color: #334155;
            --table-header-bg: #0f172a;
            --badge-bg: rgba(14, 165, 233, 0.15);
            --badge-text: #38bdf8;
        }}

        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
            font-family: 'Inter', sans-serif;
        }}

        body {{
            background-color: var(--bg-primary);
            color: var(--text-primary);
            min-height: 100vh;
            padding: 2rem;
        }}

        .container {{
            max-width: 1650px;
            margin: 0 auto;
        }}

        header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 2rem;
            padding-bottom: 1.5rem;
            border-bottom: 1px solid var(--border-color);
        }}

        .header-title h1 {{
            font-size: 1.8rem;
            font-weight: 700;
            background: linear-gradient(135deg, #38bdf8 0%, #818cf8 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            display: flex;
            align-items: center;
            gap: 0.75rem;
        }}

        .header-title p {{
            color: var(--text-secondary);
            font-size: 0.9rem;
            margin-top: 0.3rem;
        }}

        .stats-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
            gap: 1.25rem;
            margin-bottom: 2rem;
        }}

        .stat-card {{
            background: var(--bg-card);
            padding: 1.25rem;
            border-radius: 12px;
            border: 1px solid var(--border-color);
            display: flex;
            align-items: center;
            gap: 1rem;
            transition: transform 0.2s ease, border-color 0.2s ease;
        }}

        .stat-card:hover {{
            transform: translateY(-2px);
            border-color: var(--accent-primary);
        }}

        .stat-icon {{
            width: 48px;
            height: 48px;
            border-radius: 10px;
            background: rgba(14, 165, 233, 0.1);
            color: var(--accent-primary);
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 1.4rem;
        }}

        .stat-info h3 {{
            font-size: 1.5rem;
            font-weight: 700;
            color: var(--text-primary);
        }}

        .stat-info p {{
            font-size: 0.85rem;
            color: var(--text-secondary);
        }}

        /* Controls Section */
        .controls-card {{
            background: var(--bg-card);
            border-radius: 12px;
            padding: 1.25rem;
            border: 1px solid var(--border-color);
            margin-bottom: 1.5rem;
            display: flex;
            flex-wrap: wrap;
            gap: 1rem;
            align-items: center;
            justify-content: space-between;
        }}

        .search-box {{
            position: relative;
            flex: 1;
            min-width: 300px;
        }}

        .search-box i {{
            position: absolute;
            left: 1rem;
            top: 50%;
            transform: translateY(-50%);
            color: var(--text-secondary);
        }}

        .search-box input {{
            width: 100%;
            padding: 0.75rem 1rem 0.75rem 2.8rem;
            background: var(--bg-primary);
            border: 1px solid var(--border-color);
            border-radius: 8px;
            color: var(--text-primary);
            font-size: 0.95rem;
            outline: none;
            transition: border-color 0.2s ease;
        }}

        .search-box input:focus {{
            border-color: var(--accent-primary);
        }}

        .filter-group {{
            display: flex;
            gap: 1rem;
            flex-wrap: wrap;
            align-items: center;
        }}

        select {{
            background: var(--bg-primary);
            border: 1px solid var(--border-color);
            border-radius: 8px;
            color: var(--text-primary);
            padding: 0.75rem 1rem;
            font-size: 0.9rem;
            outline: none;
            cursor: pointer;
        }}

        .btn {{
            padding: 0.75rem 1.25rem;
            border-radius: 8px;
            border: none;
            font-weight: 500;
            font-size: 0.9rem;
            cursor: pointer;
            display: inline-flex;
            align-items: center;
            gap: 0.5rem;
            transition: all 0.2s ease;
            text-decoration: none;
        }}

        .btn-primary {{
            background: linear-gradient(135deg, var(--accent-primary), var(--accent-secondary));
            color: white;
        }}

        .btn-primary:hover {{
            opacity: 0.9;
            box-shadow: 0 4px 12px rgba(14, 165, 233, 0.3);
        }}

        .btn-link {{
            padding: 0.35rem 0.65rem;
            background: rgba(14, 165, 233, 0.1);
            color: var(--badge-text);
            border: 1px solid rgba(14, 165, 233, 0.3);
            border-radius: 6px;
            font-size: 0.8rem;
            font-weight: 500;
            display: inline-flex;
            align-items: center;
            gap: 0.4rem;
            text-decoration: none;
            transition: all 0.2s ease;
        }}

        .btn-link:hover {{
            background: rgba(14, 165, 233, 0.25);
            border-color: var(--accent-primary);
            color: white;
        }}

        /* Table Styling */
        .table-container {{
            background: var(--bg-card);
            border-radius: 12px;
            border: 1px solid var(--border-color);
            overflow: hidden;
            box-shadow: 0 10px 30px rgba(0, 0, 0, 0.2);
        }}

        table {{
            width: 100%;
            border-collapse: collapse;
            text-align: left;
            font-size: 0.9rem;
        }}

        th {{
            background: var(--table-header-bg);
            color: var(--text-secondary);
            font-weight: 600;
            padding: 1rem;
            border-bottom: 1px solid var(--border-color);
            text-transform: uppercase;
            font-size: 0.75rem;
            letter-spacing: 0.05em;
            cursor: pointer;
            user-select: none;
        }}

        th:hover {{
            color: var(--accent-primary);
        }}

        td {{
            padding: 1rem;
            border-bottom: 1px solid var(--border-color);
            color: var(--text-primary);
            vertical-align: top;
        }}

        tr:hover td {{
            background-color: rgba(255, 255, 255, 0.02);
        }}

        .badge-oc {{
            background: var(--badge-bg);
            color: var(--badge-text);
            padding: 0.25rem 0.6rem;
            border-radius: 6px;
            font-weight: 600;
            font-size: 0.85rem;
            display: inline-block;
        }}

        .badge-cant {{
            background: rgba(16, 185, 129, 0.15);
            color: #34d399;
            padding: 0.25rem 0.6rem;
            border-radius: 6px;
            font-weight: 600;
            font-size: 0.85rem;
            display: inline-block;
        }}

        .text-desc {{
            color: var(--text-secondary);
            font-size: 0.85rem;
            white-space: pre-wrap;
            max-width: 380px;
            line-height: 1.4;
        }}

        .btn-view-desc {{
            color: var(--accent-primary);
            background: none;
            border: none;
            cursor: pointer;
            font-size: 0.8rem;
            margin-top: 0.4rem;
            text-decoration: underline;
        }}

        /* Pagination */
        .pagination {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 1rem 1.25rem;
            background: var(--bg-card);
            border-top: 1px solid var(--border-color);
        }}

        .page-info {{
            color: var(--text-secondary);
            font-size: 0.85rem;
        }}

        .page-buttons {{
            display: flex;
            gap: 0.5rem;
        }}

        .page-btn {{
            padding: 0.4rem 0.8rem;
            background: var(--bg-primary);
            border: 1px solid var(--border-color);
            border-radius: 6px;
            color: var(--text-primary);
            cursor: pointer;
            font-size: 0.85rem;
        }}

        .page-btn.active {{
            background: var(--accent-primary);
            border-color: var(--accent-primary);
        }}

        .page-btn:disabled {{
            opacity: 0.4;
            cursor: not-allowed;
        }}

        /* Modal */
        .modal-overlay {{
            position: fixed;
            top: 0; left: 0; right: 0; bottom: 0;
            background: rgba(0, 0, 0, 0.7);
            backdrop-filter: blur(4px);
            display: none;
            align-items: center;
            justify-content: center;
            z-index: 1000;
        }}

        .modal-card {{
            background: var(--bg-card);
            border: 1px solid var(--border-color);
            border-radius: 14px;
            width: 90%;
            max-width: 650px;
            padding: 1.5rem;
            box-shadow: 0 20px 40px rgba(0,0,0,0.5);
        }}

        .modal-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 1rem;
            padding-bottom: 0.75rem;
            border-bottom: 1px solid var(--border-color);
        }}

        .modal-header h3 {{
            font-size: 1.1rem;
            color: var(--accent-primary);
        }}

        .modal-close {{
            background: none;
            border: none;
            color: var(--text-secondary);
            font-size: 1.2rem;
            cursor: pointer;
        }}

        .modal-body {{
            color: var(--text-primary);
            font-size: 0.95rem;
            line-height: 1.6;
            white-space: pre-wrap;
            max-height: 60vh;
            overflow-y: auto;
        }}

        @media (max-width: 768px) {{
            body {{ padding: 1rem; }}
            .controls-card {{ flex-direction: column; align-items: stretch; }}
            .search-box {{ min-width: 100%; }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <header>
            <div class="header-title">
                <h1><i class="fa-solid fa-file-invoice-dollar"></i> Órdenes de Compra TELCOM (OC &ge; 20174)</h1>
                <p>Reporte consolidado con base de proveedores y enlace a comprobantes PDF</p>
            </div>
            <button class="btn btn-primary" onclick="exportToCSV()">
                <i class="fa-solid fa-file-excel"></i> Exportar a Excel (CSV)
            </button>
        </header>

        <!-- Stats Grid -->
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-icon"><i class="fa-solid fa-receipt"></i></div>
                <div class="stat-info">
                    <h3 id="stat-total-ocs">0</h3>
                    <p>Órdenes de Compra (&ge; 20174)</p>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-icon" style="color: #10b981; background: rgba(16,185,129,0.1);"><i class="fa-solid fa-boxes-stacked"></i></div>
                <div class="stat-info">
                    <h3 id="stat-total-items">0</h3>
                    <p>Total de Artículos Extraídos</p>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-icon" style="color: #6366f1; background: rgba(99,102,241,0.1);"><i class="fa-solid fa-building-user"></i></div>
                <div class="stat-info">
                    <h3 id="stat-total-proveedores">0</h3>
                    <p>Proveedores Únicos (CUIT Base)</p>
                </div>
            </div>
        </div>

        <!-- Controls Card -->
        <div class="controls-card">
            <div class="search-box">
                <i class="fa-solid fa-magnifying-glass"></i>
                <input type="text" id="searchInput" placeholder="Buscar por N° OC, Proveedor, Artículo, Solicitante, Descripción..." oninput="onSearchChange()">
            </div>
            <div class="filter-group">
                <select id="supplierFilter" onchange="onFilterChange()">
                    <option value="">Todos los Proveedores</option>
                </select>
                <select id="itemsPerPage" onchange="onPageSizeChange()">
                    <option value="25">25 por pág.</option>
                    <option value="50" selected>50 por pág.</option>
                    <option value="100">100 por pág.</option>
                    <option value="500">500 por pág.</option>
                </select>
            </div>
        </div>

        <!-- Table Container -->
        <div class="table-container">
            <table>
                <thead>
                    <tr>
                        <th onclick="sortTable('nro_oc')">N° OC <i class="fa-solid fa-sort"></i></th>
                        <th onclick="sortTable('fecha')">Fecha <i class="fa-solid fa-sort"></i></th>
                        <th onclick="sortTable('proveedor')">Proveedor <i class="fa-solid fa-sort"></i></th>
                        <th onclick="sortTable('info_adicional')">Info Adicional <i class="fa-solid fa-sort"></i></th>
                        <th onclick="sortTable('cantidad')">Cant. <i class="fa-solid fa-sort"></i></th>
                        <th onclick="sortTable('articulo')">Artículo <i class="fa-solid fa-sort"></i></th>
                        <th>Descripción Debajo</th>
                        <th>Ver OC</th>
                    </tr>
                </thead>
                <tbody id="tableBody">
                    <!-- Dynamic Rows -->
                </tbody>
            </table>

            <div class="pagination">
                <span class="page-info" id="pageInfo">Mostrando 0 a 0 de 0 registros</span>
                <div class="page-buttons" id="pageButtons">
                    <!-- Dynamic Buttons -->
                </div>
            </div>
        </div>
    </div>

    <!-- Detail Modal -->
    <div class="modal-overlay" id="detailModal">
        <div class="modal-card">
            <div class="modal-header">
                <h3 id="modalTitle">Detalle Adicional del Artículo</h3>
                <button class="modal-close" onclick="closeModal()"><i class="fa-solid fa-xmark"></i></button>
            </div>
            <div class="modal-body" id="modalBody"></div>
        </div>
    </div>

    <script>
        const RAW_DATA = {json_data};
        let filteredData = [...RAW_DATA];
        let currentPage = 1;
        let pageSize = 50;
        let sortColumn = 'nro_oc';
        let sortAsc = false;

        document.addEventListener('DOMContentLoaded', () => {{
            populateSupplierFilter();
            updateStats();
            sortTable('nro_oc', false);
        }});

        function populateSupplierFilter() {{
            const suppliers = [...new Set(RAW_DATA.map(d => d.proveedor).filter(Boolean))].sort();
            const select = document.getElementById('supplierFilter');
            suppliers.forEach(sup => {{
                const opt = document.createElement('option');
                opt.value = sup;
                opt.textContent = sup;
                select.appendChild(opt);
            }});
        }}

        function updateStats() {{
            const uniqueOCs = new Set(filteredData.map(d => d.nro_oc)).size;
            const uniqueSuppliers = new Set(filteredData.map(d => d.proveedor).filter(Boolean)).size;
            
            document.getElementById('stat-total-ocs').textContent = uniqueOCs.toLocaleString();
            document.getElementById('stat-total-items').textContent = filteredData.length.toLocaleString();
            document.getElementById('stat-total-proveedores').textContent = uniqueSuppliers.toLocaleString();
        }}

        function renderTable() {{
            const tbody = document.getElementById('tableBody');
            tbody.innerHTML = '';

            const start = (currentPage - 1) * pageSize;
            const end = start + pageSize;
            const pageData = filteredData.slice(start, end);

            if (pageData.length === 0) {{
                tbody.innerHTML = `<tr><td colspan="8" style="text-align: center; padding: 2rem; color: var(--text-secondary);">No se encontraron registros que coincidan con la búsqueda.</td></tr>`;
                document.getElementById('pageInfo').textContent = 'Mostrando 0 de 0 registros';
                renderPagination(0);
                return;
            }}

            pageData.forEach(row => {{
                const tr = document.createElement('tr');
                
                const descText = row.descripcion_debajo || '';
                let descHTML = '-';
                if (descText) {{
                    const shortText = descText.length > 110 ? descText.substring(0, 110) + '...' : descText;
                    descHTML = `<div class="text-desc">${{escapeHTML(shortText)}}</div>`;
                    if (descText.length > 110) {{
                        descHTML += `<button class="btn-view-desc" onclick="openModal('${{escapeHTML(row.articulo)}}', \\`${{escapeModalText(descText)}}\\`)">Ver completo</button>`;
                    }}
                }}

                const pdfLink = row.archivo ? `<a href="${{encodeURI(row.archivo)}}" target="_blank" class="btn-link"><i class="fa-solid fa-file-pdf"></i> PDF</a>` : '-';

                tr.innerHTML = `
                    <td><span class="badge-oc">${{escapeHTML(row.nro_oc)}}</span></td>
                    <td style="white-space: nowrap;">${{escapeHTML(row.fecha || '-')}}</td>
                    <td><strong>${{escapeHTML(row.proveedor || '-')}}</strong></td>
                    <td style="color: var(--text-secondary); font-size: 0.85rem;">${{escapeHTML(row.info_adicional || '-')}}</td>
                    <td><span class="badge-cant">${{escapeHTML(row.cantidad || '1')}}</span></td>
                    <td style="font-weight: 500;">${{escapeHTML(row.articulo || '-')}}</td>
                    <td>${{descHTML}}</td>
                    <td>${{pdfLink}}</td>
                `;
                tbody.appendChild(tr);
            }});

            const total = filteredData.length;
            const currentStart = start + 1;
            const currentEnd = Math.min(end, total);
            document.getElementById('pageInfo').textContent = `Mostrando ${{currentStart}} a ${{currentEnd}} de ${{total}} registros`;

            renderPagination(Math.ceil(total / pageSize));
        }}

        function renderPagination(totalPages) {{
            const container = document.getElementById('pageButtons');
            container.innerHTML = '';

            if (totalPages <= 1) return;

            const prevBtn = document.createElement('button');
            prevBtn.className = 'page-btn';
            prevBtn.innerHTML = '<i class="fa-solid fa-chevron-left"></i>';
            prevBtn.disabled = currentPage === 1;
            prevBtn.onclick = () => {{ currentPage--; renderTable(); }};
            container.appendChild(prevBtn);

            let startPage = Math.max(1, currentPage - 2);
            let endPage = Math.min(totalPages, startPage + 4);
            if (endPage - startPage < 4) startPage = Math.max(1, endPage - 4);

            for (let i = startPage; i <= endPage; i++) {{
                const btn = document.createElement('button');
                btn.className = `page-btn ${{i === currentPage ? 'active' : ''}}`;
                btn.textContent = i;
                btn.onclick = () => {{ currentPage = i; renderTable(); }};
                container.appendChild(btn);
            }}

            const nextBtn = document.createElement('button');
            nextBtn.className = 'page-btn';
            nextBtn.innerHTML = '<i class="fa-solid fa-chevron-right"></i>';
            nextBtn.disabled = currentPage === totalPages;
            nextBtn.onclick = () => {{ currentPage++; renderTable(); }};
            container.appendChild(nextBtn);
        }}

        function applyFilters() {{
            const search = document.getElementById('searchInput').value.toLowerCase().trim();
            const supplier = document.getElementById('supplierFilter').value;

            filteredData = RAW_DATA.filter(row => {{
                const matchSupplier = !supplier || row.proveedor === supplier;
                const matchSearch = !search || 
                    (row.nro_oc && row.nro_oc.toLowerCase().includes(search)) ||
                    (row.proveedor && row.proveedor.toLowerCase().includes(search)) ||
                    (row.articulo && row.articulo.toLowerCase().includes(search)) ||
                    (row.info_adicional && row.info_adicional.toLowerCase().includes(search)) ||
                    (row.descripcion_debajo && row.descripcion_debajo.toLowerCase().includes(search));

                return matchSupplier && matchSearch;
            }});

            currentPage = 1;
            updateStats();
            sortData();
            renderTable();
        }}

        function onSearchChange() {{ applyFilters(); }}
        function onFilterChange() {{ applyFilters(); }}
        function onPageSizeChange() {{
            pageSize = parseInt(document.getElementById('itemsPerPage').value);
            currentPage = 1;
            renderTable();
        }}

        function sortTable(column, toggle = true) {{
            if (toggle) {{
                if (sortColumn === column) sortAsc = !sortAsc;
                else {{ sortColumn = column; sortAsc = true; }}
            }}
            sortData();
            renderTable();
        }}

        function sortData() {{
            filteredData.sort((a, b) => {{
                let valA = a[sortColumn] || '';
                let valB = b[sortColumn] || '';

                if (sortColumn === 'nro_oc') {{
                    valA = parseInt(valA) || 0;
                    valB = parseInt(valB) || 0;
                }} else if (sortColumn === 'cantidad') {{
                    valA = parseFloat(valA.replace('.', '').replace(',', '.')) || 0;
                    valB = parseFloat(valB.replace('.', '').replace(',', '.')) || 0;
                }} else if (sortColumn === 'fecha') {{
                    valA = parseDateToISO(valA);
                    valB = parseDateToISO(valB);
                }}

                if (valA < valB) return sortAsc ? -1 : 1;
                if (valA > valB) return sortAsc ? 1 : -1;
                return 0;
            }});
        }}

        function parseDateToISO(dStr) {{
            if (!dStr || typeof dStr !== 'string') return '';
            const parts = dStr.trim().split('/');
            if (parts.length === 3) {{
                return parts[2] + parts[1].padStart(2, '0') + parts[0].padStart(2, '0');
            }}
            return dStr;
        }}

        function openModal(title, content) {{
            document.getElementById('modalTitle').textContent = title;
            document.getElementById('modalBody').textContent = content;
            document.getElementById('detailModal').style.display = 'flex';
        }}

        function closeModal() {{
            document.getElementById('detailModal').style.display = 'none';
        }}

        function escapeHTML(str) {{
            if (!str) return '';
            return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
        }}

        function escapeModalText(str) {{
            if (!str) return '';
            return str.replace(/\\\\/g, '\\\\\\\\').replace(/`/g, '\\\\`').replace(/\\$/g, '\\\\$');
        }}

        function exportToCSV() {{
            let csv = '\\uFEFF';
            csv += 'Nro OC;Fecha;Proveedor;Info Adicional;Cantidad;Articulo;Descripcion Debajo;Archivo\\n';
            filteredData.forEach(row => {{
                const clean = (text) => `"${{(text || '').replace(/"/g, '""')}}"`;
                csv += `${{clean(row.nro_oc)}};${{clean(row.fecha)}};${{clean(row.proveedor)}};${{clean(row.info_adicional)}};${{clean(row.cantidad)}};${{clean(row.articulo)}};${{clean(row.descripcion_debajo)}};${{clean(row.archivo)}}\\n`;
            }});
            const blob = new Blob([csv], {{ type: 'text/csv;charset=utf-8;' }});
            const url = URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.setAttribute('href', url);
            link.setAttribute('download', `Reporte_OC_Antina_20174_${{new Date().toISOString().slice(0,10)}}.csv`);
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }}
    </script>
</body>
</html>"""
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    print(f"[OK] Reporte HTML generado exitosamente en: {output_path}", flush=True)

def main():
    print(f"[*] Iniciando escaneo de Órdenes de Compra (>= {MIN_OC_NUMBER}) en: {NETWORK_FOLDER}", flush=True)
    start_time = time.time()
    
    if not os.path.exists(NETWORK_FOLDER):
        print(f"[ERROR] La carpeta {NETWORK_FOLDER} no es accesible.", flush=True)
        sys.exit(1)
        
    all_files = [os.path.join(NETWORK_FOLDER, f) for f in os.listdir(NETWORK_FOLDER) if f.lower().endswith('.pdf')]
    total_files = len(all_files)
    print(f"[*] Archivos PDF totales en carpeta: {total_files}", flush=True)
    
    all_records = []
    processed_files = 0
    oc_files_count = 0
    
    print("[*] Procesando PDFs con 30 hilos concurrentes...", flush=True)
    with ThreadPoolExecutor(max_workers=30) as executor:
        futures = {executor.submit(parse_pdf_file, filepath): filepath for filepath in all_files}
        
        for future in as_completed(futures):
            processed_files += 1
            if processed_files % 1000 == 0 or processed_files == total_files:
                print(f"  -> Escaneados {processed_files}/{total_files} PDFs... ({len(all_records)} registros de OC >= 20174 encontrados)", flush=True)
                
            res = future.result()
            if res:
                oc_files_count += 1
                all_records.extend(res)
                
    elapsed = time.time() - start_time
    print(f"\n[+] Escaneo completado en {elapsed:.2f} segundos.", flush=True)
    print(f"[+] Total PDFs analizados: {total_files}", flush=True)
    print(f"[+] Archivos de OC (>= {MIN_OC_NUMBER}) procesados: {oc_files_count}", flush=True)
    print(f"[+] Total de registros de artículos extraídos: {len(all_records)}", flush=True)
    
    # Save raw JSON data
    with open(OUTPUT_JSON, 'w', encoding='utf-8') as f:
        json.dump(all_records, f, ensure_ascii=False, indent=2)
    print(f"[OK] Datos guardados en JSON: {OUTPUT_JSON}", flush=True)
    
    # Generate standalone HTML report locally and in network folder
    generate_html_report(all_records, OUTPUT_HTML)
    network_html = os.path.join(NETWORK_FOLDER, "Reporte_OC_Antina.html")
    try:
        generate_html_report(all_records, network_html)
        print(f"[OK] Reporte HTML copiado a red: {network_html}", flush=True)
    except Exception as e:
        print(f"[WARNING] No se pudo copiar a red: {e}", flush=True)

if __name__ == "__main__":
    main()
