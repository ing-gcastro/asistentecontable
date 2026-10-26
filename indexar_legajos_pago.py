import os
import re
import sys
from pathlib import Path
from datetime import datetime
import pandas as pd
import fitz  # PyMuPDF

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

# ─── RUTAS Y CONFIGURACIÓN ───────────────────────────────────────────────────
CARPETAS_LEGAJOS = [
    r"\\10.10.10.210\AyF_Trabajoadistancia\Facturas a pagar\Periodo Actual",
    r"\\10.10.10.210\AyF_Trabajoadistancia\Facturas a pagar\2026"
]

HTML_SALIDA = r"\\10.10.10.210\AyF_Trabajoadistancia\Facturas a pagar\Reporte_Legajos_Pago.html"

def parsear_liquidacion_pdf(ruta_pdf):
    """Lee y extrae la información estructurada del PDF de Liquidación de Pago."""
    try:
        doc = fitz.open(ruta_pdf)
        if len(doc) == 0:
            doc.close()
            return None
        
        texto_pagina1 = doc[0].get_text("text")
        doc.close()
        
        if "Liquidación de Pago" not in texto_pagina1 and "Liquidacion de Pago" not in texto_pagina1:
            return None
            
        lines = [line.strip() for line in texto_pagina1.split("\n") if line.strip()]
        
        # 1. Empresa Pagadora
        empresa = "TELCOM VENTURES"
        if "EUROSAT" in texto_pagina1.upper():
            empresa = "EUROSAT"
            
        # 2. Número de Liquidación (4 a 5 dígitos, NUNCA 2 dígitos)
        nro_liq = ""
        m_liq_before = re.search(r'(\d{4,5})\s*\n?\s*Liquidaci[oó]n\s*de\s*Pago', texto_pagina1, re.IGNORECASE)
        if m_liq_before:
            nro_liq = m_liq_before.group(1)
        else:
            m_liq_after = re.search(r'Liquidaci[oó]n\s*de\s*Pago\s*N[oº\.]*\s*\n?\s*(\d{4,5})', texto_pagina1, re.IGNORECASE)
            if m_liq_after:
                nro_liq = m_liq_after.group(1)
            else:
                m_fname = re.search(r'\b(\d{4,5})\b', os.path.basename(ruta_pdf))
                if m_fname:
                    nro_liq = m_fname.group(1)
                else:
                    m_folder = re.search(r'\b(\d{4,5})\b', os.path.basename(os.path.dirname(ruta_pdf)))
                    if m_folder:
                        nro_liq = m_folder.group(1)

        if not nro_liq or len(nro_liq) < 4:
            nro_liq = "S/N"

        # 3. Fecha de Emisión Liquidación
        fecha_emision = ""
        m_fecha = re.search(r'(\d{2}/\d{2}/\d{4})\s*\n?\s*Fecha\s*de\s*Emisi[oó]n', texto_pagina1, re.IGNORECASE)
        if m_fecha:
            fecha_emision = m_fecha.group(1)
        else:
            m_f2 = re.search(r'Fecha\s*de\s*Emisi[oó]n\s*[:\.]?\s*(\d{2}/\d{2}/\d{4})', texto_pagina1, re.IGNORECASE)
            if m_f2:
                fecha_emision = m_f2.group(1)
            else:
                m_f3 = re.search(r'\b(\d{2}/\d{2}/\d{4})\b', texto_pagina1)
                if m_f3: fecha_emision = m_f3.group(1)

        # 4. Proveedor y CUIT Proveedor
        prov_nombre = "DESCONOCIDO"
        cuit_prov = ""
        
        m_co = re.search(r'Co\.\s*Cuenta:\s*\n?\s*(\d+)\s*\n?\s*([^\n\r]+)', texto_pagina1)
        if m_co:
            prov_nombre = m_co.group(2).strip()
        else:
            m_p2 = re.search(r'Co\.\s*Cuenta:\s*([^\n\r]+)', texto_pagina1)
            if m_p2: prov_nombre = m_p2.group(1).strip()
            
        cuits_encontrados = re.findall(r'\b((?:20|23|24|27|30|33|34)\d{9})\b', texto_pagina1.replace("-", "").replace(" ", ""))
        for c in cuits_encontrados:
            if c not in ["30696428898", "33676549639"]:  # Excluir empresas propias
                cuit_prov = c
                break

        # 5. Extraer Comprobantes Incluidos (Tabla FACT, NCRE, NDEB)
        comprobantes = []
        rx_comp = r'\b(FACT|NCRE|NDEB|ND|NC|FC)\s+(\d{2}/\d{2}/\d{4})\s+(\d{2}/\d{2}/\d{4})\s+([A-Z0-9\-]+)\s+([\d\.,]+)'
        matches_comp = re.findall(rx_comp, texto_pagina1)
        
        for t_comp, f_comp, f_venc, nro_comp_raw, monto_comp in matches_comp:
            nro_limpio = nro_comp_raw
            if "-" in nro_comp_raw:
                partes_nro = nro_comp_raw.split("-")
                nro_limpio = str(int(partes_nro[-1]))
            elif nro_comp_raw.isdigit():
                nro_limpio = str(int(nro_comp_raw))
                
            t_std = "FC" if t_comp in ["FACT", "FC"] else ("NC" if t_comp in ["NCRE", "NC"] else "ND")
            comprobantes.append({
                "tipo": t_std,
                "nro_original": nro_comp_raw,
                "nro_corto": nro_limpio,
                "fecha_emision": f_comp,
                "fecha_vencimiento": f_venc,
                "monto": monto_comp
            })

        # 6. Extraer Totales (Bruto, Retenciones, Neto)
        bruto_str = ""
        m_bruto = re.search(r'Son\s*Pesos[^\n\r]*\n?\s*Total\s*\n?\s*\$\s*([\d\.,]+)', texto_pagina1, re.IGNORECASE)
        if m_bruto:
            bruto_str = m_bruto.group(1)
        else:
            m_b2 = re.search(r'Total\s*\n?\s*\$\s*([\d\.,]+)\s*\n?\s*(?:Retenciones|Firmante|Prepar)', texto_pagina1, re.IGNORECASE)
            if m_b2: bruto_str = m_b2.group(1)

        neto_str = bruto_str
        m_neto = re.search(r'Total\s*Neto\s*:\s*\n?\s*\$\s*([\d\.,]+)', texto_pagina1, re.IGNORECASE)
        if m_neto:
            neto_str = m_neto.group(1)

        # Sumar todas las retenciones presentes
        ret_Ganancias = re.search(r'Retenciones\s*Ganancias\s*:\s*\n?\s*\$\s*([\d\.,]+)', texto_pagina1)
        ret_IVA = re.search(r'Retenciones\s*I\.V\.A\.\s*:\s*\n?\s*\$\s*([\d\.,]+)', texto_pagina1)
        ret_IIBB = re.search(r'Retenciones\s*Ing\.?\s*Brutos\s*:\s*\n?\s*\$\s*([\d\.,]+)', texto_pagina1)
        ret_SUSS = re.search(r'Retenciones\s*Seg\.?\s*Social\s*:\s*\n?\s*\$\s*([\d\.,]+)', texto_pagina1)
        ret_Otros = re.search(r'Retenciones\s*Otros\s*:\s*\n?\s*\$\s*([\d\.,]+)', texto_pagina1)

        tot_ret_val = 0.0
        for r_m in [ret_Ganancias, ret_IVA, ret_IIBB, ret_SUSS, ret_Otros]:
            if r_m:
                v_str = r_m.group(1).replace(".", "").replace(",", ".")
                try: tot_ret_val += float(v_str)
                except: pass

        retenciones_str = f"{tot_ret_val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if tot_ret_val > 0 else "0,00"

        # 7. Medio de Pago / Banco
        banco = "N/A"
        m_banco = re.search(r'\b(BANCO[^\n\r]+|DCPE[^\n\r]+|CHEQUE[^\n\r]+)', texto_pagina1, re.IGNORECASE)
        if m_banco:
            banco = m_banco.group(1).strip()

        return {
            "Nro Liquidación": nro_liq,
            "Fecha": fecha_emision,
            "Empresa": empresa,
            "Proveedor": prov_nombre,
            "CUIT": cuit_prov,
            "Total Bruto": bruto_str if bruto_str else "0.00",
            "Retenciones": retenciones_str,
            "Total Neto": neto_str if neto_str else "0.00",
            "Banco": banco,
            "Comprobantes": comprobantes,
            "Cantidad Facturas": len(comprobantes),
            "Archivo PDF": os.path.basename(ruta_pdf),
            "Ruta Legajo": os.path.dirname(ruta_pdf)
        }
    except Exception as e:
        print(f"⚠️ Error leyendo PDF {ruta_pdf}: {e}")
        return None

def generar_reporte_legajos():
    print("🚀 Iniciando indexación de Legajos de Pago 2026...")
    registros = []
    procesados_set = set()

    for base_folder in CARPETAS_LEGAJOS:
        if not os.path.exists(base_folder):
            print(f"⚠️ Ruta no encontrada: {base_folder}")
            continue

        print(f"📂 Escaneando: {base_folder}...")
        for root, dirs, files in os.walk(base_folder):
            for f in files:
                if f.lower().endswith(".pdf"):
                    ruta_pdf = os.path.join(root, f)
                    if ruta_pdf in procesados_set:
                        continue

                    # Inspeccionar si es una Liquidación de Pago
                    data = parsear_liquidacion_pdf(ruta_pdf)
                    if data:
                        procesados_set.add(ruta_pdf)
                        registros.append(data)

    print(f"✅ Se encontraron {len(registros)} Liquidaciones de Pago válidas.")
    if not registros:
        print("❌ No se encontraron datos para generar el reporte.")
        return

    df = pd.DataFrame(registros)
    
    # Escribir HTML interactivo
    _escribir_html_legajos(df, HTML_SALIDA)

def _escribir_html_legajos(df, ruta_salida):
    timestamp_actual = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    
    html_content = f"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
    <title>Reporte de Legajos y Liquidaciones de Pago 2026</title>
    <!-- Bootstrap 5 CSS -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <!-- DataTables CSS -->
    <link href="https://cdn.datatables.net/1.13.6/css/dataTables.bootstrap5.min.css" rel="stylesheet">
    <style>
        body {{ font-family: 'Segoe UI', Arial, sans-serif; background-color: #f4f7f6; margin: 0; padding: 15px; }}
        .container-fluid {{ max-width: 98%; margin: auto; background: white; padding: 25px; border-radius: 10px; box-shadow: 0px 4px 15px rgba(0,0,0,0.08); }}
        table {{ width: 100%; border-collapse: collapse; margin-top: 15px; }}
        th, td {{ padding: 8px 12px; text-align: left; border-bottom: 1px solid #e9ecef; font-size: 0.85rem; }}
        th {{ background-color: #0d6efd; color: white; }}
        tr:hover {{ background-color: #f8f9fa; }}
        .badge-empresa {{ font-size: 0.75rem; padding: 4px 8px; border-radius: 4px; }}
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
        .comp-list {{ font-size: 0.78rem; font-family: monospace; background: #f8f9fa; padding: 4px 8px; border-radius: 4px; border: 1px solid #dee2e6; }}
    </style>
</head>
<body>
    <div class="container-fluid">
        <div class="d-flex justify-content-between align-items-center mb-3">
            <div>
                <h2 class="text-primary fw-bold">💳 Reporte de Legajos de Pago y Liquidaciones (2026)</h2>
                <p class="text-muted mb-0">Última actualización: <strong>{timestamp_actual}</strong> | Total de Liquidaciones: <strong>{len(df)}</strong></p>
            </div>
            <div>
                <span class="badge bg-primary fs-6 p-2">Legajos Procesados: {len(df)}</span>
            </div>
        </div>
        <hr>
        <div class="table-responsive">
            <table class="table table-striped table-bordered table-hover align-middle" id="tabla_legajos">
                <thead>
                    <tr>
                        <th>Fecha Pago</th>
                        <th>Empresa</th>
                        <th>N° Liq.</th>
                        <th>Proveedor</th>
                        <th>CUIT</th>
                        <th>Comprobantes / Facturas Incluidas</th>
                        <th>Total Bruto ($)</th>
                        <th>Retenciones ($)</th>
                        <th>Total Neto ($)</th>
                        <th>Medio / Banco</th>
                        <th>Ruta del Legajo (Carpeta)</th>
                    </tr>
                    <tr>
                        <th><input type="text" placeholder="Filtrar Fecha" /></th>
                        <th><input type="text" placeholder="Filtrar Empresa" /></th>
                        <th><input type="text" placeholder="Filtrar N° Liq" /></th>
                        <th><input type="text" placeholder="Filtrar Proveedor" /></th>
                        <th><input type="text" placeholder="Filtrar CUIT" /></th>
                        <th><input type="text" placeholder="Filtrar Facturas" /></th>
                        <th><input type="text" placeholder="Filtrar Bruto" /></th>
                        <th><input type="text" placeholder="Filtrar Retenciones" /></th>
                        <th><input type="text" placeholder="Filtrar Neto" /></th>
                        <th><input type="text" placeholder="Filtrar Banco" /></th>
                        <th><input type="text" placeholder="Filtrar Ruta Legajo" /></th>
                    </tr>
                </thead>
                <tbody>
"""

    for _, row in df.iterrows():
        empresa_badge = '<span class="badge bg-secondary badge-empresa">TELCOM VENTURES</span>' if row['Empresa'] == 'TELCOM VENTURES' else '<span class="badge bg-warning text-dark badge-empresa">EUROSAT</span>'
        
        # Generar data-order para ordenamiento cronológico real (YYYYMMDD)
        fecha_str = str(row['Fecha']).strip()
        data_order_attr = ""
        if "/" in fecha_str:
            p = fecha_str.split("/")
            if len(p) == 3:
                data_order_attr = f'data-order="{p[2]}{p[1].zfill(2)}{p[0].zfill(2)}"'

        # Formatear lista de comprobantes
        comps_txt = []
        for c in row['Comprobantes']:
            comps_txt.append(f"{c['tipo']} {c['nro_corto']} (${c['monto']})")
            
        comps_str = "<br>".join(comps_txt) if comps_txt else "<span class='text-muted'>Sin detalle</span>"
        
        ruta_legajo = str(row['Ruta Legajo'])
        nombre_carpeta_legajo = os.path.basename(ruta_legajo)
        ruta_link = f'<a href="file:///{ruta_legajo.replace("#", "%23")}" target="_blank" class="text-decoration-none" title="Abrir carpeta legajo">📁 {nombre_carpeta_legajo}</a>'

        html_content += f"""
                    <tr>
                        <td class="fw-bold" {data_order_attr}>{row['Fecha']}</td>
                        <td>{empresa_badge}</td>
                        <td class="fw-bold text-primary">{row['Nro Liquidación']}</td>
                        <td class="fw-bold">{row['Proveedor']}</td>
                        <td style="font-family: monospace;">{row['CUIT']}</td>
                        <td><div class="comp-list">{comps_str}</div></td>
                        <td class="text-end fw-bold text-dark">{row['Total Bruto']}</td>
                        <td class="text-end text-danger">{row['Retenciones']}</td>
                        <td class="text-end fw-bold text-success">{row['Total Neto']}</td>
                        <td><small>{row['Banco']}</small></td>
                        <td style="font-size: 0.75rem; word-break: break-all;">{ruta_link}</td>
                    </tr>
"""

    html_content += f"""
                </tbody>
            </table>
        </div>
        <div class="footer text-muted text-end mt-3" style="font-size: 0.8rem;">Sistema de Gestión Documental - Reporte de Legajos de Pago</div>
    </div>

    <!-- Scripts -->
    <script src="https://code.jquery.com/jquery-3.7.0.min.js"></script>
    <script src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
    <script src="https://cdn.datatables.net/1.13.6/js/dataTables.bootstrap5.min.js"></script>
    <script>
        $(document).ready(function() {{
            var table = $('#tabla_legajos').DataTable({{
                "language": {{
                    "url": "//cdn.datatables.net/plug-ins/1.13.6/i18n/es-ES.json"
                }},
                "pageLength": 25,
                "order": [[0, "desc"]]
            }});

            $('#tabla_legajos thead tr:eq(1) input').on('keyup change clear', function () {{
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

    os.makedirs(os.path.dirname(ruta_salida), exist_ok=True)
    with open(ruta_salida, "w", encoding="utf-8") as f:
        f.write(html_content)

    print(f"🎉 ¡Reporte HTML de Legajos generado con éxito en: {ruta_salida}!")

if __name__ == "__main__":
    generar_reporte_legajos()
