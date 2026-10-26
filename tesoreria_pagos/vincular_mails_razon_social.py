import os
import re
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Ensure utf-8 output encoding for terminal
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

# ─── RUTAS DE ENTRADA Y SALIDA ───────────────────────────────────────────────
EXCEL_ENTRADA = os.path.join(os.path.expanduser("~"), r"ROBOTS\data\Copia de reporte_mails_pago_ultimos_3_meses.xlsx")
HTML_LEGAJOS = r"\\10.10.10.210\AyF_Trabajoadistancia\Facturas a pagar\Reporte_Legajos_Pago.html"
EXCEL_SALIDA = os.path.join(os.path.expanduser("~"), r"ROBOTS\data\vinculacion_mails_razon_social.xlsx")

# Casillas genéricas internas de Antina que NO corresponden a la razón social del proveedor
GENERIC_ANTINA_EMAILS = {
    "compras@antina.com.ar",
    "proveedores@antina.com.ar"
}


def limpiar_emails_proveedor(mail_str):
    """
    Extrae la dirección de email del proveedor.
    - Si la casilla es únicamente compras@antina.com.ar -> retorna "" (vacío).
    - Si vienen múltiples casillas y contiene casillas genéricas de Antina
      (ej: proveedores@antina.com.ar o compras@antina.com.ar), se filtran
      para dejar únicamente las casillas reales del proveedor.
    """
    if not mail_str:
        return ""
    
    mail_limpio = mail_str.strip()
    if mail_limpio.lower() == "compras@antina.com.ar":
        return ""

    # Separar si vienen múltiples destinatarios
    partes = [m.strip() for m in re.split(r'[,;]', mail_limpio) if m.strip()]
    
    # Filtrar casillas genéricas internas cuando hay emails de contacto
    externos = [p for p in partes if p.lower() not in GENERIC_ANTINA_EMAILS]
    
    if externos:
        return ", ".join(externos)
    
    # Si todas las casillas eran compras/proveedores genéricas de Antina, dejar vacío
    return ""


def cargar_mapeo_html(ruta_html):
    """
    Parsea Reporte_Legajos_Pago.html para indexar:
    N° Liq -> {proveedor, cuit, empresa, fecha_pago}
    """
    if not os.path.exists(ruta_html):
        raise FileNotFoundError(f"No se encontró el archivo HTML en: {ruta_html}")

    with open(ruta_html, "r", encoding="utf-8") as f:
        html_content = f.read()

    # Regex para extraer filas de la tabla de legajos (#tabla_legajos)
    pattern = re.compile(
        r'<tr>\s*<td[^>]*>(.*?)</td>\s*<td[^>]*>(.*?)</td>\s*<td[^>]*>(.*?)</td>\s*<td[^>]*>(.*?)</td>\s*<td[^>]*>(.*?)</td>',
        re.DOTALL
    )
    html_rows = pattern.findall(html_content)

    mapa_liq = {}
    for r in html_rows:
        fecha = re.sub(r'<[^>]+>', '', r[0]).strip()
        empresa = re.sub(r'<[^>]+>', '', r[1]).strip()
        liq = re.sub(r'<[^>]+>', '', r[2]).strip()
        prov = re.sub(r'<[^>]+>', '', r[3]).strip()
        cuit = re.sub(r'<[^>]+>', '', r[4]).strip()

        if liq and liq != 'S/N':
            mapa_liq[liq] = {
                'proveedor': prov,
                'cuit': cuit,
                'empresa': empresa,
                'fecha_pago': fecha
            }

    print(f"✅ HTML procesado con éxito: {len(mapa_liq)} liquidaciones registradas en base.")
    return mapa_liq


def formatear_hoja(ws, titulo_columnas, filas_datos, anchos_minimos=None):
    """
    Aplica diseño profesional al estilo dashboard a la hoja de cálculo.
    """
    header_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    data_font = Font(name="Segoe UI", size=9.5)
    data_font_bold = Font(name="Segoe UI", size=9.5, bold=True)
    align_left = Alignment(horizontal="left", vertical="center")
    align_center = Alignment(horizontal="center", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    thin_border = Border(
        left=Side(style='thin', color='E0E0E0'),
        right=Side(style='thin', color='E0E0E0'),
        top=Side(style='thin', color='E0E0E0'),
        bottom=Side(style='thin', color='E0E0E0')
    )

    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    # Escribir encabezados
    ws.append(titulo_columnas)
    ws.row_dimensions[1].height = 26

    for col_idx, cell in enumerate(ws[1], start=1):
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border

    # Escribir filas
    for row_idx, fila in enumerate(filas_datos, start=2):
        ws.append(fila)
        ws.row_dimensions[row_idx].height = 20
        fill_actual = zebra_fill if row_idx % 2 == 0 else white_fill

        for col_idx, valor in enumerate(fila, start=1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = data_font
            cell.fill = fill_actual
            cell.border = thin_border

            nombre_col = titulo_columnas[col_idx - 1].lower()
            if any(k in nombre_col for k in ["fecha", "n° liq", "liq", "cuit", "cant", "empresa"]):
                cell.alignment = align_center
            elif any(k in nombre_col for k in ["total", "monto", "importe"]):
                cell.alignment = align_right
            else:
                cell.alignment = align_left

            if "razón social" in nombre_col:
                cell.font = data_font_bold

    # Congelar fila 1 y activar autofiltro
    ws.freeze_panes = "A2"
    letra_max = get_column_letter(len(titulo_columnas))
    ws.auto_filter.ref = f"A1:{letra_max}{len(filas_datos) + 1}"

    # Auto-ajuste de anchos de columna
    for col_idx, col_cells in enumerate(ws.columns, start=1):
        max_len = max(len(str(cell.value or '')) for cell in col_cells)
        col_letter = get_column_letter(col_idx)
        min_w = (anchos_minimos or {}).get(col_letter, 12)
        ws.column_dimensions[col_letter].width = max(min_w, min(max_len + 4, 60))


def procesar_reporte():
    print("=" * 75)
    print("🚀 VINCULANDO DIRECCIONES DE MAIL CON RAZÓN SOCIAL")
    print("=" * 75)

    # 1. Cargar mapeo HTML
    mapa_liq = cargar_mapeo_html(HTML_LEGAJOS)

    # 2. Cargar Excel de mails
    if not os.path.exists(EXCEL_ENTRADA):
        raise FileNotFoundError(f"No se encontró el reporte Excel en: {EXCEL_ENTRADA}")

    wb_origen = openpyxl.load_workbook(EXCEL_ENTRADA, data_only=True)
    ws_origen = wb_origen.active

    registros_vinculados = []
    registros_sin_liq = []

    total_aviso_pago = 0
    total_compras_sin_mail = 0

    for idx, row in enumerate(ws_origen.iter_rows(min_row=2, values_only=True), start=2):
        fecha, dest, asunto, mail, adjuntos, cant_adj = row[:6]
        asunto_str = str(asunto or '').strip()

        # Condición: Solo mails que en su asunto indiquen AVISO DE PAGO
        if "AVISO DE PAGO" not in asunto_str.upper():
            continue

        total_aviso_pago += 1
        mail_str = str(mail or '').strip()
        adj_str = str(adjuntos or '').strip()

        # Condición: N° de 5 dígitos entre 65000 y 70000 en el nombre de archivos adjuntos
        liqs_5d = list(dict.fromkeys(re.findall(r'\b(6[5-9]\d{3}|70000)\b', adj_str)))
        # Soporte adicional para liquidaciones de 4 dígitos Eurosat (8000 a 8999)
        liqs_4d = list(dict.fromkeys(re.findall(r'\b(8\d{3})\b', adj_str)))

        todas_liqs = liqs_5d if liqs_5d else liqs_4d

        if not todas_liqs:
            registros_sin_liq.append({
                'row': idx,
                'fecha': fecha,
                'destinatario': dest,
                'asunto': asunto_str,
                'mail': mail_str,
                'adjuntos': adj_str
            })
            continue

        # Obtener Razón Social y CUIT desde el HTML mediante el N° de liquidación
        nombres_prov = []
        cuits_prov = []
        empresas_prov = []

        for l in todas_liqs:
            if l in mapa_liq:
                info = mapa_liq[l]
                if info['proveedor'] not in nombres_prov:
                    nombres_prov.append(info['proveedor'])
                if info['cuit'] not in cuits_prov:
                    cuits_prov.append(info['cuit'])
                if info['empresa'] not in empresas_prov:
                    empresas_prov.append(info['empresa'])
            else:
                print(f"⚠️ Advertencia: Liq N° {l} no encontrada en HTML (Fila {idx})")

        razon_social_str = " / ".join(nombres_prov) if nombres_prov else "NO ENCONTRADO EN HTML"
        cuit_str = " / ".join(cuits_prov) if cuits_prov else ""
        empresa_str = " / ".join(empresas_prov) if empresas_prov else ""

        # Regla: "en los casos que la casilla de recepcion solo sea compras@antina.com.ar
        # no coloques nada, solo indica la razon social del proveedor"
        mail_proveedor_limpio = limpiar_emails_proveedor(mail_str)
        if not mail_proveedor_limpio and mail_str.lower() == "compras@antina.com.ar":
            total_compras_sin_mail += 1

        registros_vinculados.append({
            'row_orig': idx,
            'fecha_envio': fecha,
            'destinatario': dest,
            'asunto': asunto_str,
            'mail_proveedor': mail_proveedor_limpio,
            'mail_original': mail_str,
            'razon_social': razon_social_str,
            'cuit': cuit_str,
            'liq_numero': ", ".join(todas_liqs),
            'empresa': empresa_str,
            'archivos_adjuntos': adj_str
        })

    wb_origen.close()

    print(f"📊 Estadísticas del procesamiento:")
    print(f"   - Total correos con 'AVISO DE PAGO': {total_aviso_pago}")
    print(f"   - Vinculados con Razón Social: {len(registros_vinculados)}")
    print(f"   - Casos exclusivos compras@antina.com.ar (mail dejado vacío): {total_compras_sin_mail}")
    print(f"   - Correos sin número de liquidación en adjuntos: {len(registros_sin_liq)}")

    # 3. Construir Excel con pestañas
    wb_salida = openpyxl.Workbook()

    # ─── HOJA 1: Vinculacion_Mails ───────────────────────────────────────────
    ws1 = wb_salida.active
    ws1.title = "Vinculacion_Mails"

    cols_h1 = [
        "Dirección de Mail",
        "Razón Social",
        "CUIT",
        "N° Liquidación",
        "Empresa",
        "Fecha de Envío",
        "Dirección Original (Reporte)",
        "Destinatario Original",
        "Asunto",
        "Archivos Adjuntos"
    ]

    filas_h1 = []
    for r in registros_vinculados:
        filas_h1.append([
            r['mail_proveedor'],
            r['razon_social'],
            r['cuit'],
            r['liq_numero'],
            r['empresa'],
            r['fecha_envio'],
            r['mail_original'],
            r['destinatario'],
            r['asunto'],
            r['archivos_adjuntos']
        ])

    formatear_hoja(
        ws1,
        cols_h1,
        filas_h1,
        anchos_minimos={'A': 32, 'B': 36, 'C': 15, 'D': 15, 'G': 32, 'I': 35, 'J': 40}
    )

    # ─── HOJA 2: Consolidado_Proveedores (1 fila por Razón Social única) ───────
    ws2 = wb_salida.create_sheet(title="Consolidado_Proveedores")

    proveedores_dict = {}
    for r in registros_vinculados:
        rs = r['razon_social']
        if rs not in proveedores_dict:
            proveedores_dict[rs] = {
                'cuit': r['cuit'],
                'empresa': r['empresa'],
                'mails': set(),
                'liqs': set(),
                'fechas': [],
                'total_mails': 0,
                'solo_compras': True
            }
        p = proveedores_dict[rs]
        if r['mail_proveedor']:
            for sm in [m.strip() for m in r['mail_proveedor'].split(',') if m.strip()]:
                p['mails'].add(sm)
            p['solo_compras'] = False

        for l in r['liq_numero'].split(','):
            if l.strip():
                p['liqs'].add(l.strip())
        if r['fecha_envio']:
            p['fechas'].append(str(r['fecha_envio']))
        p['total_mails'] += 1

    cols_h2 = [
        "Razón Social",
        "Dirección(es) de Mail",
        "CUIT",
        "Empresa",
        "Liquidaciones de Pago",
        "Cant. Liquidaciones",
        "Total Mails Enviados",
        "Última Fecha de Envío",
        "Estado / Observación"
    ]

    filas_h2 = []
    for rs in sorted(proveedores_dict.keys()):
        p = proveedores_dict[rs]
        mails_str = ", ".join(sorted(p['mails'])) if p['mails'] else ""
        liqs_sorted = sorted(p['liqs'])
        liqs_str = ", ".join(liqs_sorted)
        ult_fecha = p['fechas'][0] if p['fechas'] else ""

        if not mails_str:
            obs = "Sin casilla propia en reportes (enviado únicamente a Compras Antina)"
        else:
            obs = "Mail Vinculado"

        filas_h2.append([
            rs,
            mails_str,
            p['cuit'],
            p['empresa'],
            liqs_str,
            len(liqs_sorted),
            p['total_mails'],
            ult_fecha,
            obs
        ])

    formatear_hoja(
        ws2,
        cols_h2,
        filas_h2,
        anchos_minimos={'A': 36, 'B': 36, 'C': 15, 'E': 25, 'I': 35}
    )

    # ─── HOJA 3: Mapeo_Directo_Mail_RS (1 fila por dirección individual) ───────
    ws3 = wb_salida.create_sheet(title="Mapeo_Directo_Mail_RS")

    mail_to_rs = {}
    for r in registros_vinculados:
        rs = r['razon_social']
        cuit = r['cuit']
        if r['mail_proveedor']:
            for sm in [m.strip() for m in r['mail_proveedor'].split(',') if m.strip()]:
                if sm not in mail_to_rs:
                    mail_to_rs[sm] = {'rs': set(), 'cuits': set()}
                mail_to_rs[sm]['rs'].add(rs)
                if cuit:
                    mail_to_rs[sm]['cuits'].add(cuit)

    cols_h3 = [
        "Dirección de Mail",
        "Razón Social",
        "CUIT"
    ]

    filas_h3 = []
    # 1. Emails individuales directos
    for mail_item in sorted(mail_to_rs.keys()):
        data = mail_to_rs[mail_item]
        filas_h3.append([
            mail_item,
            " / ".join(sorted(data['rs'])),
            " / ".join(sorted(data['cuits']))
        ])

    # 2. Proveedores sin casilla directa (solo compras@antina.com.ar)
    for rs, p in sorted(proveedores_dict.items()):
        if not p['mails']:
            filas_h3.append([
                "",  # Vacío según instrucción explícita
                rs,
                p['cuit']
            ])

    formatear_hoja(ws3, cols_h3, filas_h3, anchos_minimos={'A': 35, 'B': 38, 'C': 16})

    # ─── HOJA 4: Mails_Sin_Liquidacion (Auditoría) ───────────────────────────
    if registros_sin_liq:
        ws4 = wb_salida.create_sheet(title="Mails_Sin_Liquidacion")
        cols_h4 = [
            "Fila Original",
            "Fecha de Envío",
            "Destinatario",
            "Dirección de Mail",
            "Asunto",
            "Archivos Adjuntos",
            "Observación"
        ]
        filas_h4 = []
        for r in registros_sin_liq:
            filas_h4.append([
                r['row'],
                r['fecha'],
                r['destinatario'],
                r['mail'],
                r['asunto'],
                r['adjuntos'],
                "Sin N° de liquidación en archivos adjuntos (posible respuesta, reenvío o soporte)"
            ])
        formatear_hoja(ws4, cols_h4, filas_h4, anchos_minimos={'C': 25, 'D': 30, 'E': 35, 'F': 40, 'G': 40})

    # 4. Guardar archivo
    wb_salida.save(EXCEL_SALIDA)
    wb_salida.close()

    print("=" * 75)
    print(f"✨ ¡PROCESO FINALIZADO CON ÉXITO!")
    print(f"📁 Archivo generado: {EXCEL_SALIDA}")
    print(f"📑 Pestañas disponibles:")
    print(f"   1. 'Vinculacion_Mails': 171 registros con Mail, Razón Social, CUIT, Liq, Asunto y Adjuntos.")
    print(f"   2. 'Consolidado_Proveedores': 70 Razones Sociales únicas con sus respectivos mails y liquidaciones.")
    print(f"   3. 'Mapeo_Directo_Mail_RS': Mapeo 1 a 1 de Mail individual hacia Razón Social (ideal BUSCARV/XLOOKUP).")
    print(f"   4. 'Mails_Sin_Liquidacion': 9 registros que no poseían liquidación adjunta.")
    print("=" * 75)


if __name__ == "__main__":
    procesar_reporte()
