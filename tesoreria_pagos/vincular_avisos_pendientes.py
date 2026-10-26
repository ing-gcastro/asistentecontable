import os
import re
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Configurar salida UTF-8 para consola Windows
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

# ─── RUTAS DE ENTRADA Y SALIDA ───────────────────────────────────────────────
EXCEL_PENDIENTES = r"\\10.10.10.210\AyF_Trabajoadistancia\Facturas a pagar\AUTOMATIZACIONES IA\reporte_avisos_pago_pendientes.xlsx"
HTML_LEGAJOS = r"\\10.10.10.210\AyF_Trabajoadistancia\Facturas a pagar\Reporte_Legajos_Pago.html"

# Salidas específicas del reporte de pendientes
SALIDA_PENDIENTES_RED = r"\\10.10.10.210\AyF_Trabajoadistancia\Facturas a pagar\AUTOMATIZACIONES IA\vinculacion_mails_razon_social_pendientes.xlsx"
SALIDA_PENDIENTES_LOCAL = os.path.join(os.path.expanduser("~"), r"ROBOTS\data\vinculacion_mails_razon_social_pendientes.xlsx")

# Archivo Maestro Consolidado (histórico + pendientes)
SALIDA_MASTER_RED = r"\\10.10.10.210\AyF_Trabajoadistancia\Facturas a pagar\AUTOMATIZACIONES IA\vinculacion_mails_razon_social.xlsx"
SALIDA_MASTER_LOCAL = os.path.join(os.path.expanduser("~"), r"ROBOTS\data\vinculacion_mails_razon_social.xlsx")

# Casillas internas de Antina que NO corresponden a la razón social de un proveedor
GENERIC_ANTINA_EMAILS = {
    "compras@antina.com.ar",
    "proveedores@antina.com.ar"
}


def limpiar_emails_proveedor(mail_str):
    """
    Limpia y extrae las direcciones de email reales del proveedor:
    - Si la casilla es únicamente compras@antina.com.ar o proveedores@antina.com.ar -> retorna "" (vacío).
    - Si vienen múltiples casillas que incluyen buzones internos de Antina, se filtran
      para dejar únicamente las casillas externas del proveedor.
    """
    if not mail_str:
        return ""
    
    mail_limpio = mail_str.strip()
    if mail_limpio.lower() in GENERIC_ANTINA_EMAILS:
        return ""

    partes = [m.strip() for m in re.split(r'[,;]', mail_limpio) if m.strip()]
    externos = [p for p in partes if p.lower() not in GENERIC_ANTINA_EMAILS]
    
    if externos:
        return ", ".join(externos)
    
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

    # Caso especial conocido de liquidación Eurosat con nombre en archivo
    if '8306' not in mapa_liq:
        mapa_liq['8306'] = {
            'proveedor': 'BECHER Y ASOCIADOS SRL',
            'cuit': '30659199811',
            'empresa': 'EUROSAT',
            'fecha_pago': '14/05/2026'
        }

    print(f"✅ HTML procesado con éxito: {len(mapa_liq)} liquidaciones registradas en base.")
    return mapa_liq


def formatear_hoja(ws, titulo_columnas, filas_datos, anchos_minimos=None):
    """
    Aplica diseño profesional al estilo dashboard corporativo a la hoja de cálculo.
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
        ws.column_dimensions[col_letter].width = max(min_w, min(max_len + 4, 65))


def extraer_datos_fila(row, idx_mail, idx_adj, idx_fecha, idx_asunto, idx_dest):
    """Extrae valores de fila considerando índices dinámicos."""
    mail = row[idx_mail] if idx_mail is not None and idx_mail < len(row) else ""
    adj = row[idx_adj] if idx_adj is not None and idx_adj < len(row) else ""
    fecha = row[idx_fecha] if idx_fecha is not None and idx_fecha < len(row) else ""
    asunto = row[idx_asunto] if idx_asunto is not None and idx_asunto < len(row) else ""
    dest = row[idx_dest] if idx_dest is not None and idx_dest < len(row) else ""
    return str(mail or '').strip(), str(adj or '').strip(), fecha, str(asunto or '').strip(), str(dest or '').strip()


def resolver_indices_columnas(headers):
    """
    Detecta automáticamente las columnas según los nombres de encabezado,
    con fallback posicional si no coincide.
    """
    idx_mail = None
    idx_adj = None
    idx_fecha = None
    idx_asunto = None
    idx_dest = None

    for i, h in enumerate(headers):
        h_str = str(h or '').lower()
        if "mail" in h_str or "correo" in h_str:
            if idx_mail is None: idx_mail = i
        elif "adjunto" in h_str or "archivo" in h_str:
            if idx_adj is None: idx_adj = i
        elif "fecha" in h_str or "senton" in h_str:
            if idx_fecha is None: idx_fecha = i
        elif "asunto" in h_str or "subject" in h_str:
            if idx_asunto is None: idx_asunto = i
        elif "destinatario" in h_str or "contacto" in h_str or "display" in h_str:
            if idx_dest is None: idx_dest = i

    # Fallback predeterminado para 'reporte_avisos_pago_pendientes.xlsx'
    if idx_mail is None: idx_mail = 0
    if idx_adj is None: idx_adj = 1
    if idx_fecha is None: idx_fecha = 2
    if idx_asunto is None: idx_asunto = 3
    if idx_dest is None: idx_dest = 4

    return idx_mail, idx_adj, idx_fecha, idx_asunto, idx_dest


def procesar_archivo_excel(ruta_excel, mapa_liq):
    """
    Lee el archivo de avisos de pago y lo vincula con el mapa de liquidaciones HTML.
    """
    if not os.path.exists(ruta_excel):
        raise FileNotFoundError(f"No se encontró el archivo: {ruta_excel}")

    wb_origen = openpyxl.load_workbook(ruta_excel, data_only=True)
    ws_origen = wb_origen.active

    rows = list(ws_origen.iter_rows(values_only=True))
    wb_origen.close()

    if not rows:
        return [], []

    headers = rows[0]
    idx_mail, idx_adj, idx_fecha, idx_asunto, idx_dest = resolver_indices_columnas(headers)

    registros_vinculados = []
    registros_sin_liq = []

    for idx, row in enumerate(rows[1:], start=2):
        mail_str, adj_str, fecha, asunto_str, dest_str = extraer_datos_fila(
            row, idx_mail, idx_adj, idx_fecha, idx_asunto, idx_dest
        )

        # Buscar liquidaciones:
        # 1. 5 dígitos entre 60000 y 79999 (Telcom)
        liqs_5d = list(dict.fromkeys(re.findall(r'\b(6\d{4}|7\d{4})\b', adj_str)))
        # 2. 4 dígitos entre 8000 y 8999 (Eurosat)
        liqs_4d = list(dict.fromkeys(re.findall(r'\b(8\d{3})\b', adj_str)))

        todas_liqs = liqs_5d + liqs_4d

        if not todas_liqs:
            # Observación sobre posible razón social deducida del asunto para control
            registros_sin_liq.append({
                'row': idx,
                'fecha': fecha,
                'destinatario': dest_str,
                'mail': mail_str,
                'asunto': asunto_str,
                'adjuntos': adj_str
            })
            continue

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

        razon_social_str = " / ".join(nombres_prov) if nombres_prov else "NO ENCONTRADO EN HTML"
        cuit_str = " / ".join(cuits_prov) if cuits_prov else ""
        empresa_str = " / ".join(empresas_prov) if empresas_prov else ""

        mail_proveedor_limpio = limpiar_emails_proveedor(mail_str)

        registros_vinculados.append({
            'row_orig': idx,
            'fecha_envio': fecha,
            'destinatario': dest_str,
            'asunto': asunto_str,
            'mail_proveedor': mail_proveedor_limpio,
            'mail_original': mail_str,
            'razon_social': razon_social_str,
            'cuit': cuit_str,
            'liq_numero': ", ".join(todas_liqs),
            'empresa': empresa_str,
            'archivos_adjuntos': adj_str
        })

    return registros_vinculados, registros_sin_liq


def generar_libro_vinculacion(registros_vinculados, registros_sin_liq, ruta_destino):
    """
    Construye y guarda el libro Excel profesional con 4 hojas clave.
    """
    wb_salida = openpyxl.Workbook()

    # ─── HOJA 1: Vinculacion_Mails (Detalle completo) ─────────────────────────
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
        anchos_minimos={'A': 32, 'B': 36, 'C': 15, 'D': 15, 'G': 32, 'I': 35, 'J': 45}
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
                'total_mails': 0
            }
        p = proveedores_dict[rs]
        if r['mail_proveedor']:
            for sm in [m.strip() for m in r['mail_proveedor'].split(',') if m.strip()]:
                p['mails'].add(sm)

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
            obs = "Sin casilla externa en reportes (enviado a buzón interno Antina)"
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
    # 1. Mails individuales directos
    for mail_item in sorted(mail_to_rs.keys()):
        data = mail_to_rs[mail_item]
        filas_h3.append([
            mail_item,
            " / ".join(sorted(data['rs'])),
            " / ".join(sorted(data['cuits']))
        ])

    # 2. Proveedores sin casilla directa
    for rs, p in sorted(proveedores_dict.items()):
        if not p['mails']:
            filas_h3.append([
                "",
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
                "Sin N° de liquidación en archivos adjuntos (reenvío, soporte o comprobante externo)"
            ])
        formatear_hoja(ws4, cols_h4, filas_h4, anchos_minimos={'C': 25, 'D': 30, 'E': 35, 'F': 40, 'G': 40})

    os.makedirs(os.path.dirname(ruta_destino), exist_ok=True)
    wb_salida.save(ruta_destino)
    wb_salida.close()
    print(f" -> Guardado exitoso: {ruta_destino}")


def fusionar_con_maestro(registros_nuevos, registros_nuevos_sin_liq, mapa_liq):
    """
    Carga el archivo maestro histórico (si existe) y le suma los registros nuevos,
    evitando duplicar exactamente los mismos avisos de pago.
    """
    registros_historicos = []
    registros_hist_sin_liq = []

    if os.path.exists(SALIDA_MASTER_RED):
        ruta_maestro = SALIDA_MASTER_RED
    elif os.path.exists(SALIDA_MASTER_LOCAL):
        ruta_maestro = SALIDA_MASTER_LOCAL
    else:
        ruta_maestro = None

    if ruta_maestro:
        print(f"\n📂 Leyendo maestro histórico existente: {ruta_maestro}")
        wb = openpyxl.load_workbook(ruta_maestro, data_only=True)
        if "Vinculacion_Mails" in wb.sheetnames:
            ws = wb["Vinculacion_Mails"]
            for row in ws.iter_rows(min_row=2, values_only=True):
                # Cols: Mail, Razon Social, CUIT, Liq, Empresa, Fecha, Mail Orig, Dest Orig, Asunto, Adj
                if any(row):
                    registros_historicos.append({
                        'mail_proveedor': str(row[0] or '').strip(),
                        'razon_social': str(row[1] or '').strip(),
                        'cuit': str(row[2] or '').strip(),
                        'liq_numero': str(row[3] or '').strip(),
                        'empresa': str(row[4] or '').strip(),
                        'fecha_envio': row[5],
                        'mail_original': str(row[6] or '').strip(),
                        'destinatario': str(row[7] or '').strip(),
                        'asunto': str(row[8] or '').strip(),
                        'archivos_adjuntos': str(row[9] or '').strip(),
                        'row_orig': 0
                    })
        wb.close()

    # Combinar sin duplicados por clave (Fecha + Asunto + Liq)
    claves_existentes = set()
    for r in registros_historicos:
        k = (str(r['fecha_envio']), str(r['asunto']).upper(), str(r['liq_numero']))
        claves_existentes.add(k)

    agregados = 0
    for r in registros_nuevos:
        k = (str(r['fecha_envio']), str(r['asunto']).upper(), str(r['liq_numero']))
        if k not in claves_existentes:
            registros_historicos.append(r)
            claves_existentes.add(k)
            agregados += 1

    print(f"📊 Registros consolidados en Maestro: {len(registros_historicos)} ({agregados} nuevos agregados)")

    for destino in [SALIDA_MASTER_RED, SALIDA_MASTER_LOCAL]:
        try:
            generar_libro_vinculacion(registros_historicos, registros_hist_sin_liq, destino)
        except Exception as e:
            print(f" [!] Error al guardar maestro en {destino}: {e}")


def main():
    print("=" * 75)
    print("🚀 PROCESANDO REPORTE DE AVISOS DE PAGO PENDIENTES")
    print("=" * 75)

    # 1. Cargar mapeo HTML
    mapa_liq = cargar_mapeo_html(HTML_LEGAJOS)

    # 2. Procesar Excel de pendientes
    print(f"\n📂 Leyendo archivo de pendientes: {EXCEL_PENDIENTES}")
    registros_vinculados, registros_sin_liq = procesar_archivo_excel(EXCEL_PENDIENTES, mapa_liq)

    print(f"\n📊 Resultados del archivo de pendientes:")
    print(f"   - Total avisos procesados: {len(registros_vinculados) + len(registros_sin_liq)}")
    print(f"   - Vinculados exitosamente con Razón Social: {len(registros_vinculados)}")
    print(f"   - Sin N° de liquidación en adjuntos: {len(registros_sin_liq)}")

    # 3. Guardar archivo específico de pendientes en RED y LOCAL
    print(f"\n💾 Generando archivo de resultados para pendientes:")
    for dest in [SALIDA_PENDIENTES_RED, SALIDA_PENDIENTES_LOCAL]:
        try:
            generar_libro_vinculacion(registros_vinculados, registros_sin_liq, dest)
        except Exception as e:
            print(f" [!] Error al guardar en {dest}: {e}")

    # 4. Actualizar el archivo Maestro Consolidado (histórico + pendientes)
    print(f"\n🔄 Actualizando archivo Maestro de vinculación general:")
    fusionar_con_maestro(registros_vinculados, registros_sin_liq, mapa_liq)

    print("\n" + "=" * 75)
    print("✨ ¡PROCESO DE VINCULACIÓN COMPLETADO CON ÉXITO!")
    print("=" * 75)


if __name__ == "__main__":
    main()
