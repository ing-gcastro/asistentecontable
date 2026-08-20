import os
import sys
import argparse
from datetime import datetime, timezone, timedelta
import win32com.client
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def resolver_email_destinatario(recipient):
    """
    Obtiene la dirección de correo electrónico SMTP real de un destinatario,
    incluso si es un usuario o lista de distribución interna de Exchange.
    """
    # 1. Intentar obtener la propiedad MAPI PR_SMTP_ADDRESS directamente
    try:
        prop_smtp = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
        smtp = recipient.PropertyAccessor.GetProperty(prop_smtp)
        if smtp and "@" in smtp:
            return smtp.strip()
    except Exception:
        pass

    # 2. Intentar a través de AddressEntry (Exchange User o Distribution List)
    try:
        ae = recipient.AddressEntry
        if ae:
            if ae.Type == "EX":
                ex_user = ae.GetExchangeUser()
                if ex_user and ex_user.PrimarySmtpAddress:
                    return ex_user.PrimarySmtpAddress.strip()
                ex_dl = ae.GetExchangeDistributionList()
                if ex_dl and ex_dl.PrimarySmtpAddress:
                    return ex_dl.PrimarySmtpAddress.strip()
            elif ae.Address and "@" in ae.Address:
                return ae.Address.strip()
    except Exception:
        pass

    # 3. Fallback a la propiedad Address o al Nombre
    addr = recipient.Address or ""
    if "@" in addr:
        return addr.strip()
    return recipient.Name or "Desconocido"

def extraer_mails_enviados(fecha_desde=None, fecha_hasta=None, max_items=None):
    """
    Lee la carpeta de 'Elementos enviados' de Outlook y extrae los datos requeridos.
    """
    print("Conectando a Microsoft Outlook...")
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        # 5 = olFolderSentMail (Elementos enviados)
        sent_folder = outlook.GetDefaultFolder(5)
    except Exception as e:
        print(f"Error al conectar con Outlook: {e}")
        return []

    items = sent_folder.Items
    items.Sort("[SentOn]", True)  # Ordenar de más reciente a más antiguo

    total_disponibles = items.Count
    print(f"Buscando en '{sent_folder.Name}' ({total_disponibles} elementos totales)...")

    registros = []
    procesados = 0

    for msg in items:
        try:
            # 43 = olMail (MailItem)
            if msg.Class != 43:
                continue

            # Fecha de envío
            sent_on = msg.SentOn
            # Normalizar timezone para comparar fechas
            fecha_item = sent_on.date() if hasattr(sent_on, "date") else None

            if fecha_desde and fecha_item and fecha_item < fecha_desde:
                # Como están ordenados descendentemente, si es menor a fecha_desde ya no habrá más dentro del rango
                break

            if fecha_hasta and fecha_item and fecha_item > fecha_hasta:
                continue

            # 1. Fecha formateada
            fecha_str = sent_on.strftime("%d/%m/%Y %H:%M:%S") if sent_on else "Sin Fecha"

            # 2. Asunto
            asunto = msg.Subject or "(Sin Asunto)"

            # 3. Mail Destinatarios (To, CC, BCC)
            destinatarios_lista = []
            try:
                for r in msg.Recipients:
                    email_dest = resolver_email_destinatario(r)
                    if email_dest not in destinatarios_lista:
                        destinatarios_lista.append(email_dest)
            except Exception:
                pass
            destinatarios_str = ", ".join(destinatarios_lista) if destinatarios_lista else (msg.To or "")

            # 4. Cuerpo del mail
            cuerpo = msg.Body or ""
            # Límite de caracteres en celda de Excel es 32,767
            if len(cuerpo) > 32000:
                cuerpo = cuerpo[:32000] + "... [TRUNCADO]"

            # 5. Nombre de archivos adjuntos
            adjuntos_lista = []
            try:
                for att in msg.Attachments:
                    adjuntos_lista.append(att.FileName)
            except Exception:
                pass
            adjuntos_str = ", ".join(adjuntos_lista) if adjuntos_lista else "Sin adjuntos"

            registros.append({
                "fecha": fecha_str,
                "asunto": asunto,
                "destinatarios": destinatarios_str,
                "cuerpo": cuerpo,
                "adjuntos": adjuntos_str
            })

            procesados += 1
            if procesados % 25 == 0:
                print(f"  -> Procesados {procesados} correos...")

            if max_items and procesados >= max_items:
                break

        except Exception as err:
            # Continuar en caso de algún correo dañado
            continue

    print(f"Extracción finalizada: {len(registros)} correos recopilados.\n")
    return registros

def guardar_en_excel(registros, ruta_salida):
    """
    Genera un archivo Excel con diseño profesional y estilizado.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Mails Enviados"

    # Encabezados solicitados
    headers = [
        "Fecha",
        "Asunto",
        "Mail Destinatarios",
        "Cuerpo del Mail",
        "Nombre de los Archivos Adjuntos"
    ]
    ws.append(headers)

    # Estilos
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    zebra_fill = PatternFill(start_color="F2F4F7", end_color="F2F4F7", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    thin_border_side = Side(style="thin", color="D9D9D9")
    cell_border = Border(
        left=thin_border_side,
        right=thin_border_side,
        top=thin_border_side,
        bottom=thin_border_side
    )

    # Aplicar estilos a encabezados
    for col_num, _ in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = cell_border
    ws.row_dimensions[1].height = 28

    # Insertar filas
    for row_idx, item in enumerate(registros, start=2):
        row_data = [
            item["fecha"],
            item["asunto"],
            item["destinatarios"],
            item["cuerpo"],
            item["adjuntos"]
        ]
        ws.append(row_data)
        
        # Color alterno para filas (zebra)
        row_fill = zebra_fill if row_idx % 2 == 0 else white_fill

        for col_idx in range(1, 6):
            c = ws.cell(row=row_idx, column=col_idx)
            c.fill = row_fill
            c.border = cell_border
            c.font = Font(name="Calibri", size=10)

            # Alineación según columna
            if col_idx == 1: # Fecha
                c.alignment = Alignment(horizontal="center", vertical="top")
            elif col_idx in (2, 3, 5): # Asunto, Destinatarios, Adjuntos
                c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            elif col_idx == 4: # Cuerpo
                c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

    # Configurar anchos de columna recomendados
    column_widths = {
        "A": 20, # Fecha
        "B": 35, # Asunto
        "C": 35, # Destinatarios
        "D": 60, # Cuerpo del Mail
        "E": 35  # Adjuntos
    }
    for col_letter, width in column_widths.items():
        ws.column_dimensions[col_letter].width = width

    # Habilitar autofiltros
    ws.auto_filter.ref = f"A1:E{max(len(registros) + 1, 2)}"
    
    # Congelar panel superior
    ws.freeze_panes = "A2"

    # Guardar archivo
    os.makedirs(os.path.dirname(os.path.abspath(ruta_salida)), exist_ok=True)
    wb.save(ruta_salida)
    print(f"Excel generado exitosamente en:\n  {ruta_salida}")

def main():
    parser = argparse.ArgumentParser(description="Exportar correos enviados desde Outlook a Excel.")
    parser.add_argument("--dias", type=int, default=None, help="Número de días hacia atrás a exportar (ej: 30). Por defecto exporta todos.")
    parser.add_argument("--desde", type=str, default=None, help="Fecha inicio en formato DD/MM/YYYY")
    parser.add_argument("--hasta", type=str, default=None, help="Fecha fin en formato DD/MM/YYYY")
    parser.add_argument("--limite", type=int, default=None, help="Cantidad máxima de correos a exportar")
    parser.add_argument("--salida", type=str, default=None, help="Ruta o nombre del archivo Excel de salida")
    parser.add_argument("--abrir", action="store_true", help="Abrir el archivo Excel automáticamente al finalizar")

    args = parser.parse_args()

    fecha_desde = None
    fecha_hasta = None

    if args.dias is not None and args.dias > 0:
        fecha_desde = (datetime.now() - timedelta(days=args.dias)).date()
    elif args.desde:
        try:
            fecha_desde = datetime.strptime(args.desde, "%d/%m/%Y").date()
        except ValueError:
            print("Error: Formato de fecha '--desde' inválido. Use DD/MM/YYYY (ej: 01/08/2026)")
            sys.exit(1)

    if args.hasta:
        try:
            fecha_hasta = datetime.strptime(args.hasta, "%d/%m/%Y").date()
        except ValueError:
            print("Error: Formato de fecha '--hasta' inválido. Use DD/MM/YYYY (ej: 14/08/2026)")
            sys.exit(1)

    # Determinar ruta de salida
    if args.salida:
        ruta_salida = args.salida
    else:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        directorio_actual = os.path.dirname(os.path.abspath(__file__))
        ruta_salida = os.path.join(directorio_actual, f"Reporte_Mails_Enviados_{timestamp}.xlsx")

    print("=" * 60)
    print(" EXPORTADOR DE CORREOS ENVIADOS (OUTLOOK -> EXCEL)")
    print("=" * 60)
    if fecha_desde:
        print(f"Filtro Desde: {fecha_desde.strftime('%d/%m/%Y')}")
    if fecha_hasta:
        print(f"Filtro Hasta: {fecha_hasta.strftime('%d/%m/%Y')}")
    print("-" * 60)

    registros = extraer_mails_enviados(fecha_desde, fecha_hasta, max_items=args.limite)

    if not registros:
        print("No se encontraron correos enviados para exportar con los criterios especificados.")
        return

    guardar_en_excel(registros, ruta_salida)

    if args.abrir or sys.stdin.isatty():
        try:
            os.startfile(ruta_salida)
        except Exception:
            pass

if __name__ == "__main__":
    main()
