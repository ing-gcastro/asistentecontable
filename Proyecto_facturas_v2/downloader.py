# downloader.py — Descarga de adjuntos de facturas desde Outlook (carpeta "Compras").
import os
import shutil
from datetime import datetime

try:
    import pymupdf as fitz  # PyMuPDF >= 1.24
except ImportError:  # versiones viejas
    import fitz

from config import CARPETA_A, ESTADO_EMAILS
from estado import agregar_remitentes, cargar_json, guardar_json

CARPETA_OUTLOOK = "Compras"
REMITENTES_IGNORADOS = ("gire",)
REMITENTES_FUSION_TODO = ("facturacion@marinosa.com.ar",)
REMITENTES_EXCEL_A_PDF = ("facturacion@redguard.com.ar", "facturacion@aqualine.com.ar")


def obtener_ruta_unica(ruta):
    if not os.path.exists(ruta):
        return ruta
    nombre, ext = os.path.splitext(ruta)
    c = 1
    while os.path.exists(f"{nombre} ({c}){ext}"):
        c += 1
    return f"{nombre} ({c}){ext}"


def fusionar_pdfs(rutas_entrada, ruta_salida):
    doc_salida = fitz.open()
    for r in rutas_entrada:
        try:
            doc_parcial = fitz.open(r)
            doc_salida.insert_pdf(doc_parcial)
            doc_parcial.close()
        except Exception as e:
            print(f"⚠️ Error fusionando {r}: {e}")
    doc_salida.save(ruta_salida)
    doc_salida.close()


def excel_a_pdf(excel_path, pdf_path):
    import win32com.client
    excel = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")  # instancia propia: no toca los Excel del usuario
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(os.path.abspath(excel_path))
        wb.ExportAsFixedFormat(0, os.path.abspath(pdf_path))
        wb.Close(False)
        return True
    except Exception as e:
        print(f"Error convirtiendo Excel a PDF {excel_path}: {e}")
        return False
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass


def cargar_procesados():
    return set(cargar_json(ESTADO_EMAILS, []))


def guardar_procesado(entry_id, procesados_set):
    procesados_set.add(entry_id)
    guardar_json(ESTADO_EMAILS, sorted(procesados_set))


def _email_remitente(msg):
    try:
        if msg.SenderEmailType == "EX":
            usuario = msg.Sender.GetExchangeUser()
            return (usuario.PrimarySmtpAddress if usuario else "desconocido").lower()
        return (msg.SenderEmailAddress or "desconocido").lower()
    except Exception:
        return "desconocido"


def _guardar_adjuntos(msg, temp_dir):
    pdfs, excels = [], []
    for adj in msg.Attachments:
        nombre = (adj.FileName or "").strip()
        ext = os.path.splitext(nombre)[1].lower()
        if ext not in (".pdf", ".xls", ".xlsx"):
            continue
        r = obtener_ruta_unica(os.path.join(temp_dir, nombre))  # dos adjuntos con el mismo nombre no se pisan
        adj.SaveAsFile(r)
        (pdfs if ext == ".pdf" else excels).append(r)
    return pdfs, excels


def _mover_a_entrada(origen, nombre_final):
    ruta_final = obtener_ruta_unica(os.path.join(CARPETA_A, nombre_final))
    shutil.move(origen, ruta_final)
    return ruta_final


def descargar_facturas_outlook(fecha_inicio="", filtro_proveedor=""):
    print(f"Conectando a Outlook (Carpeta: {CARPETA_OUTLOOK})...")
    try:
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        carpeta = outlook.Folders[CARPETA_OUTLOOK]
    except Exception as e:
        print(f"ERROR: No se encontro la carpeta '{CARPETA_OUTLOOK}' en Outlook. Detalles: {e}")
        return {}

    mensajes = carpeta.Items
    mensajes.Sort("[ReceivedTime]", True)

    os.makedirs(CARPETA_A, exist_ok=True)
    descargados = 0
    errores = 0
    f_ini = datetime.strptime(fecha_inicio, "%d/%m/%Y").date() if fecha_inicio else None
    f_correo = filtro_proveedor.strip().lower()

    mapa_remitentes = {}
    print(f"Buscando correos en Outlook y descargando adjuntos a: {CARPETA_A}\n")

    temp_dir = os.path.join(CARPETA_A, "temp_descargas")
    os.makedirs(temp_dir, exist_ok=True)
    procesados = cargar_procesados()

    try:
        for msg in mensajes:
            try:
                if msg.Class != 43:
                    continue
                if f_ini and msg.ReceivedTime.date() < f_ini:
                    break  # ordenados del más nuevo al más viejo

                email = _email_remitente(msg)
                if f_correo and f_correo not in email:
                    continue
                if any(ign in email for ign in REMITENTES_IGNORADOS):
                    continue

                asunto_limpio = msg.Subject.replace('\r', '').replace('\n', '')[:50] if msg.Subject else "sin_asunto"
                entry_id = f"{msg.ReceivedTime.strftime('%Y%m%d%H%M%S')}_{email}_{asunto_limpio}"
                if entry_id in procesados:
                    continue

                if msg.Attachments.Count == 0:
                    guardar_procesado(entry_id, procesados)
                    continue

                adjuntos_pdf, adjuntos_excel = _guardar_adjuntos(msg, temp_dir)
                if not adjuntos_pdf:
                    guardar_procesado(entry_id, procesados)
                    continue

                nuevos = []
                if any(r in email for r in REMITENTES_FUSION_TODO):
                    # REGLA MARINO: todos los PDF del mail en uno (la factura primero)
                    base_pdf = next((p for p in adjuntos_pdf if "fac" in os.path.basename(p).lower()), adjuntos_pdf[0])
                    orden = [base_pdf] + [p for p in adjuntos_pdf if p != base_pdf]
                    nombre_final = os.path.basename(base_pdf)
                    if len(orden) > 1:
                        ruta_final = obtener_ruta_unica(os.path.join(CARPETA_A, nombre_final))
                        fusionar_pdfs(orden, ruta_final)
                        print(f"[Descargado/Fusionado MARINO] {nombre_final}")
                    else:
                        ruta_final = _mover_a_entrada(base_pdf, nombre_final)
                        print(f"[Descargado MARINO] {nombre_final}")
                    nuevos.append(ruta_final)

                elif any(r in email for r in REMITENTES_EXCEL_A_PDF):
                    # REGLA REDGUARD / AQUALINE: cada PDF + los Excel convertidos van al primer PDF
                    pdfs_a_fusionar = list(adjuntos_pdf)
                    for exc in adjuntos_excel:
                        pdf_gen = exc + ".pdf"
                        if excel_a_pdf(exc, pdf_gen):
                            pdfs_a_fusionar.append(pdf_gen)
                    nombre_final = os.path.basename(adjuntos_pdf[0])
                    if len(pdfs_a_fusionar) > 1:
                        ruta_final = obtener_ruta_unica(os.path.join(CARPETA_A, nombre_final))
                        fusionar_pdfs(pdfs_a_fusionar, ruta_final)
                        print(f"[Descargado/Fusionado con Excel] {nombre_final}")
                    else:
                        ruta_final = _mover_a_entrada(adjuntos_pdf[0], nombre_final)
                        print(f"[Descargado] {nombre_final}")
                    nuevos.append(ruta_final)

                else:
                    for p in adjuntos_pdf:
                        nombre_final = os.path.basename(p)
                        nuevos.append(_mover_a_entrada(p, nombre_final))
                        print(f"[Descargado] {nombre_final}")

                for ruta in nuevos:
                    mapa_remitentes[os.path.basename(ruta)] = email
                descargados += len(nuevos)
                agregar_remitentes({os.path.basename(r): email for r in nuevos})
                guardar_procesado(entry_id, procesados)
            except Exception as e:
                # No se marca como procesado: se reintenta en la próxima corrida
                errores += 1
                try:
                    asunto = msg.Subject
                except Exception:
                    asunto = "?"
                print(f"⚠️ Error con el mail '{asunto}': {e}")
            finally:
                for f in os.listdir(temp_dir):
                    try:
                        os.remove(os.path.join(temp_dir, f))
                    except OSError:
                        pass
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

    print(f"\n¡Descarga finalizada! Se obtuvieron {descargados} PDFs nuevos." + (f" Mails con error: {errores}." if errores else ""))
    return mapa_remitentes
