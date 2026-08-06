import os

new_code = r"""# downloader.py
import os
import win32com.client
from datetime import datetime
import fitz  # PyMuPDF
from config import CARPETA_A

def obtener_ruta_unica(ruta):
    if not os.path.exists(ruta): return ruta
    nombre, ext = os.path.splitext(ruta)
    c = 1
    while os.path.exists(f"{nombre} ({c}){ext}"): c += 1
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
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(excel_path)
        wb.ExportAsFixedFormat(0, pdf_path)
        wb.Close(False)
        excel.Quit()
        return True
    except Exception as e:
        print(f"⚠️ Error convirtiendo Excel a PDF {excel_path}: {e}")
        try:
            excel.Quit()
        except:
            pass
        return False

def descargar_facturas_outlook(fecha_inicio="", filtro_proveedor=""):
    print(f"🔄 Conectando a Outlook (Carpeta: Compras)...")
    
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        carpeta = outlook.Folders["Compras"]
    except Exception as e:
        print(f"❌ ERROR: No se encontró la carpeta 'Compras' en Outlook. Detalles: {e}")
        return {}
        
    mensajes = carpeta.Items
    mensajes.Sort("[ReceivedTime]", True)
    
    if not os.path.exists(CARPETA_A):
        os.makedirs(CARPETA_A, exist_ok=True)
        
    descargados = 0
    f_ini = datetime.strptime(fecha_inicio, "%d/%m/%Y").date() if fecha_inicio else None
    f_correo = filtro_proveedor.strip().lower()

    mapa_remitentes = {}
    print(f"📥 Buscando correos en Outlook y descargando adjuntos a: {CARPETA_A}\n")

    # Carpeta temporal para procesamiento
    temp_dir = os.path.join(CARPETA_A, "temp_descargas")
    os.makedirs(temp_dir, exist_ok=True)

    for msg in mensajes:
        try:
            if msg.Class != 43: continue 
            
            fecha_msg = msg.ReceivedTime.date()
            if f_ini and fecha_msg < f_ini:
                break 
                
            email = (msg.SenderEmailAddress if msg.SenderEmailType != "EX" else (msg.Sender.GetExchangeUser().PrimarySmtpAddress if msg.Sender.GetExchangeUser() else "desconocido")).lower()
            
            if f_correo and f_correo not in email:
                continue
                
            if msg.Attachments.Count > 0:
                adjuntos_pdf = []
                adjuntos_excel = []
                
                # Guardar adjuntos temporalmente
                for adj in msg.Attachments:
                    nombre = adj.FileName.lower()
                    if nombre.endswith('.pdf'):
                        r = os.path.join(temp_dir, adj.FileName)
                        adj.SaveAsFile(r)
                        adjuntos_pdf.append(r)
                    elif nombre.endswith('.xls') or nombre.endswith('.xlsx'):
                        r = os.path.join(temp_dir, adj.FileName)
                        adj.SaveAsFile(r)
                        adjuntos_excel.append(r)

                if not adjuntos_pdf and not adjuntos_excel:
                    continue

                # REGLA MARINO: Unir múltiples PDFs
                if "facturacion@marinosa.com.ar" in email and len(adjuntos_pdf) > 0:
                    # Buscar el que tiene FAC para usar de nombre base
                    base_pdf = adjuntos_pdf[0]
                    for p in adjuntos_pdf:
                        if "fac" in os.path.basename(p).lower():
                            base_pdf = p
                            break
                            
                    nombre_final = os.path.basename(base_pdf)
                    ruta_final = obtener_ruta_unica(os.path.join(CARPETA_A, nombre_final))
                    
                    if len(adjuntos_pdf) > 1:
                        fusionar_pdfs(adjuntos_pdf, ruta_final)
                        print(f"[📥 Descargado/Fusionado MARINO] {nombre_final}")
                    else:
                        os.rename(adjuntos_pdf[0], ruta_final)
                        print(f"[📥 Descargado MARINO] {nombre_final}")
                        
                    mapa_remitentes[os.path.basename(ruta_final)] = email
                    descargados += 1

                # REGLA REDGUARD / AQUALINE: Excel a PDF y fusionar
                elif ("facturacion@redguard.com.ar" in email or "facturacion@aqualine.com.ar" in email) and len(adjuntos_pdf) > 0:
                    base_pdf = adjuntos_pdf[0]
                    pdfs_a_fusionar = [base_pdf]
                    
                    # Convertir Excel
                    for exc in adjuntos_excel:
                        pdf_gen = exc + ".pdf"
                        if excel_a_pdf(exc, pdf_gen):
                            pdfs_a_fusionar.append(pdf_gen)
                            
                    nombre_final = os.path.basename(base_pdf)
                    ruta_final = obtener_ruta_unica(os.path.join(CARPETA_A, nombre_final))
                    
                    if len(pdfs_a_fusionar) > 1:
                        fusionar_pdfs(pdfs_a_fusionar, ruta_final)
                        print(f"[📥 Descargado/Fusionado con Excel] {nombre_final}")
                    else:
                        os.rename(base_pdf, ruta_final)
                        print(f"[📥 Descargado] {nombre_final}")
                        
                    mapa_remitentes[os.path.basename(ruta_final)] = email
                    descargados += 1

                # DESCARGA NORMAL
                else:
                    for p in adjuntos_pdf:
                        nombre_final = os.path.basename(p)
                        ruta_final = obtener_ruta_unica(os.path.join(CARPETA_A, nombre_final))
                        os.rename(p, ruta_final)
                        mapa_remitentes[os.path.basename(ruta_final)] = email
                        print(f"[📥 Descargado] {nombre_final}")
                        descargados += 1
                        
                # Limpiar temporales
                for f in os.listdir(temp_dir):
                    try:
                        os.remove(os.path.join(temp_dir, f))
                    except:
                        pass
        except Exception:
            continue
            
    # Borrar carpeta temporal si está vacía
    try:
        os.rmdir(temp_dir)
    except:
        pass
        
    print(f"\n✅ ¡Descarga finalizada! Se obtuvieron {descargados} PDFs nuevos.")
    return mapa_remitentes
"""

with open(r"c:\Users\gcastro\ROBOTS\asistentecontable\Proyecto_facturas_v2\downloader.py", "w", encoding="utf-8") as f:
    f.write(new_code)
