# downloader.py
import os
import win32com.client
from datetime import datetime
from config import CARPETA_A

def descargar_facturas_outlook(fecha_inicio="", filtro_proveedor=""):
    print(f"🔄 Conectando a Outlook (Carpeta: Compras)...")
    
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    try:
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

    # Diccionario para mapear qué archivo PDF vino de qué correo (ej: {"factura.pdf": "facturacion@ecosan.com.ar"})
    mapa_remitentes = {}

    print(f"📥 Buscando correos en Outlook y descargando adjuntos a: {CARPETA_A}\n")

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
                for adj in msg.Attachments:
                    nombre_archivo = adj.FileName.lower()
                    if nombre_archivo.endswith('.pdf'):
                        ruta_destino = os.path.join(CARPETA_A, adj.FileName)
                        
                        if os.path.exists(ruta_destino):
                            base, ext = os.path.splitext(adj.FileName)
                            c = 1
                            while os.path.exists(os.path.join(CARPETA_A, f"{base} ({c}){ext}")):
                                c += 1
                            ruta_destino = os.path.join(CARPETA_A, f"{base} ({c}){ext}")
                            
                        adj.SaveAsFile(ruta_destino)
                        nombre_final_guardado = os.path.basename(ruta_destino)
                        mapa_remitentes[nombre_final_guardado] = email
                        
                        print(f"[📥 Descargado] {nombre_final_guardado} (Remitente: {email})")
                        descargados += 1
        except Exception:
            continue
            
    print(f"\n✅ ¡Descarga finalizada! Se obtuvieron {descargados} PDFs nuevos.")
    return mapa_remitentes