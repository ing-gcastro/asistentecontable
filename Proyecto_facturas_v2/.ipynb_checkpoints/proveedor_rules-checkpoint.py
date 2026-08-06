# proveedor_rules.py
import fitz

def es_factura_por_remitente(email_remitente):
    """
    Define si por el correo del remitente debemos dar la factura por válida de forma automática.
    """
    if not email_remitente:
        return False
    
    remitentes_validos = [
        "facturacion@ecosan.com.ar",
        "facturacion@marinosa.com.ar",
        "administracion@contributionsrl.com.ar", # Por si acaso
        "contributionsrl" # Comodín parcial por si el dominio varía
    ]
    
    email_lower = email_remitente.strip().lower()
    for rem in remitentes_validos:
        if rem in email_lower:
            return True
            
    return False

def analizar_pdf_ecosan(ruta_pdf):
    """
    Regla específica para extraer datos detallados de Ecosan (Tipo y Locación).
    """
    tipo = "MANTENIMIENTO"
    locacion = "GENERAL"
    try:
        doc = fitz.open(ruta_pdf)
        texto = "\n".join([pagina.get_text("text", sort=True) for pagina in doc]).lower()
        doc.close()
        
        if "alquiler" in texto: 
            tipo = "ALQUILER"
        elif "mantenimiento" in texto or "servicio" in texto: 
            tipo = "MANTENIMIENTO"
            
        if "marcos paz y vidt" in texto or "mariano acosta" in texto or "la huella" in texto: 
            locacion = "MARIANO ACOSTA"
        elif "florida 9563" in texto or "ramal pilar" in texto or "pilar" in texto: 
            locacion = "PILAR"
            
        return tipo, locacion
    except Exception:
        return tipo, locacion