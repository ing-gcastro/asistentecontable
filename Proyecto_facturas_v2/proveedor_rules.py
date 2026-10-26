# proveedor_rules.py
try:
    import pymupdf as fitz  # PyMuPDF >= 1.24
except ImportError:  # versiones viejas
    import fitz

# Remitentes cuyas facturas se aceptan aunque no se detecte el CAE en el texto.
# OJO (BUG-02): esto SOLO relaja la exigencia del CAE. Un documento que dice
# explícitamente "no válido como factura" o que es un recibo se rechaza igual.
REMITENTES_CONFIABLES = [
    "facturacion@ecosan.com.ar",
    "facturacion@marinosa.com.ar",
    "administracion@contributionsrl.com.ar",
    "contributionsrl",  # Comodín parcial por si el dominio varía
]


def es_factura_por_remitente(email_remitente):
    """
    Indica si el remitente es de confianza (se acepta sin CAE detectable).
    No alcanza por sí solo para aprobar: ver procesador_core.es_factura_valida().
    """
    if not email_remitente:
        return False

    email_lower = email_remitente.strip().lower()
    return any(rem in email_lower for rem in REMITENTES_CONFIABLES)


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
