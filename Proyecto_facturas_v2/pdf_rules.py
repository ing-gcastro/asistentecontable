# pdf_rules.py
REGLAS_EXTRACCION = {
    "empresa_propia": "TELCOM VENTURES DE ARGENTINA",
    "palabras_prohibidas": [
        "ANULADA", 
        "PRESUPUESTO", 
        "NOTA DE PEDIDO", 
        "DUPLICADO DE RECIBO"
    ],
    # Tolerancia total a saltos de línea y símbolos entre la palabra CAE y los 14-15 dígitos
    "rx_cae": r"C[.\s]*A[.\s]*E[.\s]*[:\s\|N°]*\n?\s*\|?\s*(\d{14,15})",
    "rx_nota_credito": r"NOTA\s*DE\s*CREDITO",
    "rx_nota_debito": r"NOTA\s*DE\s*DEBITO",
    "rx_numero_comp": r"(?:A|B|C)?[-]?\d{4}[-]?(\d{8})",
    "rx_numero_guion": r"N[°º]?\s*[:]?\s*(\d{4,8})",
    "rx_razon_social": r"^([A-ZÁÉÍÓÚÑ\s\.\,\-]+(?:S\.A\.|S\.R\.L\.|S\.C\.A\.|S\.A\.|SA|SRL))\b",
    "basura_afip": [
        "DOMICILIO FISCAL", "RESPONSABLE INSCRIPTO", "IVA", 
        "PUNTO DE VENTA", "FACTURA", "CUIT"
    ],
    "rx_cuit": r"C\.?U\.?I\.?T\.?[\s:\.\|]*\n?\s*\|?\s*((?:20|23|24|27|30|33|34)[-\s]?\d{2}[-\s]?\d{6}[-\s]?\d{1})",
    "rx_fecha_emision": r"Fecha\s*[:\.]?\s*(\d{2}[-/]\d{2}[-/]\d{4})",
    "cuit_eurosat": "33676549639",
    "nombre_eurosat": "EUROSAT"
}