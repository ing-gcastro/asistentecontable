# pdf_rules.py — Reglas de extracción y clasificación de comprobantes.
# Todas las comparaciones de texto se hacen en minúsculas y SIN acentos
# (ver procesador_core.normalizar_texto), así que las listas van en ese formato.

# Monto con 2 decimales: 1.234.567,89 | 1234,89 | 1,234,567.89 | 1234.89
# (el orden de las alternativas importa; el lookahead evita cortar "1.234" en "1.23")
_MONTO_SIMPLE = r"(?:\d{1,3}(?:\.\d{3})+,\d{2}|\d{1,3}(?:,\d{3})+\.\d{2}|\d+,\d{2}|\d+\.\d{2})(?![\d.,]?\d)"
_MONTO = r"(" + _MONTO_SIMPLE + r")"
# Separador entre etiqueta y monto: hasta 25 caracteres sin dígitos y, a lo sumo, un salto de línea.
_SEP = r"[^\d\n]{0,25}(?:\n[^\d\n]{0,10})?"

REGLAS_EXTRACCION = {
    "empresa_propia": "TELCOM VENTURES DE ARGENTINA",

    # Rechazo DURO: el documento se descarta siempre, aunque tenga algo parecido
    # a un CAE o venga de un remitente de confianza (BUG-02).
    "rechazo_explicito": [
        "documento no valido como factura",
        "no valido como factura",
        "duplicado de recibo",
        "recibo de pago",
        "recibo de cobro",
        "recibo oficial",
        "recibo:",
        "recibo n",       # "Recibo N° 0002-..." / "Recibo Nro"
        "keops",          # sistema de recibos de Volquetes Marino (Keops 360)
    ],

    # Rechazo BLANDO: se descarta solo si el documento NO tiene un CAE válido.
    "palabras_prohibidas": [
        "anulada",
        "presupuesto",
        "nota de pedido",
        "orden de compra",
        "orden de compras",
        "vista de deuda",
        "resumen de cuenta",
        "estado de cuenta",
        "cuenta corriente de clientes",
        "cuenta corriente",
        "para cotizar",
        "cotizacion",
        "solicitud de pedido",
        "certificado de cobertura",
    ],

    # Frases de facturas reales que contienen palabras prohibidas; se borran
    # del texto antes de buscar palabras prohibidas.
    "frases_seguras": [
        "presupuesto economico",
        "orden de compra:",
        "orden de compra :",
        "orden de compra n",
        "orden de compra-",
    ],

    # Tolerancia total a saltos de línea y símbolos entre la palabra CAE y los 14-15 dígitos
    "rx_cae": r"C[.\s]*A[.\s]*E[.\s]*[:\s\|N°]*\n?\s*\|?\s*(\d{14,15})",
    "rx_nota_credito": r"NOTA\s*DE\s*CREDITO",
    "rx_nota_debito": r"NOTA\s*DE\s*DEBITO",
    "rx_numero_comp": r"(?:A|B|C)?[-]?\d{4}[-]?(\d{8})",
    "rx_numero_guion": r"N[°º]?\s*[:]?\s*(\d{4,8})",
    # Línea que termina en forma societaria (se evalúa línea por línea, ver procesador_core)
    "rx_razon_social": r"^[ \t]*([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ \t\.\,\-&]+?[ \t]+(?:S\.?R\.?L\.?|S\.?A\.?S\.?|S\.?C\.?A\.?|S\.?A\.?))(?![A-Za-z])",
    "rx_razon_social_etiqueta": r"(?:Razón\s*Social|Razon\s*Social|Denominación|Denominacion|Nombre\s*o\s*Razón\s*Social|Apellido\s*y\s*Nombre\s*/?\s*Razón\s*Social)\s*[:\.]?\s*([^\n\r]+)",
    "basura_afip": [
        "DOMICILIO FISCAL", "RESPONSABLE INSCRIPTO", "IVA",
        "PUNTO DE VENTA", "FACTURA", "CUIT", "CONDICION DE IVA", "INICIO DE ACTIVIDADES"
    ],
    "rx_cuit": r"(?:C\.?U\.?I\.?T\.?|C\.?U\.?I\.?L\.?|CUIT/CUIL/DNI)?[\s:\.\|N°º]*\n?\s*\|?\s*((?:20|23|24|27|30|33|34)[-\s]?\d{2}[-\s]?\d{6}[-\s]?\d{1})",
    # "Fecha:", "Fecha de Emisión:", "FECHA EMISION:" (no toma "Fecha de Vto." porque exige el número enseguida)
    "rx_fecha_emision": r"Fecha(?:\s*(?:de\s*)?Emisi[oó]n)?\s*[:\.]?\s*(\d{2}[-/]\d{2}[-/]\d{4})",
    "cuit_propio": "30696428898",
    "cuit_eurosat": "33676549639",
    "nombre_eurosat": "EUROSAT",
    "rx_oc": r"OC[\s:-]*(\d{2,6})",

    # ─── IMPORTES (MEJ-02) ───────────────────────────────────────────────────
    # Etiqueta + hasta 25 caracteres no numéricos (":", "$", "%", espacios) +
    # opcionalmente un salto de línea + MONTO (siempre con 2 decimales, para no
    # confundir cantidades o números de línea con importes).
    "rx_subtotal": r"(?:Sub\s*total|Importe\s*Neto\s*Gravado|Neto\s*Gravado)" + _SEP + _MONTO,
    "rx_iva_21": r"I\.?V\.?A\.?[^\d\n]{0,6}21(?:[,.]0{1,2})?\s*%" + _SEP + _MONTO,
    "rx_iva_105": r"I\.?V\.?A\.?[^\d\n]{0,6}10[,.]50?\s*%" + _SEP + _MONTO,
    "rx_iva_27": r"I\.?V\.?A\.?[^\d\n]{0,6}27(?:[,.]0{1,2})?\s*%" + _SEP + _MONTO,
    "rx_percepciones": r"(?:Percepci[oó]n(?:es)?|Perc\.|IIBB|Ingresos\s*Brutos|Otros\s*Tributos|Imp(?:uestos)?\.?\s*Internos)" + _SEP + _MONTO,
    # Total: primero etiquetas específicas; si no hay, "Total" suelto (no "Subtotal").
    "rx_total_especifico": r"(?:Importe\s*Total|Total\s*a\s*Pagar|Total\s*Factura|Total\s*Comprobante)" + _SEP + _MONTO,
    "rx_total": r"(?<![A-Za-z])(?<!Sub\s)Total\b" + _SEP + _MONTO,

    # Línea de detalle: [código] descripción  cantidad  [u.medida]  precio unit.  [bonif %]  subtotal
    # Heurística (formato libre según proveedor): texto + cantidad + dos montos.
    "rx_item": (
        r"^\s*(?:[A-Z0-9][A-Z0-9\-\.]{0,14}\s+)?"
        r"(?P<desc>[A-Za-zÁÉÍÓÚÑáéíóúñ][^\n]{2,80}?)\s+"
        r"(?P<cant>\d{1,6}(?:[\.,]\d{1,3})?)\s+"
        r"(?:[A-Za-z]{1,10}\.?\s+)?\$?\s*"
        r"(?P<unit>" + _MONTO_SIMPLE + r")\s+"
        r"(?:\d{1,3}(?:[,\.]\d{1,2})?\s*%\s+)?\$?\s*"
        r"(?P<sub>" + _MONTO_SIMPLE + r")\s*$"
    ),
}

# ─── CÓDIGOS DE COMPROBANTE AFIP (QR, MEJ-01) ────────────────────────────────
# Tipos aceptados como factura / nota de crédito / nota de débito.
TIPOS_AFIP_FACTURA = {1, 6, 11, 19, 51, 201, 206, 211}
TIPOS_AFIP_NOTA_DEBITO = {2, 7, 12, 20, 52, 202, 207, 212}
TIPOS_AFIP_NOTA_CREDITO = {3, 8, 13, 21, 53, 203, 208, 213}
TIPOS_AFIP_ACEPTADOS = TIPOS_AFIP_FACTURA | TIPOS_AFIP_NOTA_DEBITO | TIPOS_AFIP_NOTA_CREDITO
# Recibos A/B/C/M: son comprobantes AFIP pero NO son facturas.
TIPOS_AFIP_RECIBO = {4, 9, 15, 54}
