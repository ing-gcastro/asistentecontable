---
name: cae-invoice-validator
description: >
  Mejora la detección y validación de CAE (Código de Autorización Electrónica) y
  clasificación de facturas AFIP en el proyecto. Usar cuando: hay facturas que se
  clasifican incorrectamente como "no es factura", cuando se menciona que el CAE
  no se detecta, al trabajar en es_factura_valida() o extraer_datos_pdf(),
  cuando se habla de AFIP o comprobantes ARCA, o cuando facturas válidas van
  a CARPETA_B incorrectamente.
---

# CAE Invoice Validator — Skill de implementación

## Contexto del problema

### Bug crítico: rx_cae definida pero nunca usada

`pdf_rules.py` define una regex robusta para detectar el CAE:
```python
"rx_cae": r"C\.?[A-ZÀ-ÿ]\.?\.?E\.?[\sN°:\.\|]*\n?\s*\|?\s*(\d{14,15})"
```

Pero `es_factura_valida()` en `procesador_core.py` solo hace:
```python
texto_sin_espacios_raros = re.sub(r'[\s\n\r\t:\-\._\|]+', '', texto_lower)
if "cae" in texto_sin_espacios_raros:
    return True
```

## Implementación: reemplazar es_factura_valida() completo

```python
def es_factura_valida(texto_lower, remitente=""):
    if es_factura_por_remitente(remitente):
        return True
    texto_seguro = (texto_lower
                    .replace("presupuesto económico", "")
                    .replace("presupuesto economico", ""))
    for prohibida in REGLAS_EXTRACCION["palabras_prohibidas"]:
        if prohibida.lower() in texto_seguro:
            return False
    if re.search(r'\b8\d{13}\b', texto_lower):
        return True
    if re.search(REGLAS_EXTRACCION["rx_cae"], texto_lower, re.IGNORECASE | re.MULTILINE):
        return True
    if "cae" in re.sub(r'[\s\n\r\t:\-\._\|]+', '', texto_lower):
        return True
    if re.search(r'\bfactura\b', texto_lower) and re.search(r'\b\d{4,5}[-\s]\d{7,8}\b', texto_lower):
        return True
    return False
```

## Mejora adicional: regex CAE más robusta en pdf_rules.py

```python
"rx_cae": r"C[.\s]*A[.\s]*E[.\s]*[:\s\|N°]*\n?\s*\|?\s*(\d{14,15})"
```

## Criterios de éxito

- Facturas con "C.A.E." con puntos se clasifican correctamente
- Facturas con CAE en columna separada (tabulaciones) se detectan
- No hay falsos positivos en documentos sin CAE (presupuestos, remitos)
