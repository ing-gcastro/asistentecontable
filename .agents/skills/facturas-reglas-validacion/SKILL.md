---
name: facturas-reglas-validacion
description: >
  Ajusta cómo el sistema decide si un PDF es factura/NC/ND o no (CAE, QR AFIP,
  recibos, palabras prohibidas, remitentes confiables) y cómo extrae número, CUIT, fecha e importes.
  Usar cuando una factura válida termina en "no es factura", cuando un recibo/presupuesto entra como factura,
  cuando un dato sale mal en el Excel del Paso 1, o al tocar pdf_rules.py / procesador_core.py / qr_afip.py.
---

# Reglas de validación y extracción

## Dónde está cada cosa
- Datos de reglas: `Proyecto_facturas_v2/pdf_rules.py` (`rechazo_explicito`, `palabras_prohibidas`, `frases_seguras`, regex, tipos AFIP).
- Lógica: `procesador_core.es_factura_valida()` (orden de reglas en SDD §4.2) y `extraer_datos_pdf()`.
- QR: `qr_afip.py` (si hay QR AFIP, manda sobre el texto).
- Remitentes confiables: `proveedor_rules.REMITENTES_CONFIABLES` (solo relajan el CAE, ADR-0005).

## Procedimiento (siempre con test)
1. Pedí el texto del PDF problemático (o un extracto): `python -c "import fitz;print(fitz.open(r'RUTA')[0].get_text('text',sort=True))"`.
2. Agregá un caso en `tests/test_sistema.py` (`TestReglas`) que reproduzca el problema → debe fallar.
3. Ajustá **datos** en `pdf_rules.py` antes que lógica. Las listas van en minúsculas y sin acentos (el texto se normaliza).
   - Recibo que entra → agregar frase a `rechazo_explicito`.
   - Factura válida rechazada por una palabra → agregar la frase completa a `frases_seguras`.
   - Importe con otra etiqueta → ampliar la regex correspondiente (usa `_SEP` + `_MONTO`).
4. `VERIFICAR.bat --sin-pausa` → `TESTS OK`. Anotá el cambio en `CHANGELOG.md`.

## Cuidado
- `rechazo_explicito` rechaza aunque haya CAE: usar frases inequívocas.
- El rechazo explícito se evalúa solo en la 1ª página (PDF fusionados Marino).
- Montos: siempre con 2 decimales para no confundir cantidades con importes.
