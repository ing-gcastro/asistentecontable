# ADR-0005: Rechazo duro vs. blando; el remitente confiable solo relaja el CAE (BUG-02)

- **Estado:** Aceptado
- **Fecha:** 2026-10-26

## Contexto
Un recibo de Volquetes Marino (Keops 360, "DOCUMENTO NO VALIDO COMO FACTURA") se aprobaba porque el remitente estaba en la lista blanca y eso salteaba todas las validaciones; además, números de 14 dígitos del importe parecían un CAE.

## Decisión
Dos listas en `pdf_rules.py`: `rechazo_explicito` (rechaza **siempre**: no válido como factura, recibo:, recibo de pago, duplicado de recibo, keops…) y `palabras_prohibidas` (rechaza solo si no hay CAE). El remitente confiable se evalúa **después** del rechazo explícito. El rechazo explícito mira solo la primera página, para que un recibo adjunto a una factura fusionada (regla Marino) no la tumbe. Todo el texto se normaliza (sin acentos) antes de comparar; esto además arregla las NC con "CRÉDITO" acentuado, que se clasificaban como FC.

## Alternativas consideradas
- Quitar a Marino de la lista blanca: perdería sus facturas sin CAE legible.

## Consecuencias
- Agregar una frase nueva = una línea en `pdf_rules.py` + un test.
