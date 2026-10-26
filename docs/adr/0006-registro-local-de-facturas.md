# ADR-0006: Registro local de datos por factura + caché de lecturas (BUG-04, MEJ-02)

- **Estado:** Aceptado
- **Fecha:** 2026-10-26

## Contexto
La fecha del reporte era la de modificación del archivo (`st_mtime`) y no la de emisión. El reporte detallado necesita importes e ítems que el nombre del archivo no contiene.

## Decisión
El Paso 1 guarda lo extraído en `lote_datos.json`; el Paso 2 consolida Excel (lo editado por el usuario) + lote en `registro_facturas.json`, indexado por nombre final del archivo. El reporter resuelve los datos de cada PDF así: registro → `cache_datos_pdf.json` → lectura liviana del PDF (sin OCR, con tope `REPORTE_MAX_LECTURAS_PDF` por corrida). El maestro muestra `Fecha Emisión` y `Fecha Carga` por separado.

## Alternativas consideradas
- Codificar la fecha en el nombre del archivo: cambia una convención que usan otras áreas.
- Base SQLite: más robusta pero agrega complejidad; se reconsidera arriba de ~50 000 facturas.

## Consecuencias
- `registro_facturas.json` es estado de negocio: **respaldarlo** (RUNBOOK).
- La primera regeneración completa puede tardar más mientras completa la caché.
