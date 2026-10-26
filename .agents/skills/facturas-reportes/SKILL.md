---
name: facturas-reportes
description: >
  Mantiene los reportes HTML/Excel del ciclo de compras (Reporte_Maestro, reportes por sector,
  Reportesector, Reporte_Detallado con evolución de precios). Usar al agregar/quitar columnas, cuando un reporte
  muestra datos vacíos o duplicados, cuando la etapa contable o el estado no coinciden, o al tocar reporter.py.
---

# Reportes

## Piezas
- `reporter.leer_tabla_html()` — ÚNICA forma de leer un reporte existente (todo texto, sin miles, sin fila de filtros). Nunca usar `pd.read_html` directo (ADR-0007).
- `COLS_MAESTRO` / `COLS_SECTOR` — columnas y orden. `_celda()` — formato por columna.
- `generar_reporte_html()` — completo (escanea red: etapas + estados). `agregar_al_reporte_maestro()` — rápido (Paso 2).
- `generar_reportes_sector()` — por ubicación física en `AUTORIZACION`.
- `generar_reporte_detallado()` — desde `registro_facturas.json` + `cache_datos_pdf.json`.
- `LectorDatosFactura` — fecha de emisión e importes: registro → caché → PDF (tope `REPORTE_MAX_LECTURAS_PDF`).

## Agregar una columna
1. Sumarla a `COLS_MAESTRO` (y a `COLS_SECTOR` si corresponde, que son las primeras 12).
2. Poblarla en `_escanear_compras()`, `datos_para_html` de `2_procesar_lote.py` y/o `_completar_con_datos()`.
3. Si es de tipo texto numérico, agregarla a `_COLS_TEXTO` (ya incluye todas las de `COLS_MAESTRO`).
4. Test en `TestReporter` o en `test_flujo`.

## Diagnóstico
- Fecha de emisión vacía en facturas viejas: normal hasta que la caché se complete (varias corridas).
- Duplicados en el maestro: la clave es (Razón Social en mayúsculas, Nro Comprobante sin `.0`).
