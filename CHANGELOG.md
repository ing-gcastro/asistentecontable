# CHANGELOG

## [3.0.0] — 2026-10-26

### Corregido
- **BUG-01** El mes del nombre del archivo sale de la fecha de emisión de la factura (antes: mes actual).
- **BUG-02** Los recibos ("Documento no válido como factura", Keops, QR de recibo) se rechazan siempre; el remitente confiable solo relaja la exigencia del CAE.
- **BUG-03/MEJ-03** Las aprobadas van directo a `\Compras` (antes `\Compras\fc a subir`).
- **BUG-04** El reporte muestra la fecha de emisión real (`Fecha Emisión`) y aparte la `Fecha Carga`.
- **BUG-05** Eliminado `update_reporter.py` (editaba `reporter.py` en tiempo de ejecución).
- La detección de duplicados contra el historial no funcionaba (`47336` vs `47336.0`) — ADR-0007.
- Faltaba `lxml` en las dependencias (historial vacío en instalaciones limpias).
- Los Pasos 1/2 ya no cierran todos los Excel con `taskkill`; `excel_a_pdf` no cierra el Excel del usuario.
- Downloader: se perdían PDFs extra de Redguard/Aqualine y adjuntos con el mismo nombre; los errores ya no se ocultan.
- El remitente del mail llega al Paso 1 (lista blanca operativa en el flujo por lotes).
- NC con "CRÉDITO" acentuado, extracción de razón social, CUIT con dígito verificador, sector editado en el nombre.
- HTML con escape; JSON de estado con escritura atómica; búsqueda de OC sin regex.

### Agregado
- **MEJ-01** Lectura del QR AFIP (`qr_afip.py`, zxing-cpp).
- **MEJ-02** `Reporte_Detallado.html`: neto, IVA 21/10,5/27, otros impuestos, total, ítems y evolución de precios por proveedor.
- **MEJ-04** El reporte y el monitor consideran solo los PDF sueltos de `\Compras`.
- **MEJ-05** `monitor_compras.py` (sondeo apto para rutas de red).
- Aviso por mail a cada sector (`notificador_sectores.py`), conectado pero desactivado hasta definir el SDD §8.
- `MENU.bat`, `INSTALAR.bat`, `VERIFICAR.bat`, lanzador `scripts\_python.bat`, `.venv`.
- Log de auditoría mensual de Pasos 1 y 2 en `logs\`.
- Tests automáticos (28) y CI en GitHub Actions.
- Documentación: SDD, 13 ADR, RUNBOOK, estado del QA, AGENTS.md, skills y workflows.

### Cambiado
- Repo organizado por dominio: `compras_oc/`, `tesoreria_pagos/`, `datos_maestros/`, `flota_seguros/` (ADR-0002).
- Rutas sin usuario hardcodeado; overrides por PC en `config_local.py`.
- Estado de ejecución y datos generados fuera de git (ADR-0010).
- `REPORTE DE AUTORIZACION/indexar_pdfs.py` y `monitor.py` reemplazados por `reporter.py` / `monitor_compras.py`.
