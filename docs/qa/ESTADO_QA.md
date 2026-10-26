# Estado de los ítems del QA Report (26/10/2026)

Reporte original: [`QA_Reporte_Mejoras.md`](QA_Reporte_Mejoras.md) (y `.docx`, mismo contenido).

| ID | Ítem | Estado | Dónde | Test |
|---|---|---|---|---|
| BUG-01 | Mes del nombre desde la fecha de la factura | ✅ Resuelto | `procesador_core.nombre_mes_desde_fecha()`, usado en Paso 1 y modo interactivo | `test_bug01_*`, `test_flujo` |
| BUG-02 | Recibos clasificados como factura | ✅ Resuelto | `pdf_rules.rechazo_explicito`, `es_factura_valida()` (orden de reglas), QR de recibo | `test_bug02_*`, `test_recibo_*` |
| BUG-03 | Reporter dependía de `fc a subir` | ✅ Resuelto | `config.CARPETA_COMPRAS`, `reporter._escanear_compras()` | `test_flujo` |
| BUG-04 | Fecha del reporte = `st_mtime` | ✅ Resuelto | Columnas `Fecha Emisión` + `Fecha Carga`; `registro_facturas.json` + `LectorDatosFactura` | `test_flujo` |
| BUG-05 | `update_reporter.py` editaba el código en runtime | ✅ Eliminado | La lógica vive en `reporter.py` | — |
| MEJ-01 | QR AFIP sin IA | ✅ Implementado | `qr_afip.py` (zxing-cpp) | `TestQR` (incluye un QR real dentro de un PDF) |
| MEJ-02 | Reporte detallado con impuestos y evolución de precios | ✅ Implementado | `reporter.generar_reporte_detallado()` → `Reporte_Detallado.html` (Chart.js) | `test_importes_e_items`, `test_flujo` |
| MEJ-03 | Destino directo a `\Compras` | ✅ Implementado | `config.CARPETA_COMPRAS` | `test_flujo` |
| MEJ-04 | Solo PDF sueltos de `\Compras` | ✅ Implementado | `glob('*.pdf')` sin recursión en reporter y monitor | `test_monitor_solo_pdf_sueltos` |
| MEJ-05 | Actualización automática del reporte | ✅ Implementado | `monitor_compras.py` + `ejecutar_monitor_compras.bat` | `test_monitor_*` |

## Hallazgos adicionales corregidos (no estaban en el QA)

| # | Problema | Impacto | Corrección |
|---|---|---|---|
| A1 | `pd.read_html` convertía `47336` en `47336.0` (por la fila de filtros) | **La detección de duplicados contra el historial nunca funcionaba** | `reporter.leer_tabla_html()` (ADR-0007) |
| A2 | `lxml` no estaba en `requirements.txt` | En una instalación limpia el historial quedaba vacío, sin aviso | Agregado + `VERIFICAR.bat` |
| A3 | `taskkill /f /im excel.exe` en Pasos 1 y 2 | Cerraba todos los Excel abiertos sin guardar | Espera a que se cierre (ADR-0012) |
| A4 | Redguard/Aqualine con 2+ PDF: los extra se borraban con la carpeta temporal | Pérdida de adjuntos | Se fusionan todos |
| A5 | Adjuntos con el mismo nombre en un mail se pisaban | Pérdida de adjuntos | Nombres únicos en temporal |
| A6 | Errores de un mail se tragaban en silencio | Mails sin descargar sin aviso | Se informa y se reintenta |
| A7 | El remitente del mail nunca llegaba al Paso 1 | La lista blanca no funcionaba en el flujo por lotes | `remitentes_descargas.json` |
| A8 | NC con "CRÉDITO" acentuado se clasificaban como FC | Tipo incorrecto en el nombre | Texto normalizado sin acentos |
| A9 | Razón social: regex anclada al inicio del documento y `\b` tras "S.R.L." | Casi nunca se extraía del texto | Búsqueda por línea |
| A10 | Cambiar el Sector en el Excel no cambiaba el nombre del archivo | Archivos "… GENERAL.pdf" dentro de otro sector | `ajustar_sector_en_nombre()` |
| A11 | `excel_a_pdf` con `Dispatch` cerraba el Excel del usuario | Pérdida de trabajo | `DispatchEx` |
| A12 | Rutas `C:\Users\gcastro\…` en 20+ lugares | No funcionaba en otra PC/usuario | `config.py` dinámico + `scripts\_python.bat` |
| A13 | `indexar_pdfs.py`/`monitor.py` y `reporter.py` escribían el mismo `Reportesector.html` con formatos distintos | El reporte cambiaba según quién corría último | Un solo generador |
| A14 | HTML sin escape (`&`, `<` en razón social) | Reporte roto / inyección | `html.escape` |
| A15 | JSON de estado escritos sin atomicidad | Corrupción ante un corte | `estado.guardar_json()` |
| A16 | `oc_manager` usaba regex con el nombre del proveedor | Error con nombres con `(`, `+`, etc. | `regex=False` |
| A17 | CUIT: se tomaba el primer número de 11 dígitos | CUIT equivocado si había otro número antes | Dígito verificador |
| A18 | `reporter_run.py` sin import (no funcionaba), `fix_nb.py` de un solo uso, `__pycache__` y checkpoints versionados | Ruido | Eliminados |
| A19 | `Auditoria_Facturas_2026.ipynb` importa `extraer_datos_factura`, que ya no existe | El notebook no corre | Documentado como legado: usar `extraer_datos_pdf()` |
