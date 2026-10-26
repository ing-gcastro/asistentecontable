# RUNBOOK — Operación y mantenimiento

## Día a día (doble clic en `MENU.bat`)

1. **Descargar facturas de Outlook** (opción 1). También corre solo como tarea programada (`Proyecto_facturas_v2\run_downloader.bat`).
2. **Paso 1** (opción 2): genera `Validacion_Lote.xlsx` en la Carpeta A.
3. Abrir el Excel, revisar **Acción** (APROBAR / RECHAZAR / REVISION), **Sector**, **OC**, guardar y **cerrar**.
4. **Paso 2** (opción 3): archiva, actualiza los reportes y genera el aviso a los sectores.
5. Reportes: `\\10.10.10.210\AyF_Trabajoadistancia\Compras\Reporte_Maestro.html` y `Reporte_Detallado.html`.

## Instalación en una PC nueva
1. Instalar Git y Anaconda (o Python 3.11+ con "Add to PATH").
2. `git clone https://github.com/ing-gcastro/asistentecontable.git %USERPROFILE%\ROBOTS\asistentecontable`
3. Doble clic en `INSTALAR.bat` (crea `.venv`, instala dependencias y corre los tests).
4. Copiar los maestros a `%USERPROFILE%\ROBOTS\data\` (`proveedores (2).xls`, `Sectores.xlsx`, OCs, `consumidas.xlsx`).
5. Si alguna carpeta es distinta: copiar `Proyecto_facturas_v2\config_local.example.py` como `config_local.py` y editarla.
6. `VERIFICAR.bat`: tiene que decir `TESTS OK` y `OK` en las rutas.

## Tareas programadas recomendadas (Programador de tareas de Windows)

| Tarea | Programa | Frecuencia |
|---|---|---|
| Descarga de Outlook | `Proyecto_facturas_v2\run_downloader.bat` | cada 30 min en horario laboral |
| Reportes | `ejecutar_reportes.bat` | cada 10 min, o dejar abierto `ejecutar_monitor_compras.bat` |
| Recordatorio a sectores | `ejecutar_recordatorio_sectores.bat` | diario 9:00 (cuando se active, SDD §8) |

"Iniciar en": la carpeta del `.bat`. "Ejecutar solo cuando el usuario haya iniciado sesión" (Outlook necesita la sesión).

## Accesos directos que cambiaron de lugar (v3)

| Antes (raíz) | Ahora |
|---|---|
| `ejecutar_reporte_oc.bat` | `compras_oc\ejecutar_reporte_oc.bat` |
| `ejecutar_reporte_legajos.bat`, `ejecutar_vincular_pendientes.bat`, `ejecutar_exportar_mails.bat` | `tesoreria_pagos\…` |
| `ejecutar_dividir_talones.bat` | `flota_seguros\…` |
| `REPORTE DE AUTORIZACION\indexar_pdfs.py`, `monitor.py` | eliminados: `reporter.py` / `monitor_compras.py` (`actualizar.bat` sigue funcionando) |

`ejecutar_paso1_…`, `ejecutar_paso2_…`, `ejecutar_reportes.bat` y `run_downloader.bat` **no** cambiaron de lugar.

## Respaldos
Copiar semanalmente a la red (por ejemplo a `AyF_Trabajoadistancia\Compras\_respaldo_robot\`):
`Proyecto_facturas_v2\registro_facturas.json`, `estado_oc.json`, `emails_procesados.json`.

## Problemas frecuentes

| Síntoma | Causa probable | Solución |
|---|---|---|
| "El archivo está abierto en otro programa" | Excel de validación abierto | Guardarlo, cerrarlo y apretar ENTER |
| Una factura válida va a "no es factura" | Sin CAE legible / palabra prohibida / frase de rechazo | Ver la columna **Motivo** del Excel; si es un caso nuevo, ajustar `pdf_rules.py` y agregar un test (skill `facturas-reglas-validacion`) |
| Un recibo entra como factura | Frase de recibo no contemplada | Agregarla a `rechazo_explicito` en `pdf_rules.py` |
| La fecha de emisión aparece vacía en el maestro | Factura vieja sin leer todavía | Se completa sola en las próximas regeneraciones (tope por corrida: `REPORTE_MAX_LECTURAS_PDF`) |
| `QR_AFIP = NO` en todas | Falta zxing-cpp | `INSTALAR.bat` |
| "No se encontró Python" | Python/Anaconda no instalado | Instalarlo y correr `INSTALAR.bat` |
| Duplicados no detectados | Reporte maestro inaccesible o falta `lxml` | `VERIFICAR.bat` |

## Volver atrás una actualización
`git log --oneline -5` para ver el commit, y luego `git revert <commit>` (deja historial) o
`git reset --hard respaldo/antes-v3-<fecha>` (rama creada por `APLICAR.bat`, solo local).
