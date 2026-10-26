# ADR-0008: Monitor de \Compras por sondeo propio (MEJ-05)

- **Estado:** Aceptado
- **Fecha:** 2026-10-26

## Contexto
Hace falta regenerar los reportes cuando entran o se mueven archivos en `\Compras` (ruta UNC). Los eventos nativos (`ReadDirectoryChangesW`) no son confiables sobre SMB; el QA propuso `watchdog.PollingObserver`.

## Decisión
`monitor_compras.py` hace lo mismo que `PollingObserver` sin la dependencia: cada `MONITOR_INTERVALO_SEG` compara una foto (nombre, tamaño, fecha) de los PDF de `\Compras` (raíz) y de `AUTORIZACION` (recursivo); ante cambios espera `MONITOR_ESPERA_SEG` sin cambios (agrupa ráfagas) y regenera todo.

## Alternativas consideradas
- `watchdog`: una dependencia más para el mismo resultado.
- Solo una tarea programada cada 5-10 min: más simple, pero con más demora. Se puede combinar (RUNBOOK).

## Consecuencias
- El `monitor.py` viejo (Observer nativo, que además pisaba `Reportesector.html` con otro formato) se elimina.
