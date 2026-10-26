# ADR-0010: Estado de ejecución fuera de git y escritura atómica

- **Estado:** Aceptado
- **Fecha:** 2026-10-26

## Contexto
`emails_procesados.json` (con mails y asuntos) y `estado_oc.json` estaban versionados; un zip o pull viejo podía pisar el estado actual y volver a descargar mails o liberar OCs consumidas. Las escrituras directas podían dejar un JSON a medio escribir ante un corte.

## Decisión
El estado y los datos generados se ignoran en git (`.gitignore`) y se dejan de versionar con `git rm --cached` (quedan en disco). `estado.guardar_json()` escribe en un `.tmp` y reemplaza con `os.replace` (atómico).

## Alternativas consideradas
- Mantenerlos versionados como "backup": git no es un backup de datos operativos y expone datos personales.

## Consecuencias
- Otra PC que clone el repo arranca sin estado (el downloader usa 30 días hacia atrás y los duplicados se frenan por el historial del reporte).
- Respaldo de `registro_facturas.json` y `estado_oc.json` en el RUNBOOK.
