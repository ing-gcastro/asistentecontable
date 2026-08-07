---
name: smart-downloader
description: >
  Mejora el sistema de descarga automática de facturas desde Outlook.
  Usar cuando: se habla de auto_downloader.py, la descarga no baja todos los
  correos, se menciona fecha hardcodeada o fecha fija, cuando se quiere que
  la descarga sea inteligente sobre desde qué fecha buscar, al configurar
  tareas programadas de Windows para la descarga automática, o cuando el
  sistema descarga más de lo necesario o no descarga lo suficiente.
---

# Smart Downloader — Skill de implementación

## Contexto del problema

### Bug: fecha hardcodeada

`auto_downloader.py` actual tiene fecha hardcodeada. Cuando se ejecuta como tarea
programada de Windows siempre usa la misma fecha de inicio.

### Solución: fecha dinámica + logging + argumento CLI

La lógica correcta es obtener la fecha del último email descargado desde
`emails_procesados.json` (formato `YYYYMMDDHHMMSS_email_asunto`).

## Implementación

Reemplazar `auto_downloader.py` completo con versión inteligente que:
- Calcula la fecha de inicio automáticamente desde el tracking JSON
- Acepta argumentos CLI (`--desde DD/MM/YYYY`, `--dias-atras N`)
- Genera logs en `logs/auto_downloader.log`
- Funciona como Tarea Programada de Windows sin intervención

## Criterios de éxito

- Al ejecutar sin argumentos, usa la fecha del tracking
- El archivo `logs/auto_downloader.log` se crea y registra cada ejecución
- Si el tracking está vacío, usa el fallback de N días atrás
