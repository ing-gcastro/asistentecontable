---
name: config-centralizer
description: >
  Centraliza todas las rutas y configuraciones hardcodeadas dispersas en
  múltiples archivos del proyecto hacia config.py. Usar cuando: se encuentran
  rutas absolutas hardcodeadas en downloader.py, oc_manager.py o cualquier otro
  archivo que no sea config.py, cuando se va a cambiar el usuario de Windows
  (gcastro), cuando se habla de configuración centralizada, o cuando se necesita
  que el proyecto sea portable entre equipos.
---

# Config Centralizer — Skill de implementación

## Problema: rutas dispersas en múltiples archivos

Rutas hardcodeadas en downloader.py, oc_manager.py y procesador_core.py
deben moverse a config.py como única fuente de verdad.

## Implementación

Reemplazar `config.py` con versión centralizada que define:
- `_USUARIO`, `_BASE_LOCAL`, `_BASE_REPO`, `_BASE_DATA`, `_SERVIDOR`
- `CARPETA_A`, `CARPETA_B`, `CARPETA_C`, `CARPETA_OC`
- `ARCHIVO_PROVEEDORES`, `ARCHIVO_SECTORES`, `OC_DB_PATH`, `CONSUMIDAS_PATH`
- `ESTADO_EMAILS`, `ESTADO_OC_PATH`
- `REPORTE_MAESTRO_HTML`, `LOGS_DIR`
- `MESES_ESPANOL`

Luego actualizar imports en cada archivo consumidor.

## Criterios de éxito

- config.py es la ÚNICA fuente de verdad para rutas
- Ningún otro archivo tiene cadenas `r"C:\Users\gcastro\"` o `r"\\10.10.10.210\"`
- Para cambiar de usuario solo se edita `_USUARIO` en config.py
