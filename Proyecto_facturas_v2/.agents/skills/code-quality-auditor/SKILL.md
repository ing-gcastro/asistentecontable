---
name: code-quality-auditor
description: >
  Audita y mejora la calidad del código del proyecto: elimina archivos de código
  muerto, agrega caching para operaciones costosas, mejora el manejo de errores,
  y agrega requirements.txt. Usar cuando: se menciona código muerto, se quieren
  eliminar prueba_fase1.py o update_downloader_script.py, cuando el proceso es
  lento con muchas facturas, al hacer refactoring general, o cuando se pide
  mejorar la calidad o mantenibilidad del código.
---

# Code Quality Auditor — Skill de implementación

## Acciones

1. Eliminar `prueba_fase1.py` y `update_downloader_script.py` (código muerto)
2. Crear `requirements.txt` con todas las dependencias
3. Agregar caché a `cargar_sectores()` en `proveedores.py`
4. Agregar logging básico a `procesador_core.py`
5. Resolver tipo hardcodeado "S" en `oc_manager.py`

## Criterios de éxito

- Archivos de código muerto eliminados
- `requirements.txt` presente y correcto
- `cargar_sectores()` usa caché en memoria
- Directorio `logs/` con archivos de log
