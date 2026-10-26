---
name: notificaciones-sectores
description: >
  Activa, configura o modifica el aviso por mail a cada sector con las facturas para
  autorizar (notificador_sectores.py). Usar cuando se habla de mails a sectores, recordatorios de pendientes,
  destinatarios por sector o del SDD §8.
---

# Notificaciones a sectores

Estado: **desactivado** hasta cerrar las decisiones N1–N6 de `docs/SDD.md` §8.

## Activar (solo configuración, `config.py` o `config_local.py`)
1. `NOTIFICACION_DESTINATARIOS = {"RRHH": ["rrhh@empresa.com"], ...}` (sector sin lista = no se notifica).
2. `NOTIFICACION_MODO = "borrador"` → correr `ejecutar_recordatorio_sectores.bat` y revisar Borradores de Outlook.
3. Cuando el formato esté aprobado: `NOTIFICACION_MODO = "enviar"` y tarea programada diaria del recordatorio.
4. Registrar la decisión en un ADR nuevo que actualice ADR-0009, y completar SDD §8.

## Reglas
- Nunca poner `"enviar"` sin pedido explícito de una persona.
- Para cambiar el contenido: `construir_email()` (es pura, tiene test).
- Vista previa sin Outlook: modo `desactivado` → `logs\notificaciones\*.html`.
