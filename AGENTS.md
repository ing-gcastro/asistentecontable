# AGENTS.md — Instrucciones para agentes de IA (Antigravity/Gemini, Claude, Codex)

Leé esto primero. Está pensado para gastar **pocos tokens**: no explores el repo entero, usá este mapa.

## Qué es
Sistema de recepción y procesamiento de facturas de compras (Windows + Outlook + carpetas de red UNC).
Núcleo en `Proyecto_facturas_v2/`. Diseño: `docs/SDD.md`. Decisiones: `docs/adr/`. Operación: `docs/RUNBOOK.md`.

## Comandos (no inventes otros)
| Para | Ejecutar |
|---|---|
| Instalar dependencias + tests | `INSTALAR.bat --sin-pausa` |
| Verificar (librerías, rutas, tests) | `VERIFICAR.bat --sin-pausa` |
| Solo tests | `.venv\Scripts\python.exe -m unittest discover -s Proyecto_facturas_v2/tests -q` |
| Operar | `MENU.bat` (interactivo: es para humanos) |

## Reglas
1. **Toda ruta o constante va en `Proyecto_facturas_v2/config.py`.** Nunca hardcodear `C:\Users\...` ni `\\10.10.10.210` en otro archivo.
2. Las reglas de clasificación viven en `pdf_rules.py` (datos), la lógica en `procesador_core.es_factura_valida()`.
3. No ejecutes `1_preparar_lote.py`, `2_procesar_lote.py`, `downloader.py` ni `notificador_sectores.py` para "probar": mueven archivos reales, leen Outlook o mandan mails. Probá con los tests (usan carpetas temporales).
4. No cambies `NOTIFICACION_MODO` a `"enviar"` sin pedido explícito (SDD §8).
5. No renombres `Proyecto_facturas_v2/` (lo usan tareas programadas).
6. No versiones estado ni datos (`*.json` de estado, `*.xlsx`, logs): ya están en `.gitignore`.
7. Todo cambio de comportamiento: test nuevo o ajustado + entrada en `CHANGELOG.md`; si es una decisión de diseño, ADR nuevo (`docs/adr/plantilla.md`).
8. Antes de commitear: `VERIFICAR.bat --sin-pausa` tiene que terminar en `TESTS OK`.
9. Respondé en español.

## Mapa rápido (dónde tocar)
| Necesidad | Archivo |
|---|---|
| Una factura válida se rechaza / un recibo entra | `pdf_rules.py` (`rechazo_explicito`, `palabras_prohibidas`) + test en `tests/test_sistema.py` |
| Proveedor nuevo con regla especial | `proveedor_rules.py` (remitentes), `downloader.py` (fusión de adjuntos) |
| Columnas del Excel de validación | `1_preparar_lote.py` (`COLUMNAS`) y su lectura en `2_procesar_lote.py` |
| Reportes HTML | `reporter.py` (`COLS_MAESTRO`, plantillas) |
| Carpetas, servidores, tiempos | `config.py` / `config_local.py` |
| Mails a sectores | `notificador_sectores.py` + `config.NOTIFICACION_*` |

## Skills y workflows
- Skills: `.agents/skills/<nombre>/SKILL.md` (reglas de validación, reportes, operación, proveedor nuevo, notificaciones, ADR/SDD).
- Workflows: `.agents/workflows/` (`verificar.md`, `publicar-cambios.md`, `nuevo-adr.md`).
