---
name: facturas-operacion
description: >
  Guía para operar, instalar y diagnosticar el sistema en una PC (MENU.bat, INSTALAR.bat,
  VERIFICAR.bat, tareas programadas, rutas, Python). Usar cuando algo "no anda", al instalar en una PC nueva,
  al cambiar de usuario de Windows, o cuando hay errores de Python, rutas de red, Outlook o Excel abierto.
---

# Operación y diagnóstico

1. Correr `VERIFICAR.bat --sin-pausa` y leer: librerías (FALTA → `INSTALAR.bat`), rutas (FALTA → red/VPN o `config_local.py`), tests.
2. Python: lo elige `scripts\_python.bat` (.venv → Anaconda del usuario → py → python). Si no hay, instalar Anaconda.
3. Rutas distintas en esta PC: copiar `config_local.example.py` → `config_local.py` y redefinir solo lo necesario (si cambia una base, redefinir también las derivadas).
4. Outlook: el downloader necesita Outlook abierto con la carpeta "Compras" y la sesión del usuario iniciada.
5. "Archivo abierto": cerrar el Excel; el sistema ya no mata Excel (ADR-0012).
6. Logs: `logs\paso1_preparar_lote_AAAA-MM.log`, `paso2_…`, `auto_downloader.log`, `notificaciones\`.
7. Detalle completo: `docs/RUNBOOK.md`.
