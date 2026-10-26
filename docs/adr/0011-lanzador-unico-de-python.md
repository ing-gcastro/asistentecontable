# ADR-0011: Lanzador único de Python y .venv con paquetes del sistema

- **Estado:** Aceptado
- **Fecha:** 2026-10-26

## Contexto
Cada `.bat` tenía `C:\Users\gcastro\AppData\Local\anaconda3\python.exe` hardcodeado; en otra PC o usuario nada funcionaba.

## Decisión
`scripts/_python.bat` resuelve el intérprete: `.venv` del repo → Anaconda del usuario actual → `py -3` → `python`. `INSTALAR.bat` crea `.venv --system-site-packages` (reutiliza pandas/pywin32 de Anaconda) e instala solo lo que falta. Todos los `.bat` usan el lanzador.

## Alternativas consideradas
- Instalar todo en la base de Anaconda: funciona, pero mezcla versiones con otros usos de la PC.

## Consecuencias
- Un único punto a tocar si cambia el Python de la PC.
- `MENU.bat` es la puerta de entrada.
