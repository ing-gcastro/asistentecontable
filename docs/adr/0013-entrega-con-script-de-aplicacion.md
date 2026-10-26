# ADR-0013: Entrega de la actualización con un script de aplicación

- **Estado:** Aceptado
- **Fecha:** 2026-10-26

## Contexto
La actualización se envía como zip a alguien con poco tiempo y una IA con pocos tokens. Descomprimir un zip reestructurado sobre el repo deja duplicados (los archivos viejos no se borran) y puede pisar estado local.

## Decisión
El zip trae `APLICAR.bat` + `aplicar.ps1` + `contenido/`. El script verifica git y el repo, crea una rama de respaldo, hace `git mv` de los archivos que cambian de carpeta, borra los obsoletos, deja de versionar los datos generados, copia el contenido nuevo, instala dependencias, corre los tests, hace commit y push a `main`. Si falla algo antes del commit, vuelve al estado anterior. `REVERTIR` documentado en el LEEME.

## Alternativas consideradas
- Instrucciones manuales: propensas a error.
- Pull request desde un fork: requiere más pasos y conocimiento de git.

## Consecuencias
- Un clic o una orden a la IA ("ejecutá APLICAR.bat").
- El script es idempotente: correrlo dos veces no rompe nada.
