# ADR-0012: No cerrar Excel a la fuerza

- **Estado:** Aceptado
- **Fecha:** 2026-10-26

## Contexto
Los Pasos 1 y 2 ejecutaban `taskkill /f /im excel.exe`, que cierra **todos** los Excel abiertos sin guardar. `excel_a_pdf` usaba `Dispatch`, que se engancha al Excel del usuario y lo cierra con `Quit()`.

## Decisión
`estado.esperar_archivo_libre()` detecta si el archivo está bloqueado y le pide al usuario que lo cierre (ENTER para reintentar, `c` para cancelar). `excel_a_pdf` usa `DispatchEx` (instancia propia).

## Alternativas consideradas
- Seguir con `taskkill`: pérdida de trabajo.

## Consecuencias
- Un paso más para el usuario si dejó el Excel abierto, pero sin pérdida de datos.
