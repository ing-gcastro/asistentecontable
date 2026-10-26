# ADR-0003: Destino directo a \Compras y reporte solo con PDF sueltos (MEJ-03/04)

- **Estado:** Aceptado
- **Fecha:** 2026-10-26

## Contexto
Las aprobadas iban a `\Compras\fc a subir`. El negocio pidió que vayan a la raíz de `\Compras`, donde otras personas las mueven a los pocos minutos.

## Decisión
`CARPETA_COMPRAS = \\10.10.10.210\AyF_Trabajoadistancia\Compras` y `CARPETA_C` es un alias. El reporte y el monitor escanean **solo el nivel raíz** (`glob('*.pdf')`, sin subcarpetas). La búsqueda de etapas usa `CARPETAS_ETAPAS` definida en `config.py`, también sin recursión para `\Compras` y `\Contabilidad`.

## Alternativas consideradas
- Mantener `fc a subir` y copiar a ambos lados: duplica archivos.

## Consecuencias
- Lo que quedó en `fc a subir` no aparece como "nuevo" en el reporte, pero el histórico se conserva porque se lee del propio `Reporte_Maestro.html`.
- Cambiar la ruta en un solo lugar (`config.py`) propaga a todos.
