# ADR-0007: Lectura robusta del historial HTML (duplicados que no se detectaban)

- **Estado:** Aceptado
- **Fecha:** 2026-10-26

## Contexto
`pd.read_html` con valores por defecto: (1) la fila de filtros del `<thead>` volvía como una fila vacía, que convertía la columna `Nro Comprobante` en float (`47336.0`); el Paso 1 comparaba contra `"47336"`, así que **la detección de duplicados contra el historial nunca coincidía**; (2) `thousands=','` puede corromper montos; (3) sin `lxml` instalado fallaba en silencio y el historial quedaba vacío.

## Decisión
Una sola función `reporter.leer_tabla_html()`: lee todo como texto, sin separador de miles, descarta filas vacías, normaliza encabezados viejos (`Razon Social`, `Fecha` → `Fecha Carga`) y quita `.0`. `lxml` pasa a ser dependencia obligatoria. Todos los módulos usan esta función.

## Alternativas consideradas
- Cambiar el HTML para no tener fila de filtros: no arregla los reportes ya generados.

## Consecuencias
- Compatibilidad total con los reportes v2 (test con un fixture real generado por el código anterior).
