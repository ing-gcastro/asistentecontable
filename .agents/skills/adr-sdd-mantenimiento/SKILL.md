---
name: adr-sdd-mantenimiento
description: >
  Mantiene la documentación de diseño: crear un ADR, actualizar el SDD, el CHANGELOG y el
  estado del QA. Usar después de cualquier cambio de comportamiento, estructura, dependencias o datos, o cuando se
  pide documentar una decisión.
---

# ADR / SDD / CHANGELOG

1. **ADR** si hubo una decisión con alternativas: copiar `docs/adr/plantilla.md` → `docs/adr/NNNN-slug.md` (número siguiente), completar y sumar la fila a `docs/adr/README.md`. Si reemplaza a otro, marcar el viejo `Reemplazado por ADR-NNNN`.
2. **SDD** (`docs/SDD.md`): actualizar la sección afectada (módulos §3.1, flujo §4, datos §5, backlog §9). El SDD describe el sistema **como es hoy**.
3. **CHANGELOG.md**: entrada bajo la versión en curso (Corregido / Agregado / Cambiado).
4. Si cierra un ítem de QA: `docs/qa/ESTADO_QA.md`.
5. Mantener breve: 1 ADR ≈ 20-40 líneas.
