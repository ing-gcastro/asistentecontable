# ADR-0001: Registrar decisiones de arquitectura con ADR

- **Estado:** Aceptado
- **Fecha:** 2026-10-26

## Contexto
El proyecto creció con cambios hechos por distintas personas y herramientas de IA, sin registro del porqué. El reporte de QA mostró decisiones implícitas (rutas, reglas de validación) difíciles de mantener.

## Decisión
Cada decisión que cambie comportamiento, datos o estructura se registra en `docs/adr/NNNN-titulo.md` con este formato (Contexto, Decisión, Alternativas, Consecuencias). El diseño vigente vive en `docs/SDD.md` y los cambios por versión en `CHANGELOG.md`. Un ADR no se edita para cambiar la decisión: se crea otro que lo reemplaza (`Estado: Reemplazado por ADR-XXXX`).

## Alternativas consideradas
- Documentar solo en el código: se pierde el porqué.
- Wiki externa: se desincroniza del código.

## Consecuencias
- Cualquier persona o agente (Gemini, Claude) puede entender el sistema leyendo `AGENTS.md` → `SDD.md` → ADRs.
- Costo: unos minutos por decisión. Skill de apoyo: `.agents/skills/adr-sdd-mantenimiento`.
