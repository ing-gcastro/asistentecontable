---
name: architecture-reviewer
description: >
  Revisa y propone mejoras de arquitectura y diseño del proyecto asistentecontable.
  Usar cuando: se pide revisar la arquitectura general, cuando se habla de separar
  responsabilidades, cuando se quiere hacer el proyecto más testeable o mantenible,
  cuando se menciona diseño de código o patrones, o cuando se quiere preparar el
  proyecto para escalar o agregar nuevas funcionalidades.
---

# Architecture Reviewer — Skill de implementación

## Diagnóstico arquitectónico actual

### Problemas detectados
1. Acoplamiento fuerte entre descarga y procesamiento
2. Side effects en módulos de datos
3. Interactividad mezclada con lógica de negocio
4. No hay capa de abstracción para acceso a datos

## Propuestas

### Propuesta A: Separación de capas (recomendada)
### Propuesta B: Modo batch sin interactividad (alta prioridad)
### Propuesta C: Tests unitarios básicos

## Prioridades

| Prioridad | Mejora | Complejidad | Impacto |
|-----------|--------|-------------|---------|
| P1 | Modo batch sin input() | Media | Alto |
| P1 | Centralizar rutas (config.py) | Baja | Alto |
| P2 | Caché en cargar_sectores() | Baja | Medio |
| P2 | Separar extracción en módulo | Media | Medio |
| P3 | Tests unitarios básicos | Alta | Alto |
| P3 | Separación completa de capas | Alta | Alto |
