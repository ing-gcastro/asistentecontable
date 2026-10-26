# ADR-0002: Monorepo con carpetas por dominio (y no 3 repos)

- **Estado:** Aceptado
- **Fecha:** 2026-10-26

## Contexto
La raíz del repo mezclaba el núcleo de facturas con scripts de tesorería, flota y limpieza de datos, más archivos generados pesados (`datos_oc_telcom.json` de 770 KB, `reporte_oc_telcom.html` de 730 KB, `.xlsx`). Se evaluó separar en 3 repositorios. El análisis de dependencias mostró: (a) `procesar_oc_telcom.py` importa `config` y `proveedores` del núcleo y escanea la misma carpeta de OC que `oc_manager.py`, así que es del dominio **Compras**, no de Tesorería; (b) la cadena de Tesorería (`indexar_legajos_pago` → `Reporte_Legajos_Pago.html` → `vincular_*`) es cerrada; (c) `ejecutar_reportes.bat` encadena Compras y Tesorería; (d) hay rutas absolutas a `Proyecto_facturas_v2` (tarea programada del downloader).

## Decisión
Un solo repo con carpetas por dominio: `Proyecto_facturas_v2/` (núcleo, **no se renombra**), `compras_oc/`, `tesoreria_pagos/`, `datos_maestros/`, `flota_seguros/`. Los movimientos se hacen con `git mv` desde `APLICAR.bat` para conservar el historial. Los datos generados dejan de versionarse (`git rm --cached`, el archivo queda en disco). `REPORTE DE AUTORIZACION/` queda solo con `actualizar.bat`, que redirige al reporter nuevo, por si una tarea de inicio de Windows lo usa.

## Alternativas consideradas
- **3 repos (A):** correcto a largo plazo, pero hoy obliga a crear repos, filtrar historial y coordinar rutas relativas entre repos; demasiado riesgo para un equipo chico. Queda habilitado: `git subtree split -P <carpeta>`.
- **Promover el núcleo a la raíz:** rompe la tarea programada y los notebooks sin aviso.

## Consecuencias
- El QA ve claramente qué es núcleo y qué es satélite.
- Los accesos directos a los `.bat` movidos (legajos, OC, mails, talones) hay que actualizarlos (lista en RUNBOOK).
- Separar luego es un comando.
