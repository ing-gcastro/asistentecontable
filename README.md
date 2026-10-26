# Asistente Contable — Recepción y procesamiento de facturas de compras

Automatiza el ciclo de facturas de proveedores: **Outlook → validación (QR AFIP / CAE) → Excel de revisión →
archivo en red por sector → reportes → aviso a los sectores**. 100% local, sin IA ni servicios pagos.

## Uso rápido
| Qué | Cómo |
|---|---|
| Instalar o actualizar dependencias | doble clic en **`INSTALAR.bat`** |
| Operar (descarga, Paso 1, Paso 2, reportes, monitor) | doble clic en **`MENU.bat`** |
| Comprobar que todo está bien | doble clic en **`VERIFICAR.bat`** |

Paso a paso diario, tareas programadas, respaldos y problemas frecuentes: [`docs/RUNBOOK.md`](docs/RUNBOOK.md).

## Estructura
```
asistentecontable/
├── MENU.bat · INSTALAR.bat · VERIFICAR.bat      ← puntos de entrada
├── ejecutar_*.bat                               ← accesos directos del ciclo de compras
├── Proyecto_facturas_v2/                        ← NÚCLEO (no renombrar: lo usan tareas programadas)
│   ├── config.py              rutas y constantes (overrides en config_local.py)
│   ├── pdf_rules.py           reglas de clasificación y regex
│   ├── qr_afip.py             lector de QR AFIP
│   ├── procesador_core.py     extracción y validación
│   ├── 1_preparar_lote.py     PASO 1 → Validacion_Lote.xlsx
│   ├── 2_procesar_lote.py     PASO 2 → archivo + reportes + aviso
│   ├── reporter.py            Reporte_Maestro / sectores / Reporte_Detallado
│   ├── monitor_compras.py     regenera reportes ante cambios en \Compras
│   ├── notificador_sectores.py mail por sector (desactivado hasta definir SDD §8)
│   ├── downloader.py · auto_downloader.py · oc_manager.py · proveedores.py · estado.py
│   └── tests/                 tests automáticos (sin red ni Outlook)
├── compras_oc/                ← índice de Órdenes de Compra TELCOM
├── tesoreria_pagos/           ← legajos de pago y vinculación de avisos
├── datos_maestros/            ← limpieza del maestro de proveedores (ad-hoc)
├── flota_seguros/             ← división de pólizas por patente
├── REPORTE DE AUTORIZACION/   ← solo actualizar.bat (compatibilidad)
├── scripts/_python.bat        ← detecta el Python a usar
├── docs/                      ← SDD, ADR, RUNBOOK, QA
└── .agents/                   ← skills y workflows para agentes de IA (Antigravity/Gemini, Claude)
```

## Documentación
- Diseño: [`docs/SDD.md`](docs/SDD.md) · Decisiones: [`docs/adr/`](docs/adr/README.md) · Cambios: [`CHANGELOG.md`](CHANGELOG.md)
- Estado del QA: [`docs/qa/ESTADO_QA.md`](docs/qa/ESTADO_QA.md)
- Para agentes de IA: [`AGENTS.md`](AGENTS.md)
