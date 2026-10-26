---
name: facturas-nuevo-proveedor
description: >
  Incorpora reglas para un proveedor nuevo o con formato especial: remitente confiable,
  fusión de varios adjuntos, conversión de Excel adjunto a PDF, ubicación/sector especial (tipo Ecosan) o lista negra.
  Usar cuando se menciona un proveedor concreto que llega mal, en varios archivos o con un formato raro.
---

# Proveedor nuevo / regla especial

| Caso | Dónde |
|---|---|
| Sus facturas no tienen CAE legible pero son confiables | `proveedor_rules.REMITENTES_CONFIABLES` (email completo) |
| Manda factura + anexos en PDFs separados (unir todo) | `downloader.REMITENTES_FUSION_TODO` |
| Manda PDF + Excel de detalle (convertir y unir) | `downloader.REMITENTES_EXCEL_A_PDF` |
| Sector/locación según el contenido (como Ecosan) | `proveedor_rules.analizar_pdf_ecosan` (copiar el patrón) y su uso en `procesador_core` |
| No se le deben aprobar facturas | Maestro de proveedores: la palabra "LISTA NEGRA" en el nombre |
| Sector por defecto | `Sectores.xlsx` (columnas `Numero_Documento`, `Sector`) |

Siempre: test con un texto de ejemplo del proveedor + `CHANGELOG.md`.
