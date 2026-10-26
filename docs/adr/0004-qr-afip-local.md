# ADR-0004: QR AFIP local como fuente primaria de datos (MEJ-01)

- **Estado:** Aceptado
- **Fecha:** 2026-10-26

## Contexto
Las regex sobre el texto del PDF son frágiles (formatos distintos por proveedor, PDF escaneados). Los comprobantes electrónicos traen un QR con un JSON (CUIT, tipo, punto de venta, número, fecha, importe, CAE). Restricción: sin costo de IA.

## Decisión
`qr_afip.py` renderiza las 2 primeras páginas con PyMuPDF (200 dpi) y decodifica con **zxing-cpp** (wheel puro para Windows, sin DLL externas). Fallbacks: pyzbar, OpenCV. Si hay QR: manda sobre el texto para tipo, número, CUIT, fecha y total; tipos de recibo (4, 9, 15, 54) se rechazan. Si no hay QR o no hay decodificador instalado, el flujo sigue con regex sin error.

## Alternativas consideradas
- **pyzbar:** necesita `libzbar` y a veces el redistribuible VC++ 2013 en Windows.
- **OpenCV:** ~40 MB y menos tolerante con QR chicos.
- **Consultar AFIP online:** requiere certificado y conexión; queda en el backlog.

## Consecuencias
- Datos exactos cuando hay QR; el Excel muestra `QR_AFIP = SI/NO`.
- Costo: ~0,2-0,5 s por PDF.
