---
name: ocr-pdf-processor
description: >
  Implementa soporte OCR automático para PDFs escaneados o imagen en el sistema
  de procesamiento de facturas. Usar cuando: el archivo procesador_core.py lanza
  "PDF vacío o imagen", cuando se menciona pytesseract o pdf2image, cuando hay
  facturas que no se pueden leer, o cuando se pide agregar OCR al proyecto.
  También aplica cuando se detecta texto vacío en fitz.get_text().
---

# OCR PDF Processor — Skill de implementación

Este skill implementa OCR automático como fallback en `procesador_core.py`.

## Contexto del problema

`fitz.get_text()` (PyMuPDF) solo extrae texto de PDFs con capa de texto digital.
PDFs escaneados o enviados como imagen devuelven string vacío.
El sistema actual lanza un `input()` manual, colgando el proceso si no hay operador.

## Stack tecnológico a usar

- `pdf2image` — convierte páginas PDF a imágenes PIL
- `pytesseract` — wrapper de Python para Tesseract OCR
- Tesseract con paquete de español (`spa.traineddata`)
- Windows: Tesseract portable en `C:\Program Files\Tesseract-OCR\`

## Pasos de implementación

### Paso 1: Agregar dependencias a requirements.txt

Añadir al archivo `requirements.txt`:
```
pymupdf>=1.24.0
pytesseract>=0.3.13
pdf2image>=1.17.0
pandas>=2.2.0
pywin32>=306
openpyxl>=3.1.0
Pillow>=10.0.0
```

### Paso 2: Agregar función OCR a procesador_core.py

Insertar DESPUÉS de los imports existentes, ANTES de `es_factura_valida()`:

```python
def extraer_texto_con_ocr(ruta_pdf: str) -> str:
    try:
        from pdf2image import convert_from_path
        import pytesseract
        print("   📷 Activando OCR (pytesseract + pdf2image)...")
        paginas = convert_from_path(ruta_pdf, dpi=300)
        texto = "\n".join([pytesseract.image_to_string(p, lang="spa") for p in paginas])
        return texto.strip()
    except ImportError:
        print("   ⚠️ OCR no disponible. pip install pytesseract pdf2image")
        return ""
    except Exception as e:
        print(f"   ⚠️ Error OCR: {e}")
        return ""
```

### Paso 3: Modificar extraer_datos_pdf() en procesador_core.py

Reemplazar el check de texto vacío:
```python
if not texto_lower.strip():
    print(f"   ℹ️ Sin texto digital. Intentando OCR automático...")
    texto_original = extraer_texto_con_ocr(ruta_pdf)
    if not texto_original.strip():
        return False, "PDF vacío o imagen (OCR también falló)"
    texto_lower = texto_original.lower()
    print(f"   ✅ OCR exitoso: {len(texto_original)} caracteres")
```

## Criterios de éxito

- PDFs escaneados se procesan sin intervención manual
- El mensaje "PDF vacío o imagen (OCR también falló)" solo aparece para PDFs realmente corruptos
- El log muestra "OCR exitoso: N caracteres extraídos" para facturas escaneadas
- La extracción de CAE, CUIT y número de comprobante funciona después del OCR
