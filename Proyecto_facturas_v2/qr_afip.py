# qr_afip.py — Lectura del QR de comprobantes electrónicos AFIP/ARCA (MEJ-01).
#
# El QR de AFIP es una URL del tipo https://www.afip.gob.ar/fe/qr/?p=<BASE64>
# (o arca.gob.ar), donde <BASE64> es un JSON con los datos del comprobante:
#   {"ver":1,"fecha":"2026-10-01","cuit":30712345678,"ptoVta":2,"tipoCmp":1,
#    "nroCmp":47336,"importe":12100.00,"moneda":"PES","ctz":1,
#    "tipoDocRec":80,"nroDocRec":30696428898,"tipoCodAut":"E","codAut":70417054367476}
#
# Todo es local y gratuito: no se consulta a AFIP ni a ninguna IA.
# Decodificadores soportados (se usa el primero que esté instalado):
#   1) zxing-cpp  (pip install zxing-cpp)    ← recomendado, sin DLLs externas
#   2) pyzbar     (pip install pyzbar)
#   3) OpenCV     (pip install opencv-python)
# Si no hay ninguno, leer_qr_afip() devuelve None y el sistema sigue con regex.
import base64
import json
import re
from datetime import datetime
from urllib.parse import parse_qs, urlparse

try:
    import pymupdf as fitz  # PyMuPDF >= 1.24
except ImportError:  # versiones viejas
    import fitz

from pdf_rules import (TIPOS_AFIP_ACEPTADOS, TIPOS_AFIP_NOTA_CREDITO,
                       TIPOS_AFIP_NOTA_DEBITO, TIPOS_AFIP_RECIBO)

_DPI_QR = 200


def _decodificar_imagen(pix):
    """Devuelve la lista de textos de los QR encontrados en un pixmap de PyMuPDF."""
    try:
        import numpy as np
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
    except Exception:
        return []

    try:
        import zxingcpp
        return [r.text for r in zxingcpp.read_barcodes(img) if r.text]
    except ImportError:
        pass
    except Exception:
        return []

    try:
        from pyzbar.pyzbar import decode
        from PIL import Image
        modo = "RGB" if pix.n == 3 else "L"
        pil = Image.frombytes(modo, (pix.w, pix.h), pix.samples)
        return [r.data.decode("utf-8", "ignore") for r in decode(pil)]
    except ImportError:
        pass
    except Exception:
        return []

    try:
        import cv2
        ok, textos, _, _ = cv2.QRCodeDetector().detectAndDecodeMulti(img)
        return [t for t in textos if t] if ok else []
    except Exception:
        return []


def hay_decodificador_qr():
    for modulo in ("zxingcpp", "pyzbar.pyzbar", "cv2"):
        try:
            __import__(modulo)
            return True
        except Exception:
            continue
    return False


def parsear_url_qr(texto_qr):
    """
    Convierte el texto de un QR AFIP en un dict con los datos del comprobante.
    Devuelve None si el texto no es un QR de AFIP/ARCA válido.
    """
    if not texto_qr:
        return None
    texto_qr = texto_qr.strip()
    if "afip.gob.ar" not in texto_qr.lower() and "arca.gob.ar" not in texto_qr.lower():
        return None

    try:
        params = parse_qs(urlparse(texto_qr).query)
        p = params.get("p", [""])[0]
        if not p:
            m = re.search(r"[?&]p=([A-Za-z0-9+/=_\-]+)", texto_qr)
            p = m.group(1) if m else ""
        if not p:
            return None
        p = p.replace(" ", "+")  # parse_qs convierte '+' en espacio
        p += "=" * (-len(p) % 4)
        crudo = base64.urlsafe_b64decode(p) if ("-" in p or "_" in p) else base64.b64decode(p)
        data = json.loads(crudo.decode("utf-8", "ignore"))
    except Exception:
        return None

    if not isinstance(data, dict) or "tipoCmp" not in data:
        return None
    return data


def interpretar_qr(data):
    """Normaliza el JSON del QR a los campos que usa el procesador."""
    try:
        tipo_cmp = int(data.get("tipoCmp"))
    except (TypeError, ValueError):
        tipo_cmp = None

    fecha = ""
    f_raw = str(data.get("fecha", "")).strip()
    for fmt in ("%Y-%m-%d", "%Y%m%d", "%d/%m/%Y"):
        try:
            fecha = datetime.strptime(f_raw[:10], fmt).strftime("%d/%m/%Y")
            break
        except ValueError:
            continue

    if tipo_cmp in TIPOS_AFIP_NOTA_CREDITO:
        tipo = "NC"
    elif tipo_cmp in TIPOS_AFIP_NOTA_DEBITO:
        tipo = "ND"
    elif tipo_cmp in TIPOS_AFIP_RECIBO:
        tipo = "RECIBO"
    else:
        tipo = "FC"

    cuit = "".join(ch for ch in str(data.get("cuit", "")) if ch.isdigit())
    nro = str(data.get("nroCmp", "")).strip()
    try:
        nro = str(int(float(nro)))
    except ValueError:
        pass

    importe = data.get("importe", "")
    try:
        importe = float(importe)
    except (TypeError, ValueError):
        importe = None

    return {
        "tipo_cmp": tipo_cmp,
        "tipo": tipo,
        "es_aceptado": tipo_cmp in TIPOS_AFIP_ACEPTADOS,
        "es_recibo": tipo_cmp in TIPOS_AFIP_RECIBO,
        "cuit": cuit,
        "punto_venta": str(data.get("ptoVta", "")),
        "numero": nro,
        "fecha_emision": fecha,
        "importe_total": importe,
        "moneda": str(data.get("moneda", "")),
        "cae": str(data.get("codAut", "")),
        "cuit_receptor": "".join(ch for ch in str(data.get("nroDocRec", "")) if ch.isdigit()),
    }


def leer_qr_afip(ruta_pdf, max_paginas=2):
    """
    Busca un QR AFIP en las primeras páginas del PDF.
    Devuelve el dict de interpretar_qr() o None si no hay QR / no hay decodificador.
    """
    if not hay_decodificador_qr():
        return None
    try:
        doc = fitz.open(ruta_pdf)
    except Exception:
        return None
    try:
        for i in range(min(len(doc), max_paginas)):
            pix = doc[i].get_pixmap(dpi=_DPI_QR)
            for texto in _decodificar_imagen(pix):
                data = parsear_url_qr(texto)
                if data:
                    return interpretar_qr(data)
    except Exception:
        return None
    finally:
        doc.close()
    return None
