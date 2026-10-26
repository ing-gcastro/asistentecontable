# procesador_core.py — Extracción y validación de comprobantes (núcleo compartido).
#
# Lo usan: 1_preparar_lote.py (Paso 1), 2_procesar_lote.py (Paso 2), reporter.py
# y el modo interactivo legado (iniciar_proceso_interactivo, usado desde notebooks).
# Este módulo NO debe importar Outlook (win32com) ni el reporter al cargarse:
# esos imports se hacen dentro de las funciones que los necesitan.
import os
import re
import shutil
import unicodedata
from datetime import datetime

try:
    import pymupdf as fitz  # PyMuPDF >= 1.24
except ImportError:  # versiones viejas
    import fitz
import pandas as pd

from config import (ARCHIVO_PROVEEDORES, CARPETA_A, CARPETA_AUTORIZACION, CARPETA_B,
                    CARPETA_C, MESES_ESPANOL, REPORTE_MAESTRO_HTML)
from pdf_rules import REGLAS_EXTRACCION
from proveedor_rules import analizar_pdf_ecosan, es_factura_por_remitente
from proveedores import buscar_proveedor, cargar_excel_proveedores
import oc_manager
import qr_afip

_RX_CUIT_LIBRE = r"\b((?:20|23|24|27|30|33|34)[-\s]?\d{2}[-\s]?\d{6}[-\s]?\d{1})\b"
_CUITS_CLIENTE = {REGLAS_EXTRACCION["cuit_propio"], REGLAS_EXTRACCION["cuit_eurosat"]}


# ─── UTILIDADES ──────────────────────────────────────────────────────────────
def normalizar_texto(texto):
    """Minúsculas y sin acentos ('Válido' -> 'valido'). Conserva saltos de línea."""
    if not texto:
        return ""
    sin_acentos = unicodedata.normalize("NFKD", texto)
    sin_acentos = "".join(c for c in sin_acentos if not unicodedata.combining(c))
    return sin_acentos.lower()


def parsear_monto(valor):
    """
    Convierte un importe en texto a float. Acepta formato argentino (1.234,56),
    anglosajón (1,234.56) y simple (1234.56 / 1234,56). Devuelve None si no puede.
    """
    if valor is None:
        return None
    if isinstance(valor, (int, float)):
        return None if pd.isna(valor) else float(valor)
    s = str(valor).strip().replace("$", "").replace(" ", "")
    if not s or s.lower() == "nan":
        return None
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".") if len(s.split(",")[-1]) != 3 else s.replace(",", "")
    elif s.count(".") > 1 or (s.count(".") == 1 and len(s.split(".")[-1]) == 3):
        s = s.replace(".", "")
    try:
        return round(float(s), 2)
    except ValueError:
        return None


def nombre_mes_desde_fecha(fecha_emision, defecto=None):
    """
    BUG-01: el mes del nombre de archivo sale de la fecha de EMISIÓN de la factura.
    Si la fecha no se pudo extraer, usa `defecto` (por omisión, el mes actual).
    """
    if defecto is None:
        defecto = MESES_ESPANOL.get(datetime.now().month, "MES")
    if fecha_emision:
        for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
            try:
                return MESES_ESPANOL.get(datetime.strptime(str(fecha_emision).strip()[:10], fmt).month, defecto)
            except ValueError:
                continue
    return defecto


def _buscar_monto(clave, texto, ultimo=False):
    patron = REGLAS_EXTRACCION.get(clave)
    if not patron:
        return None, ""
    hallados = re.findall(patron, texto, re.IGNORECASE | re.MULTILINE)
    if not hallados:
        return None, ""
    crudo = hallados[-1] if ultimo else hallados[0]
    return parsear_monto(crudo), crudo.strip()


def extraer_items(texto, max_items=30):
    """
    Heurística para líneas de detalle (descripción, cantidad, precio unitario, subtotal).
    El formato varía por proveedor: si no reconoce ninguna línea devuelve [].
    """
    items = []
    for m in re.finditer(REGLAS_EXTRACCION["rx_item"], texto, re.MULTILINE):
        desc = re.sub(r"\s{2,}", " ", m.group("desc")).strip(" .-:")
        desc_norm = normalizar_texto(desc)
        if any(p in desc_norm for p in ("subtotal", "total", "iva", "percep", "neto", "importe", "bonif")):
            continue
        cant = parsear_monto(m.group("cant"))
        unit = parsear_monto(m.group("unit"))
        sub = parsear_monto(m.group("sub"))
        if not cant or unit is None:
            continue
        items.append({"descripcion": desc[:80], "cantidad": cant, "precio_unitario": unit, "subtotal": sub})
        if len(items) >= max_items:
            break
    return items


# ─── OCR FALLBACK ────────────────────────────────────────────────────────────
def extraer_texto_con_ocr(ruta_pdf: str, max_paginas: int = 3) -> str:
    """Fallback OCR para PDFs escaneados o imagen usando RapidOCR + PyMuPDF (limitado a max_paginas)."""
    print("   📷 Activando OCR (RapidOCR + PyMuPDF)...")
    texto_extraido = []
    try:
        from rapidocr_onnxruntime import RapidOCR
        engine = RapidOCR()
        doc = fitz.open(ruta_pdf)
        num_paginas = min(len(doc), max_paginas)
        for page_num in range(num_paginas):
            pix = doc[page_num].get_pixmap(dpi=200)
            result, _ = engine(pix.tobytes("png"))
            if result:
                texto_extraido.append("\n".join([line[1] for line in result]))
        doc.close()
        res = "\n".join(texto_extraido).strip()
        if res:
            print(f"   ✅ RapidOCR exitoso: {len(res)} caracteres extraídos")
            return res
    except Exception as e:
        print(f"   ⚠️ RapidOCR no disponible o falló: {e}")

    # Fallback secundario pytesseract (requiere Tesseract + Poppler instalados en Windows)
    try:
        from pdf2image import convert_from_path
        import pytesseract
        paginas = convert_from_path(ruta_pdf, dpi=200, first_page=1, last_page=max_paginas)
        texto = "\n".join([pytesseract.image_to_string(p, lang="spa") for p in paginas])
        return texto.strip()
    except Exception as e:
        print(f"   ⚠️ Error OCR secundario: {e}")
        return ""
# ─────────────────────────────────────────────────────────────────────────────


def tiene_rechazo_explicito(texto):
    """True si el texto dice explícitamente que NO es factura (recibos, etc.)."""
    t = normalizar_texto(texto)
    return any(frase in t for frase in REGLAS_EXTRACCION["rechazo_explicito"])


def tiene_cae(texto):
    t = normalizar_texto(texto)
    if re.search(r"\bc[.\s]*a[.\s]*e[.\s]*[a\s:\.\|n°º]*\n?\s*\|?\s*(\d{14,15})\b", t):
        return True
    if re.search(r"\b[678]\d{13}\b", t):
        return True
    return bool(re.search(r"\bcae\b\s*[:\.]?\s*\d{14}", t, re.MULTILINE))


def es_factura_valida(texto_lower, remitente="", qr=None, texto_rechazo=None):
    """
    Decide si el texto corresponde a una factura / NC / ND válida.

    Orden de decisión (BUG-02 + MEJ-01):
      1. QR AFIP leído: manda el tipo de comprobante (factura/NC/ND sí; recibo no).
      2. Rechazo explícito ("documento no válido como factura", "recibo:" ...): NO,
         aunque el remitente sea de confianza o haya números parecidos a un CAE.
      3. Remitente de confianza: SÍ (relaja solo la exigencia de CAE).
      4. Palabras prohibidas sin CAE: NO.
      5. CAE detectado: SÍ.
      6. Fallback: dice "factura" y tiene número de comprobante AFIP.

    `texto_rechazo` permite evaluar el rechazo explícito solo sobre la primera
    página (en PDFs fusionados, un recibo adjunto no debe tumbar la factura).
    """
    if qr:
        if qr.get("es_aceptado"):
            return True
        if qr.get("es_recibo"):
            return False

    t = normalizar_texto(texto_lower)
    t_rechazo = normalizar_texto(texto_rechazo) if texto_rechazo is not None else t

    if any(frase in t_rechazo for frase in REGLAS_EXTRACCION["rechazo_explicito"]):
        return False

    if es_factura_por_remitente(remitente):
        return True

    texto_seguro = t
    for frase in REGLAS_EXTRACCION["frases_seguras"]:
        texto_seguro = texto_seguro.replace(frase, "")

    has_cae_valido = tiene_cae(t)

    if not has_cae_valido and any(p in texto_seguro for p in REGLAS_EXTRACCION["palabras_prohibidas"]):
        return False

    if has_cae_valido:
        return True

    if re.search(r"\bfactura\b", t) and re.search(r"\b\d{4,5}[-\s]\d{7,8}\b", t):
        if "vista de deuda" not in t and "cuenta corriente" not in t:
            return True

    return False


def _extraer_numero(texto_original):
    m_exacto = re.search(r"\b\d{4,5}\s*[-]\s*(\d{8})\b", texto_original)
    if m_exacto:
        return str(int(m_exacto.group(1)))
    m_letra = re.search(r"\b\d{4,5}[A-Za-z](\d{8})\b", texto_original)
    if m_letra:
        return str(int(m_letra.group(1)))
    m_arca = re.search(r"(?:Comp\.?\s*Nro\.?|Nro\.?|Número)\s*[:\.]?\s*(\d{4,8})", texto_original, re.IGNORECASE)
    if m_arca:
        return str(int(m_arca.group(1)))
    m_comp = re.search(REGLAS_EXTRACCION["rx_numero_comp"], texto_original, re.IGNORECASE)
    if m_comp:
        nro_crudo = m_comp.group(1).replace("-", "").replace(" ", "")
        return str(int(nro_crudo[-8:]))
    m_guion = re.search(REGLAS_EXTRACCION["rx_numero_guion"], texto_original)
    if m_guion:
        return str(int(m_guion.group(1)))
    return "SinNumero"


def _extraer_razon_social(texto_original):
    razon_social = ""
    # Primera línea con forma societaria que no sea la empresa propia / Eurosat (clientes)
    for m_razon in re.finditer(REGLAS_EXTRACCION["rx_razon_social"], texto_original, re.IGNORECASE | re.MULTILINE):
        rs_raw = m_razon.group(1).strip().upper()
        for basura in REGLAS_EXTRACCION["basura_afip"]:
            if basura in rs_raw:
                rs_raw = rs_raw.split(basura)[0]
        candidato = rs_raw.strip(" .,-")
        if len(candidato) < 4 or "TELCOM VENTURES" in candidato or "EUROSAT" in candidato:
            continue
        razon_social = rs_raw.strip()
        break

    if not razon_social:
        m_etiqueta = re.search(REGLAS_EXTRACCION["rx_razon_social_etiqueta"], texto_original, re.IGNORECASE)
        if m_etiqueta:
            rs_raw = m_etiqueta.group(1).strip().upper()
            for basura in REGLAS_EXTRACCION["basura_afip"]:
                if basura in rs_raw:
                    rs_raw = rs_raw.split(basura)[0]
            rs_candidate = rs_raw.strip()
            if rs_candidate and "TELCOM VENTURES" not in rs_candidate and "EUROSAT" not in rs_candidate:
                razon_social = rs_candidate
    return razon_social


def cuit_valido(cuit):
    """Dígito verificador de CUIT/CUIL (módulo 11)."""
    c = "".join(filter(str.isdigit, str(cuit)))
    if len(c) != 11:
        return False
    suma = sum(int(d) * p for d, p in zip(c[:10], (5, 4, 3, 2, 7, 6, 5, 4, 3, 2)))
    dv = 11 - (suma % 11)
    dv = 0 if dv == 11 else (9 if dv == 10 else dv)
    return dv == int(c[10])


def _cuits_proveedor(texto):
    """CUITs candidatos (sin los del cliente), primero los que tienen dígito verificador válido."""
    encontrados = []
    for c in re.findall(_RX_CUIT_LIBRE, texto):
        c_limpio = "".join(filter(str.isdigit, c))
        if len(c_limpio) == 11 and c_limpio not in _CUITS_CLIENTE and c_limpio not in encontrados:
            encontrados.append(c_limpio)
    return sorted(encontrados, key=lambda x: not cuit_valido(x))  # sort estable: conserva el orden de aparición


def extraer_importes(texto):
    """MEJ-02: neto, IVA 21/10.5/27, otros impuestos y total (floats o None)."""
    neto, subtotal_crudo = _buscar_monto("rx_subtotal", texto)
    iva_21, _ = _buscar_monto("rx_iva_21", texto)
    iva_105, _ = _buscar_monto("rx_iva_105", texto)
    iva_27, _ = _buscar_monto("rx_iva_27", texto)

    otros = None
    for crudo in re.findall(REGLAS_EXTRACCION["rx_percepciones"], texto, re.IGNORECASE | re.MULTILINE):
        v = parsear_monto(crudo)
        if v:
            otros = round((otros or 0) + v, 2)

    total, _ = _buscar_monto("rx_total_especifico", texto, ultimo=True)
    if total is None:
        total, _ = _buscar_monto("rx_total", texto, ultimo=True)

    return {
        "subtotal": subtotal_crudo,      # texto tal cual (compatibilidad con la columna histórica)
        "neto": neto,
        "iva_21": iva_21,
        "iva_105": iva_105,
        "iva_27": iva_27,
        "otros_impuestos": otros,
        "total": total,
    }


def extraer_datos_pdf(ruta_pdf, remitente="", leer_qr=True):
    """
    Lee un PDF y devuelve (True, datos) si es un comprobante válido,
    o (False, motivo) si no lo es / no se pudo leer.
    """
    try:
        doc = fitz.open(ruta_pdf)
        num_paginas = len(doc)
        paginas = [p.get_text("text", sort=True) for p in doc]
        doc.close()
    except Exception as e:
        return False, f"Error de lectura: {e}"

    try:
        texto_p1 = paginas[0] if paginas else ""
        texto_cabecera = "\n".join(paginas[:3])
        texto_original = "\n".join(paginas)

        qr = qr_afip.leer_qr_afip(ruta_pdf) if leer_qr else None
        if qr:
            print(f"   🔳 QR AFIP leído: tipo {qr['tipo_cmp']} | CUIT {qr['cuit']} | Nro {qr['numero']} | {qr['fecha_emision']}")

        ocr_ejecutado = False
        es_valida = (es_factura_valida(texto_cabecera, remitente, qr, texto_rechazo=texto_p1)
                     or es_factura_valida(texto_original, remitente, qr, texto_rechazo=texto_p1))

        texto_escaso = len(texto_cabecera.strip()) < 50 or len(texto_original.strip()) < 50
        if not es_valida and texto_escaso and not tiene_rechazo_explicito(texto_p1):
            print("   ℹ️ Texto escaso o sin validación previa. Intentando OCR automático...")
            texto_ocr = extraer_texto_con_ocr(ruta_pdf)
            if texto_ocr.strip():
                texto_original = texto_original + "\n" + texto_ocr
                texto_cabecera = texto_cabecera + "\n" + texto_ocr
                texto_p1 = texto_p1 + "\n" + texto_ocr
                ocr_ejecutado = True
                es_valida = es_factura_valida(texto_original, remitente, qr, texto_rechazo=texto_p1)
            elif not texto_original.strip():
                return False, "PDF vacío o imagen (OCR también falló)"

        if not es_valida:
            if qr and qr.get("es_recibo"):
                return False, "Documento descartado (el QR AFIP indica que es un RECIBO, no una factura)"
            if tiene_rechazo_explicito(texto_p1):
                return False, "Documento descartado (dice 'no válido como factura' o es un recibo)"
            if num_paginas > 3:
                return False, "Documento descartado (Tiene más de 3 páginas y no contiene CAE/Factura - Catálogo/Manual)"
            return False, "Documento descartado (No tiene CAE o contiene palabras prohibidas)"

        # ─── TIPO Y NÚMERO ────────────────────────────────────────────────
        if qr and qr.get("tipo") in ("FC", "NC", "ND"):
            tipo = qr["tipo"]
        elif re.search(REGLAS_EXTRACCION["rx_nota_credito"], normalizar_texto(texto_original), re.IGNORECASE):
            tipo = "NC"
        elif re.search(REGLAS_EXTRACCION["rx_nota_debito"], normalizar_texto(texto_original), re.IGNORECASE):
            tipo = "ND"
        else:
            tipo = "FC"

        nro = qr["numero"] if qr and qr.get("numero") else _extraer_numero(texto_original)
        razon_social = _extraer_razon_social(texto_original)

        # ─── CUIT DEL PROVEEDOR ───────────────────────────────────────────
        cuit_proveedor = qr["cuit"] if qr and qr.get("cuit") and qr["cuit"] not in _CUITS_CLIENTE else ""
        if not cuit_proveedor:
            posibles = _cuits_proveedor(texto_original)
            cuit_proveedor = posibles[0] if posibles else ""

        if not cuit_proveedor and not ocr_ejecutado:
            print("   ℹ️ CUIT del proveedor no hallado en texto digital. Forzando OCR completo...")
            texto_ocr = extraer_texto_con_ocr(ruta_pdf)
            if texto_ocr.strip():
                texto_original += "\n" + texto_ocr
                posibles = _cuits_proveedor(texto_original)
                cuit_proveedor = posibles[0] if posibles else ""

        # ─── FECHA DE EMISIÓN Y EUROSAT ───────────────────────────────────
        fecha_emision_str = qr["fecha_emision"] if qr and qr.get("fecha_emision") else ""
        if not fecha_emision_str:
            m_fecha = re.search(REGLAS_EXTRACCION["rx_fecha_emision"], texto_original, re.IGNORECASE)
            if m_fecha:
                fecha_emision_str = m_fecha.group(1).replace("-", "/")

        es_eurosat = bool(re.search(r"\bEUROSAT\b", texto_original, re.IGNORECASE)) or \
            bool(re.search(r"33[-\s]?67654963[-\s]?9", texto_original))
        if qr and qr.get("cuit_receptor") == REGLAS_EXTRACCION["cuit_eurosat"]:
            es_eurosat = True

        # ─── OC, IMPORTES E ÍTEMS ─────────────────────────────────────────
        m_oc = re.search(REGLAS_EXTRACCION["rx_oc"], texto_original, re.IGNORECASE)
        oc_encontrada = m_oc.group(1).strip() if m_oc else ""

        importes = extraer_importes(texto_original)
        if qr and qr.get("importe_total") is not None:
            importes["total"] = round(qr["importe_total"], 2)

        es_tarjeta = bool(re.search(r"\b(visa\s*corpo|visacorpo|visa|tarjeta\s*de\s*credito|pago\s*con\s*tarjeta)\b",
                                    texto_original, re.IGNORECASE))

        datos = {
            "tipo": tipo,
            "numero": nro,
            "razon_social": razon_social,
            "cuit": cuit_proveedor,
            "fecha_emision": fecha_emision_str,
            "es_eurosat": es_eurosat,
            "es_tarjeta": es_tarjeta,
            "oc_extraida": oc_encontrada,
            "items": extraer_items(texto_original),
            "qr_afip": bool(qr),
            "cae": qr.get("cae", "") if qr else "",
        }
        datos.update(importes)
        return True, datos
    except Exception as e:
        return False, f"Error de lectura: {str(e)}"


def extraer_datos_livianos(ruta_pdf, leer_qr_si_falta=True):
    """
    Lectura rápida para reportes de PDFs ya archivados (sin validar, sin OCR):
    fecha de emisión, CUIT, importes e ítems. Devuelve {} si no se puede leer.
    """
    try:
        doc = fitz.open(ruta_pdf)
        texto = "\n".join(doc[i].get_text("text", sort=True) for i in range(min(3, len(doc))))
        doc.close()
    except Exception:
        return {}

    datos = {"fecha_emision": "", "cuit": ""}
    m_fecha = re.search(REGLAS_EXTRACCION["rx_fecha_emision"], texto, re.IGNORECASE)
    if m_fecha:
        datos["fecha_emision"] = m_fecha.group(1).replace("-", "/")
    posibles = _cuits_proveedor(texto)
    datos["cuit"] = posibles[0] if posibles else ""
    datos.update(extraer_importes(texto))
    datos["items"] = extraer_items(texto)

    if leer_qr_si_falta and (not datos["fecha_emision"] or datos.get("total") is None):
        qr = qr_afip.leer_qr_afip(ruta_pdf)
        if qr:
            datos["fecha_emision"] = datos["fecha_emision"] or qr.get("fecha_emision", "")
            datos["cuit"] = datos["cuit"] or qr.get("cuit", "")
            if datos.get("total") is None and qr.get("importe_total") is not None:
                datos["total"] = round(qr["importe_total"], 2)
            datos["qr_afip"] = True
    return datos


def fusionar_pdfs_adjuntos(lista_rutas_pdf, ruta_salida):
    doc_salida = fitz.open()
    for ruta in lista_rutas_pdf:
        try:
            doc_parcial = fitz.open(ruta)
            if doc_parcial.is_pdf:
                doc_salida.insert_pdf(doc_parcial)
            else:
                pdfbytes = doc_parcial.convert_to_pdf()
                img_pdf = fitz.open("pdf", pdfbytes)
                doc_salida.insert_pdf(img_pdf)
                img_pdf.close()
            doc_parcial.close()
        except Exception as e:
            print(f"⚠️ No se pudo fusionar {ruta}: {e}")
    doc_salida.save(ruta_salida)
    doc_salida.close()


def obtener_ruta_unica(ruta):
    if not os.path.exists(ruta):
        return ruta
    nombre, ext = os.path.splitext(ruta)
    c = 1
    while os.path.exists(f"{nombre} ({c}){ext}"):
        c += 1
    return f"{nombre} ({c}){ext}"


def obtener_ultima_fecha_procesada(ruta_html):
    """Última fecha de CARGA registrada en el reporte (para acotar la búsqueda en Outlook)."""
    if not os.path.exists(ruta_html):
        return None
    try:
        for df_existente in pd.read_html(ruta_html):
            columnas = [str(c) for c in df_existente.columns]
            col_fecha = next((c for c in columnas if c.lower() == "fecha carga"), None) or \
                next((c for c in columnas if c.lower() == "fecha"), None)
            if col_fecha and not df_existente.empty:
                fechas = pd.to_datetime(df_existente[col_fecha], dayfirst=True, errors="coerce")
                ultima = fechas.max()
                if pd.notna(ultima):
                    return ultima.strftime('%d/%m/%Y')
    except Exception as e:
        print(f"⚠️ Error leyendo fecha del reporte: {e}")
    return None


def iniciar_proceso_interactivo(ejecutar_descarga=False, usar_filtro_fecha=True, fecha_inicio="01/07/2026", fecha_fin="", filtro_correo=""):
    """Modo interactivo LEGADO (consola, uno por uno). El flujo recomendado es Paso 1 + Paso 2."""
    from reporter import generar_reporte_html, obtener_facturas_procesadas

    ruta_html_reporte = REPORTE_MAESTRO_HTML

    ultima_fecha_historica = obtener_ultima_fecha_procesada(ruta_html_reporte)
    fecha_a_utilizar = ultima_fecha_historica if ultima_fecha_historica else fecha_inicio

    print(f"📅 Rango de búsqueda optimizado desde: {fecha_a_utilizar}")

    mapa_remitentes = {}
    if ejecutar_descarga:
        from downloader import descargar_facturas_outlook
        mapa_remitentes = descargar_facturas_outlook(fecha_inicio=fecha_a_utilizar, filtro_proveedor=filtro_correo)
        print("-" * 50)

    print("📥 Cargando bases de datos de proveedores y sectores...")
    base_prov = cargar_excel_proveedores(ARCHIVO_PROVEEDORES)

    facturas_procesadas_historico = obtener_facturas_procesadas(ruta_html_reporte)
    print(f"🛡️ Filtro activo: Se encontraron {len(facturas_procesadas_historico)} facturas previas registradas en el historial.")

    if not os.path.exists(CARPETA_A):
        print("❌ La Carpeta A no existe.")
        return

    archivos = [f for f in os.listdir(CARPETA_A) if f.lower().endswith(".pdf")]
    if not archivos:
        print("📂 No hay PDFs en la Carpeta A para procesar.")
        return

    os.makedirs(CARPETA_B, exist_ok=True)
    os.makedirs(CARPETA_C, exist_ok=True)

    procesados_contador = 0
    duplicados_omitidos = 0
    processed_files = set()

    for archivo in archivos:
        if archivo in processed_files:
            continue

        ruta_completa = os.path.join(CARPETA_A, archivo)
        remitente_correo = mapa_remitentes.get(archivo, "")

        match_marino = re.search(r'(?:VOLQUETES.*?MARINO|MARINO).*?(\d+)', archivo, re.IGNORECASE)
        archivos_a_fusionar = [ruta_completa]
        if match_marino:
            nro_detectado = match_marino.group(1)
            for otro_arch in archivos:
                if otro_arch != archivo and nro_detectado in otro_arch:
                    archivos_a_fusionar.append(os.path.join(CARPETA_A, otro_arch))
                    processed_files.add(otro_arch)

        es_valido, datos = extraer_datos_pdf(ruta_completa, remitente_correo)

        if not es_valido:
            if "vacío" in datos or "imagen" in datos:
                print(f"\n⚠️ ATENCIÓN: {archivo} parece ser un PDF escaneado o imagen (no se pudo leer el texto).")
                opc = input("¿Deseas procesarlo manualmente como factura? (s/n): ").strip().lower()
                if opc == 's':
                    es_valido = True
                    datos = {"tipo": "FC", "numero": "SinNumero", "razon_social": "PROVEEDOR_DESCONOCIDO", "cuit": ""}
                    m_exacto = re.search(r"\b\d{4,5}\s*[-_]\s*(\d{8})\b", archivo)
                    if m_exacto:
                        datos["numero"] = str(int(m_exacto.group(1)))
                    else:
                        m_arca = re.search(r"(\d{4,8})\.pdf$", archivo, re.IGNORECASE)
                        if m_arca:
                            datos["numero"] = str(int(m_arca.group(1)))

            if not es_valido:
                for r_fusa in archivos_a_fusionar:
                    destino_b = obtener_ruta_unica(os.path.join(CARPETA_B, os.path.basename(r_fusa)))
                    shutil.move(r_fusa, destino_b)
                print(f"[-] {archivo} -> NO ES FACTURA. Movido automáticamente a Carpeta B.")
                continue

        nombre_final, sector, es_negra = buscar_proveedor(datos.get("cuit", ""), datos["razon_social"], base_prov)

        if es_negra:
            for r_fusa in archivos_a_fusionar:
                destino_b = obtener_ruta_unica(os.path.join(CARPETA_B, os.path.basename(r_fusa)))
                shutil.move(r_fusa, destino_b)
            print(f"[!] {archivo} -> LISTA NEGRA. Movido a Carpeta B.")
            continue

        if "ECOSAN" in nombre_final.upper() or "facturacion@ecosan.com.ar" in remitente_correo:
            _, locacion_ecosan = analizar_pdf_ecosan(ruta_completa)
            if not sector or sector == "GENERAL":
                sector = locacion_ecosan

        # --- 🛡️ VALIDACIÓN ANTICIPADA DE DUPLICADOS (RAZÓN SOCIAL + NÚMERO, como el historial) ---
        nro_str = str(datos.get('numero', '')).strip()
        identidad_factura = (re.sub(r'^(\([^)]+\)\s*)+', '', nombre_final).strip().upper(), nro_str)

        if identidad_factura in facturas_procesadas_historico or nro_str == "SinNumero":
            for r_fusa in archivos_a_fusionar:
                destino_duplicado = obtener_ruta_unica(os.path.join(CARPETA_B, os.path.basename(r_fusa)))
                shutil.move(r_fusa, destino_duplicado)
            print(f"[⏭️ DUPLICADO OMITIDO] {archivo} (Comprobante {nro_str}) -> Ya procesado anteriormente. Movido a Carpeta B.")
            duplicados_omitidos += 1
            continue

        nombre_mes = nombre_mes_desde_fecha(datos.get("fecha_emision"))
        prefijo_empresa = "(EUROSAT) " if datos.get("es_eurosat") else ""

        sec_str = f" {sector}" if sector else ""
        nombre_sugerido = f"{prefijo_empresa}({nombre_mes}) {nombre_final} {datos['tipo']} {datos['numero']}{sec_str}.pdf"

        print("\n" + "=" * 60)
        print(f"📄 Archivo principal: {archivo}")
        if len(archivos_a_fusionar) > 1:
            print(f"   📎 Se fusionarán {len(archivos_a_fusionar)-1} adjuntos en un único PDF.")
        print(f"   ↳ Proveedor detectado : {nombre_final}")
        print(f"   ↳ CUIT                  : {datos.get('cuit', 'No detectado')}")
        print(f"   ↳ Fecha Emisión         : {datos.get('fecha_emision', 'No detectada')}")
        print(f"   ↳ Tipo y Número         : {datos['tipo']} - {datos['numero']}")
        print(f"   ↳ Sector                : {sector if sector else 'GENERAL'}")
        print(f"   ✨ Nombre Propuesto     : {nombre_sugerido}")
        print("=" * 60)

        opcion = input("¿Qué deseas hacer?\n  [Enter] ➔ Aprobar y mover a 'Compras'\n  [e]     ➔ Editar nombre manualmente\n  [n]     ➔ Rechazar (Mover a 'no es factura')\nTu opción: ").strip().lower()

        if opcion == 'n':
            for r_fusa in archivos_a_fusionar:
                destino_b = obtener_ruta_unica(os.path.join(CARPETA_B, os.path.basename(r_fusa)))
                shutil.move(r_fusa, destino_b)
            print("❌ Rechazado. Movido a Carpeta B.")
            continue

        elif opcion == 'e':
            nuevo_input = input(f"Escribe el nuevo nombre (sin .pdf) o presiona Enter para usar '{nombre_sugerido}': ").strip()
            if nuevo_input:
                nombre_sugerido = nuevo_input if nuevo_input.lower().endswith('.pdf') else f"{nuevo_input}.pdf"

        # --- INTEGRACIÓN DE ÓRDENES DE COMPRA (OC) ---
        ocs_disponibles = oc_manager.leer_oc_disponibles(datos.get("cuit", ""), nombre_final)

        oc_elegida = ""
        oc_extraida = datos.get("oc_extraida", "")

        if oc_extraida:
            print(f"📌 Se detectó automáticamente la leyenda OC {oc_extraida} en el PDF.")
            if any(o['nro_oc'].lower() == oc_extraida.lower() for o in ocs_disponibles):
                print(f"✅ La OC {oc_extraida} está disponible en la base de datos.")
            oc_elegida = oc_extraida

        if ocs_disponibles and not oc_elegida:
            print(f"\n🛒 OCs disponibles para {nombre_final}:")
            for i, oc in enumerate(ocs_disponibles, 1):
                print(f"  {i}) OC {oc['nro_oc']} | {oc['fecha']} | {oc['articulo']} | Tipo: {oc['tipo']}")
            oc_elegida = input("Escribe el Nro de OC a vincular (o presiona Enter para omitir): ").strip()

        if oc_elegida:
            oc_encontrada = next((o for o in ocs_disponibles if o['nro_oc'].lower() == oc_elegida.lower()), None)
            nro_oc_final = oc_encontrada['nro_oc'] if oc_encontrada else oc_elegida

            ruta_pdf_oc = oc_manager.buscar_pdf_oc(nro_oc_final)
            if ruta_pdf_oc:
                print(f"✅ Se adjuntará el PDF de la OC: {os.path.basename(ruta_pdf_oc)}")
                archivos_a_fusionar.append(ruta_pdf_oc)
                base_n, ext_n = os.path.splitext(nombre_sugerido)
                nombre_sugerido = f"{base_n} -OC {nro_oc_final}{ext_n}"
                if oc_encontrada:
                    oc_manager.registrar_consumo_oc(nro_oc_final, nombre_sugerido)
            else:
                print(f"⚠️ No se encontró el archivo PDF para la OC {nro_oc_final} en la red. Se omitirá la fusión de la OC.")

        nombre_limpio = re.sub(r'[\\/*?:"<>|]', "", nombre_sugerido)
        ruta_destino_c = obtener_ruta_unica(os.path.join(CARPETA_C, nombre_limpio))

        sector_final = sector if sector else "GENERAL"
        carpeta_sector = os.path.join(CARPETA_AUTORIZACION, sector_final)
        os.makedirs(carpeta_sector, exist_ok=True)
        ruta_destino_sector = obtener_ruta_unica(os.path.join(carpeta_sector, nombre_limpio))

        if len(archivos_a_fusionar) > 1:
            fusionar_pdfs_adjuntos(archivos_a_fusionar, ruta_destino_c)
            shutil.copy2(ruta_destino_c, ruta_destino_sector)
            for r_fusa in archivos_a_fusionar:
                # No borrar la OC original de la red: solo adjuntos de la Carpeta A
                if os.path.exists(r_fusa) and r_fusa != ruta_completa and os.path.dirname(r_fusa) == CARPETA_A:
                    os.remove(r_fusa)
            if os.path.exists(ruta_completa):
                os.remove(ruta_completa)
            print(f"✅ ¡Fusionados y guardados en 'Compras' y '{sector_final}'!")
        else:
            shutil.copy2(ruta_completa, ruta_destino_c)
            shutil.move(ruta_completa, ruta_destino_sector)
            print(f"✅ ¡Aprobado y movido a 'Compras' y '{sector_final}'!")

        procesados_contador += 1

    print(f"\n🎯 Proceso finalizado. Aprobadas: {procesados_contador} | Duplicados omitidos: {duplicados_omitidos}")
    generar_reporte_html()
