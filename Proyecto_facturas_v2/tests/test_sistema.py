# test_sistema.py — Tests del núcleo. Correr con VERIFICAR.bat o:
#   python -m unittest discover -s Proyecto_facturas_v2/tests -v
# No tocan la red ni Outlook: todo corre en carpetas temporales (ver _entorno.py).
import base64
import glob
import importlib
import json
import os
import shutil
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _entorno  # noqa: E402  (debe ir antes de importar el proyecto)

try:
    import pymupdf as fitz  # PyMuPDF >= 1.24
except ImportError:  # versiones viejas
    import fitz
import pandas as pd  # noqa: E402

import config  # noqa: E402
import estado  # noqa: E402
import notificador_sectores  # noqa: E402
import procesador_core as pc  # noqa: E402
import qr_afip  # noqa: E402
import reporter  # noqa: E402

paso1 = importlib.import_module("1_preparar_lote")
paso2 = importlib.import_module("2_procesar_lote")

FIXTURES = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fixtures")


# ─── Generadores de PDFs de prueba ───────────────────────────────────────────
def crear_pdf(ruta, paginas, imagen_png=None):
    doc = fitz.open()
    for texto in paginas:
        page = doc.new_page()
        y = 60
        for linea in texto.split("\n"):
            page.insert_text((50, y), linea, fontsize=10)
            y += 14
        if imagen_png and page.number == 0:
            page.insert_image(fitz.Rect(400, 600, 550, 750), stream=imagen_png)
    doc.save(ruta)
    doc.close()
    return ruta


def url_qr(tipo_cmp=1, fecha="2026-03-10", cuit=30711111111, nro=1234, importe=12100.0):
    data = {"ver": 1, "fecha": fecha, "cuit": cuit, "ptoVta": 2, "tipoCmp": tipo_cmp, "nroCmp": nro,
            "importe": importe, "moneda": "PES", "ctz": 1, "tipoDocRec": 80,
            "nroDocRec": 30696428898, "tipoCodAut": "E", "codAut": 70417054367476}
    p = base64.b64encode(json.dumps(data).encode()).decode()
    return f"https://www.afip.gob.ar/fe/qr/?p={p}"


def png_qr(texto):
    """PNG de un QR usando zxing-cpp (o qrcode si está). None si no hay generador."""
    try:
        import zxingcpp
        import numpy as np
        escritor = getattr(zxingcpp, "write_barcode", None)
        if escritor is not None:
            img = np.asarray(escritor(zxingcpp.BarcodeFormat.QRCode, texto, 400, 400))
        else:
            img = np.asarray(zxingcpp.create_barcode(texto, zxingcpp.BarcodeFormat.QRCode).to_image(scale=8))
        img = np.ascontiguousarray(img.astype("uint8"))
        if img.ndim == 3:
            img = np.ascontiguousarray(img[:, :, 0])
        pix = fitz.Pixmap(fitz.csGRAY, img.shape[1], img.shape[0], img.tobytes(), 0)
        return pix.tobytes("png")
    except Exception:
        pass
    try:
        import io
        import qrcode
        buf = io.BytesIO()
        qrcode.make(texto).save(buf, format="PNG")
        return buf.getvalue()
    except Exception:
        return None


TEXTO_FACTURA = """FACTURA A    N° 0003-00012345
PROVEEDOR EJEMPLO S.R.L.
CUIT: 30-71111111-1
Fecha de Emisión: 15/08/2025
Cliente: TELCOM VENTURES DE ARGENTINA S.A.  CUIT 30-69642889-8
Cod  Descripcion                 Cant   P.Unit      Subtotal
A1   Alquiler de volquete 6m3    2      5.000,00    10.000,00
Subtotal: $ 10.000,00
IVA 21%: $ 2.100,00
Percepción IIBB: $ 300,00
Importe Total: $ 12.400,00
CAE N°: 70417054367476   Fecha Vto. CAE: 25/08/2025"""

TEXTO_RECIBO = """VOLQUETES Y TRANSPORTES MARINO S.A.
Recibo: 0002-00032472
DOCUMENTO NO VALIDO COMO FACTURA
Keops 360
Importe 70417054367476
Recibimos de TELCOM VENTURES la suma de pesos..."""


class TestReglas(unittest.TestCase):
    def test_parsear_monto(self):
        self.assertEqual(pc.parsear_monto("1.234.567,89"), 1234567.89)
        self.assertEqual(pc.parsear_monto("1,234.56"), 1234.56)
        self.assertEqual(pc.parsear_monto("12.100"), 12100.0)
        self.assertEqual(pc.parsear_monto("99,5"), 99.5)
        self.assertEqual(pc.parsear_monto("12400.0"), 12400.0)
        self.assertIsNone(pc.parsear_monto("nan"))

    def test_bug01_mes_desde_fecha_emision(self):
        self.assertEqual(pc.nombre_mes_desde_fecha("15/08/2025"), "AGOSTO")
        self.assertEqual(pc.nombre_mes_desde_fecha("2026-03-10"), "MARZO")
        self.assertEqual(pc.nombre_mes_desde_fecha("", defecto="X"), "X")
        self.assertEqual(pc.nombre_mes_desde_fecha("basura", defecto="X"), "X")

    def test_bug02_recibo_rechazado_aunque_remitente_confiable(self):
        self.assertFalse(pc.es_factura_valida(TEXTO_RECIBO, "facturacion@marinosa.com.ar"))
        self.assertFalse(pc.es_factura_valida(TEXTO_RECIBO))

    def test_remitente_confiable_relaja_cae(self):
        self.assertTrue(pc.es_factura_valida("Factura de servicio mensual", "facturacion@marinosa.com.ar"))
        self.assertFalse(pc.es_factura_valida("Factura de servicio mensual"))

    def test_factura_con_cae(self):
        self.assertTrue(pc.es_factura_valida(TEXTO_FACTURA))

    def test_palabra_prohibida_sin_cae(self):
        self.assertFalse(pc.es_factura_valida("PRESUPUESTO 0001-00000001 para cotizar"))

    def test_nota_credito_con_acento(self):
        ruta = crear_pdf(os.path.join(_entorno.BASE, "nc.pdf"), [TEXTO_FACTURA.replace("FACTURA A", "NOTA DE CRÉDITO A")])
        ok, datos = pc.extraer_datos_pdf(ruta)
        self.assertTrue(ok)
        self.assertEqual(datos["tipo"], "NC")

    def test_importes_e_items(self):
        imp = pc.extraer_importes(TEXTO_FACTURA)
        self.assertEqual(imp["neto"], 10000.0)
        self.assertEqual(imp["iva_21"], 2100.0)
        self.assertEqual(imp["otros_impuestos"], 300.0)
        self.assertEqual(imp["total"], 12400.0)
        items = pc.extraer_items(TEXTO_FACTURA)
        self.assertEqual(len(items), 1)
        self.assertEqual(items[0]["precio_unitario"], 5000.0)
        self.assertEqual(items[0]["cantidad"], 2.0)

    def test_total_no_confunde_subtotal(self):
        self.assertIsNone(pc.extraer_importes("Subtotal: $ 10.000,00")["total"])

    def test_extraccion_pdf_factura(self):
        ruta = crear_pdf(os.path.join(_entorno.BASE, "fc.pdf"), [TEXTO_FACTURA])
        ok, datos = pc.extraer_datos_pdf(ruta)
        self.assertTrue(ok, datos)
        self.assertEqual(datos["numero"], "12345")
        self.assertEqual(datos["cuit"], "30711111111")
        self.assertEqual(datos["fecha_emision"], "15/08/2025")
        self.assertEqual(datos["total"], 12400.0)

    def test_recibo_adjunto_no_tumba_factura(self):
        ruta = crear_pdf(os.path.join(_entorno.BASE, "fc_mas_recibo.pdf"), [TEXTO_FACTURA, TEXTO_RECIBO])
        ok, datos = pc.extraer_datos_pdf(ruta)
        self.assertTrue(ok, datos)

    def test_recibo_pdf_rechazado(self):
        ruta = crear_pdf(os.path.join(_entorno.BASE, "recibo.pdf"), [TEXTO_RECIBO])
        ok, motivo = pc.extraer_datos_pdf(ruta, remitente="facturacion@marinosa.com.ar")
        self.assertFalse(ok)
        self.assertIn("no válido como factura", motivo)


class TestQR(unittest.TestCase):
    def test_parsear_url(self):
        d = qr_afip.interpretar_qr(qr_afip.parsear_url_qr(url_qr()))
        self.assertTrue(d["es_aceptado"])
        self.assertEqual(d["fecha_emision"], "10/03/2026")
        self.assertEqual(d["numero"], "1234")
        self.assertEqual(d["cuit"], "30711111111")

    def test_recibo_por_qr(self):
        d = qr_afip.interpretar_qr(qr_afip.parsear_url_qr(url_qr(tipo_cmp=9)))
        self.assertTrue(d["es_recibo"])
        self.assertFalse(pc.es_factura_valida("Factura CAE 70417054367476", qr=d))

    def test_url_no_afip(self):
        self.assertIsNone(qr_afip.parsear_url_qr("https://example.com/?p=abc"))

    def test_qr_en_pdf(self):
        png = png_qr(url_qr())
        if png is None or not qr_afip.hay_decodificador_qr():
            self.skipTest("Sin generador/decodificador de QR instalado")
        ruta = crear_pdf(os.path.join(_entorno.BASE, "qr.pdf"), ["Comprobante sin texto de CAE"], imagen_png=png)
        qr = qr_afip.leer_qr_afip(ruta)
        self.assertIsNotNone(qr)
        ok, datos = pc.extraer_datos_pdf(ruta)
        self.assertTrue(ok, datos)
        self.assertEqual(datos["fecha_emision"], "10/03/2026")
        self.assertEqual(datos["total"], 12100.0)


class TestReporter(unittest.TestCase):
    def test_lee_reporte_viejo_sin_corromper(self):
        df = reporter.leer_tabla_html(os.path.join(FIXTURES, "reporte_maestro_v2.html"))
        self.assertEqual(len(df), 2)                        # sin la fila vacía de filtros
        self.assertIn("Fecha Carga", df.columns)            # "Fecha" vieja = fecha de carga
        self.assertEqual(df.iloc[0]["Nro Comprobante"], "47336")
        self.assertEqual(df.iloc[0]["OC"], "21095")
        self.assertEqual(df.iloc[0]["Subtotal"], "10.000,00")

    def test_historial_detecta_duplicados(self):
        hechas = reporter.obtener_facturas_procesadas(os.path.join(FIXTURES, "reporte_maestro_v2.html"))
        self.assertIn(("NUAGE S.A.", "47336"), hechas)      # antes devolvía "47336.0" y nunca coincidía

    def test_parsear_nombre(self):
        info = reporter.parsear_nombre_aprobado("(VISA CORPO) (AGOSTO) PROVEEDOR SA FC 1234 RRHH -OC 21095.pdf")
        self.assertEqual(info["Razón Social"], "PROVEEDOR SA")
        self.assertEqual(info["Nro Comprobante"], "1234")
        self.assertEqual(info["OC"], "21095")
        self.assertEqual(info["Sector"], "RRHH")

    def test_html_escapa(self):
        ruta = os.path.join(_entorno.BASE, "esc.html")
        reporter._escribir_html(pd.DataFrame([{"Razón Social": "A & B <SRL>", "Nro Comprobante": "1"}]), ruta, "t")
        contenido = open(ruta, encoding="utf-8").read()
        self.assertIn("A &amp; B &lt;SRL&gt;", contenido)


class TestFlujoCompleto(unittest.TestCase):
    """Paso 1 -> Excel -> Paso 2 -> reportes -> notificación, en carpetas temporales."""

    @classmethod
    def setUpClass(cls):
        a = config.CARPETA_A
        crear_pdf(os.path.join(a, "factura_agosto.pdf"), [TEXTO_FACTURA])
        crear_pdf(os.path.join(a, "recibo_marino.pdf"), [TEXTO_RECIBO])
        estado.agregar_remitentes({"recibo_marino.pdf": "facturacion@marinosa.com.ar"})

    def test_flujo(self):
        paso1.generar_excel_lote()
        df = pd.read_excel(paso1.EXCEL_VALIDACION, dtype=str)
        filas = {r["Archivo_Original"]: r for _, r in df.iterrows()}
        self.assertEqual(filas["recibo_marino.pdf"]["Acción"], "RECHAZAR")
        fc = filas["factura_agosto.pdf"]
        self.assertEqual(fc["Acción"], "APROBAR")
        self.assertIn("(AGOSTO)", fc["Nombre_Propuesto"])   # BUG-01: mes de la factura, no el actual
        self.assertEqual(fc["Fecha_Emision"], "15/08/2025")

        # El usuario edita el sector en el Excel
        df.loc[df["Archivo_Original"] == "factura_agosto.pdf", "Sector"] = "ADM"
        df.to_excel(paso1.EXCEL_VALIDACION, index=False)

        paso2.procesar_excel_lote()

        compras = [f for f in os.listdir(config.CARPETA_COMPRAS) if f.lower().endswith(".pdf")]
        self.assertEqual(len(compras), 1)                    # MEJ-03: directo a \Compras
        self.assertTrue(compras[0].startswith("(AGOSTO)"))
        self.assertTrue(os.listdir(os.path.join(config.CARPETA_AUTORIZACION, "ADM")))
        self.assertIn("recibo_marino.pdf", os.listdir(config.CARPETA_B))

        registro = estado.cargar_registro()
        self.assertEqual(registro[compras[0]]["fecha_emision"], "15/08/2025")
        self.assertEqual(registro[compras[0]]["total"], 12400.0)
        self.assertEqual(len(registro[compras[0]]["items"]), 1)

        maestro = reporter.leer_tabla_html(config.REPORTE_MAESTRO_HTML)
        self.assertEqual(maestro.iloc[0]["Fecha Emisión"], "15/08/2025")   # BUG-04
        self.assertTrue(os.path.exists(config.REPORTE_DETALLADO_HTML))      # MEJ-02
        self.assertTrue(os.path.exists(config.REPORTE_MAESTRO_EXCEL))
        self.assertTrue(glob.glob(os.path.join(config.LOGS_DIR, "notificaciones", "*ADM*.html")))

        # Regeneración completa: conserva datos y no duplica
        reporter.generar_todos_los_reportes()
        maestro = reporter.leer_tabla_html(config.REPORTE_MAESTRO_HTML)
        self.assertEqual(len(maestro), 1)
        self.assertEqual(maestro.iloc[0]["Fecha Emisión"], "15/08/2025")
        self.assertEqual(maestro.iloc[0]["Total"], "12.400,00")

        # La misma factura vuelve a entrar: debe marcarse como duplicada del historial
        crear_pdf(os.path.join(config.CARPETA_A, "factura_agosto_reenvio.pdf"), [TEXTO_FACTURA])
        paso1.generar_excel_lote()
        df2 = pd.read_excel(paso1.EXCEL_VALIDACION, dtype=str)
        fila = df2[df2["Archivo_Original"] == "factura_agosto_reenvio.pdf"].iloc[0]
        self.assertEqual(fila["Acción"], "RECHAZAR")
        self.assertIn("HISTORIAL", fila["Estado_Historial"])


class TestVarios(unittest.TestCase):
    def test_clasificar_accion(self):
        c = paso2.clasificar_accion
        self.assertEqual(c("Aprobar"), "APROBAR")
        self.assertEqual(c("REVISIÓN"), "REVISION")
        self.assertEqual(c("rechazar"), "RECHAZAR")
        self.assertEqual(c("no"), "RECHAZAR")
        self.assertIsNone(c("nan"))
        self.assertIsNone(c("quizas"))

    def test_sector_editado_se_refleja_en_nombre(self):
        a = paso2.ajustar_sector_en_nombre
        self.assertEqual(a("(AGOSTO) X SA FC 12 GENERAL.pdf", "ADM"), "(AGOSTO) X SA FC 12 ADM.pdf")
        self.assertEqual(a("(AGOSTO) X SA FC 12 GENERAL -OC 5.pdf", "RRHH"), "(AGOSTO) X SA FC 12 RRHH -OC 5.pdf")
        self.assertEqual(a("nombre libre.pdf", "ADM"), "nombre libre.pdf")

    def test_razon_social_desde_texto(self):
        self.assertEqual(pc._extraer_razon_social(TEXTO_FACTURA), "PROVEEDOR EJEMPLO S.R.L.")
        self.assertEqual(pc._extraer_razon_social("Alquiler de casa\nTELCOM VENTURES DE ARGENTINA S.A."), "")

    def test_cuit_digito_verificador(self):
        self.assertTrue(pc.cuit_valido("30-69642889-8"))      # TELCOM
        self.assertTrue(pc.cuit_valido("33714232199"))
        self.assertFalse(pc.cuit_valido("30711111118"))
        # Un número de 11 dígitos inválido no le gana al CUIT real aunque aparezca antes
        self.assertEqual(pc._cuits_proveedor("Ref 30123456780\nCUIT 33-71423219-9")[0], "33714232199")

    def test_notificador_construye_mail(self):
        asunto, cuerpo = notificador_sectores.construir_email("RRHH", [
            {"razon_social": "A & B", "numero": "1", "tipo": "FC", "total": 1500.5, "fecha_emision": "01/01/2026"}])
        self.assertIn("[RRHH]", asunto)
        self.assertIn("A &amp; B", cuerpo)
        self.assertIn("1.500,50", cuerpo)

    def test_estado_json_atomico(self):
        ruta = os.path.join(_entorno.BASE, "x.json")
        estado.guardar_json(ruta, {"a": 1})
        self.assertEqual(estado.cargar_json(ruta, {}), {"a": 1})
        self.assertFalse(os.path.exists(ruta + ".tmp"))

    def test_monitor_solo_pdf_sueltos(self):
        import monitor_compras
        sub = os.path.join(config.CARPETA_COMPRAS, "fc a subir")
        os.makedirs(sub, exist_ok=True)
        crear_pdf(os.path.join(sub, "viejo.pdf"), ["x"])
        foto = monitor_compras.foto_actual()
        self.assertFalse(any("fc a subir" in p for p in foto))   # MEJ-04


_STDOUT_ORIGINAL = sys.stdout


def setUpModule():
    # Los procesos imprimen mucho; se silencia salvo TESTS_VERBOSE=1 (unittest informa por stderr).
    if not os.environ.get("TESTS_VERBOSE"):
        sys.stdout = open(os.devnull, "w", encoding="utf-8")


def tearDownModule():
    if sys.stdout is not _STDOUT_ORIGINAL:
        sys.stdout.close()
        sys.stdout = _STDOUT_ORIGINAL
    shutil.rmtree(_entorno.BASE, ignore_errors=True)


if __name__ == "__main__":
    unittest.main(verbosity=2)
