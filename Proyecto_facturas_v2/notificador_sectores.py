# notificador_sectores.py — Aviso por mail a cada sector con sus facturas a autorizar.
#
# ESTADO: lógica conectada, ACTIVACIÓN PENDIENTE DE SDD (docs/SDD.md §8).
# Por defecto (NOTIFICACION_MODO = "desactivado") no toca Outlook: solo deja una
# vista previa HTML por sector en logs\notificaciones\ para validar el formato.
#
# Uso:
#   - Automático: 2_procesar_lote.py llama a notificar_lote() con lo recién aprobado.
#   - Recordatorio: python notificador_sectores.py --pendientes
#       (junta TODO lo que sigue pendiente en cada carpeta de sector; apto para tarea programada)
import html
import os
import re
import sys
from collections import defaultdict
from datetime import datetime

from config import (CARPETA_AUTORIZACION, LOGS_DIR, NOTIFICACION_ADJUNTAR_PDF, NOTIFICACION_CC,
                    NOTIFICACION_DESTINATARIOS, NOTIFICACION_MAX_ADJUNTOS_MB, NOTIFICACION_MODO)

MODOS_VALIDOS = ("desactivado", "borrador", "enviar")
DIR_PREVIEW = os.path.join(LOGS_DIR, "notificaciones")


def agrupar_por_sector(registros):
    grupos = defaultdict(list)
    for r in registros or []:
        grupos[(r.get("sector") or "GENERAL").upper()].append(r)
    return dict(grupos)


def _monto(v):
    try:
        return f"$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (TypeError, ValueError):
        return ""


def construir_email(sector, registros, titulo="Facturas pendientes de autorización"):
    """Devuelve (asunto, cuerpo_html). No depende de Outlook: se puede testear."""
    fecha = datetime.now().strftime("%d/%m/%Y")
    asunto = f"[{sector}] {titulo} - {len(registros)} comprobante(s) - {fecha}"

    filas = []
    total_general = 0.0
    for r in sorted(registros, key=lambda x: (str(x.get("razon_social", "")), str(x.get("numero", "")))):
        try:
            total_general += float(r.get("total") or 0)
        except (TypeError, ValueError):
            pass
        ruta = r.get("ruta_sector") or ""
        link = f'<a href="file:///{html.escape(ruta.replace(chr(92), "/"))}">abrir</a>' if ruta else ""
        filas.append(
            "<tr>"
            f"<td>{html.escape(str(r.get('fecha_emision', '') or ''))}</td>"
            f"<td><b>{html.escape(str(r.get('razon_social', '') or ''))}</b></td>"
            f"<td>{html.escape(str(r.get('tipo', '') or ''))} {html.escape(str(r.get('numero', '') or ''))}</td>"
            f"<td>{html.escape(str(r.get('oc', '') or '-'))}</td>"
            f"<td style='text-align:right'>{_monto(r.get('total'))}</td>"
            f"<td>{link}</td>"
            "</tr>"
        )

    cuerpo = f"""<html><body style="font-family:Segoe UI,Arial,sans-serif;font-size:13px">
<p>Hola, equipo de <b>{html.escape(sector)}</b>:</p>
<p>Tienen <b>{len(registros)}</b> comprobante(s) para revisar y autorizar en
<code>{html.escape(os.path.join(CARPETA_AUTORIZACION, sector))}</code>.
Para autorizar, mover el PDF a la subcarpeta <b>AUTORIZADA</b>.</p>
<table cellpadding="6" cellspacing="0" border="1" style="border-collapse:collapse;border-color:#ccc">
<tr style="background:#0d6efd;color:#fff"><th>Fecha emisión</th><th>Proveedor</th><th>Comprobante</th><th>OC</th><th>Total</th><th>PDF</th></tr>
{''.join(filas)}
<tr><td colspan="4" style="text-align:right"><b>Total</b></td><td style="text-align:right"><b>{_monto(total_general)}</b></td><td></td></tr>
</table>
<p style="color:#777">Mensaje generado automáticamente por el Asistente Contable. Los importes se extraen del PDF: ante dudas, verificar el original.</p>
</body></html>"""
    return asunto, cuerpo


def _guardar_preview(sector, asunto, cuerpo):
    os.makedirs(DIR_PREVIEW, exist_ok=True)
    nombre = re.sub(r"[^A-Za-z0-9_-]+", "_", f"{datetime.now():%Y%m%d_%H%M%S}_{sector}") + ".html"
    ruta = os.path.join(DIR_PREVIEW, nombre)
    with open(ruta, "w", encoding="utf-8") as f:
        f.write(f"<!-- Asunto: {html.escape(asunto)} -->\n{cuerpo}")
    return ruta


def _adjuntos(registros):
    if not NOTIFICACION_ADJUNTAR_PDF:
        return []
    rutas, total = [], 0
    for r in registros:
        ruta = r.get("ruta_sector")
        if ruta and os.path.exists(ruta):
            total += os.path.getsize(ruta)
            if total > NOTIFICACION_MAX_ADJUNTOS_MB * 1024 * 1024:
                print(f"   ⚠️ Se alcanzó el tope de {NOTIFICACION_MAX_ADJUNTOS_MB} MB de adjuntos; el resto va solo como link.")
                break
            rutas.append(ruta)
    return rutas


def _outlook(destinatarios, asunto, cuerpo, adjuntos, modo):
    import win32com.client  # import tardío: solo Windows con Outlook
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = "; ".join(destinatarios)
    if NOTIFICACION_CC:
        mail.CC = "; ".join(NOTIFICACION_CC)
    mail.Subject = asunto
    mail.HTMLBody = cuerpo
    for a in adjuntos:
        mail.Attachments.Add(a)
    if modo == "enviar":
        mail.Send()
    else:
        mail.Save()  # queda en Borradores


def notificar(grupos, modo=None, titulo="Facturas pendientes de autorización"):
    """grupos: {sector: [registros]}. Devuelve la lista de (sector, resultado)."""
    modo = (modo or NOTIFICACION_MODO or "desactivado").lower()
    if modo not in MODOS_VALIDOS:
        print(f"⚠️ NOTIFICACION_MODO inválido ('{modo}'). Se usa 'desactivado'.")
        modo = "desactivado"

    resultados = []
    for sector, registros in grupos.items():
        if not registros:
            continue
        asunto, cuerpo = construir_email(sector, registros, titulo)
        destinatarios = [d for d in NOTIFICACION_DESTINATARIOS.get(sector, []) if d]

        if modo == "desactivado" or not destinatarios:
            ruta = _guardar_preview(sector, asunto, cuerpo)
            motivo = "notificación desactivada (pendiente de SDD)" if modo == "desactivado" else "sector sin destinatarios"
            print(f"   ✉️  [{sector}] {motivo}: vista previa en {ruta}")
            resultados.append((sector, "preview"))
            continue
        try:
            _outlook(destinatarios, asunto, cuerpo, _adjuntos(registros), modo)
            print(f"   ✉️  [{sector}] {'enviado' if modo == 'enviar' else 'borrador creado'} para {', '.join(destinatarios)}")
            resultados.append((sector, modo))
        except Exception as e:
            ruta = _guardar_preview(sector, asunto, cuerpo)
            print(f"   ⚠️ [{sector}] Outlook falló ({e}). Vista previa en {ruta}")
            resultados.append((sector, "error"))
    return resultados


def notificar_lote(registros, modo=None):
    """Llamado por Paso 2 con las facturas recién aprobadas (estado PENDIENTE)."""
    grupos = agrupar_por_sector(registros)
    if grupos:
        print("📨 Notificación a sectores:")
        return notificar(grupos, modo, "Nuevas facturas para autorizar")
    return []


def pendientes_por_sector():
    """Escanea AUTORIZACION DE FACTURAS: PDFs en la raíz de cada sector = pendientes de autorizar."""
    from reporter import parsear_nombre_aprobado, LectorDatosFactura
    grupos = {}
    if not os.path.exists(CARPETA_AUTORIZACION):
        return grupos
    lector = LectorDatosFactura(max_lecturas=0)  # solo datos ya conocidos: no lee PDFs por la red
    for sec in sorted(os.listdir(CARPETA_AUTORIZACION)):
        dir_sector = os.path.join(CARPETA_AUTORIZACION, sec)
        if not os.path.isdir(dir_sector):
            continue
        regs = []
        for f in os.listdir(dir_sector):
            ruta = os.path.join(dir_sector, f)
            if not f.lower().endswith(".pdf") or not os.path.isfile(ruta):
                continue
            info = parsear_nombre_aprobado(f)
            d = lector.datos(f, ruta)
            regs.append({"sector": sec.upper(), "razon_social": info["Razón Social"], "tipo": info["Tipo"],
                         "numero": info["Nro Comprobante"], "oc": info["OC"],
                         "fecha_emision": d.get("fecha_emision", ""), "total": d.get("total"),
                         "ruta_sector": ruta})
        if regs:
            grupos[sec.upper()] = regs
    return grupos


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if "--pendientes" in sys.argv:
        g = pendientes_por_sector()
        print(f"Sectores con pendientes: {len(g)}")
        notificar(g, titulo="Recordatorio: facturas pendientes de autorización")
    else:
        print(__doc__ if __doc__ else "Uso: python notificador_sectores.py --pendientes")
