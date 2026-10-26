# 1_preparar_lote.py — PASO 1: analiza los PDFs de la Carpeta A y arma el Excel de validación.
import os
import re
import sys

import openpyxl
import pandas as pd
from openpyxl.worksheet.datavalidation import DataValidation

from config import ARCHIVO_PROVEEDORES, CARPETA_A, LOTE_DATOS_PATH, REPORTE_MAESTRO_HTML
from estado import esperar_archivo_libre, guardar_json, podar_remitentes
from procesador_core import extraer_datos_pdf, nombre_mes_desde_fecha
from proveedores import buscar_proveedor, cargar_excel_proveedores
from reporter import obtener_facturas_procesadas

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

# 📍 El Excel se ubica directamente en la carpeta de entrada
EXCEL_VALIDACION = os.path.join(CARPETA_A, "Validacion_Lote.xlsx")
NUMEROS_INVALIDOS = ["SinNumero", "N/A", "S_N"]

# Orden de columnas del Excel. Las editables por el usuario: Sector, OC_Detectada,
# Adjuntos_Extra, Tarjeta, Autorizar_Directo, Nombre_Propuesto, Acción.
COLUMNAS = [
    "Archivo_Original", "Ver_PDF", "Empresa", "Razón_Social", "Sector", "Tipo",
    "Nro_Comprobante", "Fecha_Emision", "Subtotal", "Total", "OC_Detectada",
    "Adjuntos_Extra", "Tarjeta", "Autorizar_Directo", "Nombre_Propuesto",
    "Estado_Historial", "Acción",
    # Informativas (MEJ-02 / MEJ-01)
    "CUIT", "Neto", "IVA_21", "IVA_10_5", "IVA_27", "Otros_Impuestos",
    "Items_Detectados", "QR_AFIP", "Motivo",
]


def _fila_base(archivo, link_formula):
    fila = {c: "" for c in COLUMNAS}
    fila.update({"Archivo_Original": archivo, "Ver_PDF": link_formula, "Tarjeta": "NO",
                 "Autorizar_Directo": "NO"})
    return fila


def _fila_rechazo(archivo, link_formula, motivo):
    fila = _fila_base(archivo, link_formula)
    fila.update({
        "Empresa": "TELCOM VENTURES",
        "Razón_Social": "NO FACTURA",
        "Sector": "GENERAL",
        "Tipo": "N/A",
        "Nro_Comprobante": "N/A",
        "Nombre_Propuesto": "NO FACTURA",
        "Estado_Historial": "NO ES FACTURA / RECHAZAR",
        "Acción": "RECHAZAR",
        "Motivo": motivo,
    })
    return fila


def _formatear_excel(worksheet):
    font_hyperlink = openpyxl.styles.Font(color="0000FF", underline="single")
    indices = {str(c.value).strip(): i for i, c in enumerate(worksheet[1], 1)}
    ultima = max(worksheet.max_row, 2)

    if "Ver_PDF" in indices:
        for r_idx in range(2, worksheet.max_row + 1):
            worksheet.cell(row=r_idx, column=indices["Ver_PDF"]).font = font_hyperlink

    listas = {"Acción": '"APROBAR,RECHAZAR,REVISION"', "Tarjeta": '"NO,SI"', "Autorizar_Directo": '"NO,SI"'}
    for col, formula in listas.items():
        if col in indices:
            letra = openpyxl.utils.get_column_letter(indices[col])
            dv = DataValidation(type="list", formula1=formula, allow_blank=False)
            worksheet.add_data_validation(dv)
            dv.add(f"{letra}2:{letra}{ultima}")

    for col in worksheet.columns:
        max_length = 0
        for cell in col:
            val_str = "" if cell.value is None else str(cell.value)
            val_len = 10 if val_str.startswith("=HYPERLINK") else len(val_str)
            max_length = max(max_length, val_len)
        worksheet.column_dimensions[col[0].column_letter].width = min(max(max_length + 2, 10), 50)

    worksheet.freeze_panes = "B2"


def generar_excel_lote():
    print("Iniciando Preparación de Lote (Paso 1)...")

    if not os.path.exists(CARPETA_A):
        print(f"❌ La carpeta de entrada no existe: {CARPETA_A}")
        return

    archivos = sorted(f for f in os.listdir(CARPETA_A) if f.lower().endswith(".pdf"))
    if not archivos:
        print("📂 No hay facturas pendientes en la bandeja de entrada para procesar.")
        return

    if not esperar_archivo_libre(EXCEL_VALIDACION):
        print("❌ Cancelado: el Excel de validación sigue abierto.")
        return

    print("📥 Cargando bases de proveedores, sectores e historial...")
    base_prov = cargar_excel_proveedores(ARCHIVO_PROVEEDORES)
    facturas_procesadas_historico = obtener_facturas_procesadas(REPORTE_MAESTRO_HTML)
    print(f"🛡️ Historial cargado: {len(facturas_procesadas_historico)} facturas registradas previamente.")
    remitentes = podar_remitentes(archivos)
    if remitentes:
        print(f"📧 Remitentes de Outlook conocidos para {len(remitentes)} PDF(s).")

    datos_excel = []
    datos_lote = {}
    vistos_en_lote = set()

    for archivo in archivos:
        ruta_completa = os.path.join(CARPETA_A, archivo)
        ruta_abs = os.path.abspath(ruta_completa)
        link_formula = f'=HYPERLINK("{ruta_abs}", "Ver PDF")'
        url_pdf = "file:///" + ruta_abs.replace("\\", "/")
        print(f"  Analizando: {archivo}")
        print(f"   └─ Link al PDF: {url_pdf}")

        exito, datos = extraer_datos_pdf(ruta_completa, remitente=remitentes.get(archivo, ""))

        # ─── CASO 1: NO ES FACTURA O ES CATÁLOGO (>3 PÁGS) ────────────────
        if not exito:
            print(f"  ⚠️ Documento no válido/descartado ({datos}). Marcando para RECHAZAR.")
            datos_excel.append(_fila_rechazo(archivo, link_formula, str(datos)))
            continue

        # ─── CASO 2: DOCUMENTO VÁLIDO -> BUSCAR PROVEEDOR ─────────────────
        nombre_final, sector, es_lista_negra = buscar_proveedor(datos["cuit"], datos["razon_social"], base_prov)

        tipo_str = datos["tipo"] or "FC"
        nro_str = str(datos["numero"] or "S_N").strip()

        # BUG-01: el mes sale de la fecha de emisión (fallback: mes actual)
        nombre_mes = nombre_mes_desde_fecha(datos.get("fecha_emision"))
        sector_str = sector if sector else "GENERAL"

        oc_extraida_limpia = str(datos.get("oc_extraida", "")).strip()
        if oc_extraida_limpia.endswith('.0'):
            oc_extraida_limpia = oc_extraida_limpia[:-2]

        es_tarjeta = datos.get("es_tarjeta", False) or bool(re.search(r"\b(visa|visacorpo|visa corpo|tarjeta)\b", archivo, re.IGNORECASE))
        prefijo_tarjeta = "(VISA CORPO) " if es_tarjeta else ""
        es_eurosat = datos.get("es_eurosat", False)
        empresa_str = "EUROSAT" if es_eurosat else "TELCOM VENTURES"
        prefijo_empresa = "(EUROSAT) " if es_eurosat else ""

        nombre_sugerido = f"{prefijo_tarjeta}{prefijo_empresa}({nombre_mes}) {nombre_final} {tipo_str} {nro_str} {sector_str}"
        if oc_extraida_limpia:
            nombre_sugerido += f" -OC {oc_extraida_limpia}"

        # ─── CASO 3: DETECCIÓN DE DUPLICADAS (HISTORIAL + EN MISMO LOTE) ──
        razon_limpia = re.sub(r'^(\([^)]+\)\s*)+', '', nombre_final).strip().upper()
        identificador_prov = str(datos.get("cuit", "")).strip() or razon_limpia
        numero_valido = nro_str not in NUMEROS_INVALIDOS

        es_duplicada_historico = numero_valido and any(
            n_hist == nro_str and (r_hist in razon_limpia or razon_limpia in r_hist)
            for (r_hist, n_hist) in facturas_procesadas_historico
        )

        clave_lote = (identificador_prov, nro_str)
        es_duplicada_lote = numero_valido and clave_lote in vistos_en_lote
        if numero_valido:
            vistos_en_lote.add(clave_lote)

        fila = _fila_base(archivo, link_formula)
        fila.update({
            "Empresa": empresa_str,
            "Razón_Social": nombre_final,
            "Sector": sector_str,
            "Tipo": tipo_str,
            "Nro_Comprobante": nro_str,
            "Fecha_Emision": datos.get("fecha_emision", ""),
            "Subtotal": datos.get("subtotal", ""),
            "Total": datos.get("total"),
            "OC_Detectada": oc_extraida_limpia,
            "Tarjeta": "SI" if es_tarjeta else "NO",
            "CUIT": datos.get("cuit", ""),
            "Neto": datos.get("neto"),
            "IVA_21": datos.get("iva_21"),
            "IVA_10_5": datos.get("iva_105"),
            "IVA_27": datos.get("iva_27"),
            "Otros_Impuestos": datos.get("otros_impuestos"),
            "Items_Detectados": len(datos.get("items", [])),
            "QR_AFIP": "SI" if datos.get("qr_afip") else "NO",
        })

        if es_duplicada_historico:
            print(f"  🛑 DUPLICADA EN HISTORIAL: {nombre_final} {nro_str} ya fue procesada anteriormente.")
            fila.update({"Nombre_Propuesto": f"{nombre_sugerido} [DUPLICADA HISTORIAL]",
                         "Estado_Historial": "YA PROCESADA (HISTORIAL)", "Acción": "RECHAZAR"})
        elif es_duplicada_lote:
            print(f"  🛑 DUPLICADA EN LOTE: {nombre_final} {nro_str} aparece varias veces en la misma carpeta.")
            fila.update({"Nombre_Propuesto": f"{nombre_sugerido} [DUPLICADA EN LOTE]",
                         "Estado_Historial": "DUPLICADA EN MISMO LOTE", "Acción": "RECHAZAR"})
        else:
            fila.update({"Nombre_Propuesto": nombre_sugerido, "Estado_Historial": "NUEVA",
                         "Acción": "RECHAZAR" if es_lista_negra else "APROBAR",
                         "Motivo": "PROVEEDOR EN LISTA NEGRA" if es_lista_negra else ""})

        datos_excel.append(fila)
        datos_lote[archivo] = {k: datos.get(k) for k in
                               ("cuit", "fecha_emision", "neto", "iva_21", "iva_105", "iva_27",
                                "otros_impuestos", "total", "items", "qr_afip", "cae")}

    df = pd.DataFrame(datos_excel, columns=COLUMNAS)

    try:
        with pd.ExcelWriter(EXCEL_VALIDACION, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name="Lote_Pendiente")
            _formatear_excel(writer.sheets['Lote_Pendiente'])
        guardar_json(LOTE_DATOS_PATH, datos_lote)

        print("\n✅ ¡Lote generado con éxito!")
        print(f"👉 Excel generado en: '{EXCEL_VALIDACION}'")
        print("👉 Opciones de Acción disponibles en Excel: [APROBAR, RECHAZAR, REVISION]")
        print("👉 Revisa las Acciones, edita Sector u OC si hace falta, guarda y cierra el Excel.")
        print("👉 Luego, ejecuta el Paso 2 (Aplicar Lote).")
    except Exception as e:
        print(f"❌ Error al guardar el Excel: {e}")


if __name__ == "__main__":
    from estado import iniciar_log_auditoria
    iniciar_log_auditoria("paso1_preparar_lote")
    generar_excel_lote()
    input("\nPresiona ENTER para salir...")
