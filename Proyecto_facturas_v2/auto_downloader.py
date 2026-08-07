#!/usr/bin/env python3
"""
auto_downloader.py — Descarga automática inteligente de facturas desde Outlook.

Uso:
    python auto_downloader.py                    # Fecha automática desde tracking
    python auto_downloader.py --desde 01/07/2026 # Override manual de fecha
    python auto_downloader.py --dias-atras 60    # Bajar últimos N días

Diseñado para ejecutarse como Tarea Programada de Windows.
"""
import os
import sys
import json
import logging
import argparse
from datetime import datetime, timedelta

# Permite importar módulos del proyecto desde cualquier ubicación
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from config import ESTADO_EMAILS, CARPETA_A, LOGS_DIR
from downloader import descargar_facturas_outlook

# ─── CONFIGURACIÓN DE LOGGING ─────────────────────────────────────────────────
LOG_PATH = os.path.join(LOGS_DIR, "auto_downloader.log")
os.makedirs(os.path.dirname(LOG_PATH), exist_ok=True)

logging.basicConfig(
    handlers=[
        logging.FileHandler(LOG_PATH, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger(__name__)
# ─────────────────────────────────────────────────────────────────────────────

DIAS_ATRAS_FALLBACK = 30  # Valor por defecto si no hay tracking


def obtener_fecha_desde_tracking():
    """
    Lee el tracking de emails procesados y devuelve la fecha del email
    más reciente descargado, en formato dd/mm/yyyy.
    Retorna None si el archivo no existe o está vacío.
    """
    try:
        if not os.path.exists(ESTADO_EMAILS):
            log.info("Archivo de tracking no encontrado. Se usara fallback.")
            return None

        with open(ESTADO_EMAILS, "r", encoding="utf-8") as f:
            procesados = json.load(f)

        if not procesados:
            log.info("Tracking vacio. Se usara fallback.")
            return None

        # Los entry_id tienen formato: YYYYMMDDHHMMSS_email_asunto
        fechas = []
        for entry_id in procesados:
            try:
                fecha_str = entry_id.split("_")[0]
                if len(fecha_str) == 14:
                    fechas.append(datetime.strptime(fecha_str, "%Y%m%d%H%M%S"))
            except (ValueError, IndexError):
                continue

        if not fechas:
            log.warning("No se pudo parsear ninguna fecha del tracking.")
            return None

        ultima_fecha = max(fechas)
        resultado = ultima_fecha.strftime("%d/%m/%Y")
        log.info(f"Fecha extraida del tracking: {resultado} ({len(procesados)} emails registrados)")
        return resultado

    except json.JSONDecodeError:
        log.error("El archivo de tracking esta corrupto (JSON invalido).")
        return None
    except Exception as e:
        log.error(f"Error leyendo tracking: {e}")
        return None


def obtener_fecha_inicio(dias_atras=DIAS_ATRAS_FALLBACK):
    """
    Determina la fecha de inicio para la búsqueda de correos.
    Prioridad: tracking JSON > fallback N días atrás.
    """
    fecha_tracking = obtener_fecha_desde_tracking()
    if fecha_tracking:
        return fecha_tracking

    fallback = datetime.now() - timedelta(days=dias_atras)
    fecha_fallback = fallback.strftime("%d/%m/%Y")
    log.info(f"Usando fallback: ultimos {dias_atras} dias -> desde {fecha_fallback}")
    return fecha_fallback


def main():
    parser = argparse.ArgumentParser(
        description="Descarga automatica inteligente de facturas desde Outlook"
    )
    parser.add_argument(
        "--desde",
        default="",
        metavar="DD/MM/YYYY",
        help="Fecha de inicio manual. Si se omite, se calcula automaticamente.",
    )
    parser.add_argument(
        "--dias-atras",
        type=int,
        default=DIAS_ATRAS_FALLBACK,
        metavar="N",
        help=f"Dias hacia atras como fallback si no hay tracking (default: {DIAS_ATRAS_FALLBACK})",
    )
    args = parser.parse_args()

    log.info("=" * 60)
    log.info("Iniciando descarga automatica de facturas")

    # Determinar fecha de inicio
    if args.desde:
        try:
            datetime.strptime(args.desde, "%d/%m/%Y")  # validar formato
            fecha_inicio = args.desde
            log.info(f"Fecha de inicio (manual): {fecha_inicio}")
        except ValueError:
            log.error(f"Formato de fecha invalido: '{args.desde}'. Usar DD/MM/YYYY.")
            sys.exit(1)
    else:
        fecha_inicio = obtener_fecha_inicio(dias_atras=args.dias_atras)

    log.info(f"Buscando correos desde: {fecha_inicio}")
    log.info(f"Destino: {CARPETA_A}")

    try:
        mapa = descargar_facturas_outlook(fecha_inicio=fecha_inicio)
        log.info(f"Descarga finalizada. PDFs obtenidos: {len(mapa)}")
    except Exception as e:
        log.error(f"Error critico en la descarga: {e}")
        sys.exit(1)

    log.info("=" * 60)


if __name__ == "__main__":
    main()
