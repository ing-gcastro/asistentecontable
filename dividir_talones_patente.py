"""
Script para dividir PDFs de talones de circulación de vehículos en archivos individuales de 2 páginas nombrados por su patente / dominio.
"""

import os
import sys
import re
import fitz  # PyMuPDF

if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


def dividir_talones_por_patente(input_pdf_path: str, output_folder: str = None) -> list:
    if not os.path.exists(input_pdf_path):
        raise FileNotFoundError(f"No se encontró el archivo: {input_pdf_path}")

    if output_folder is None:
        output_folder = r"C:\Users\gcastro\ROBOTS\SEGUROS"

    os.makedirs(output_folder, exist_ok=True)

    doc = fitz.open(input_pdf_path)
    total_pages = len(doc)
    print(f"Abriendo archivo: {input_pdf_path}")
    print(f"Total de páginas encontradas: {total_pages}")
    print("-" * 60)

    # Detectar páginas de frente de talón
    talon_front_pages = []
    for i in range(total_pages):
        text = doc[i].get_text().upper()
        if (
            "CERTIFICADO DE PÓLIZA ÚNICA DE SEGURO" in text
            or "SEGURO OBLIGATORIO AUTOMOTOR" in text
            or "CERTIFICADO DE APÓLICE ÚNICA" in text
            or "DATOS REQUERIDOS EN CASO DE SINIESTRO" in text
        ):
            talon_front_pages.append(i)

    # Si no se detectaron encabezados específicos, asumir pares consecutivos de 2 páginas
    if not talon_front_pages:
        pairs = [(i, min(i + 1, total_pages - 1)) for i in range(0, total_pages, 2)]
    else:
        pairs = [(p, min(p + 1, total_pages - 1)) for p in talon_front_pages]

    resultados = []
    used_names = {}

    for page_start, page_end in pairs:
        page_text = doc[page_start].get_text()

        # Buscar patente (formato anterior: AAA123 o formato Mercosur: AA123AA)
        matches = re.findall(r"\b([A-Z]{2}\s?[0-9]{3}\s?[A-Z]{2}|[A-Z]{3}\s?[0-9]{3})\b", page_text)
        unique_matches = list(dict.fromkeys([m.replace(" ", "").upper() for m in matches]))

        if unique_matches:
            plate = unique_matches[0]
        else:
            plate = f"TALON_PAG_{page_start+1}_{page_end+1}"

        # Manejar nombres duplicados si existieran
        count = used_names.get(plate, 0) + 1
        used_names[plate] = count
        filename_plate = plate if count == 1 else f"{plate}_{count}"

        out_filename = f"{filename_plate}.pdf"
        out_filepath = os.path.join(output_folder, out_filename)

        # Crear nuevo documento PDF de 2 páginas
        out_pdf = fitz.open()
        out_pdf.insert_pdf(doc, from_page=page_start, to_page=page_end)
        out_pdf.save(out_filepath)
        out_pdf.close()

        resultados.append((out_filename, page_start + 1, page_end + 1, out_filepath))
        print(f"  [OK] {out_filename:<15} -> Páginas {page_start+1:03d} a {page_end+1:03d}")

    doc.close()
    print("-" * 60)
    print(f"Proceso finalizado exitosamente.")
    print(f"Total de archivos PDF generados: {len(resultados)}")
    print(f"Ubicación: {output_folder}")
    return resultados


if __name__ == "__main__":
    if len(sys.argv) > 1:
        pdf_path = sys.argv[1]
    else:
        pdf_path = r"C:\Users\gcastro\Desktop\Frente_Poliza_11101290_0.pdf"

    if len(sys.argv) > 2:
        dest_folder = sys.argv[2]
    else:
        dest_folder = r"C:\Users\gcastro\ROBOTS\SEGUROS"

    dividir_talones_por_patente(pdf_path, dest_folder)

