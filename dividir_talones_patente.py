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
        base_dir = os.path.dirname(input_pdf_path)
        output_folder = os.path.join(base_dir, "Talones_Por_Patente")

    os.makedirs(output_folder, exist_ok=True)

    doc = fitz.open(input_pdf_path)
    total_pages = len(doc)
    print(f"Abriendo archivo: {input_pdf_path}")
    print(f"Total de páginas encontradas: {total_pages}")
    print("-" * 60)

    resultados = []

    for i in range(0, total_pages, 2):
        page_start = i
        page_end = min(i + 1, total_pages - 1)

        # Extraer texto de la primera página del par
        page_text = doc[page_start].get_text()

        # Buscar patente (formato anterior: AAA123 o formato nuevo Mercosur: AA123AA)
        matches = re.findall(r"\b([A-Z]{2}\s?[0-9]{3}\s?[A-Z]{2}|[A-Z]{3}\s?[0-9]{3})\b", page_text)
        unique_matches = list(dict.fromkeys(matches))

        if unique_matches:
            plate = unique_matches[0].replace(" ", "").upper()
        else:
            plate = f"TALON_PAG_{page_start+1}_{page_end+1}"

        # Crear nuevo documento de 2 páginas
        out_pdf = fitz.open()
        out_pdf.insert_pdf(doc, from_page=page_start, to_page=page_end)

        out_filename = f"{plate}.pdf"
        out_filepath = os.path.join(output_folder, out_filename)

        out_pdf.save(out_filepath)
        out_pdf.close()

        resultados.append((out_filename, page_start + 1, page_end + 1, out_filepath))
        print(f"  [OK] {out_filename:<12} -> Páginas {page_start+1:02d} a {page_end+1:02d}")

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
        pdf_path = r"C:\Users\gcastro\Desktop\Talon_Circulacion_10983577_0.pdf"

    dividir_talones_por_patente(pdf_path)
