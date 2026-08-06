import json

notebook_path = r"c:\Users\gcastro\ROBOTS\asistentecontable\Proyecto_facturas_v2\prueba_fase1 note.ipynb"
with open(notebook_path, "r", encoding="utf-8") as f:
    nb = json.load(f)

for cell in nb.get("cells", []):
    if cell.get("cell_type") == "code":
        source = cell.get("source", [])
        for i, line in enumerate(source):
            if "Reporte_FC_A_Subir.html" in line:
                source[i] = line.replace('os.path.join(CARPETA_C, "Reporte_FC_A_Subir.html")', 'r"\\\\10.10.10.210\\\\AyF_Trabajoadistancia\\\\Compras\\\\Reporte_Maestro.html"')

with open(notebook_path, "w", encoding="utf-8") as f:
    json.dump(nb, f, indent=1, ensure_ascii=False)
