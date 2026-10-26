import openpyxl
import re
import sys
from unicodedata import normalize as uni_norm

# Ensure utf-8 output printing
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

def clean_name(text):
    if not text:
        return ""
    text = uni_norm('NFD', str(text)).encode('ascii', 'ignore').decode('utf-8')
    text = text.upper().strip()
    # Remove common company types
    text = re.sub(r'\b(S\.?R\.?L\.?|S\.?A\.?|S\.?H\.?|S\.?A\.?C\.?I\.?F\.?I\.?|S\.?A\.?I\.?C\.?|S\.?A\.?U\.?|S\.?A\.?S\.?|INC|LLC|S\.A\.U\.)\b', '', text)
    # Remove punctuation & symbols
    text = re.sub(r'[^A-Z0-9\s]', ' ', text)
    # Remove extra spaces
    text = ' '.join(text.split())
    return text

excel_path = r'C:\Users\gcastro\ROBOTS\data\proveedores (2).xls.xlsx'

wb = openpyxl.load_workbook(excel_path)

if 'Hoja5' not in wb.sheetnames or 'Hoja1' not in wb.sheetnames:
    print("Error: No se encontraron las hojas necesarias ('Hoja1' y 'Hoja5')")
    sys.exit(1)

ws5 = wb['Hoja5']
ws1 = wb['Hoja1']

# 1. Read Hoja1 target rows
hoja1_rows = []
for r in range(2, ws1.max_row + 1):
    desc_val = ws1.cell(r, 2).value  # Col B
    cur_email = ws1.cell(r, 12).value # Col L
    if desc_val:
        desc_str = str(desc_val).strip()
        hoja1_rows.append({
            'row': r,
            'desc_raw': desc_str,
            'desc_clean': clean_name(desc_str),
            'email_cell': ws1.cell(r, 12)
        })

print(f"Hoja1: {len(hoja1_rows)} proveedores cargados.")

# 2. Read Hoja5 source rows
hoja5_providers = []
for r in range(2, ws5.max_row + 1):
    col_d = ws5.cell(r, 4).value   # Col D (Razon Social)
    col_e = ws5.cell(r, 5).value   # Col E (Direccion de Mail)
    col_k = ws5.cell(r, 11).value  # Col K (Clean extracted name)
    
    email = str(col_e).strip() if col_e else ""
    if not email or email.lower() == 'none':
        continue
        
    names_to_try = []
    if col_k and str(col_k).strip():
        names_to_try.append(str(col_k).strip())
    if col_d and str(col_d).strip():
        names_to_try.append(str(col_d).strip())
        
    for raw_name in names_to_try:
        c_name = clean_name(raw_name)
        # Skip generic terms / emails used as name
        if not c_name or c_name in ['COMPRAS', 'PAGO A PROVEEDORES', 'LISTA NEGRA', 'NO FACTURA'] or '@' in raw_name:
            continue
        hoja5_providers.append({
            'row_h5': r,
            'raw_name': raw_name,
            'clean_name': c_name,
            'email': email
        })

print(f"Hoja5: {len(hoja5_providers)} registros de proveedores con mail encontrados.")

# Group Hoja5 providers by clean_name to aggregate emails if a provider appears multiple times
hoja5_grouped = {}
for item in hoja5_providers:
    cn = item['clean_name']
    if cn not in hoja5_grouped:
        hoja5_grouped[cn] = {
            'raw_names': set(),
            'emails': set(),
            'rows_h5': []
        }
    hoja5_grouped[cn]['raw_names'].add(item['raw_name'])
    hoja5_grouped[cn]['emails'].add(item['email'])
    hoja5_grouped[cn]['rows_h5'].append(item['row_h5'])

# 3. Match and Update Hoja1
updated_count = 0
matched_h5_keys = set()
unmatched_h5 = []

for cn, data in hoja5_grouped.items():
    # Attempt matching in Hoja1
    # Strategy A: Clean exact match
    targets = [h for h in hoja1_rows if h['desc_clean'] == cn]
    
    # Strategy B: Clean substring match (if cn length >= 4)
    if not targets and len(cn) >= 4:
        for h in hoja1_rows:
            h_cn = h['desc_clean']
            if not h_cn: continue
            if cn == h_cn or (len(cn) >= 5 and cn in h_cn) or (len(h_cn) >= 5 and h_cn in cn):
                ratio = len(cn) / len(h_cn) if len(h_cn) > 0 else 0
                if 0.45 < ratio < 2.2:
                    targets.append(h)
                    
    if targets:
        matched_h5_keys.add(cn)
        mail_text = ", ".join(sorted(data['emails']))
        for target in targets:
            existing = target['email_cell'].value
            # Only update if empty or different
            if not existing or str(existing).strip() != mail_text:
                target['email_cell'].value = mail_text
                updated_count += 1
    else:
        unmatched_h5.append(data)

# Save workbook
wb.save(excel_path)
print(f"✅ Se actualizaron {updated_count} celdas en la columna L de la Hoja1.")
print(f"📊 De {len(hoja5_grouped)} proveedores únicos de la Hoja5:")
print(f"   - Encontrados en Hoja1: {len(matched_h5_keys)}")
print(f"   - NO encontrados en Hoja1: {len(unmatched_h5)}")

# Save report of unmatched to text/markdown
with open(r'c:\Users\gcastro\ROBOTS\asistentecontable\proveedores_no_encontrados.txt', 'w', encoding='utf-8') as f:
    f.write("PROVEEDORES DE LA HOJA 5 NO ENCONTRADOS EN LA HOJA 1:\n")
    f.write("=" * 65 + "\n\n")
    for idx, item in enumerate(unmatched_h5, 1):
        raw_str = " / ".join(sorted(item['raw_names']))
        mails_str = ", ".join(sorted(item['emails']))
        rows_str = ", ".join(map(str, sorted(set(item['rows_h5']))))
        f.write(f"{idx}. Proveedor / Razón Social: {raw_str}\n")
        f.write(f"   Mail(s): {mails_str}\n")
        f.write(f"   Filas en Hoja5: {rows_str}\n")
        f.write("-" * 65 + "\n")

print("Reporte de no encontrados guardado en 'proveedores_no_encontrados.txt'")
