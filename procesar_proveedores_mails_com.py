import openpyxl
import win32com.client
import re
import sys
import os
from unicodedata import normalize as uni_norm

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
    text = re.sub(r'\b(S\.?R\.?L\.?|S\.?A\.?|S\.?H\.?|S\.?A\.?C\.?I\.?F\.?I\.?|S\.?A\.?I\.?C\.?|S\.?A\.?U\.?|S\.?A\.?S\.?|INC|LLC)\b', '', text)
    text = re.sub(r'[^A-Z0-9\s]', ' ', text)
    text = ' '.join(text.split())
    return text

excel_path = os.path.abspath(r'C:\Users\gcastro\ROBOTS\data\proveedores (2).xls.xlsx')

wb_read = openpyxl.load_workbook(excel_path, data_only=True)
ws1 = wb_read['Hoja1']
ws5 = wb_read['Hoja5']

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
            'cur_email': str(cur_email).strip() if cur_email and str(cur_email) != 'None' else ''
        })

hoja5_providers = []
for r in range(2, ws5.max_row + 1):
    col_d = ws5.cell(r, 4).value   # Col D
    col_e = ws5.cell(r, 5).value   # Col E
    col_k = ws5.cell(r, 11).value  # Col K
    
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
        if not c_name or c_name in ['COMPRAS', 'PAGO A PROVEEDORES', 'LISTA NEGRA', 'NO FACTURA'] or '@' in raw_name:
            continue
        hoja5_providers.append({
            'row_h5': r,
            'raw_name': raw_name,
            'clean_name': c_name,
            'email': email
        })

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

wb_read.close()

# Enhanced matching logic
updates = {}
matched_h5_keys = set()
unmatched_h5 = []

for cn, data in hoja5_grouped.items():
    # 1. Clean exact match
    targets = [h for h in hoja1_rows if h['desc_clean'] == cn]
    
    # 2. Token set match / Substring match
    if not targets:
        cn_tokens = set(cn.split())
        for h in hoja1_rows:
            h_cn = h['desc_clean']
            if not h_cn: continue
            h_tokens = set(h_cn.split())
            
            # If all tokens of cn are in h_cn (and cn length >= 3)
            if len(cn_tokens) >= 1 and len(cn) >= 3:
                if cn_tokens.issubset(h_tokens) or cn in h_cn or h_cn in cn:
                    ratio = len(cn) / len(h_cn) if len(h_cn) > 0 else 0
                    if 0.35 < ratio < 2.5:
                        targets.append(h)
                        
    if targets:
        matched_h5_keys.add(cn)
        mail_text = ", ".join(sorted(data['emails']))
        for target in targets:
            curr = target['cur_email']
            if not curr:
                updates[target['row']] = mail_text
            elif mail_text not in curr:
                combined = ", ".join(sorted(set(curr.split(', ') + list(data['emails']))))
                updates[target['row']] = combined
    else:
        unmatched_h5.append(data)

print(f"Updates prepared for {len(updates)} rows in Hoja1.")
print(f"Matched Hoja5 providers: {len(matched_h5_keys)}")
print(f"Unmatched Hoja5 providers: {len(unmatched_h5)}")

# Perform updates via COM
excel_app = win32com.client.Dispatch("Excel.Application")
excel_app.DisplayAlerts = False
excel_app.Visible = False

try:
    wb_com = excel_app.Workbooks.Open(excel_path)
    sh1 = wb_com.Sheets("Hoja1")
    
    updated_cells = 0
    for row_idx, mail_val in updates.items():
        sh1.Cells(row_idx, 12).Value = mail_val
        updated_cells += 1
        
    wb_com.Save()
    wb_com.Close(True)
    print(f"✅ Éxito: Se actualizaron {updated_cells} filas en la Hoja1 (Columna L).")
except Exception as e:
    print(f"Error escribiendo con COM: {e}")
finally:
    excel_app.Quit()

# Filter out internal/employee names or irrelevant entries from the unmatched report to keep it clean
report_path = r'c:\Users\gcastro\ROBOTS\asistentecontable\proveedores_no_encontrados.txt'
with open(report_path, 'w', encoding='utf-8') as f:
    f.write("PROVEEDORES DE LA HOJA 5 NO ENCONTRADOS EN LA HOJA 1:\n")
    f.write("=" * 65 + "\n\n")
    count = 0
    for item in unmatched_h5:
        raw_str = " / ".join(sorted(item['raw_names']))
        mails_str = ", ".join(sorted(item['emails']))
        rows_str = ", ".join(map(str, sorted(set(item['rows_h5']))))
        
        # Skip if raw_str is empty or just generic recipient
        count += 1
        f.write(f"{count}. Razón Social / Nombre en Hoja5: {raw_str}\n")
        f.write(f"   Mail(s): {mails_str}\n")
        f.write(f"   Filas en Hoja5: {rows_str}\n")
        f.write("-" * 65 + "\n")

print(f"Reporte completo guardado en: {report_path}")
