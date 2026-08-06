import os

with open(r"C:\Users\gcastro\ROBOTS\asistentecontable\Proyecto_facturas_v2\reporter.py", "r", encoding="utf-8") as f:
    lines = f.readlines()

new_content = []
for line in lines[:165]:
    new_content.append(line)

new_logic = """def generar_reporte_html(datos_nuevos=None, ruta_html=HTML_SALIDA):
    estados_por_archivo, estados_por_datos = obtener_estados_autorizacion()
    cache_rutas = cargar_cache()
    
    # 1. Obtener histórico (si existe)
    if os.path.exists(ruta_html):
        try:
            df_historico = pd.read_html(ruta_html)[0]
        except Exception:
            df_historico = pd.DataFrame()
    else:
        df_historico = pd.DataFrame()

    # 2. Escanear 'fc a subir' para encontrar nuevos archivos
    nuevos_datos = []
    ruta_fc = Path(CARPETA_C)
    if ruta_fc.exists():
        for root, dirs, files in os.walk(ruta_fc):
            for file in files:
                if file.lower().endswith('.pdf'):
                    ruta_completa = str(Path(root) / file)
                    info = parsear_nombre_aprobado(file)
                    try:
                        fecha_mod = datetime.fromtimestamp(os.path.getmtime(ruta_completa)).strftime('%d/%m/%Y %H:%M:%S')
                    except Exception:
                        fecha_mod = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
                    
                    nuevos_datos.append({
                        "Fecha": fecha_mod,
                        "Empresa": info["Empresa"],
                        "Razón Social": info["Razón Social"],
                        "Tipo": info["Tipo"],
                        "Nro Comprobante": info["Nro Comprobante"],
                        "OC": info["OC"],
                        "Información Adicional": info["Información Adicional"],
                        "Sector": info["Sector"],
                        "Archivo": info["Archivo Original"],
                        "Ruta Completa": ruta_completa
                    })
    
    df_nuevos = pd.DataFrame(nuevos_datos)
    
    # 3. Unir histórico y nuevos para tener la lista de "rastreados"
    if not df_historico.empty and not df_nuevos.empty:
        for col in df_nuevos.columns:
            if col not in df_historico.columns:
                df_historico[col] = ''
        df_tracked = pd.concat([df_historico, df_nuevos], ignore_index=True)
    elif not df_nuevos.empty:
        df_tracked = df_nuevos
    else:
        df_tracked = df_historico

    if df_tracked.empty:
        print("⚠️ No hay datos en histórico ni en fc a subir para mostrar.")
        return

    df_tracked['Razón Social'] = df_tracked['Razón Social'].fillna('DESCONOCIDO')
    df_tracked['Nro Comprobante'] = df_tracked['Nro Comprobante'].fillna('SinNumero')
    
    # Desduplicar preliminar
    df_tracked.drop_duplicates(subset=["Razón Social", "Nro Comprobante"], keep="last", inplace=True)

    # 4. Crear un diccionario de rastreados para búsqueda rápida
    rastreados = set()
    for _, row in df_tracked.iterrows():
        rs_key = str(row.get('Razón Social', '')).strip().upper()
        nro_key = str(row.get('Nro Comprobante', '')).strip()
        rastreados.add((rs_key, nro_key))

    # 5. Buscar gemelos en carpetas contables
    etapas_encontradas = {} 
    
    ahora = datetime.now().timestamp()
    limite_antiguedad = 60 * 24 * 60 * 60

    for carpeta in CARPETAS_BUSQUEDA:
        ruta_base = Path(carpeta)
        if not ruta_base.exists(): continue
        for root, dirs, files in os.walk(ruta_base):
            root_path = Path(root)
            try:
                if root_path != ruta_base and (ahora - root_path.stat().st_mtime) > limite_antiguedad:
                    dirs.clear()
            except:
                pass
            for file in files:
                if file.lower().endswith('.pdf'):
                    ruta_completa = str(root_path / file)
                    info = parsear_nombre_aprobado(file)
                    
                    rs_key = str(info["Razón Social"]).strip().upper()
                    nro_key = str(info["Nro Comprobante"]).strip()
                    
                    if (rs_key, nro_key) in rastreados:
                        ruta_upper = ruta_completa.upper()
                        etapa_contable = "OTRA"
                        nivel_etapa = 0

                        if r"FACTURAS A PAGAR\\2026" in ruta_upper:
                            etapa_contable = "PAGADA"
                            nivel_etapa = 4
                        elif r"FACTURAS A PAGAR\\PERIODO ACTUAL" in ruta_upper:
                            etapa_contable = "PARA PAGAR"
                            nivel_etapa = 3
                        elif r"\\CONTABILIDAD" in ruta_upper:
                            etapa_contable = "CONTABILIZADA"
                            nivel_etapa = 2
                        elif r"\\COMPRAS" in ruta_upper:
                            etapa_contable = "ENTRADA"
                            nivel_etapa = 1

                        if (rs_key, nro_key) not in etapas_encontradas or nivel_etapa > etapas_encontradas[(rs_key, nro_key)][0]:
                            etapas_encontradas[(rs_key, nro_key)] = (nivel_etapa, etapa_contable, ruta_completa)

    # 6. Actualizar dataframe con etapas y estados
    df_tracked['Nivel Etapa'] = 0
    df_tracked['Etapa Contable'] = 'ENTRADA'
    df_tracked['Estado'] = 'N/A'

    for idx, row in df_tracked.iterrows():
        rs_key = str(row.get('Razón Social', '')).strip().upper()
        tipo_key = str(row.get('Tipo', '')).strip().upper()
        nro_key = str(row.get('Nro Comprobante', '')).strip()
        nombre_arch = str(row.get('Archivo', '')).strip()

        # Actualizar Etapa Contable desde gemelos
        if (rs_key, nro_key) in etapas_encontradas:
            nivel, etapa, ruta = etapas_encontradas[(rs_key, nro_key)]
            df_tracked.at[idx, 'Nivel Etapa'] = nivel
            df_tracked.at[idx, 'Etapa Contable'] = etapa
            df_tracked.at[idx, 'Ruta Completa'] = ruta
        else:
            df_tracked.at[idx, 'Etapa Contable'] = 'ENTRADA'
            df_tracked.at[idx, 'Nivel Etapa'] = 1
            # Si no encontró gemelo en otras carpetas, asumimos que su ruta es la que teníamos
            df_tracked.at[idx, 'Ruta Completa'] = row.get('Ruta Completa', '')

        # Actualizar Estado desde Autorización (N/A por defecto)
        estado = "N/A"
        if nombre_arch in estados_por_archivo:
            estado = estados_por_archivo[nombre_arch]
        elif (rs_key, tipo_key, nro_key) in estados_por_datos:
            estado = estados_por_datos[(rs_key, tipo_key, nro_key)]
        df_tracked.at[idx, 'Estado'] = estado

    df_final = df_tracked.sort_values(by=["Razón Social", "Nro Comprobante", "Nivel Etapa"], ascending=[True, True, True])
    df_final.drop_duplicates(subset=["Razón Social", "Nro Comprobante"], keep="last", inplace=True)
    
    if "Fecha" in df_final.columns:
        df_final = df_final.sort_values(by="Fecha", ascending=False)

    timestamp_actual = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
"""
new_content.append(new_logic)

for line in lines[292:]:
    # Ajustar para que el badge de N/A no se rompa
    if 'badge_estado = \'<span class="badge bg-secondary">PENDIENTE</span>\'' in line:
        line = line.replace('badge_estado = \'<span class="badge bg-secondary">PENDIENTE</span>\'', 
                            'badge_estado = \'<span class="badge bg-danger">N/A</span>\' if estado_val == "N/A" else \'<span class="badge bg-secondary">PENDIENTE</span>\'')
    new_content.append(line)

with open(r"C:\Users\gcastro\ROBOTS\asistentecontable\Proyecto_facturas_v2\reporter.py", "w", encoding="utf-8") as f:
    f.writelines(new_content)
