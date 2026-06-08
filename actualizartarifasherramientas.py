import streamlit as st
import pandas as pd
import io
import zipfile
from openpyxl.styles import Alignment, Font

# Configuración de la página de Streamlit
st.set_page_config(page_title="Actualizador de Tarifas e Ingestas", layout="wide", page_icon="📊")

# =========================================================================
# FUNCIONES AUXILIARES DE EXTRACCIÓN Y LIMPIEZA
# =========================================================================
def formatear_sku(val):
    if pd.isna(val):
        return ""
    val_str = str(val).strip()
    if val_str.endswith('.0'):
        val_str = val_str[:-2]
    
    if any(c.isalpha() for c in val_str):
        return val_str
        
    num_str = "".join(c for c in val_str if c.isdigit())
    if num_str:
        return num_str.zfill(5)
    return val_str

def limpiar_y_convertir_precio(val):
    if pd.isna(val):
        return None
    val_str = str(val).replace('€', '').replace('$', '').replace(' ', '').replace('zł', '').strip()
    if not val_str:
        return None
    
    if ',' in val_str and '.' in val_str:
        val_str = val_str.replace(',', '')
    elif ',' in val_str:
        val_str = val_str.replace(',', '.')
        
    try:
        return round(float(val_str), 2)
    except ValueError:
        return None

def extraer_precios_por_posicion(df, col_ref_idx, col_precio_idx):
    mapping = {}
    if df is None or df.empty:
        return mapping
        
    for index, row in df.iterrows():
        if col_ref_idx >= len(row) or pd.isna(row.iloc[col_ref_idx]):
            continue
            
        ref_raw = str(row.iloc[col_ref_idx]).strip()
        if ref_raw.upper() in ["REFERENCIA", "SKU", "NOMBRE COMPLETO", "REFERENCE", "REF", ""]:
            continue
            
        sku = formatear_sku(ref_raw)
        if not sku:
            continue
            
        precio_val = None
        if col_precio_idx is not None and col_precio_idx < len(row):
            precio_val = limpiar_y_convertir_precio(row.iloc[col_precio_idx])
            
        mapping[sku] = precio_val
    return mapping

def exportar_caracteristica_excel(df_datos, nombre_caracteristica):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_datos.to_excel(writer, sheet_name='Sheet1', index=False)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        
        for row in range(2, worksheet.max_row + 1):
            cell = worksheet.cell(row=row, column=1)
            if cell.value is not None:
                cell.number_format = '@'
                cell.value = str(cell.value)
                
    return output.getvalue()

def exportar_excel_con_cabecera_herramientas(df_datos, columnas_plantilla):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_datos.to_excel(writer, sheet_name='Sheet1', index=False, startrow=1)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        
        worksheet['A1'] = "Herramientas"
        worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(columnas_plantilla))
        
        worksheet['A1'].alignment = Alignment(horizontal='center', vertical='center')
        worksheet['A1'].font = Font(bold=True, size=11)
        
        for row in range(3, worksheet.max_row + 1):
            cell = worksheet.cell(row=row, column=1)
            if cell.value is not None:
                cell.number_format = '@'
                cell.value = str(cell.value)
                
    return output.getvalue()

def letra_columna(n):
    string = ""
    while n >= 0:
        n, remainder = divmod(n, 26)
        string = chr(65 + remainder) + string
        n -= 1
    return string

# =========================================================================
# BARRA LATERAL - CARGA DE ARCHIVOS MAESTROS Y VALIDACIÓN SINCRO
# =========================================================================
st.sidebar.header("📂 Archivos Maestros de Origen")
archivo_nac = st.sidebar.file_uploader("1. Tarifa Nacional (Excel)", type=["xlsx"])
archivo_int = st.sidebar.file_uploader("2. Tarifa Internacional (Excel)", type=["xlsx"])

pestañas_nac = {}
pestañas_int = {}

# Sincronización estricta con tus nombres reales de pestañas (Mayúsculas completas)
pestañas_nac_esperadas = ["AMAZON", "MIRAVIA", "MEDIAMARKT", "CARREFOUR", "PRIVALIA"]
pestañas_int_esperadas = [
    "FRANCIA (ES-FR)", "FRANCIA (FR-FR)", "ITALIA (ES-IT)", "ITALIA (IT-IT)", 
    "ALEMANAI (ES-DE)", "ALEMANIA (DE-DE)", "PORTUGAL", "HOLANDA", "BELGICA", "POLONIA", "SUECIA"
]

if archivo_nac:
    try:
        xl_nac = pd.ExcelFile(archivo_nac)
        pestañas_detectadas = xl_nac.sheet_names
        for sheet in pestañas_detectadas:
            pestañas_nac[sheet] = xl_nac.parse(sheet, header=None)
        
        pestañas_faltantes = [p for p in pestañas_nac_esperadas if p not in pestañas_detectadas]
        pestañas_nuevas = [p for p in pestañas_detectadas if p not in pestañas_nac_esperadas]
        
        if pestañas_faltantes:
            st.sidebar.error(f"⚠️ Faltan pestañas críticas Nacionales: {pestañas_faltantes}")
        if pestañas_nuevas:
            st.sidebar.warning(f"💡 Pestañas adicionales detectadas: {pestañas_nuevas}")
        if not pestañas_faltantes:
            st.sidebar.success("✅ Estructura Nacional validada correctamente.")
    except Exception as e:
        st.sidebar.error(f"Error al leer Tarifa Nacional: {e}")

if archivo_int:
    try:
        xl_int = pd.ExcelFile(archivo_int)
        pestañas_detectadas_int = xl_int.sheet_names
        for sheet in pestañas_detectadas_int:
            pestañas_int[sheet] = xl_int.parse(sheet, header=None)
            
        pestañas_faltantes_int = [p for p in pestañas_int_esperadas if p not in pestañas_detectadas_int]
        pestañas_nuevas_int = [p for p in pestañas_detectadas_int if p not in pestañas_int_esperadas]
        
        if pestañas_faltantes_int:
            st.sidebar.error(f"⚠️ Faltan pestañas críticas Internacionales: {pestañas_faltantes_int}")
        if pestañas_nuevas_int:
            st.sidebar.warning(f"💡 Pestañas adicionales detectadas: {pestañas_nuevas_int}")
        if not pestañas_faltantes_int:
            st.sidebar.success("✅ Estructura Internacional validada correctamente.")
    except Exception as e:
        st.sidebar.error(f"Error al leer Tarifa Internacional: {e}")

# =========================================================================
# PANEL DE CONTROL DINÁMICO DE MAPEO
# =========================================================================
st.write("---")
with st.expander("🛠️ Panel Avanzado de Mapeo y Remapeo de Columnas"):
    st.write("Si cambias el orden de los archivos Excel, puedes ajustar los índices aquí de forma visual:")
    c_map1, c_map2 = st.columns(2)
    
    with c_map1:
        st.markdown("**📌 Canales Nacionales**")
        sample_nac = pestañas_nac.get("AMAZON") if "AMAZON" in pestañas_nac else (next(iter(pestañas_nac.values())) if pestañas_nac else None)
        max_cols_nac = len(sample_nac.columns) if sample_nac is not None else 20
        opciones_cols_nac = [f"Columna {letra_columna(i)} (Índice {i})" for i in range(max_cols_nac)]
        
        re_ref_nac = st.selectbox("Columna de Referencia / SKU (Nacional):", opciones_cols_nac, index=0)
        re_pvp_amz = st.selectbox("Columna PVP PUB para AMAZON, MIRAVIA, CARREFOUR:", opciones_cols_nac, index=15)
        re_pvp_mm  = st.selectbox("Columna PVP PUB para MEDIAMARKT:", opciones_cols_nac, index=16)
        
        idx_ref_nac = opciones_cols_nac.index(re_ref_nac)
        idx_pvp_nac = opciones_cols_nac.index(re_pvp_amz)
        idx_mm_pub  = opciones_cols_nac.index(re_pvp_mm)

    with c_map2:
        st.markdown("**📌 Canales Internacionales**")
        sample_int = pestañas_int.get("FRANCIA (ES-FR)") if "FRANCIA (ES-FR)" in pestañas_int else (next(iter(pestañas_int.values())) if pestañas_int else None)
        max_cols_int = len(sample_int.columns) if sample_int is not None else 20
        opciones_cols_int = [f"Columna {letra_columna(i)} (Índice {i})" for i in range(max_cols_int)]
        
        re_ref_int = st.selectbox("Columna de Referencia / SKU (Internacional):", opciones_cols_int, index=2)
        re_pvp_int = st.selectbox("Columna PVP PUB Estándar (FR, IT, DE, PT, NL, BE):", opciones_cols_int, index=12)
        re_pvp_div = st.selectbox("Columna PVP Divisas Especiales (PL, SE):", opciones_cols_int, index=13)
        
        idx_ref_int = opciones_cols_int.index(re_ref_int)
        idx_pvp_int = opciones_cols_int.index(re_pvp_int)
        idx_div_pub = opciones_cols_int.index(re_pvp_div)

columnas_plantilla = [
    'reference', 'price_france', 'price_italy', 'price_germany', 'price_portugal',
    'price_spain', 'price_poland', 'price_holand', 'price_tradeinn_es', 'price_aliexpress_es',
    'price_makro_es', 'price_mediamarkt_es', 'price_aurgi_es', 'price_elcorteingles_es',
    'price_makro_de', 'price_makro_it', 'price_carrefour', 'price_pccomponentes'
]
columnas_de_precio_totales = [c for c in columnas_plantilla if c != 'reference']

# =========================================================================
# VISTA DE PESTAÑAS DE TRABAJO (BLOQUES)
# =========================================================================
tab1, tab2 = st.tabs(["📦 Bloque 1: Características", "🚀 Bloque 2: Cargador de Precios (3 Ficheros)"])

# -------------------------------------------------------------------------
# BLOQUE 1: PROCESAMIENTO MASIVO POR CARACTERÍSTICAS
# -------------------------------------------------------------------------
with tab1:
    st.header("Generación Masiva por Características")
    st.write("Genera los archivos individuales organizados y limpios en formato **Excel (.xlsx)**.")

    reglas_caracteristicas = {
        "PVP_LEROYES": ("Nacional", ["AMAZON"], idx_ref_nac, idx_pvp_nac), 
        "PVP_PcComponentes": ("Nacional", ["MEDIAMARKT"], idx_ref_nac, idx_mm_pub),
        "PVPR": ("Nacional", ["AMAZON"], idx_ref_nac, idx_pvp_nac),
        "PVP ESPANA": ("Nacional", ["AMAZON"], idx_ref_nac, idx_pvp_nac),
        "Pvp mediamarkt es": ("Nacional", ["MEDIAMARKT"], idx_ref_nac, idx_mm_pub),
        "PVP_SHEIN_ES": ("Nacional", ["AMAZON"], idx_ref_nac, idx_pvp_nac),
        "Prix_France": ("Internacional", ["FRANCIA (ES-FR)"], idx_ref_int, idx_pvp_int),
        "PVP_SHEIN_FR": ("Internacional", ["FRANCIA (ES-FR)"], idx_ref_int, idx_pvp_int),
        "Prix_Italia": ("Internacional", ["ITALIA (ES-IT)"], idx_ref_int, idx_pvp_int),
        "PVP_SHEIN_IT": ("Internacional", ["ITALIA (ES-IT)"], idx_ref_int, idx_pvp_int),
        "prix_Alemania": ("Internacional", ["ALEMANAI (ES-DE)"], idx_ref_int, idx_pvp_int),
        "PVP_SHEIN_DE": ("Internacional", ["ALEMANAI (ES-DE)"], idx_ref_int, idx_pvp_int),
        "PVP_SHEIN_PT": ("Internacional", ["PORTUGAL"], idx_ref_int, idx_pvp_int),
        "PVP PORTUGAL": ("Internacional", ["PORTUGAL"], idx_ref_int, idx_pvp_int),
        "PVP_BE": ("Internacional", ["BELGICA"], idx_ref_int, idx_pvp_int),
        "PRIXHOLANDA": ("Internacional", ["HOLANDA"], idx_ref_int, idx_pvp_int),
        "PVP_SHEIN_NL": ("Internacional", ["HOLANDA"], idx_ref_int, idx_pvp_int),
        "PVP_SHEIN_PL": ("Internacional", ["POLONIA"], idx_ref_int, idx_div_pub),
        "PRIX_POLONIA": ("Internacional", ["POLONIA"], idx_ref_int, idx_div_pub),
        "PVP Suecia": ("Internacional", ["SUECIA"], idx_ref_int, idx_div_pub)
    }

    if st.button("📦 Generar todas las Características"):
        if not archivo_nac or not archivo_int:
            st.error("Por favor, asegúrate de subir ambos archivos maestros en la barra lateral.")
        else:
            diccionario_archivos = {}
            for nombre_carac, (tipo, posibles_pestañas, col_ref, col_precio) in reglas_caracteristicas.items():
                df_trabajo = None
                mapping_precios = {}
                
                if tipo == "Nacional":
                    df_trabajo = pestañas_nac.get(posibles_pestañas[0])
                    mapping_precios = extraer_precios_por_posicion(df_trabajo, col_ref, col_precio)
                else:
                    for p in posibles_pestañas:
                        if p in pestañas_int:
                            df_trabajo = pestañas_int[p]
                            break
                    mapping_precios = extraer_precios_por_posicion(df_trabajo, col_ref, col_precio)
                
                if mapping_precios:
                    df_out = pd.DataFrame(list(mapping_precios.items()), columns=['sku', 'valor'])
                    df_out = df_out[df_out['valor'].notna()]
                    df_out['valor'] = df_out['valor'].apply(lambda x: "{:.2f}".format(x) if isinstance(x, (int, float)) else str(x))
                    df_out = df_out[df_out['valor'] != ""]
                    
                    if not df_out.empty:
                        diccionario_archivos[nombre_carac] = exportar_caracteristica_excel(df_out, nombre_carac)

            if diccionario_archivos:
                st.session_state["archivos_caracteristicas"] = diccionario_archivos
                st.success(f"¡Procesamiento completo! {len(diccionario_archivos)} archivos de características preparados.")
            else:
                st.error("No se pudo extraer información. Revisa los índices del panel de mapeo.")

    if "archivos_caracteristicas" in st.session_state:
        archivos = st.session_state["archivos_caracteristicas"]
        st.write("### ⬇️ Descarga de Ficheros Individuales (.xlsx)")
        cols = st.columns(3)
        for idx, (nombre_fichero, datos_excel) in enumerate(archivos.items()):
            col_actual = cols[idx % 3]
            with col_actual:
                st.download_button(
                    label=f"📊 {nombre_fichero}.xlsx", data=datos_excel,
                    file_name=f"{nombre_fichero}.xlsx", key=f"btn_{nombre_fichero}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        st.write("---")
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for nombre_fichero, datos_excel in archivos.items():
                zip_file.writestr(f"{nombre_fichero}.xlsx", datos_excel)
        st.download_button(
            label="📥 Descargar TODOS los archivos de Características (.ZIP)",
            data=zip_buffer.getvalue(), file_name="caracteristicas_actualizadas_excel.zip",
            mime="application/zip", type="primary"
        )

# -------------------------------------------------------------------------
# BLOQUE 2: CARGADOR DE PRECIOS
# -------------------------------------------------------------------------
with tab2:
    st.header("Generación del Cargador de Precios")
    st.write("Presiona el botón para procesar las tarifas horizontales.")
    
    if st.button("🚀 Procesar y Preparar Ficheros del Cargador"):
        df_amz = pestañas_nac.get("AMAZON")
        if not archivo_nac or not archivo_int or df_amz is None:
            st.error("Asegúrate de subir ambos archivos y que el Nacional contenga la pestaña 'AMAZON'.")
        else:
            # --- PROCESAMIENTO FICHERO 1: ESPAÑA ---
            map_amz = extraer_precios_por_posicion(df_amz, idx_ref_nac, idx_pvp_nac)
            map_mir = extraer_precios_por_posicion(pestañas_nac.get("MIRVIA"), idx_ref_nac, idx_pvp_nac)
            map_mm = extraer_precios_por_posicion(pestañas_nac.get("MEDIAMARKT"), idx_ref_nac, idx_mm_pub)
            map_c4 = extraer_precios_por_posicion(pestañas_nac.get("CARREFOUR"), idx_ref_nac, idx_pvp_nac)
            
            df_f1 = pd.DataFrame(columns=columnas_plantilla)
            df_f1['reference'] = list(map_amz.keys())
            df_f1['price_spain'] = df_f1['reference'].map(map_amz)
            df_f1['price_tradeinn_es'] = df_f1['reference'].map(map_amz)
            df_f1['price_aliexpress_es'] = df_f1['reference'].map(map_mir)
            df_f1['price_mediamarkt_es'] = df_f1['reference'].map(map_mm)
            df_f1['price_aurgi_es'] = df_f1['reference'].map(map_amz)
            df_f1['price_carrefour'] = df_f1['reference'].map(map_c4)
            df_f1['price_pccomponentes'] = df_f1['reference'].map(map_mm)
            
            df_f1 = df_f1.dropna(subset=columnas_de_precio_totales, how='all')
            df_f1 = df_f1.fillna("")
            for c in columnas_plantilla:
                if c != 'reference' and df_f1[c].astype(str).str.strip().str.len().gt(0).any():
                    df_f1[c] = df_f1[c].apply(lambda x: f"{x:.2f}" if isinstance(x, (int, float)) else x)
            st.session_state["bytes_cargador_f1"] = exportar_excel_con_cabecera_herramientas(df_f1, columnas_plantilla)

            # --- PROCESAMIENTO FICHERO 2: PORTUGAL ---
            df_pt = pestañas_int.get("PORTUGAL")
            if df_pt is not None:
                map_pt = extraer_precios_por_posicion(df_pt, idx_ref_int, idx_pvp_int)
                df_f2 = pd.DataFrame(columns=columnas_plantilla)
                df_f2['reference'] = list(map_pt.keys())
                df_f2['price_portugal'] = list(map_pt.values())
                
                df_f2 = df_f2.dropna(subset=columnas_de_precio_totales, how='all')
                df_f2 = df_f2.fillna("")
                for c in columnas_plantilla:
                    if c != 'reference': df_f2[c] = df_f2[c].apply(lambda x: f"{x:.2f}" if isinstance(x, (int, float)) else x)
                st.session_state["bytes_cargador_f2"] = exportar_excel_con_cabecera_herramientas(df_f2, columnas_plantilla)
            else:
                st.session_state["bytes_cargador_f2"] = None

            # --- PROCESAMIENTO FICHERO 3: RESTO EUROPA ---
            map_base_nac = extraer_precios_por_posicion(df_amz, idx_ref_nac, idx_ref_nac)
            df_f3 = pd.DataFrame(columns=columnas_plantilla)
            df_f3['reference'] = list(map_base_nac.keys())
            
            def buscar_y_extraer(lista_pestañas, col_ref, col_precio):
                for p in lista_pestañas:
                    if p in pestañas_int: return extraer_precios_por_posicion(pestañas_int[p], col_ref, col_precio)
                return {}

            map_fr = buscar_y_extraer(["FRANCIA (ES-FR)"], idx_ref_int, idx_pvp_int)
            map_it = buscar_y_extraer(["ITALIA (ES-IT)"], idx_ref_int, idx_pvp_int)
            map_de = buscar_y_extraer(["ALEMANAI (ES-DE)"], idx_ref_int, idx_pvp_int)
            map_nl = buscar_y_extraer(["HOLANDA"], idx_ref_int, idx_pvp_int)
            map_be = buscar_y_extraer(["BELGICA"], idx_ref_int, idx_pvp_int)
            map_pl_zlotis = buscar_y_extraer(["POLONIA"], idx_ref_int, idx_div_pub)

            df_f3['price_france'] = df_f3['reference'].map(map_fr)
            df_f3['price_italy'] = df_f3['reference'].map(map_it)
            df_f3['price_germany'] = df_f3['reference'].map(map_de)
            
            serie_nl = df_f3['reference'].map(map_nl)
            serie_be = df_f3['reference'].map(map_be)
            df_f3['price_holand'] = serie_nl.combine_first(serie_be)
            df_f3['price_poland'] = df_f3['reference'].map(map_pl_zlotis)
            
            df_f3 = df_f3.dropna(subset=columnas_de_precio_totales, how='all')
            df_f3 = df_f3.fillna("")
            for c in columnas_plantilla:
                if c != 'reference': df_f3[c] = df_f3[c].apply(lambda x: f"{x:.2f}" if isinstance(x, (int, float)) else x)
            
            st.session_state["bytes_cargador_f3"] = exportar_excel_con_cabecera_herramientas(df_f3, columnas_plantilla)
            st.success("¡Los 3 archivos del cargador se han procesado de forma limpia!")

if "bytes_cargador_f1" in st.session_state:
    st.write("### ⬇ ... Descarga de Plantillas Horizontales (.xlsx)")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.subheader("1. Tarifa Nacional España")
        st.download_button(label="📥 Descargar Tarifa Nacional", data=st.session_state["bytes_cargador_f1"], file_name="1_Tarifa_Nacional_Espana.xlsx")
    with col2:
        st.subheader("2. Tarifa Internacional Portugal")
        if st.session_state["bytes_cargador_f2"] is not None:
            st.download_button(label="📥 Descargar Tarifa Portugal", data=st.session_state["bytes_cargador_f2"], file_name="2_Tarifa_Internacional_Portugal.xlsx")
        else:
            st.caption("No se generaron datos para Portugal.")
    with col3:
        st.subheader("3. Resto de Internacional")
        st.download_button(label="📥 Descargar Tarifa Resto Europa", data=st.session_state["bytes_cargador_f3"], file_name="3_Resto_de_Internacional.xlsx")
        
    st.write("---")
    st.write("### 🗜️ Opción de Descarga Conjunta del Cargador")
    zip_cargador_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_cargador_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("1_Tarifa_Nacional_Espana.xlsx", st.session_state["bytes_cargador_f1"])
        if st.session_state["bytes_cargador_f2"] is not None:
            zf.writestr("2_Tarifa_Internacional_Portugal.xlsx", st.session_state["bytes_cargador_f2"])
        zf.writestr("3_Resto_de_Internacional.xlsx", st.session_state["bytes_cargador_f3"])
        
    st.download_button(
        label="📦 Descargar los 3 archivos del Cargador juntos (.ZIP)",
        data=zip_cargador_buffer.getvalue(), file_name="cargador_tarifas_completo.zip",
        mime="application/zip", type="primary"
    )
