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
    
    # Si contiene letras en cualquier posición (ej: A01, A90, V0416), se mantiene idéntico
    if any(c.isalpha() for c in val_str):
        return val_str
        
    # Extraer solo los dígitos numéricos
    num_str = "".join(c for c in val_str if c.isdigit())
    if num_str:
        # Rellenar con ceros a la izquierda hasta un formato estándar de 5 dígitos para TODOS
        # (Cubre el 112 -> 00112 y los de 4 dígitos como 1503 -> 01503)
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
    """
    Genera los archivos del Bloque 1 en formato .xlsx real con columnas 'sku' y 'valor'
    forzando la columna 'sku' a formato de texto para que no se pierdan los ceros.
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_datos.to_excel(writer, sheet_name='Sheet1', index=False)
        
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        
        # Forzar formato texto (@) en la columna A (SKU) para mantener los ceros a la izquierda
        for row in range(2, worksheet.max_row + 1):
            cell = worksheet.cell(row=row, column=1)
            if cell.value is not None:
                cell.number_format = '@'
                cell.value = str(cell.value)
                
    return output.getvalue()

def exportar_excel_con_cabecera_herramientas(df_datos, columnas_plantilla):
    """
    Genera un archivo .xlsx real para el Bloque 2 con la fila 1 combinada como 'Herramientas'
    y fuerza la columna 'reference' a formato TEXTO para retener los ceros.
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_datos.to_excel(writer, sheet_name='Sheet1', index=False, startrow=1)
        
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        
        # Inyectar y combinar el encabezado estricto "Herramientas" en la fila 1
        worksheet['A1'] = "Herramientas"
        worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(columnas_plantilla))
        
        worksheet['A1'].alignment = Alignment(horizontal='center', vertical='center')
        worksheet['A1'].font = Font(bold=True, size=11)
        
        # Forzamos que toda la columna A (a partir de la fila 3) sea formato Texto ('@')
        for row in range(3, worksheet.max_row + 1):
            cell = worksheet.cell(row=row, column=1)
            if cell.value is not None:
                cell.number_format = '@'
                cell.value = str(cell.value)
                
    return output.getvalue()


# =========================================================================
# BARRA LATERAL - CARGA DE ARCHIVOS MAESTROS
# =========================================================================
st.sidebar.header("📂 Archivos Maestros de Origen")
archivo_nac = st.sidebar.file_uploader("1. Tarifa Nacional (Excel)", type=["xlsx"])
archivo_int = st.sidebar.file_uploader("2. Tarifa Internacional (Excel)", type=["xlsx"])

pestañas_nac = {}
pestañas_int = {}

if archivo_nac:
    try:
        xl_nac = pd.ExcelFile(archivo_nac)
        for sheet in xl_nac.sheet_names:
            pestañas_nac[sheet] = xl_nac.parse(sheet, header=None)
        st.sidebar.success(f"Tarifa Nacional cargada ({len(pestañas_nac)} pestañas)")
    except Exception as e:
        st.sidebar.error(f"Error al leer Tarifa Nacional: {e}")

if archivo_int:
    try:
        xl_int = pd.ExcelFile(archivo_int)
        for sheet in xl_int.sheet_names:
            pestañas_int[sheet] = xl_int.parse(sheet, header=None)
        st.sidebar.success(f"Tarifa Internacional cargada ({len(pestañas_int)} pestañas)")
    except Exception as e:
        st.sidebar.error(f"Error al leer Tarifa Internacional: {e}")


# =========================================================================
# NÚCLEO DE LA APLICACIÓN - PESTAÑAS DE TRABAJO
# =========================================================================
st.title("⚙️ Sistema Automatizado de Control de Tarifas - Turaco")
tab1, tab2 = st.tabs(["📦 Bloque 1: Características", "🚀 Bloque 2: Cargador de Precios (3 Ficheros)"])

idx_ref_nac = 0
idx_pvp_nac = 15

columnas_plantilla = [
    'reference', 'price_france', 'price_italy', 'price_germany', 'price_portugal',
    'price_spain', 'price_poland', 'price_holand', 'price_tradeinn_es', 'price_aliexpress_es',
    'price_makro_es', 'price_mediamarkt_es', 'price_aurgi_es', 'price_elcorteingles_es',
    'price_makro_de', 'price_makro_it', 'price_carrefour', 'price_pccomponentes'
]

# -------------------------------------------------------------------------
# BLOQUE 1: PROCESAMIENTO MASIVO POR CARACTERÍSTICAS (.XLSX)
# -------------------------------------------------------------------------
with tab1:
    st.header("Generación Masiva por Características")
    st.write("Genera los archivos individuales organizados y limpios en formato **Excel (.xlsx)** para el módulo 'Actualizador de Características'.")

    reglas_caracteristicas = {
        "PVP_LEROYES": ("Nacional", ["T_AMZ"], idx_ref_nac, idx_pvp_nac), 
        "PVP_PcComponentes": ("Nacional", ["T_MM"], idx_ref_nac, idx_pvp_nac),
        "PVPR": ("Nacional", ["T_AMZ"], idx_ref_nac, idx_pvp_nac),
        "PVP ESPANA": ("Nacional", ["T_AMZ"], idx_ref_nac, idx_pvp_nac),
        "Pvp mediamarkt es": ("Nacional", ["T_MM"], idx_ref_nac, idx_pvp_nac),
        "PVP_SHEIN_ES": ("Nacional", ["T_AMZ"], idx_ref_nac, idx_pvp_nac),
        "Prix_France": ("Internacional", ["ES-FR", "FR-FR"], 2, 12),
        "PVP_SHEIN_FR": ("Internacional", ["ES-FR", "FR-FR"], 2, 12),
        "Prix_Italia": ("Internacional", ["ES-IT", "IT-IT"], 2, 12),
        "PVP_SHEIN_IT": ("Internacional", ["ES-IT", "IT-IT"], 2, 12),
        "prix_Alemania": ("Internacional", ["ES-DE", "DE-DE"], 2, 12),
        "PVP_SHEIN_DE": ("Internacional", ["ES-DE", "DE-DE"], 2, 12),
        "PVP_SHEIN_PT": ("Internacional", ["PT"], 2, 12),
        "PVP PORTUGAL": ("Internacional", ["PT"], 2, 12),
        "PVP_BE": ("Internacional", ["BE"], 2, 12),
        "PRIXHOLANDA": ("Internacional", ["NL"], 2, 12),
        "PVP_SHEIN_NL": ("Internacional", ["NL"], 2, 12),
        "PVP_SHEIN_PL": ("Internacional", ["PL"], 2, 13),
        "PRIX_POLONIA": ("Internacional", ["PL"], 2, 13),
        "PVP Suecia": ("Internacional", ["SE"], 2, 13)
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
                    if nombre_carac in ["PVP_PcComponentes", "Pvp mediamarkt es"]:
                        df_trabajo = pestañas_nac.get("T_MM")
                    else:
                        df_trabajo = pestañas_nac.get("T_AMZ")
                    mapping_precios = extraer_precios_por_posicion(df_trabajo, col_ref, col_precio)
                else:
                    for p in posibles_pestañas:
                        if p in pestañas_int:
                            df_trabajo = pestañas_int[p]
                            break
                    mapping_precios = extraer_precios_por_posicion(df_trabajo, col_ref, col_precio)
                
                if mapping_precios:
                    df_out = pd.DataFrame(list(mapping_precios.items()), columns=['sku', 'valor'])
                    df_out['valor'] = df_out['valor'].apply(lambda x: "{:.2f}".format(x) if x is not None else "")
                    
                    # Generar binario del archivo Excel (.xlsx) aplicando formato Texto a la columna A
                    bytes_carac_excel = exportar_caracteristica_excel(df_out, nombre_carac)
                    diccionario_archivos[nombre_carac] = bytes_carac_excel

            if diccionario_archivos:
                st.session_state["archivos_caracteristicas"] = diccionario_archivos
                st.success(f"¡Procesamiento completo! {len(diccionario_archivos)} archivos de características preparados.")
            else:
                st.error("No se pudo extraer información. Revisa la estructura de los archivos.")

    if "archivos_caracteristicas" in st.session_state:
        archivos = st.session_state["archivos_caracteristicas"]
        st.write("### ⬇️ Descarga de Ficheros Individuales (.xlsx)")
        
        cols = st.columns(3)
        for idx, (nombre_fichero, datos_excel) in enumerate(archivos.items()):
            col_actual = cols[idx % 3]
            with col_actual:
                st.download_button(
                    label=f"📊 {nombre_fichero}.xlsx",
                    data=datos_excel,
                    file_name=f"{nombre_fichero}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"btn_{nombre_fichero}"
                )
        
        st.write("---")
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for nombre_fichero, datos_excel in archivos.items():
                zip_file.writestr(f"{nombre_fichero}.xlsx", datos_excel)
                
        st.download_button(
            label="📥 Descargar TODOS los archivos de Características (.ZIP)",
            data=zip_buffer.getvalue(),
            file_name="caracteristicas_actualizadas_excel.zip",
            mime="application/zip",
            type="primary"
        )


# -------------------------------------------------------------------------
# BLOQUE 2: CARGADOR DE PRECIOS - 3 ARCHIVOS INDEPENDIENTES Y OPCIÓN ZIP
# -------------------------------------------------------------------------
with tab2:
    st.header("Generación del Cargador de Precios")
    st.write("Presiona el botón para procesar y construir los archivos correspondientes a las tarifas horizontales.")

    tipo_cambio_pln = st.number_input("💵 Tipo de cambio manual (1 EUR a PLN - Polonia):", min_value=0.01, value=4.32, step=0.01)
    
    if st.button("🚀 Procesar y Preparar Ficheros del Cargador"):
        df_amz = pestañas_nac.get("T_AMZ")
        
        if not archivo_nac or not archivo_int or df_amz is None:
            st.error("Asegúrate de subir ambos archivos y que el Nacional contenga la pestaña 'T_AMZ'.")
        else:
            # --- PROCESAMIENTO FICHERO 1: ESPAÑA ---
            map_amz = extraer_precios_por_posicion(df_amz, 0, 15)
            map_mir = extraer_precios_por_posicion(pestañas_nac.get("T_MIR"), 0, 15)
            map_mm = extraer_precios_por_posicion(pestañas_nac.get("T_MM"), 0, 15)
            map_c4 = extraer_precios_por_posicion(pestañas_nac.get("T_C4"), 0, 15)
            
            df_f1 = pd.DataFrame(columns=columnas_plantilla)
            df_f1['reference'] = list(map_amz.keys())
            df_f1['price_spain'] = df_f1['reference'].map(map_amz)
            df_f1['price_tradeinn_es'] = df_f1['reference'].map(map_amz)
            df_f1['price_aliexpress_es'] = df_f1['reference'].map(map_mir)
            df_f1['price_mediamarkt_es'] = df_f1['reference'].map(map_mm)
            df_f1['price_aurgi_es'] = df_f1['reference'].map(map_amz)
            df_f1['price_carrefour'] = df_f1['reference'].map(map_c4)
            df_f1['price_pccomponentes'] = df_f1['reference'].map(map_mm)
            df_f1 = df_f1.fillna("")
            for c in columnas_plantilla:
                if c != 'reference': df_f1[c] = df_f1[c].apply(lambda x: f"{x:.2f}" if isinstance(x, (int, float)) else x)
            
            bytes_f1 = exportar_excel_con_cabecera_herramientas(df_f1, columnas_plantilla)
            st.session_state["bytes_cargador_f1"] = bytes_f1

            # --- PROCESAMIENTO FICHERO 2: PORTUGAL ---
            df_pt = pestañas_int.get("PT")
            bytes_f2 = None
            if df_pt is not None:
                map_pt = extraer_precios_por_posicion(df_pt, 2, 12)
                df_f2 = pd.DataFrame(columns=columnas_plantilla)
                df_f2['reference'] = list(map_pt.keys())
                df_f2['price_portugal'] = list(map_pt.values())
                df_f2 = df_f2.fillna("")
                for c in columnas_plantilla:
                    if c != 'reference': df_f2[c] = df_f2[c].apply(lambda x: f"{x:.2f}" if isinstance(x, (int, float)) else x)
                bytes_f2 = exportar_excel_con_cabecera_herramientas(df_f2, columnas_plantilla)
                st.session_state["bytes_cargador_f2"] = bytes_f2

            # --- PROCESAMIENTO FICHERO 3: RESTO EUROPA ---
            map_base_nac = extraer_precios_por_posicion(df_amz, 0, 0)
            df_f3 = pd.DataFrame(columns=columnas_plantilla)
            df_f3['reference'] = list(map_base_nac.keys())
            
            def buscar_y_extraer(lista_pestañas, col_ref, col_precio):
                for p in lista_pestañas:
                    if p in pestañas_int: return extraer_precios_por_posicion(pestañas_int[p], col_ref, col_precio)
                return {}

            map_fr = buscar_y_extraer(["ES-FR", "FR-FR"], 2, 12)
            map_it = buscar_y_extraer(["ES-IT", "IT-IT"], 2, 12)
            map_de = buscar_y_extraer(["ES-DE", "DE-DE"], 2, 12)
            map_nl = buscar_y_extraer(["NL"], 2, 12)
            map_pl_eur = buscar_y_extraer(["PL"], 2, 13)

            df_f3['price_france'] = df_f3['reference'].map(map_fr)
            df_f3['price_italy'] = df_f3['reference'].map(map_it)
            df_f3['price_germany'] = df_f3['reference'].map(map_de)
            df_f3['price_holand'] = df_f3['reference'].map(map_nl)
            df_f3['price_poland'] = df_f3['reference'].map(map_pl_eur).apply(lambda x: round(x * tipo_cambio_pln, 2) if x is not None else None)
            df_f3 = df_f3.fillna("")
            for c in columnas_plantilla:
                if c != 'reference': df_f3[c] = df_f3[c].apply(lambda x: f"{x:.2f}" if isinstance(x, (int, float)) else x)
            
            bytes_f3 = exportar_excel_con_cabecera_herramientas(df_f3, columnas_plantilla)
            st.session_state["bytes_cargador_f3"] = bytes_f3
            
            st.success("✅ ¡Los 3 archivos del cargador se han procesado correctamente!")

    if "bytes_cargador_f1" in st.session_state:
        st.write("### ⬇️ Descarga de Plantillas Horizontales (.xlsx)")
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
            data=zip_cargador_buffer.getvalue(),
            file_name="cargador_tarifas_completo.zip",
            mime="application/zip",
            type="primary"
        )
