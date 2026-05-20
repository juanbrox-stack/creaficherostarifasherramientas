import streamlit as st
import pandas as pd
import io
import zipfile
from openpyxl.styles import Alignment, Font

# Configuración de página de Streamlit
st.set_page_config(page_title="Actualizador de Tarifas e Ingestas", layout="wide", page_icon="📊")

# =========================================================================
# FUNCIONES AUXILIARES DE EXTRACCIÓN (LÓGICA INTERNA DE NEGOCIO)
# =========================================================================
def extraer_precios_por_posicion(df, col_ref_idx, col_precio_idx):
    """
    Extrae un diccionario de SKU -> Precio basándose estrictamente en índices de columnas.
    Corrige los prefijos de SKU ('00', '0', etc.) según las reglas del negocio.
    """
    mapping = {}
    if df is None:
        return mapping
        
    for index, row in df.iterrows():
        # Saltamos cabeceras o filas vacías analizando la primera celda identificable
        if index < 1 or col_ref_idx >= len(row) or pd.isna(row.iloc[col_ref_idx]):
            continue
            
        ref_raw = str(row.iloc[col_ref_idx]).strip()
        if ref_raw.upper() in ["REFERENCIA", "SKU", "NOMBRE COMPLETO", ""]:
            continue
            
        # Formatear el SKU según reglas de longitud y tipo
        if ref_raw.isdigit():
            num_int = int(ref_raw)
            if num_int < 120:
                ref_formateada = f"{num_int:03d}"  # 112 -> 112
            else:
                ref_formateada = f"{num_int:05d}"  # 120 -> 00120, 1503 -> 01503
        else:
            ref_formateada = ref_raw  # A01, A90, A70 se quedan igual

        # Extraer y limpiar el precio correspondiente
        precio_val = None
        if col_precio_idx is not None and col_precio_idx < len(row):
            val = row.iloc[col_precio_idx]
            if pd.notna(val):
                try:
                    val_str = str(val).replace('€', '').replace('$', '').replace(' ', '').strip()
                    precio_val = float(val_str)
                except ValueError:
                    pass
        
        mapping[ref_formateada] = precio_val
    return mapping

def exportar_excel_con_cabecera_herramientas(df_datos, columnas_plantilla):
    """
    Genera un archivo .xlsx real con la fila 1 combinada con el texto 'Herramientas'
    y las cabeceras reales desplazadas a la fila 2.
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Desplazar los encabezados reales a la fila 2 (startrow=1 de Python)
        df_datos.to_excel(writer, sheet_name='Sheet1', index=False, startrow=1)
        
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        
        # Forzar el título "Herramientas" en la celda A1 y combinar hasta la última columna (R)
        worksheet['A1'] = "Herramientas"
        worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(columnas_plantilla))
        
        # Formato visual idéntico para que pase el validador estricto
        worksheet['A1'].alignment = Alignment(horizontal='center', vertical='center')
        worksheet['A1'].font = Font(bold=True, size=11)
    return output.getvalue()


# =========================================================================
# BARRA LATERAL - CARGA DE ARCHIVOS FUENTE
# =========================================================================
st.sidebar.header("📂 Carga de Archivos de Origen")
archivo_nac = st.sidebar.file_uploader("1. Tarifa Nacional (Excel)", type=["xlsx"])
archivo_int = st.sidebar.file_uploader("2. Tarifa Internacional (Excel)", type=["xlsx"])

pestañas_nac = {}
pestañas_int = {}

if archivo_nac:
    try:
        xl_nac = pd.ExcelFile(archivo_nac)
        for sheet in xl_nac.sheet_names:
            pestañas_nac[sheet] = xl_nac.parse(sheet, header=None)
        st.sidebar.success(f"Nacional cargado ({len(pestañas_nac)} pestañas)")
    except Exception as e:
        st.sidebar.error(f"Error al leer Nacional: {e}")

if archivo_int:
    try:
        xl_int = pd.ExcelFile(archivo_int)
        for sheet in xl_int.sheet_names:
            pestañas_int[sheet] = xl_int.parse(sheet, header=None)
        st.sidebar.success(f"Internacional cargado ({len(pestañas_int)} pestañas)")
    except Exception as e:
        st.sidebar.error(f"Error al leer Internacional: {e}")


# =========================================================================
# CUERPO PRINCIPAL - ESTRUCTURA DE PESTAÑAS (BLOQUES DE TRABAJO)
# =========================================================================
st.title("⚙️ Sistema Profesional de Procesamiento de Tarifas - Turaco")
tab1, tab2 = st.tabs(["📦 Bloque 1: Características", "🚀 Bloque 2: Cargador de Precios (3 Ficheros)"])

# Mapeo fijo de la Tarifa Nacional (T_AMZ): Col A (0) = SKU, Col P (15) = PVP PUB
idx_ref_nac = 0
idx_pvp_nac = 15

# Estructura horizontal exacta requerida por la plantilla del Cargador de Precios
columnas_plantilla = [
    'reference', 'price_france', 'price_italy', 'price_germany', 'price_portugal',
    'price_spain', 'price_poland', 'price_holand', 'price_tradeinn_es', 'price_aliexpress_es',
    'price_makro_es', 'price_mediamarkt_es', 'price_aurgi_es', 'price_elcorteingles_es',
    'price_makro_de', 'price_makro_it', 'price_carrefour', 'price_pccomponentes'
]

# -------------------------------------------------------------------------
# BLOQUE 1: GENERACIÓN MASIVA POR CARACTERÍSTICAS
# -------------------------------------------------------------------------
with tab1:
    st.header("Generación de Archivos de Características (Formatos Individuales)")
    st.write("Extrae y genera de golpe los archivos independientes para subirlos en la sección 'Actualizador de Características'.")

    reglas_caracteristicas = {
        "PVP_LEROYES": ("Nacional", ["T_AMZ"], idx_ref_nac, idx_pvp_nac), 
        "PVP_PcComponentes": ("Nacional", ["T_AMZ"], idx_ref_nac, idx_pvp_nac),
        "PVPR": ("Nacional", ["T_AMZ"], idx_ref_nac, idx_pvp_nac),
        "PVP ESPANA": ("Nacional", ["T_AMZ"], idx_ref_nac, idx_pvp_nac),
        "Pvp mediamarkt es": ("Nacional", ["T_AMZ"], idx_ref_nac, idx_pvp_nac),
        "PVP_SHEIN_ES": ("Nacional", ["T_AMZ"], idx_ref_nac, idx_pvp_nac),
        "Prix_France": ("Internacional", ["FR-FR", "ES-FR"], 2, 12),
        "PVP_SHEIN_FR": ("Internacional", ["FR-FR", "ES-FR"], 2, 12),
        "Prix_Italia": ("Internacional", ["IT-IT", "ES-IT"], 2, 12),
        "PVP_SHEIN_IT": ("Internacional", ["IT-IT", "ES-IT"], 2, 12),
        "prix_Alemania": ("Internacional", ["DE-DE", "ES-DE"], 2, 12),
        "PVP_SHEIN_DE": ("Internacional", ["DE-DE", "ES-DE"], 2, 12),
        "PVP_SHEIN_PT": ("Internacional", ["PT"], 2, 12),
        "PVP PORTUGAL": ("Internacional", ["PT"], 2, 12),
        "PVP_BE": ("Internacional", ["BE"], 2, 12),
        "PRIXHOLANDA": ("Internacional", ["NL"], 2, 12),
        "PVP_SHEIN_NL": ("Internacional", ["NL"], 2, 12),
        "PVP_SHEIN_PL": ("Internacional", ["PL"], 2, 13),
        "PRIX_POLONIA": ("Internacional", ["PL"], 2, 13),
        "PVP Suecia": ("Internacional", ["SE"], 2, 13)
    }

    if st.button("🔄 Procesar Canales y Características"):
        if not archivo_nac or not archivo_int:
            st.error("Por favor, asegúrate de subir ambos archivos maestros en la barra lateral.")
        else:
            df_nac_base = pestañas_nac.get("T_AMZ", next(iter(pestañas_nac.values())) if pestañas_nac else None)
            diccionario_archivos = {}
            
            for nombre_carac, (tipo, posibles_pestañas, col_ref, col_precio) in reglas_caracteristicas.items():
                df_trabajo = None
                mapping_precios = {}
                
                if tipo == "Nacional":
                    df_trabajo = df_nac_base
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
                    
                    csv_body = df_out.to_csv(index=False, sep=',', encoding='utf-8', lineterminator='\r\n')
                    csv_data = f"sep=,\r\n{csv_body}"
                    diccionario_archivos[nombre_carac] = csv_data

            if diccionario_archivos:
                st.session_state["archivos_caracteristicas"] = diccionario_archivos
                st.success(f"¡Procesamiento completo! Se han generado {len(diccionario_archivos)} archivos de características.")
            else:
                st.error("No se pudo extraer información. Revisa la estructura de los archivos.")

    if "archivos_caracteristicas" in st.session_state:
        archivos = st.session_state["archivos_caracteristicas"]
        st.write("### ⬇️ Descargar Ficheros Individuales (.csv)")
        
        cols = st.columns(3)
        for idx, (nombre_fichero, datos_csv) in enumerate(archivos.items()):
            col_actual = cols[idx % 3]
            with col_actual:
                st.download_button(
                    label=f"📄 {nombre_fichero}.csv",
                    data=datos_csv,
                    file_name=f"{nombre_fichero}.csv",
                    mime="text/csv",
                    key=f"btn_{nombre_fichero}"
                )
        
        st.write("---")
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for nombre_fichero, datos_csv in archivos.items():
                zip_file.writestr(f"{nombre_fichero}.csv", datos_csv)
                
        st.download_button(
            label="📥 Descargar TODOS los archivos de Características (.ZIP)",
            data=zip_buffer.getvalue(),
            file_name="caracteristicas_actualizadas.zip",
            mime="application/zip",
            type="primary"
        )

# -------------------------------------------------------------------------
# BLOQUE 2: CARGADOR DE PRECIOS - SEPARACIÓN EN 3 FICHEROS OFICIALES (.XLSX)
# -------------------------------------------------------------------------
with tab2:
    st.header("Generación del Cargador de Precios (3 Ficheros Independientes)")
    st.write("Genera los 3 archivos separados con formato **Excel (.xlsx)** real y la fila 1 combinada como **'Herramientas'**.")

    tipo_cambio_pln = st.number_input("💵 Tipo de cambio actual para Polonia (1 EUR a PLN):", min_value=0.01, value=4.32, step=0.01)
    
    col1, col2, col3 = st.columns(3)

    # Resolución del dataframe base de España
    df_nac_base = pestañas_nac.get("T_AMZ", next(iter(pestañas_nac.values())) if pestañas_nac else None)

    # --- FICHERO 1: TARIFA NACIONAL ESPAÑA ---
    with col1:
        st.subheader("1. Tarifa Nacional España")
        st.caption("Pobla exclusivamente el ecosistema nacional: price_spain, tradeinn, aliexpress, mediamarkt, aurgi, carrefour y pccomponentes.")
        if st.button("🚀 Generar Fichero 1 (Nacional)"):
            if df_nac_base is not None:
                map_spain = extraer_precios_por_posicion(df_nac_base, idx_ref_nac, idx_pvp_nac)
                
                if map_spain:
                    df_final = pd.DataFrame(columns=columnas_plantilla)
                    df_final['reference'] = list(map_spain.keys())
                    
                    precios = list(map_spain.values())
                    df_final['price_spain'] = precios
                    df_final['price_tradeinn_es'] = precios
                    df_final['price_aliexpress_es'] = precios
                    df_final['price_mediamarkt_es'] = precios
                    df_final['price_aurgi_es'] = precios
                    df_final['price_carrefour'] = precios
                    df_final['price_pccomponentes'] = precios
                    
                    df_final = df_final.fillna("")
                    bytes_excel = exportar_excel_con_cabecera_herramientas(df_final, columnas_plantilla)
                    
                    st.success("✅ Fichero 1 (Nacional) listo.")
                    st.download_button(label="📥 Descargar Tarifa Nacional", data=bytes_excel, file_name="1_Tarifa_Nacional_Espana.xlsx")
            else:
                st.warning("Carga la Tarifa Nacional en la barra lateral.")

    # --- FICHERO 2: TARIFA INTERNACIONAL PORTUGAL ---
    with col2:
        st.subheader("2. Tarifa Internacional Portugal")
        st.caption("Pobla de forma aislada la columna price_portugal (Pestaña 'PT' - Columnas C y M).")
        if st.button("🚀 Generar Fichero 2 (Portugal)"):
            df_pt = pestañas_int.get("PT")
            if df_pt is not None:
                map_pt = extraer_precios_por_posicion(df_pt, 2, 12)
                
                if map_pt:
                    df_final = pd.DataFrame(columns=columnas_plantilla)
                    df_final['reference'] = list(map_pt.keys())
                    df_final['price_portugal'] = list(map_pt.values())
                    
                    df_final = df_final.fillna("")
                    bytes_excel = exportar_excel_con_cabecera_herramientas(df_final, columnas_plantilla)
                    
                    st.success("✅ Fichero 2 (Portugal) listo.")
                    st.download_button(label="📥 Descargar Tarifa Portugal", data=bytes_excel, file_name="2_Tarifa_Internacional_Portugal.xlsx")
            else:
                st.warning("No se encontró la pestaña 'PT' en el archivo Internacional.")

    # --- FICHERO 3: RESTO DE TARIFA INTERNACIONAL ---
    with col3:
        st.subheader("3. Resto de Internacional")
        st.caption("Pobla horizontalmente los países europeos: Francia, Italia, Alemania, Holanda (Col M) y Polonia (Col N en PLN).")
        if st.button("🚀 Generar Fichero 3 (Resto Int.)"):
            if df_nac_base is not None and pestañas_int:
                map_base_nac = extraer_precios_por_posicion(df_nac_base, idx_ref_nac, idx_ref_nac)
                
                df_final = pd.DataFrame(columns=columnas_plantilla)
                df_final['reference'] = list(map_base_nac.keys())
                
                def buscar_y_extraer(lista_pestañas, col_ref, col_precio):
                    for p in lista_pestañas:
                        if p in pestañas_int:
                            return extraer_precios_por_posicion(pestañas_int[p], col_ref, col_precio)
                    return {}

                map_fr = buscar_y_extraer(["FR-FR", "ES-FR"], 2, 12)
                map_it = buscar_y_extraer(["IT-IT", "ES-IT"], 2, 12)
                map_de = buscar_y_extraer(["DE-DE", "ES-DE"], 2, 12)
                map_nl = buscar_y_extraer(["NL"], 2, 12)
                map_pl_eur = buscar_y_extraer(["PL"], 2, 13) # Polonia en columna N

                df_final['price_france'] = df_final['reference'].map(map_fr)
                df_final['price_italy'] = df_final['reference'].map(map_it)
                df_final['price_germany'] = df_final['reference'].map(map_de)
                df_final['price_holand'] = df_final['reference'].map(map_nl)
                
                # Aplicar tasa de cambio ingresada a Polonia
                df_final['price_poland'] = df_final['reference'].map(map_pl_eur).apply(
                    lambda x: round(x * tipo_cambio_pln, 2) if x is not None else ""
                )
                
                df_final = df_final.fillna("")
                bytes_excel = exportar_excel_con_cabecera_herramientas(df_final, columnas_plantilla)
                
                st.success("✅ Fichero 3 (Resto Internacional) listo.")
                st.download_button(label="📥 Descargar Fichero Resto Europa", data=bytes_excel, file_name="3_Resto_de_Internacional.xlsx")
            else:
                st.warning("Carga ambos archivos maestros para consolidar el bloque internacional.")