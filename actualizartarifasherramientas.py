import streamlit as st
import pandas as pd
import io
import zipfile

# Configuración de la interfaz de Streamlit
st.set_page_config(page_title="Generador de Tarifas - Turaco", page_icon="📦", layout="wide")

st.title("📦 Automatización de Tarifas - Turaco Herramientas")
st.write("Carga los archivos maestros de Excel para procesar masivamente todas las plantillas requeridas.")

# --- FUNCIÓN AUXILIAR DE FORMATEO DE SKU ---
def formatear_sku(val):
    if pd.isna(val):
        return ""
    val_str = str(val).strip()
    if val_str.endswith('.0'):
        val_str = val_str[:-2]
    # Si es numérico puro, rellenar con ceros a la izquierda hasta llegar a 5 dígitos
    if val_str.isdigit():
        return val_str.zfill(5)
    # Si es alfanumérico (ej. A01, A90, A70), se mantiene idéntico
    return val_str

# --- CAPTURA DE ARCHIVOS EN LA BARRA LATERAL ---
st.sidebar.header("📂 Archivos Maestros")
archivo_nac = st.sidebar.file_uploader("1. Tarifa Nacional España (.xlsx)", type=["xlsx"])
archivo_int = st.sidebar.file_uploader("2. Tarifa Internacional (.xlsx)", type=["xlsx"])

# Inicializar estructuras de datos de las pestañas
pestañas_nac = {}
pestañas_int = {}

if archivo_nac:
    try:
        xls_nac = pd.ExcelFile(archivo_nac)
        for sheet in xls_nac.sheet_names:
            pestañas_nac[sheet] = pd.read_excel(archivo_nac, sheet_name=sheet)
        st.sidebar.success("✅ Tarifa Nacional cargada.")
    except Exception as e:
        st.sidebar.error(f"Error al cargar Tarifa Nacional: {e}")

if archivo_int:
    try:
        xls_int = pd.ExcelFile(archivo_int)
        for sheet in xls_int.sheet_names:
            pestañas_int[sheet] = pd.read_excel(archivo_int, sheet_name=sheet)
        st.sidebar.success("✅ Tarifa Internacional cargada.")
    except Exception as e:
        st.sidebar.error(f"Error al cargar Tarifa Internacional: {e}")

# --- PESTAÑAS DE LA APLICACIÓN ---
tab1, tab2 = st.tabs(["📋 Bloque 1: Actualizador Características", "💰 Bloque 2: Cargador de Precios"])

# =========================================================================
# SECCIÓN 1: DESCARGA MASIVA POR CARACTERÍSTICAS
# =========================================================================
with tab1:
    st.header("Generación por Lotes de Archivos de Características")
    st.write("Esta sección extraerá automáticamente cada característica obligatoria y creará un archivo independiente con su nombre.")

    # Diccionario con la estructura de características y su pestaña/columna base esperada
    # Clave: Nombre de la característica, Valor: (Pestaña del Excel, Columna del precio en el Excel)
    mapeo_caracteristicas = {
        # España (Tarifa Nacional)
        "PVP_LEROYES": ("Nacional", "PVPR"),
        "PVP_PcComponentes": ("Nacional", "PVPR"),
        "PVPR": ("Nacional", "PVPR"),
        "PVP ESPANA": ("Nacional", "PVPR"),
        "Pvp mediamarkt es": ("Nacional", "PVPR"),
        "PVP_SHEIN_ES": ("Nacional", "PVPR"),
        # Francia (Tarifa Internacional)
        "Prix_France": ("FR-FR", "PVP PUB."),
        "PVP_SHEIN_FR": ("FR-FR", "PVP PUB."),
        # Italia
        "Prix_Italia": ("IT-IT", "PVP PUB."),
        "PVP_SHEIN_IT": ("IT-IT", "PVP PUB."),
        # Alemania
        "prix_Alemania": ("DE-DE", "PVP PUB."),
        "PVP_SHEIN_DE": ("DE-DE", "PVP PUB."),
        # Portugal
        "PVP_SHEIN_PT": ("PT", "PVP PUB."),
        "PVP_PORTUGAL": ("PT", "PVP PUB."),
        # Bélgica
        "PVP_BE": ("BE", "PVP PUB."),
        # Países Bajos
        "PRIXHOLANDA": ("NL", "PVP PUB."),
        "PVP_SHEIN_NL": ("NL", "PVP PUB."),
        # Polonia (Requiere revisión de columna específica en Divisa)
        "PVP_SHEIN_PL": ("PL", "PVP PUB. (SZL)"),
        "PRIX_POLONIA": ("PL", "PVP PUB. (SZL)"),
        # Suecia (Requiere revisión de columna específica en Divisa)
        "PVP Suecia": ("SE", "PVP PUB. (SEK)")
    }

    # Controles para ajustar nombres de pestañas dinámicamente si difieren de lo estimado
    with st.expander("⚙️ Configuración Avanzada de Mapeo de Pestañas"):
        st.caption("Si los nombres de las pestañas de tus archivos reales no coinciden con la configuración estándar, corrígelos aquí:")
        pestaña_nac_real = st.text_input("Nombre de pestaña de España en Tarifa Nacional:", value="T_PRIV")
        pestaña_fr_real = st.text_input("Nombre de pestaña de Francia:", value="FR-FR")
        pestaña_it_real = st.text_input("Nombre de pestaña de Italia:", value="IT-IT")
        pestaña_de_real = st.text_input("Nombre de pestaña de Alemania:", value="DE-DE")
        pestaña_pt_real = st.text_input("Nombre de pestaña de Portugal:", value="PT")
        pestaña_be_real = st.text_input("Nombre de pestaña de Bélgica:", value="BE")
        pestaña_nl_real = st.text_input("Nombre de pestaña de Países Bajos:", value="NL")
        pestaña_pl_real = st.text_input("Nombre de pestaña de Polonia:", value="PL")
        pestaña_se_real = st.text_input("Nombre de pestaña de Suecia:", value="SE")

    # Botón principal para generar el lote completo comprimido en ZIP
    if st.button("📦 Generar y Empaquetar todas las Características"):
        if not archivo_nac or not archivo_int:
            st.error("Por favor, asegúrate de subir ambos archivos (Nacional e Internacional) en la barra lateral para proceder.")
        else:
            # Crear un buffer en memoria para el archivo ZIP resultante
            zip_buffer = io.BytesIO()
            
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                conteo_archivos = 0
                
                # Iterar sobre la lista maestra de características obligatorias
                for nombre_carac, (grupo, col_precio_base) in mapeo_caracteristicas.items():
                    
                    # 1. Resolver a qué dataframe acudir
                    if grupo == "Nacional":
                        df_maestro = pestañas_nac.get(pestaña_nac_real)
                    else:
                        mapeo_pestañas_int = {
                            "FR-FR": pestaña_fr_real, "IT-IT": pestaña_it_real, "DE-DE": pestaña_de_real,
                            "PT": pestaña_pt_real, "BE": pestaña_be_real, "NL": pestaña_nl_real,
                            "PL": pestaña_pl_real, "SE": pestaña_se_real
                        }
                        pestaña_buscada = mapeo_pestañas_int.get(grupo)
                        df_maestro = pestañas_int.get(pestaña_buscada)
                    
                    if df_maestro is not None:
                        # Normalizar nombres de columnas a mayúsculas para evitar fallos de escritura
                        columnas_limpias = {str(c).strip().upper(): c for c in df_maestro.columns}
                        
                        if "REFERENCIA" in columnas_limpias:
                            col_ref_real = columnas_limpias["REFERENCIA"]
                            
                            # Buscar la columna del precio
                            col_precio_real = None
                            col_precio_superior = str(col_precio_base).strip().upper()
                            
                            if col_precio_superior in columnas_limpias:
                                col_precio_real = columnas_limpias[col_precio_superior]
                            else:
                                # Caída alternativa: Buscar cualquier columna parecida o usar la de por defecto "PVP PUB."
                                fallback = [c for c in df_maestro.columns if "PVP" in str(c).upper()]
                                if fallback:
                                    col_precio_real = fallback[0]
                            
                            if col_precio_real:
                                # Construcción del dataframe destino bajo formato (sku, valor)
                                df_out = pd.DataFrame()
                                df_out['sku'] = df_maestro[col_ref_real].apply(formatear_sku)
                                df_out['valor'] = df_maestro[col_precio_real]
                                
                                # Limpieza de registros sin código de producto
                                df_out = df_out[df_out['sku'] != ""]
                                
                                # Convertir DataFrame a formato CSV plano de texto
                                csv_data = df_out.to_csv(index=False, sep=',', encoding='utf-8')
                                
                                # Guardar el fichero dentro del ZIP usando el nombre exacto de la característica
                                zip_file.writestr(f"{nombre_carac}.csv", csv_data)
                                conteo_archivos += 1
                
                if conteo_archivos > 0:
                    st.success(f"¡Éxito! Se han procesado y empaquetado {conteo_archivos} archivos de características.")
                    
                    # Ofrecer el botón de descarga del ZIP completo
                    st.download_button(
                        label="📥 Descargar Pack Completo de Características (.zip)",
                        data=zip_buffer.getvalue(),
                        file_name="caracteristicas_actualizadas.zip",
                        mime="application/zip"
                    )
                else:
                    st.error("No se pudo procesar ningún archivo. Comprueba los nombres de las columnas o pestañas.")

# =========================================================================
# SECCIÓN 2: CARGADOR DE PRECIOS (3 PLANTILLAS HORIZONTALES INDEPENDIENTES)
# =========================================================================
with tab2:
    st.header("Generador del Cargador Masivo de Precios")
    st.write("Genera las 3 plantillas independientes indicadas en los requerimientos operativos.")

    # Cabecera exacta del Cargador de Precios
    columnas_cargador = [
        'reference', 'price_france', 'price_italy', 'price_germany', 'price_portugal', 
        'price_spain', 'price_poland', 'price_holand', 'price_tradeinn_es', 
        'price_aliexpress_es', 'price_makro_es', 'price_mediamarkt_es', 'price_aurgi_es', 
        'price_elcorteingles_es', 'price_makro_de', 'price_makro_it', 'price_carrefour', 
        'price_pccomponentes'
    ]

    col1, col2, col3 = st.columns(3)

    # ---------------------------------------------------------------------
    # PLANTILLA 1: TARIFA NACIONAL ESPAÑA
    # ---------------------------------------------------------------------
    with col1:
        st.subheader("1. Tarifa Nacional España")
        st.caption("Actualiza columnas del ecosistema nacional: price_spain, tradeinn, aliexpress, mediamarkt, aurgi, carrefour y pccomponentes.")
        pestaña_nac_cargador = st.text_input("Confirmar pestaña Nacional:", value="T_PRIV", key="carg_nac")
        
        if st.button("🚀 Crear Fichero 1 (Nacional)"):
            df_maestro = pestañas_nac.get(pestaña_nac_cargador)
            if df_maestro is not None:
                # Buscar la columna REFERENCIA y la columna de precio
                col_ref = [c for c in df_maestro.columns if str(c).strip().upper() == "REFERENCIA"]
                col_pvpr = [c for c in df_maestro.columns if "PVPR" in str(c).upper()]
                
                if col_ref and col_pvpr:
                    df_final = pd.DataFrame(columns=columnas_cargador)
                    df_final['reference'] = df_maestro[col_ref[0]].apply(formatear_sku)
                    
                    # Poblamos el bloque de columnas idénticas requeridas
                    precio_nacional = df_maestro[col_pvpr[0]]
                    df_final['price_spain'] = precio_nacional
                    df_final['price_tradeinn_es'] = precio_nacional
                    df_final['price_aliexpress_es'] = precio_nacional
                    df_final['price_mediamarkt_es'] = precio_nacional
                    df_final['price_aurgi_es'] = precio_nacional
                    df_final['price_carrefour'] = precio_nacional
                    df_final['price_pccomponentes'] = precio_nacional
                    
                    df_final = df_final[df_final['reference'] != ""]
                    
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_final.to_excel(writer, index=False, sheet_name='Sheet1')
                        
                    st.success("Fichero Nacional Creado.")
                    st.download_button(label="📥 Descargar Fichero 1", data=output.getvalue(), file_name="1_Tarifa_Nacional_Espana.xlsx")
                else:
                    st.error("No se encontró 'REFERENCIA' o 'PVPR' en la pestaña indicada.")
            else:
                st.warning("Carga el archivo de la Tarifa Nacional.")

    # ---------------------------------------------------------------------
    # PLANTILLA 2: TARIFA INTERNACIONAL PORTUGAL
    # ---------------------------------------------------------------------
    with col2:
        st.subheader("2. Tarifa Internacional Portugal")
        st.caption("Procedimiento específico para mapear exclusivamente la columna price_portugal.")
        pestaña_pt_cargador = st.text_input("Confirmar pestaña Portugal:", value="PT", key="carg_pt")
        
        if st.button("🚀 Crear Fichero 2 (Portugal)"):
            df_maestro = pestañas_int.get(pestaña_pt_cargador)
            if df_maestro is not None:
                col_ref = [c for c in df_maestro.columns if str(c).strip().upper() == "REFERENCIA"]
                col_pvp = [c for c in df_maestro.columns if "PVP" in str(c).upper()]
                
                if col_ref and col_pvp:
                    df_final = pd.DataFrame(columns=columnas_cargador)
                    df_final['reference'] = df_maestro[col_ref[0]].apply(formatear_sku)
                    df_final['price_portugal'] = df_maestro[col_pvp[0]]
                    df_final = df_final[df_final['reference'] != ""]
                    
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_final.to_excel(writer, index=False, sheet_name='Sheet1')
                        
                    st.success("Fichero Portugal Creado.")
                    st.download_button(label="📥 Descargar Fichero 2", data=output.getvalue(), file_name="2_Tarifa_Internacional_Portugal.xlsx")
                else:
                    st.error("No se encontró 'REFERENCIA' o columnas 'PVP' en la pestaña PT.")
            else:
                st.warning("Carga el archivo de Tarifa Internacional.")

    # ---------------------------------------------------------------------
    # PLANTILLA 3: RESTO DE TARIFA INTERNACIONAL
    # ---------------------------------------------------------------------
    with col3:
        st.subheader("3. Resto de Internacional")
        st.caption("Unifica horizontalmente las columnas faltantes: Francia, Italia, Alemania, Polonia (Zlotis) y Holanda.")
        pestaña_base_cargador = st.text_input("Pestaña de referencia base (ej. Francia):", value="FR-FR", key="carg_resto")
        
        if st.button("🚀 Crear Fichero 3 (Resto Int.)"):
            df_base = pestañas_int.get(pestaña_base_cargador)
            if df_base is not None:
                col_ref_base = [c for c in df_base.columns if str(c).strip().upper() == "REFERENCIA"][0]
                
                df_final = pd.DataFrame(columns=columnas_cargador)
                df_final['reference'] = df_base[col_ref_base].apply(formatear_sku)
                df_final = df_final[df_final['reference'] != ""]
                
                # Función interna para indexar los precios cruzando SKUs
                def extraer_precios_internacionales(nombre_pest, col_filtro_precio="PVP PUB."):
                    df_sheet = pestañas_int.get(nombre_pest)
                    if df_sheet is not None:
                        c_ref = [c for c in df_sheet.columns if str(c).strip().upper() == "REFERENCIA"][0]
                        c_price = [c for c in df_sheet.columns if col_filtro_precio.upper() in str(c).upper()][0]
                        
                        mapping = dict(zip(df_sheet[c_ref].apply(formatear_sku), df_sheet[c_price]))
                        return df_final['reference'].map(mapping)
                    return ""

                # Ejecutar cruce de datos por países correspondientes
                df_final['price_france'] = extraer_precios_internacionales(pestaña_fr_real, "PVP PUB.")
                df_final['price_italy'] = extraer_precios_internacionales(pestaña_it_real, "PVP PUB.")
                df_final['price_germany'] = extraer_precios_internacionales(pestaña_de_real, "PVP PUB.")
                df_final['price_holand'] = extraer_precios_internacionales(pestaña_nl_real, "PVP PUB.")
                df_final['price_poland'] = extraer_precios_internacionales(pestaña_pl_real, "PVP PUB. (SZL)")
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='Sheet1')
                    
                st.success("Fichero unificado de Internacional Creado.")
                st.download_button(label="📥 Descargar Fichero 3", data=output.getvalue(), file_name="3_Resto_Tarifa_Internacional.xlsx")
            else:
                st.warning("Carga el archivo de Tarifa Internacional.")