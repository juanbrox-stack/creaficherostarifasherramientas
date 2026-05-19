import streamlit as st
import pandas as pd
import io
import zipfile
import re

# Configuración de la interfaz de la aplicación de Streamlit
st.set_page_config(page_title="Generador de Tarifas - Turaco", page_icon="📦", layout="wide")

st.title("📦 Automatización de Tarifas - Turaco Herramientas")
st.write("Carga los archivos maestros de Excel para procesar masivamente todas las plantillas de subida requeridas.")

# --- FUNCIÓN AUXILIAR DE FORMATEO DE SKU ---
def formatear_sku(val):
    if pd.isna(val):
        return ""
    val_str = str(val).strip()
    if val_str.endswith('.0'):
        val_str = val_str[:-2]
    # Si es numérico puro, rellenar con ceros a la izquierda hasta llegar a 5 dígitos (ej: 120 -> 00120)
    if val_str.isdigit():
        return val_str.zfill(5)
    # Si es alfanumérico (ej. A01, A90, A70, V0416), se mantiene idéntico
    return val_str

# --- FUNCIÓN AUXILIAR PARA LIMPIAR Y CONVERTIR PRECIOS EN TEXTO/DIVISAS ---
def limpiar_y_convertir_precio(val):
    if pd.isna(val):
        return None
    val_str = str(val).strip()
    # Eliminar cualquier carácter que no sea número, signo menos, coma o punto
    val_limpio = re.sub(r'[^\d.,-]', '', val_str)
    
    if not val_limpio:
        return None
        
    # Si contiene comas y puntos (ej: 1,200.50), quitamos la coma de miles
    if ',' in val_limpio and '.' in val_limpio:
        val_limpio = val_limpio.replace(',', '')
    # Si solo tiene comas como separador decimal (ej: 32,95), la cambiamos por punto
    elif ',' in val_limpio:
        val_limpio = val_limpio.replace(',', '.')
    
    try:
        return round(float(val_limpio), 2)
    except ValueError:
        return None

# --- FUNCIÓN PARA EXTRAER PRECIOS POR POSICIÓN ESTRICTA DE COLUMNA ---
def extraer_precios_por_posicion(df, idx_col_ref, idx_col_precio):
    mapping = {}
    if df is None or df.empty:
        return mapping
        
    for _, row in df.iterrows():
        if len(row) > max(idx_col_ref, idx_col_precio):
            val_ref = row.iloc[idx_col_ref]
            val_precio = row.iloc[idx_col_precio]
            
            # Omitir si la celda es vacía o es la fila de encabezados de la tabla
            if pd.isna(val_ref) or str(val_ref).strip().upper() in ["REFERENCIA", "REFERENCE", "REF", "REF."]:
                continue
                
            sku = formatear_sku(val_ref)
            if not sku:
                continue
                
            precio = limpiar_y_convertir_precio(val_precio)
            if precio is not None:
                mapping[sku] = precio
    return mapping

# --- BARRA LATERAL: CAPTURA DE ARCHIVOS MAESTROS ---
st.sidebar.header("📂 Archivos Maestros de Origen")
archivo_nac = st.sidebar.file_uploader("1. Tarifa Nacional España (.xlsx)", type=["xlsx"])
archivo_int = st.sidebar.file_uploader("2. Tarifa Internacional (.xlsx)", type=["xlsx"])

# Inicializar diccionarios para contener las pestañas indexadas
pestañas_nac = {}
pestañas_int = {}

if archivo_nac:
    try:
        xls_nac = pd.ExcelFile(archivo_nac)
        for sheet in xls_nac.sheet_names:
            # Leemos con header=None para poder controlar las posiciones de columnas absolutas (Columna C=2, M=12, N=13)
            pestañas_nac[sheet] = pd.read_excel(archivo_nac, sheet_name=sheet, header=None)
        st.sidebar.success("✅ Tarifa Nacional cargada correctamente.")
    except Exception as e:
        st.sidebar.error(f"Error al cargar Tarifa Nacional: {e}")

if archivo_int:
    try:
        xls_int = pd.ExcelFile(archivo_int)
        for sheet in xls_int.sheet_names:
            pestañas_int[sheet] = pd.read_excel(archivo_int, sheet_name=sheet, header=None)
        st.sidebar.success("✅ Tarifa Internacional cargada correctamente.")
    except Exception as e:
        st.sidebar.error(f"Error al cargar Tarifa Internacional: {e}")

# Definición de pestañas de la interfaz
tab1, tab2 = st.tabs(["📋 Bloque 1: Actualizador Características", "💰 Bloque 2: Cargador de Precios"])

# Configuración por defecto de las pestañas reales en los Excel
pestaña_nac_defecto = "T_PRIV"

# =========================================================================
# SECCIÓN 1: DESCARGA MASIVA POR CARACTERÍSTICAS (BLOQUE 1)
# =========================================================================
with tab1:
    st.header("Generación por Lotes de Archivos de Características")
    st.write("Esta sección extraerá automáticamente cada característica obligatoria y creará un archivo independiente con su nombre.")

    # Mapeo de reglas basadas en tus indicaciones: 
    # Formato: "Nombre_Caracteristica": (Tipo_Tarifa, Lista_Posibles_Pestañas, Indice_Col_Ref, Indice_Col_Precio)
    # Columna C = 2, Columna M = 12, Columna N = 13
    reglas_caracteristicas = {
        # España (Tarifa Nacional) -> Mapeamos dinámicamente usando la columna de precio PVPR (asumimos Columna D/E o la buscamos)
        "PVP_LEROYES": ("Nacional", [pestaña_nac_defecto], 2, None), 
        "PVP_PcComponentes": ("Nacional", [pestaña_nac_defecto], 2, None),
        "PVPR": ("Nacional", [pestaña_nac_defecto], 2, None),
        "PVP ESPANA": ("Nacional", [pestaña_nac_defecto], 2, None),
        "Pvp mediamarkt es": ("Nacional", [pestaña_nac_defecto], 2, None),
        "PVP_SHEIN_ES": ("Nacional", [pestaña_nac_defecto], 2, None),
        # Francia (Col C, Col M)
        "Prix_France": ("Internacional", ["FR-FR", "ES-FR"], 2, 12),
        "PVP_SHEIN_FR": ("Internacional", ["FR-FR", "ES-FR"], 2, 12),
        # Italia (Col C, Col M)
        "Prix_Italia": ("Internacional", ["IT-IT", "ES-IT"], 2, 12),
        "PVP_SHEIN_IT": ("Internacional", ["IT-IT", "ES-IT"], 2, 12),
        # Alemania (Col C, Col M)
        "prix_Alemania": ("Internacional", ["DE-DE", "ES-DE"], 2, 12),
        "PVP_SHEIN_DE": ("Internacional", ["DE-DE", "ES-DE"], 2, 12),
        # Portugal (Col C, Col M)
        "PVP_SHEIN_PT": ("Internacional", ["PT"], 2, 12),
        "PVP PORTUGAL": ("Internacional", ["PT"], 2, 12),
        # Bélgica (Col C, Col M)
        "PVP_BE": ("Internacional", ["BE"], 2, 12),
        # Países Bajos (Col C, Col M)
        "PRIXHOLANDA": ("Internacional", ["NL"], 2, 12),
        "PVP_SHEIN_NL": ("Internacional", ["NL"], 2, 12),
        # Polonia (Col C, Col N)
        "PVP_SHEIN_PL": ("Internacional", ["PL"], 2, 13),
        "PRIX_POLONIA": ("Internacional", ["PL"], 2, 13),
        # Suecia (Col C, Col N)
        "PVP Suecia": ("Internacional", ["SE"], 2, 13)
    }

    if st.button("📦 Generar y Empaquetar todas las Características"):
        if not archivo_nac or not archivo_int:
            st.error("Por favor, asegúrate de subir ambos archivos (Nacional e Internacional) en la barra lateral para proceder.")
        else:
            zip_buffer = io.BytesIO()
            conteo_archivos = 0
            
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                
                # Obtener mapeo base de la Tarifa Nacional de forma flexible para España
                df_nac_base = None
                for name in [pestaña_nac_defecto] + list(pestañas_nac.keys()):
                    if name in pestañas_nac:
                        df_nac_base = pestañas_nac[name]
                        break
                
                # Buscar dinámicamente la columna que contiene "PVPR" y "REFERENCIA" en la hoja Nacional si no son fijas
                idx_ref_nac, idx_pvp_nac = 2, 4 # Fallbacks por defecto (C y E)
                if df_nac_base is not None:
                    for r_idx, row in df_nac_base.head(5).iterrows():
                        row_strs = [str(item).strip().upper() for item in row]
                        if "REFERENCIA" in row_strs:
                            idx_ref_nac = row_strs.index("REFERENCIA")
                        for idx, item in enumerate(row_strs):
                            if "PVPR" in item:
                                idx_pvp_nac = idx
                                break

                # Bucle de extracción masiva
                for nombre_carac, (tipo, posibles_pestañas, col_ref, col_precio) in reglas_caracteristicas.items():
                    df_trabajo = None
                    mapping_precios = {}
                    
                    if tipo == "Nacional":
                        df_trabajo = df_nac_base
                        mapping_precios = extraer_precios_por_posicion(df_trabajo, idx_ref_nac, idx_pvp_nac)
                    else:
                        # Buscar cuál pestaña internacional existe de la lista de opciones
                        for p in posibles_pestañas:
                            if p in pestañas_int:
                                df_trabajo = pestañas_int[p]
                                break
                        mapping_precios = extraer_precios_por_posicion(df_trabajo, col_ref, col_precio)
                    
                    if mapping_precios:
                        # Estructurar DataFrame para subir por características (sku, valor)
                        df_out = pd.DataFrame(list(mapping_precios.items()), columns=['sku', 'valor'])
                        
                        # Generación del string del CSV con la directiva invisible de separación para Excel
                        csv_data = "sep=,\n" + df_out.to_csv(index=False, sep=',', encoding='utf-8')
                        
                        # Escribir el fichero dentro del ZIP con el nombre exacto de la característica
                        zip_file.writestr(f"{nombre_carac}.csv", csv_data)
                        conteo_archivos += 1
                        
                if conteo_archivos > 0:
                    st.success(f"¡Éxito! Se han procesado y empaquetado {conteo_archivos} archivos de características listos para subir.")
                    st.download_button(
                        label="📥 Descargar Pack Completo de Características (.zip)",
                        data=zip_buffer.getvalue(),
                        file_name="caracteristicas_actualizadas.zip",
                        mime="application/zip"
                    )
                else:
                    st.error("No se pudo extraer información. Revisa los archivos de origen.")

# =========================================================================
# SECCIÓN 2: CARGADOR DE PRECIOS (3 PLANTILLAS HORIZONTALES CON FILA COMBINADA)
# =========================================================================
with tab2:
    st.header("Generación del Fichero Maestro de Carga de Precios")
    st.write("Genera las 3 plantillas independientes con la cabecera obligatoria unificada 'Herramientas'.")

    columnas_plantilla = [
        'reference', 'price_france', 'price_italy', 'price_germany', 'price_portugal',
        'price_spain', 'price_poland', 'price_holand', 'price_tradeinn_es', 'price_aliexpress_es',
        'price_makro_es', 'price_mediamarkt_es', 'price_aurgi_es', 'price_elcorteingles_es',
        'price_makro_de', 'price_makro_it', 'price_carrefour', 'price_pccomponentes'
    ]

    # Resolución de dataframes maestros iniciales
    df_nac_base = next(iter(pestañas_nac.values())) if pestañas_nac else None

    # Función común para inyectar la fila combinada "Herramientas" en un archivo Excel real (.xlsx)
    def exportar_excel_con_cabecera_herramientas(df_datos):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Desplazar los encabezados reales a la fila 2 (startrow=1)
            df_datos.to_excel(writer, sheet_name='Sheet1', index=False, startrow=1)
            
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']
            
            # Forzar el título "Herramientas" en la celda A1 y combinar hasta la columna R
            worksheet['A1'] = "Herramientas"
            worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(columnas_plantilla))
            
            from openpyxl.styles import Alignment, Font
            worksheet['A1'].alignment = Alignment(horizontal='center', vertical='center')
            worksheet['A1'].font = Font(bold=True, size=11)
        return output.getvalue()

    col1, col2, col3 = st.columns(3)

    # ---------------------------------------------------------------------
    # PLANTILLA 1: TARIFA NACIONAL ESPAÑA
    # ---------------------------------------------------------------------
    with col1:
        st.subheader("1. Tarifa Nacional España")
        st.caption("Actualiza columnas del ecosistema nacional usando los precios de la hoja de España.")
        
        if st.button("🚀 Generar Fichero 1 (Nacional)"):
            if df_nac_base is not None:
                # Buscar columnas de referencia y PVPR automáticamente
                idx_ref_nac, idx_pvp_nac = 2, 4
                for r_idx, row in df_nac_base.head(5).iterrows():
                    row_strs = [str(item).strip().upper() for item in row]
                    if "REFERENCIA" in row_strs: idx_ref_nac = row_strs.index("REFERENCIA")
                    for idx, item in enumerate(row_strs):
                        if "PVPR" in item: idx_pvp_nac = idx; break
                
                mapping_nac = extraer_precios_por_posicion(df_nac_base, idx_ref_nac, idx_pvp_nac)
                
                if mapping_nac:
                    df_final = pd.DataFrame(columns=columnas_plantilla)
                    df_final['reference'] = list(mapping_nac.keys())
                    
                    # Asignar los precios a todo el ecosistema nacional
                    precios = list(mapping_nac.values())
                    df_final['price_spain'] = precios
                    df_final['price_tradeinn_es'] = precios
                    df_final['price_aliexpress_es'] = precios
                    df_final['price_mediamarkt_es'] = precios
                    df_final['price_aurgi_es'] = precios
                    df_final['price_carrefour'] = precios
                    df_final['price_pccomponentes'] = precios
                    
                    # Rellenar las demás columnas con vacíos
                    df_final = df_final.fillna("")
                    
                    bytes_excel = exportar_excel_con_cabecera_herramientas(df_final)
                    st.success("✅ Fichero Nacional generado.")
                    st.download_button(label="📥 Descargar Fichero 1", data=bytes_excel, file_name="1_Tarifa_Nacional_Espana.xlsx")
            else:
                st.warning("Por favor, carga primero la Tarifa Nacional en la barra lateral.")

    # ---------------------------------------------------------------------
    # PLANTILLA 2: TARIFA INTERNACIONAL PORTUGAL
    # ---------------------------------------------------------------------
    with col2:
        st.subheader("2. Tarifa Internacional Portugal")
        st.caption("Actualiza exclusivamente el precio correspondiente a la columna price_portugal.")
        
        if st.button("🚀 Generar Fichero 2 (Portugal)"):
            df_pt = pestañas_int.get("PT")
            if df_pt is not None:
                # Extraer usando los índices fijos estipulados: Col C (2) y Col M (12)
                mapping_pt = extraer_precios_por_posicion(df_pt, 2, 12)
                
                if mapping_pt:
                    df_final = pd.DataFrame(columns=columnas_plantilla)
                    df_final['reference'] = list(mapping_pt.keys())
                    df_final['price_portugal'] = list(mapping_pt.values())
                    df_final = df_final.fillna("")
                    
                    bytes_excel = exportar_excel_con_cabecera_herramientas(df_final)
                    st.success("✅ Fichero de Portugal generado.")
                    st.download_button(label="📥 Descargar Fichero 2", data=bytes_excel, file_name="2_Tarifa_Internacional_Portugal.xlsx")
            else:
                st.warning("No se encontró la pestaña 'PT' en el archivo Internacional.")

    # ---------------------------------------------------------------------
    # PLANTILLA 3: RESTO DE TARIFA INTERNACIONAL
    # ---------------------------------------------------------------------
    with col3:
        st.subheader("3. Resto de Internacional")
        st.caption("Unifica horizontalmente los precios de: Francia, Italia, Alemania, Países Bajos y Polonia.")
        
        if st.button("🚀 Generar Fichero 3 (Resto Int.)"):
            if df_nac_base is not None and pestañas_int:
                # Usamos los SKUs de la hoja nacional como lista maestra para consolidar Europa
                idx_ref_nac = 2
                for r_idx, row in df_nac_base.head(5).iterrows():
                    if "REFERENCIA" in [str(i).strip().upper() for i in row]:
                        idx_ref_nac = [str(i).strip().upper() for i in row].index("REFERENCIA")
                        break
                        
                skus_maestros = [formatear_sku(x) for x in df_nac_base.iloc[:, idx_ref_nac].dropna() 
                                 if str(x).strip().upper() not in ["REFERENCIA", "REFERENCE", "REF"]]
                skus_maestros = [s for s in skus_maestros if s != ""]
                
                df_final = pd.DataFrame(columns=columnas_plantilla)
                df_final['reference'] = list(set(skus_maestros)) # Remover duplicados de la lista base
                
                # Definición de mapeos de datos según tus reglas de extracción de columnas
                # Estructura: columna_destino: (lista_posibles_pestañas, col_ref, col_precio)
                mapeo_paises = {
                    'price_france': (["FR-FR", "ES-FR"], 2, 12),  # Col C y M
                    'price_italy': (["IT-IT", "ES-IT"], 2, 12),   # Col C y M
                    'price_germany': (["DE-DE", "ES-DE"], 2, 12), # Col C y M
                    'price_holand': (["NL"], 2, 12),               # Col C y M
                    'price_poland': (["PL"], 2, 13)                # Col C y N (Polonia en Zlotis)
                }
                
                for col_destino, (pestañas, idx_c, idx_p) in mapeo_paises.items():
                    df_sheet = None
                    for p in pestañas:
                        if p in pestañas_int:
                            df_sheet = pestañas_int[p]
                            break
                            
                    if df_sheet is not None:
                        dict_precios = extraer_precios_por_posicion(df_sheet, idx_c, idx_p)
                        df_final[col_destino] = df_final['reference'].map(dict_precios)
                
                df_final = df_final.fillna("")
                bytes_excel = exportar_excel_con_cabecera_herramientas(df_final)
                st.success("✅ Fichero de Resto de Europa generado de forma unificada.")
                st.download_button(label="📥 Descargar Fichero 3", data=bytes_excel, file_name="3_Resto_de_Internacional.xlsx")
            else:
                st.warning("Asegúrate de cargar ambos archivos maestros para consolidar el bloque internacional.")