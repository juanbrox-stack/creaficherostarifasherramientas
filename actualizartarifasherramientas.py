import streamlit as st
import pandas as pd
import io
import zipfile

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
                ref_formateada = f"{num_int:03d}"  # 112 -> 112 (mantiene 3 dígitos)
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
                    # Eliminar símbolos monetarios o espacios si los hubiera
                    val_str = str(val).replace('€', '').replace('$', '').strip()
                    precio_val = float(val_str)
                except ValueError:
                    pass
        
        mapping[ref_formateada] = precio_val
    return mapping


# =========================================================================
# BARRA LATERAL - CARGA DE ARCHIVOS FUENTE
# =========================================================================
st.sidebar.header("📁 Carga de Archivos de Origen")
archivo_nac = st.sidebar.file_uploader("1. Tarifa Nacional (Excel)", type=["xlsx"])
archivo_int = st.sidebar.file_uploader("2. Tarifa Internacional (Excel)", type=["xlsx"])

pestañas_nac = {}
pestañas_int = {}
pestaña_nac_real = "T_AMZ"  # Forzamos la pestaña especificada por el usuario

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
st.title("⚙️ Sistema Inteligente de Procesamiento de Tarifas")
tab1, tab2 = st.tabs(["📦 Bloque 1: Características", "🚀 Bloque 2: Cargador General de Precios"])

# Mapeo fijo según especificaciones del usuario para la Tarifa Nacional:
# Columna A (Índice 0) = Referencia
# Columna P (Índice 15) = Precio (PVP PUB)
idx_ref_nac = 0
idx_pvp_nac = 15

# -------------------------------------------------------------------------
# BLOQUE 1: GENERACIÓN MASIVA POR CARACTERÍSTICAS
# -------------------------------------------------------------------------
with tab1:
    st.header("Generación por Lotes de Archivos de Características")
    st.write("Esta sección extraerá automáticamente cada característica obligatoria y creará un archivo independiente con su nombre.")

    # Reglas de negocio para características de exportación masiva
    reglas_caracteristicas = {
        # España (Tarifa Nacional -> pestaña T_AMZ)
        "PVP_LEROYES": ("Nacional", ["T_AMZ"], idx_ref_nac, idx_pvp_nac), 
        "PVP_PcComponentes": ("Nacional", ["T_AMZ"], idx_ref_nac, idx_pvp_nac),
        "PVPR": ("Nacional", ["T_AMZ"], idx_ref_nac, idx_pvp_nac),
        "PVP ESPANA": ("Nacional", ["T_AMZ"], idx_ref_nac, idx_pvp_nac),
        "Pvp mediamarkt es": ("Nacional", ["T_AMZ"], idx_ref_nac, idx_pvp_nac),
        "PVP_SHEIN_ES": ("Nacional", ["T_AMZ"], idx_ref_nac, idx_pvp_nac),
        # Francia (Col C=2, Col M=12)
        "Prix_France": ("Internacional", ["FR-FR", "ES-FR"], 2, 12),
        "PVP_SHEIN_FR": ("Internacional", ["FR-FR", "ES-FR"], 2, 12),
        # Italia (Col C=2, Col M=12)
        "Prix_Italia": ("Internacional", ["IT-IT", "ES-IT"], 2, 12),
        "PVP_SHEIN_IT": ("Internacional", ["IT-IT", "ES-IT"], 2, 12),
        # Alemania (Col C=2, Col M=12)
        "prix_Alemania": ("Internacional", ["DE-DE", "ES-DE"], 2, 12),
        "PVP_SHEIN_DE": ("Internacional", ["DE-DE", "ES-DE"], 2, 12),
        # Portugal (Col C=2, Col M=12)
        "PVP_SHEIN_PT": ("Internacional", ["PT"], 2, 12),
        "PVP PORTUGAL": ("Internacional", ["PT"], 2, 12),
        # Bélgica (Col C=2, Col M=12)
        "PVP_BE": ("Internacional", ["BE"], 2, 12),
        # Países Bajos (Col C=2, Col M=12)
        "PRIXHOLANDA": ("Internacional", ["NL"], 2, 12),
        "PVP_SHEIN_NL": ("Internacional", ["NL"], 2, 12),
        # Polonia (Col C=2, Col N=13)
        "PVP_SHEIN_PL": ("Internacional", ["PL"], 2, 13),
        "PRIX_POLONIA": ("Internacional", ["PL"], 2, 13),
        # Suecia (Col C=2, Col N=13)
        "PVP Suecia": ("Internacional", ["SE"], 2, 13)
    }

    if st.button("📦 Generar y Empaquetar todas las Características"):
        if not archivo_nac or not archivo_int:
            st.error("Por favor, asegúrate de subir ambos archivos (Nacional e Internacional) en la barra lateral para proceder.")
        else:
            zip_buffer = io.BytesIO()
            conteo_archivos = 0
            
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                
                # Obtener la hoja T_AMZ obligatoriamente, si no existe usa la primera disponible
                df_nac_base = pestañas_nac.get("T_AMZ", next(iter(pestañas_nac.values())) if pestañas_nac else None)

                # Bucle de procesamiento masivo por características
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
                        
                        # Convertir explícitamente a string con punto decimal para evitar pérdidas de formato
                        df_out['valor'] = df_out['valor'].apply(lambda x: "{:.2f}".format(x) if x is not None else "")
                        
                        # --- CORRECCIÓN DE PARÁMETRO DE PANDAS: lineterminator (sin guion bajo) ---
                        csv_body = df_out.to_csv(index=False, sep=',', encoding='utf-8', lineterminator='\r\n')
                        csv_data = f"sep=,\r\n{csv_body}"
                        
                        zip_file.writestr(f"{nombre_carac}.csv", csv_data)
                        conteo_archivos += 1
                        
                if conteo_archivos > 0:
                    st.success(f"¡Éxito! Se han procesado y empaquetado {conteo_archivos} archivos de características.")
                    st.download_button(
                        label="📥 Descargar Pack Completo de Características (.zip)",
                        data=zip_buffer.getvalue(),
                        file_name="caracteristicas_actualizadas.zip",
                        mime="application/zip"
                    )
                else:
                    st.error("No se pudo extraer información compatible. Revisa la estructura de los archivos de origen.")


# -------------------------------------------------------------------------
# BLOQUE 2: CARGADOR GENERAL DE PRECIOS
# -------------------------------------------------------------------------
with tab2:
    st.header("Generación del Fichero de Carga General de Precios")
    st.write("Genera el CSV estructurado requerido por la aplicación principal mapeando simultáneamente todos los canales internacionales y nacionales.")
    
    # Campo para ingresar de forma dinámica la tasa de cambio del Zloty Polaco (PLN)
    tipo_cambio_pln = st.number_input("💵 Tipo de cambio actual (1 EUR a PLN):", min_value=0.01, value=4.32, step=0.01)

    if st.button("🚀 Generar Fichero General de Precios (Cargador)"):
        if not archivo_nac or not archivo_int:
            st.error("Se necesitan los dos archivos cargados para realizar el cruce completo multicanal.")
        else:
            skus_totales = set()
            df_nac_base = pestañas_nac.get("T_AMZ", next(iter(pestañas_nac.values())))
            
            # Extraer el universo total de SKUs basándonos en la Columna A de T_AMZ
            map_base_nac = extraer_precios_por_posicion(df_nac_base, idx_ref_nac, idx_ref_nac)
            skus_totales.update(map_base_nac.keys())

            # 2. Construcción de mapeos específicos por canal
            map_spain = extraer_precios_por_posicion(df_nac_base, idx_ref_nac, idx_pvp_nac)
            map_pcc = map_spain
            map_mm = map_spain
            
            # Canales Internacionales
            def buscar_y_extraer(lista_pestañas, col_ref, col_precio):
                for p in lista_pestañas:
                    if p in pestañas_int:
                        return extraer_precios_por_posicion(pestañas_int[p], col_ref, col_precio)
                return {}

            map_fr = buscar_y_extraer(["FR-FR", "ES-FR"], 2, 12)
            map_it = buscar_y_extraer(["IT-IT", "ES-IT"], 2, 12)
            map_de = buscar_y_extraer(["DE-DE", "ES-DE"], 2, 12)
            map_pt = buscar_y_extraer(["PT"], 2, 12)
            map_nl = buscar_y_extraer(["NL"], 2, 12)
            map_pl_eur = buscar_y_extraer(["PL"], 2, 13)

            # 3. Ensamblar la estructura final requerida de 19 columnas
            filas_cargador = []
            for sku in sorted(skus_totales):
                precio_pln = map_pl_eur.get(sku)
                if precio_pln is not None:
                    precio_pln = round(precio_pln * tipo_cambio_pln, 2)

                fila = {
                    "reference": sku,
                    "price_france": map_fr.get(sku, ""),
                    "price_italy": map_it.get(sku, ""),
                    "price_germany": map_de.get(sku, ""),
                    "price_portugal": map_pt.get(sku, ""),
                    "price_spain": map_spain.get(sku, ""),
                    "price_poland": precio_pln if precio_pln else "",
                    "price_holand": map_nl.get(sku, ""),
                    "price_tradeinn_es": map_spain.get(sku, ""),
                    "price_aliexpress_es": map_spain.get(sku, ""),
                    "price_makro_es": map_spain.get(sku, ""),
                    "price_mediamarkt_es": map_mm.get(sku, ""),
                    "price_aurgi_es": map_spain.get(sku, ""),
                    "price_elcorteingles_es": map_spain.get(sku, ""),
                    "price_makro_de": map_de.get(sku, ""),
                    "price_makro_it": map_it.get(sku, ""),
                    "price_carrefour": map_spain.get(sku, ""),
                    "price_pccomponentes": map_pcc.get(sku, ""),
                    "Acción": ""
                }
                filas_cargador.append(fila)

            df_cargador = pd.DataFrame(filas_cargador)
            
            columnas_precios = [col for col in df_cargador.columns if col != "reference" and col != "Acción"]
            for col in columnas_precios:
                df_cargador[col] = df_cargador[col].apply(lambda x: f"{x:.2f}" if isinstance(x, (int, float)) else x)

            # Mostrar vista previa en la interfaz
            st.write("### Vista Previa del Fichero de Carga Generado")
            st.dataframe(df_cargador.head(10))

            # --- CORRECCIÓN DE PARÁMETRO DE PANDAS: lineterminator (sin guion bajo) ---
            cargador_body = df_cargador.to_csv(index=False, sep=',', encoding='utf-8', lineterminator='\r\n')
            cargador_csv_data = f"sep=,\r\n{cargador_body}"

            st.download_button(
                label="📥 Descargar Fichero General de Precios (.csv)",
                data=cargador_csv_data,
                file_name="CargadorGeneralPrecios_Procesado.csv",
                mime="text/csv"
            )
