import streamlit as st
import pandas as pd
import numpy as np
import io
import gspread
from google.oauth2.service_account import Credentials

# 1. Configuración de la página
st.set_page_config(page_title="Analítica SEC", page_icon="⚡", layout="wide")

st.title("⚡ Analítica de Consumo Específico de Energía (SEC)")
st.write("Cruza los datos de Producción, Masas y Energía para obtener el SEC por Orden de Trabajo.")

SHEET_URL = 'https://docs.google.com/spreadsheets/d/1lRg2Fc1pk3HBfXkYwXhWnFlTAGxx9gvoZ4hRnJ1AhXY/edit#gid=0'

# Columnas mínimas que DEBE tener cada hoja para ser considerada válida.
# Esto es lo que reemplaza la suposición de "la primera hoja es Masas".
# IMPORTANTE: en el archivo real, la hoja de masas es 'Material_Data' y su
# llave YA se llama 'ID_Job' (no 'OT' como asumía la versión anterior).
COLUMNAS_REQUERIDAS_PROD = {'Máquina', 'Trabajo / Orden', 'Tiempo Empezar', 'Tiempo Final'}
COLUMNAS_REQUERIDAS_MASAS = {'ID_Job', 'Total'}

# Nombres "preferidos" cuando hay más de una hoja candidata (p.ej.
# 'producción SEC' y 'producción detallada' comparten las mismas columnas
# base, así que sin esta preferencia quedarían ambiguas cada vez).
HOJA_PROD_PREFERIDA = "produccion SEC"
HOJA_MASAS_PREFERIDA = "Material_Data"

# ==========================================
# FUNCIONES DE CONEXIÓN Y DESCARGA
# ==========================================
@st.cache_resource
def conectar_sheets():
    scope = ["https://www.googleapis.com/auth/spreadsheets"]
    creds_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
    return gspread.authorize(creds)


@st.cache_data(ttl=600)
def listar_hojas_y_columnas():
    """
    Lista todas las pestañas del Sheet junto con sus encabezados.
    Esto es la base para poder detectar automáticamente cuál hoja
    es 'Producción' y cuál es 'Masas', en vez de asumir posiciones fijas.
    """
    gc = conectar_sheets()
    sh = gc.open_by_url(SHEET_URL)
    info = {}
    for ws in sh.worksheets():
        try:
            encabezados = ws.row_values(1)
        except Exception:
            encabezados = []
        info[ws.title] = encabezados
    return info


def detectar_hoja(info_hojas, columnas_requeridas):
    """
    Devuelve el nombre de la primera hoja cuyos encabezados contienen
    TODAS las columnas requeridas. None si ninguna califica.
    Si hay más de una hoja candidata, devuelve la lista completa de
    candidatas para que el usuario pueda desambiguar.
    """
    candidatas = []
    for nombre, columnas in info_hojas.items():
        columnas_set = set(c.strip() for c in columnas)
        if columnas_requeridas.issubset(columnas_set):
            candidatas.append(nombre)
    return candidatas


@st.cache_data(ttl=600)
def descargar_hoja(nombre_hoja):
    gc = conectar_sheets()
    sh = gc.open_by_url(SHEET_URL)
    ws = sh.worksheet(nombre_hoja)  # <-- SIEMPRE por nombre, nunca por posición
    return pd.DataFrame(ws.get_all_records())


# ==========================================
# FUNCIONES DE CÁLCULO (sin cambios de lógica)
# ==========================================
def normalizar_maquina(id_maquina):
    if pd.isna(id_maquina): return id_maquina
    texto = str(id_maquina).strip().upper()
    if texto.endswith('MED'): return texto[:-3]
    return texto


def calcular_metricas_ot(energia_ot, inicio_real, fin_real):
    duracion_minutos = (fin_real - inicio_real).total_seconds() / 60.0
    duracion_minutos_efectiva = max(duracion_minutos, 1.0)
    n_lecturas = len(energia_ot)

    if n_lecturas == 0:
        return {'Energia_Total_kWh': 0.0, 'Potencia_Promedio_kW': np.nan, 'Potencia_Max_kW': np.nan, 'Score_Confiabilidad': 0.0, 'Confiabilidad': 'Sin datos'}

    energia_total = energia_ot['Energia_kWh'].sum()
    potencia_promedio = energia_ot['Potencia_kW'].mean()
    potencia_max = energia_ot['Potencia_kW'].max()

    minutos_cubiertos = energia_ot['Timestamp'].nunique()
    cobertura_pct = min(minutos_cubiertos / duracion_minutos_efectiva, 1.0) * 100

    ts_ordenados = energia_ot['Timestamp'].sort_values().drop_duplicates()
    max_gap = (ts_ordenados.diff().dropna().dt.total_seconds() / 60.0).max() if len(ts_ordenados) > 1 else max(duracion_minutos_efectiva - 1, 0)

    if max_gap <= 5:
        continuidad_pct = 100.0
    else:
        exceso = max_gap - 5
        continuidad_pct = max(0.0, 100.0 * (1 - exceso / duracion_minutos_efectiva))

    score = (cobertura_pct / 100.0 * 0.5) + (continuidad_pct / 100.0 * 0.5)
    label = 'Alta' if score >= 0.85 else ('Media' if score >= 0.60 else 'Baja')

    return {
        'Energia_Total_kWh': round(float(energia_total), 3),
        'Potencia_Promedio_kW': round(float(potencia_promedio), 2) if pd.notna(potencia_promedio) else np.nan,
        'Potencia_Max_kW': round(float(potencia_max), 2) if pd.notna(potencia_max) else np.nan,
        'Score_Confiabilidad': round(score, 3),
        'Confiabilidad': label
    }


# ==========================================
# 1. RESOLUCIÓN DE HOJAS (aquí está el fix real)
# ==========================================
st.markdown("### 1. Extracción de Datos Maestros")

with st.spinner("Inspeccionando pestañas del Google Sheet..."):
    info_hojas = listar_hojas_y_columnas()

if not info_hojas:
    st.error("❌ No se pudo leer ninguna pestaña del Google Sheet. Revisa el acceso de la cuenta de servicio.")
    st.stop()

candidatas_prod = detectar_hoja(info_hojas, COLUMNAS_REQUERIDAS_PROD)
candidatas_masas = detectar_hoja(info_hojas, COLUMNAS_REQUERIDAS_MASAS)

with st.expander("🔎 Ver hojas detectadas en el archivo (diagnóstico)"):
    for nombre, columnas in info_hojas.items():
        st.write(f"**{nombre}** → columnas: {columnas}")

col_a, col_b = st.columns(2)

with col_a:
    if len(candidatas_prod) == 0:
        st.error(f"⚠️ Ninguna hoja tiene las columnas de Producción esperadas: {sorted(COLUMNAS_REQUERIDAS_PROD)}")
        hoja_prod_elegida = st.selectbox("Elige manualmente la hoja de Producción:", list(info_hojas.keys()))
    elif len(candidatas_prod) == 1:
        hoja_prod_elegida = candidatas_prod[0]
        st.success(f"✅ Hoja de Producción detectada: **{hoja_prod_elegida}**")
    else:
        # Varias hojas comparten las mismas columnas (ej. 'producción SEC' y
        # 'producción detallada'). Si el nombre preferido está entre las
        # candidatas, lo usamos como default en vez de forzar a elegir.
        st.info(f"ℹ️ Varias hojas califican como Producción: {candidatas_prod}. "
                f"Se preseleccionó '{HOJA_PROD_PREFERIDA}' por ser la hoja esperada; cambia si no aplica.")
        idx_default = candidatas_prod.index(HOJA_PROD_PREFERIDA) if HOJA_PROD_PREFERIDA in candidatas_prod else 0
        hoja_prod_elegida = st.selectbox("Confirma la hoja de Producción:", candidatas_prod, index=idx_default)

with col_b:
    if len(candidatas_masas) == 0:
        st.error(f"⚠️ Ninguna hoja tiene las columnas de Masas esperadas: {sorted(COLUMNAS_REQUERIDAS_MASAS)} (este era el bug: antes se asumía 'sh.sheet1')")
        hoja_masas_elegida = st.selectbox("Elige manualmente la hoja de Masas:", list(info_hojas.keys()))
    elif len(candidatas_masas) == 1:
        hoja_masas_elegida = candidatas_masas[0]
        st.success(f"✅ Hoja de Masas detectada: **{hoja_masas_elegida}**")
    else:
        st.info(f"ℹ️ Varias hojas califican como Masas: {candidatas_masas}. "
                f"Se preseleccionó '{HOJA_MASAS_PREFERIDA}' por ser la hoja esperada; cambia si no aplica.")
        idx_default = candidatas_masas.index(HOJA_MASAS_PREFERIDA) if HOJA_MASAS_PREFERIDA in candidatas_masas else 0
        hoja_masas_elegida = st.selectbox("Confirma la hoja de Masas:", candidatas_masas, index=idx_default)

with st.spinner(f"Descargando '{hoja_prod_elegida}' y '{hoja_masas_elegida}'..."):
    df_prod_raw = descargar_hoja(hoja_prod_elegida)
    df_masas_raw = descargar_hoja(hoja_masas_elegida)

if df_prod_raw.empty or df_masas_raw.empty:
    st.warning("No se encontraron datos de Producción o Masas en las hojas seleccionadas.")
    st.stop()

# Validación dura antes de seguir: si falta 'OT' aquí, es mejor avisar
# ahora que tronar más adelante con un KeyError críptico.
faltantes_masas = COLUMNAS_REQUERIDAS_MASAS - set(df_masas_raw.columns)
if faltantes_masas:
    st.error(f"❌ La hoja '{hoja_masas_elegida}' no tiene las columnas requeridas: {faltantes_masas}. "
             f"Columnas encontradas: {list(df_masas_raw.columns)}")
    st.stop()

faltantes_prod = COLUMNAS_REQUERIDAS_PROD - set(df_prod_raw.columns)
if faltantes_prod:
    st.error(f"❌ La hoja '{hoja_prod_elegida}' no tiene las columnas requeridas: {faltantes_prod}. "
             f"Columnas encontradas: {list(df_prod_raw.columns)}")
    st.stop()

st.success(f"Datos base descargados: {len(df_prod_raw)} OTs de Producción y {len(df_masas_raw)} registros de Materiales.")

# ==========================================
# 2. CARGA DE ENERGÍA
# ==========================================
st.markdown("### 2. Carga de Consumo Energético")
uploaded_energia = st.file_uploader("Sube tu reporte CSV de energía (Ej. Consumo_1Ene_A_1Jun_2026.csv):", type=["csv"])

if uploaded_energia is not None:
    if st.button("🚀 Iniciar Cruce y Cálculo SEC", type="primary"):
        with st.status("Procesando Motor SEC...", expanded=True) as status:
            try:
                # --- PREPARAR PRODUCCIÓN ---
                st.write("⚙️ Estandarizando Producción...")
                df_prod = df_prod_raw.copy()
                df_prod.rename(columns={'Máquina': 'ID_Maquina', 'Trabajo / Orden': 'ID_Job', 'Tiempo Empezar': 'Inicio', 'Tiempo Final': 'Fin'}, inplace=True)
                df_prod['ID_Job'] = df_prod['ID_Job'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                df_prod['Inicio_Limpio'] = pd.to_datetime(df_prod['Inicio'], errors='coerce', format='mixed', dayfirst=True)
                df_prod['Fin_Limpio'] = pd.to_datetime(df_prod['Fin'], errors='coerce', format='mixed', dayfirst=True)
                df_prod = df_prod.dropna(subset=['Inicio_Limpio', 'Fin_Limpio'])

                # --- PREPARAR MASAS ---
                # En 'Material_Data' la llave YA se llama 'ID_Job' (no 'OT').
                # Se limpia igual que el ID_Job del lado de Producción para
                # garantizar que el merge encuentre coincidencias.
                st.write("⚖️ Consolidando Masas por ID_Job...")
                df_masas = df_masas_raw.copy()
                df_masas['ID_Job'] = df_masas['ID_Job'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                df_masas['Total'] = df_masas['Total'].astype(str).str.replace(',', '.', regex=False)
                df_masas['Total'] = pd.to_numeric(df_masas['Total'], errors='coerce')
                df_masas_agg = df_masas.groupby('ID_Job', as_index=False)['Total'].sum()
                df_masas_agg.rename(columns={'Total': 'Total_Masa_Kg'}, inplace=True)

                # --- MERGE PRODUCCIÓN + MASAS ---
                st.write("🔗 Uniendo Producción y Masas...")
                df_consolidado = pd.merge(df_prod, df_masas_agg, on='ID_Job', how='inner')
                if df_consolidado.empty:
                    st.warning("⚠️ El cruce Producción↔Masas dio 0 filas. Revisa que los ID_Job/OT coincidan en formato "
                                "(la muestra de ambos lados se puede ver abajo).")
                    st.write("Ejemplos ID_Job Producción:", df_prod['ID_Job'].unique()[:5].tolist())
                    st.write("Ejemplos ID_Job Masas:", df_masas_agg['ID_Job'].unique()[:5].tolist())
                df_consolidado['ID_Maquina_Normalizado'] = df_consolidado['ID_Maquina'].apply(normalizar_maquina)

                # --- PREPARAR ENERGÍA ---
                st.write("⚡ Analizando CSV de Energía...")
                content_energia = uploaded_energia.read()
                df_energia = pd.read_csv(io.BytesIO(content_energia), encoding='utf-8', encoding_errors='replace')
                df_energia.columns = df_energia.columns.str.strip()

                renombres_nrg = {'Fecha y hora': 'Timestamp', 'maquina_o_puesto': 'ID_Maquina_Texto', 'Energía [kWh]': 'Energia_kWh', 'Potencia [kW]': 'Potencia_kW'}
                df_energia.rename(columns={k: v for k, v in renombres_nrg.items() if k in df_energia.columns}, inplace=True)

                columnas_energia_requeridas = {'Timestamp', 'ID_Maquina_Texto', 'Energia_kWh', 'Potencia_kW'}
                faltantes_nrg = columnas_energia_requeridas - set(df_energia.columns)
                if faltantes_nrg:
                    raise ValueError(f"El CSV de energía no tiene las columnas esperadas: {faltantes_nrg}. "
                                      f"Columnas encontradas: {list(df_energia.columns)}")

                df_energia['Timestamp'] = pd.to_datetime(df_energia['Timestamp'].astype(str), format='mixed', errors='coerce').dt.floor('min')
                df_energia.dropna(subset=['Timestamp', 'Energia_kWh'], inplace=True)
                df_energia['Energia_kWh'] = pd.to_numeric(df_energia['Energia_kWh'], errors='coerce')
                df_energia['Potencia_kW'] = pd.to_numeric(df_energia['Potencia_kW'], errors='coerce')

                # --- CRUCE FINAL (MOTOR SEC) ---
                st.write("🧠 Ejecutando emparejamiento con tolerancia temporal (15 min)...")
                resultados_sec = []

                for _, ot in df_consolidado.iterrows():
                    maquina_ot = ot['ID_Maquina_Normalizado']
                    inicio_real = ot['Inicio_Limpio']
                    fin_real = ot['Fin_Limpio']

                    if pd.isna(inicio_real) or pd.isna(fin_real):
                        continue

                    inicio_span = inicio_real - pd.Timedelta(minutes=15)
                    fin_span = fin_real + pd.Timedelta(minutes=15)

                    mask = (df_energia['ID_Maquina_Texto'] == maquina_ot) & (df_energia['Timestamp'] >= inicio_span) & (df_energia['Timestamp'] <= fin_span)
                    energia_ot = df_energia[mask]

                    metricas = calcular_metricas_ot(energia_ot, inicio_real, fin_real)

                    masa_kg = float(ot['Total_Masa_Kg']) if pd.notna(ot['Total_Masa_Kg']) and ot['Total_Masa_Kg'] > 0 else np.nan
                    sec_value = round(metricas['Energia_Total_kWh'] / masa_kg, 4) if pd.notna(masa_kg) else np.nan

                    fila_resultado = ot.to_dict()
                    fila_resultado.update(metricas)
                    fila_resultado['SEC_kWh_kg'] = sec_value
                    resultados_sec.append(fila_resultado)

                df_resultado_final = pd.DataFrame(resultados_sec)

                status.update(label="¡Cálculo SEC completado exitosamente!", state="complete", expanded=False)
                st.session_state['df_sec_calculado'] = df_resultado_final
                st.balloons()

            except Exception as e:
                status.update(label="Error en el procesamiento", state="error")
                st.error(f"❌ Ocurrió un error al cruzar los datos: {e}")

# ==========================================
# DASHBOARD INTERACTIVO (sin cambios)
# ==========================================
if 'df_sec_calculado' in st.session_state:
    st.divider()
    st.markdown("## 📊 Tablero de Resultados SEC")
    df_board = st.session_state['df_sec_calculado'].copy()

    col1, col2, col3 = st.columns(3)
    with col1:
        maquinas = ['Todas'] + sorted(df_board['ID_Maquina_Normalizado'].dropna().unique().tolist())
        filtro_maq = st.selectbox("🏭 Máquina:", maquinas)
    with col2:
        confiabilidades = ['Todas'] + sorted(df_board['Confiabilidad'].dropna().unique().tolist())
        filtro_conf = st.selectbox("🎯 Confiabilidad del Dato:", confiabilidades)

    if filtro_maq != 'Todas':
        df_board = df_board[df_board['ID_Maquina_Normalizado'] == filtro_maq]
    if filtro_conf != 'Todas':
        df_board = df_board[df_board['Confiabilidad'] == filtro_conf]

    st.markdown("#### Indicadores Globales (Según filtro)")
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    kpi1.metric("Órdenes Procesadas", f"{len(df_board)}")
    kpi2.metric("Masa Total (Kg)", f"{df_board['Total_Masa_Kg'].sum():,.2f}")
    kpi3.metric("Energía Total (kWh)", f"{df_board['Energia_Total_kWh'].sum():,.2f}")

    masa_total = df_board['Total_Masa_Kg'].sum()
    sec_promedio = df_board['Energia_Total_kWh'].sum() / masa_total if masa_total > 0 else 0
    kpi4.metric("SEC Promedio (kWh/kg)", f"{sec_promedio:.4f}")

    st.dataframe(
        df_board[['ID_Job', 'ID_Maquina', 'Inicio_Limpio', 'Fin_Limpio', 'Total_Masa_Kg', 'Energia_Total_kWh', 'SEC_kWh_kg', 'Confiabilidad']],
        use_container_width=True
    )
