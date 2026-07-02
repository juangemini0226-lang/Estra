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

# ==========================================
# FUNCIONES DE CONEXIÓN Y DESCARGA
# ==========================================
@st.cache_resource
def conectar_sheets():
    scope = ["https://www.googleapis.com/auth/spreadsheets"]
    creds_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
    return gspread.authorize(creds)

@st.cache_data(ttl=600) # Guarda los datos en memoria por 10 minutos para no saturar la API
def descargar_datos_base():
    gc = conectar_sheets()
    sh = gc.open_by_url(SHEET_URL)
    
    # Descargar Producción SEC
    try:
        ws_prod = sh.worksheet("produccion SEC")
        df_prod = pd.DataFrame(ws_prod.get_all_records())
    except Exception:
        df_prod = pd.DataFrame()
        
    # Descargar Masas (Material_Data u hoja principal 1)
    try:
        ws_masas = sh.sheet1
        df_masas = pd.DataFrame(ws_masas.get_all_records())
    except Exception:
        df_masas = pd.DataFrame()
        
    return df_prod, df_masas

# ==========================================
# FUNCIONES DE CÁLCULO
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
# INTERFAZ Y FLUJO PRINCIPAL
# ==========================================
st.markdown("### 1. Extracción de Datos Maestros")
with st.spinner("Descargando Producción y Masas desde Google Sheets..."):
    df_prod_raw, df_masas_raw = descargar_datos_base()

if df_prod_raw.empty or df_masas_raw.empty:
    st.warning("No se encontraron datos de Producción o Masas en Google Sheets. Asegúrate de ejecutar los extractores primero.")
    st.stop()
else:
    st.success(f"Datos base descargados: {len(df_prod_raw)} OTs de Producción y {len(df_masas_raw)} registros de Materiales.")

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
                # Filtrar órdenes terminadas
                df_prod = df_prod.dropna(subset=['Inicio_Limpio', 'Fin_Limpio'])
                
                # --- PREPARAR MASAS ---
                st.write("⚖️ Consolidando Masas por OT...")
                df_masas = df_masas_raw.copy()
                df_masas['OT'] = df_masas['OT'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip().str.upper().str.lstrip('0')
                df_masas['Total'] = df_masas['Total'].astype(str).str.replace(',', '.', regex=False)
                df_masas['Total'] = pd.to_numeric(df_masas['Total'], errors='coerce')
                df_masas_agg = df_masas.groupby('OT', as_index=False)['Total'].sum()
                df_masas_agg.rename(columns={'OT': 'ID_Job', 'Total': 'Total_Masa_Kg'}, inplace=True)
                
                # --- MERGE PRODUCCIÓN + MASAS ---
                st.write("🔗 Uniendo Producción y Masas...")
                df_consolidado = pd.merge(df_prod, df_masas_agg, on='ID_Job', how='inner')
                df_consolidado['ID_Maquina_Normalizado'] = df_consolidado['ID_Maquina'].apply(normalizar_maquina)
                
                # --- PREPARAR ENERGÍA ---
                st.write("⚡ Analizando CSV de Energía...")
                content_energia = uploaded_energia.read()
                df_energia = pd.read_csv(io.BytesIO(content_energia), encoding='utf-8', encoding_errors='replace')
                df_energia.columns = df_energia.columns.str.strip()
                
                renombres_nrg = {'Fecha y hora': 'Timestamp', 'maquina_o_puesto': 'ID_Maquina_Texto', 'Energía [kWh]': 'Energia_kWh', 'Potencia [kW]': 'Potencia_kW'}
                df_energia.rename(columns={k: v for k, v in renombres_nrg.items() if k in df_energia.columns}, inplace=True)
                
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
                    
                    # Filtrar energía para esta máquina en esta ventana
                    mask = (df_energia['ID_Maquina_Texto'] == maquina_ot) & (df_energia['Timestamp'] >= inicio_span) & (df_energia['Timestamp'] <= fin_span)
                    energia_ot = df_energia[mask]
                    
                    metricas = calcular_metricas_ot(energia_ot, inicio_real, fin_real)
                    
                    # Calcular SEC
                    masa_kg = float(ot['Total_Masa_Kg']) if pd.notna(ot['Total_Masa_Kg']) and ot['Total_Masa_Kg'] > 0 else np.nan
                    sec_value = round(metricas['Energia_Total_kWh'] / masa_kg, 4) if pd.notna(masa_kg) else np.nan
                    
                    fila_resultado = ot.to_dict()
                    fila_resultado.update(metricas)
                    fila_resultado['SEC_kWh_kg'] = sec_value
                    resultados_sec.append(fila_resultado)

                df_resultado_final = pd.DataFrame(resultados_sec)
                
                status.update(label="¡Cálculo SEC completado exitosamente!", state="complete", expanded=False)
                st.session_state['df_sec_calculado'] = df_resultado_final # Guardamos en sesión para los filtros
                st.balloons()
                
            except Exception as e:
                status.update(label="Error en el procesamiento", state="error")
                st.error(f"❌ Ocurrió un error al cruzar los datos: {e}")

# ==========================================
# DASHBOARD INTERACTIVO
# ==========================================
if 'df_sec_calculado' in st.session_state:
    st.divider()
    st.markdown("## 📊 Tablero de Resultados SEC")
    df_board = st.session_state['df_sec_calculado'].copy()
    
    # Filtros
    col1, col2, col3 = st.columns(3)
    with col1:
        maquinas = ['Todas'] + sorted(df_board['ID_Maquina_Normalizado'].dropna().unique().tolist())
        filtro_maq = st.selectbox("🏭 Máquina:", maquinas)
    with col2:
        confiabilidades = ['Todas'] + sorted(df_board['Confiabilidad'].dropna().unique().tolist())
        filtro_conf = st.selectbox("🎯 Confiabilidad del Dato:", confiabilidades)
    
    # Aplicar Filtros
    if filtro_maq != 'Todas':
        df_board = df_board[df_board['ID_Maquina_Normalizado'] == filtro_maq]
    if filtro_conf != 'Todas':
        df_board = df_board[df_board['Confiabilidad'] == filtro_conf]

    # KPIs
    st.markdown("#### Indicadores Globales (Según filtro)")
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    kpi1.metric("Órdenes Procesadas", f"{len(df_board)}")
    kpi2.metric("Masa Total (Kg)", f"{df_board['Total_Masa_Kg'].sum():,.2f}")
    kpi3.metric("Energía Total (kWh)", f"{df_board['Energia_Total_kWh'].sum():,.2f}")
    
    # SEC Promedio (Evitar división por cero)
    masa_total = df_board['Total_Masa_Kg'].sum()
    sec_promedio = df_board['Energia_Total_kWh'].sum() / masa_total if masa_total > 0 else 0
    kpi4.metric("SEC Promedio (kWh/kg)", f"{sec_promedio:.4f}")

    # Tabla
    st.dataframe(
        df_board[['ID_Job', 'ID_Maquina', 'Inicio_Limpio', 'Fin_Limpio', 'Total_Masa_Kg', 'Energia_Total_kWh', 'SEC_kWh_kg', 'Confiabilidad']],
        use_container_width=True
    )
