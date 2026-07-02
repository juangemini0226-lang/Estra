import streamlit as st
import pandas as pd
import requests
import time
from datetime import datetime, timezone, timedelta
from requests.auth import HTTPBasicAuth
import gspread
from google.oauth2.service_account import Credentials
from gspread_dataframe import set_with_dataframe, get_as_dataframe

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Extracción Energía", page_icon="🔌", layout="wide")

st.title("🔌 KERN IoP - Panel de Extracción Unificado")
st.write("Extrae consumos históricos de la API, visualízalos y envíalos a la base de datos maestra.")

# --- CREDENCIALES Y CONSTANTES ---
USERNAME = "API_estra"
PASSWORD = "API_estra*2026"
BASE_URL = "https://apps.kern-iop.tech/navigator/clientes/3"
AUTH = HTTPBasicAuth(USERNAME, PASSWORD)
SHEET_URL = "https://docs.google.com/spreadsheets/d/1lRg2Fc1pk3HBfXkYwXhWnFlTAGxx9gvoZ4hRnJ1AhXY/edit#gid=0"

MAQUINAS_DISPONIBLES = {
    "H64 (36)": 36, "H76 (37)": 37, "H61 (38)": 38, "H82 (39)": 39, "H84 (40)": 40,
    "H74 (41)": 41, "H85 (42)": 42, "S02 (43)": 43, "H73 (44)": 44, "H72 (45)": 45,
    "H71 (46)": 46, "H75 (47)": 47, "H69 (48)": 48, "H80 (49)": 49, "H81 (50)": 50,
    "H83 (51)": 51, "H62 (52)": 52, "H79 (53)": 53, "Compresores (87)": 87
}

VARIABLES_DISPONIBLES = {
    "Potencia Total (27)": 27,
    "Energía Consumida (28)": 28
}

# --- FUNCIONES AUXILIARES ---
@st.cache_resource
def conectar_sheets():
    scope = ["https://www.googleapis.com/auth/spreadsheets"]
    creds_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
    return gspread.authorize(creds)

def api_get(url: str):
    try:
        resp = requests.get(url, auth=AUTH, verify=True, timeout=30)
        if resp.status_code == 200:
            return resp.json()
    except Exception:
        pass
    return None

def fecha_a_timestamp_seguro(dt_obj) -> int:
    tz_colombia = timezone(timedelta(hours=-5))
    dt_aware = dt_obj.replace(tzinfo=tz_colombia)
    return int(dt_aware.timestamp())

# --- INTERFAZ DE USUARIO ---
st.divider()
col1, col2 = st.columns(2)

with col1:
    st.markdown("#### ⏱️ Rango de Tiempo")
    c_f1, c_h1 = st.columns(2)
    d_inicio = c_f1.date_input("Fecha Inicio", value=pd.to_datetime("today") - pd.Timedelta(days=2))
    t_inicio = c_h1.time_input("Hora Inicio", value=pd.to_datetime("00:00").time())
    
    c_f2, c_h2 = st.columns(2)
    d_fin = c_f2.date_input("Fecha Fin", value=pd.to_datetime("today"))
    t_fin = c_h2.time_input("Hora Fin", value=pd.to_datetime("23:59").time())

with col2:
    st.markdown("#### ⚙️ Activos e Indicadores")
    maq_seleccionadas = st.multiselect("Selecciona Máquinas:", options=list(MAQUINAS_DISPONIBLES.keys()), default=["H80 (49)"])
    var_seleccionadas = st.multiselect("Selecciona Variables:", options=list(VARIABLES_DISPONIBLES.keys()), default=["Potencia Total (27)", "Energía Consumida (28)"])

dt_inicio_total = datetime.combine(d_inicio, t_inicio)
dt_fin_total = datetime.combine(d_fin, t_fin)

if st.button("🚀 Extraer Datos de API KERN", type="primary", use_container_width=True):
    if not maq_seleccionadas or not var_seleccionadas:
        st.error("⚠️ Debes seleccionar al menos una máquina y una variable.")
        st.stop()
        
    if dt_inicio_total >= dt_fin_total:
        st.error("⚠️ La fecha de inicio debe ser menor a la fecha de fin.")
        st.stop()

    ids_maquinas = [MAQUINAS_DISPONIBLES[m] for m in maq_seleccionadas]
    ids_variables = [VARIABLES_DISPONIBLES[v] for v in var_seleccionadas]
    
    # --- LÓGICA DE EXTRACCIÓN (CHUNKING) ---
    with st.status("Extrayendo datos desde KERN IoP...", expanded=True) as status:
        frames_encontrados = []
        
        # Calcular pasos para barra de progreso
        total_dias = (dt_fin_total - dt_inicio_total).days + 1
        pasos_estimados = (total_dias // 5 + 1) * len(ids_maquinas) * len(ids_variables)
        paso_actual = 0
        barra_progreso = st.progress(0)
        texto_progreso = st.empty()

        for id_maq in ids_maquinas:
            for id_var in ids_variables:
                ciclo_inicio = dt_inicio_total
                while ciclo_inicio < dt_fin_total:
                    ciclo_fin = min(ciclo_inicio + timedelta(days=5), dt_fin_total)
                    
                    ts_inicio = fecha_a_timestamp_seguro(ciclo_inicio)
                    ts_fin = fecha_a_timestamp_seguro(ciclo_fin)
                    
                    rango_str = f"{ciclo_inicio.strftime('%m-%d')} al {ciclo_fin.strftime('%m-%d')}"
                    texto_progreso.write(f"🔄 Consultando: Maquina {id_maq} | Variable {id_var} | {rango_str}")
                    
                    url_data = f"{BASE_URL}/variables/{id_var}/data/maquina/{id_maq}/from/{ts_inicio}/to/{ts_fin}/"
                    data_raw = api_get(url_data)
                    
                    if data_raw and len(data_raw) > 0:
                        df_tmp = pd.DataFrame(data_raw)
                        df_tmp["id_variable_api"] = id_var
                        df_tmp["id_maquina_api"] = id_maq
                        frames_encontrados.append(df_tmp)
                    
                    ciclo_inicio = ciclo_fin + timedelta(minutes=1)
                    
                    paso_actual += 1
                    progreso = min(paso_actual / pasos_estimados, 1.0)
                    barra_progreso.progress(progreso)
                    time.sleep(0.2) # Respetar límites de la API

        # --- PIVOTEO Y CONSOLIDACIÓN ---
        st.write("🗜️ Consolidando y estructurando base de datos...")
        if frames_encontrados:
            df_consolidado = pd.concat(frames_encontrados, ignore_index=True)
            
            if "valor" in df_consolidado.columns and "Variable" in df_consolidado.columns:
                for var_name in df_consolidado['Variable'].dropna().unique():
                    mask = (df_consolidado['Variable'] == var_name) & (df_consolidado['valor'].notna())
                    if mask.any():
                        df_consolidado.loc[mask, var_name] = df_consolidado.loc[mask, 'valor']
                df_consolidado = df_consolidado.drop(columns=['valor'])

            columnas_llave = ["Fecha y hora", "maquina_o_puesto", "id_maquina_api"]
            df_final = df_consolidado.groupby(columnas_llave, as_index=False).first()
            
            for col in ["Variable", "id_variable_api"]:
                if col in df_final.columns:
                    df_final = df_final.drop(columns=[col])

            if "Fecha y hora" in df_final.columns:
                df_final["Fecha y hora"] = pd.to_datetime(df_final["Fecha y hora"]).dt.strftime('%Y-%m-%d %H:%M:%S')
            
            # Renombramientos automáticos
            renombres = {'Fecha y hora': 'Timestamp', 'maquina_o_puesto': 'ID_Maquina_Texto', 'NRG005': 'Energia_kWh', 'NRG004': 'Potencia_kW'}
            df_final.rename(columns={k: v for k, v in renombres.items() if k in df_final.columns}, inplace=True)
            
            st.session_state['df_energia_extraido'] = df_final
            status.update(label=f"¡Extracción Exitosa! {len(df_final)} registros procesados.", state="complete", expanded=False)
        else:
            status.update(label="No se encontraron datos", state="error")
            st.error("🚩 No se encontraron datos en la API para los parámetros ingresados.")

# --- RESULTADOS Y ACCIONES POST-EXTRACCIÓN ---
if 'df_energia_extraido' in st.session_state:
    df_mostrar = st.session_state['df_energia_extraido']
    st.success("✅ Datos listos para descarga o sincronización.")
    
    with st.expander("👀 Vista Previa de los Datos", expanded=True):
        st.dataframe(df_mostrar.head(100), use_container_width=True)
    
    col_dl, col_up = st.columns(2)
    
    with col_dl:
        csv_data = df_mostrar.to_csv(index=False, encoding='utf-8-sig')
        st.download_button(
            label="⬇️ Descargar como CSV",
            data=csv_data,
            file_name=f"Kern_Energia_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv",
            use_container_width=True
        )
        
    with col_up:
        if st.button("☁️ Sincronizar a Google Sheets (Upsert)", type="primary", use_container_width=True):
            with st.spinner("Subiendo registros sin duplicar..."):
                try:
                    gc = conectar_sheets()
                    sh = gc.open_by_url(SHEET_URL)
                    try:
                        ws_energia = sh.worksheet("Registro_Energía")
                        df_existente = get_as_dataframe(ws_energia).dropna(how='all')
                        
                        claves = ["Timestamp", "id_maquina_api"]
                        for col in claves:
                            df_mostrar[col] = df_mostrar[col].astype(str)
                            if col in df_existente.columns:
                                df_existente[col] = df_existente[col].astype(str)
                            else:
                                df_existente[col] = ""
                                
                        df_merged = df_mostrar.merge(df_existente[claves], on=claves, how='left', indicator=True)
                        df_nuevos = df_merged[df_merged['_merge'] == 'left_only'].drop(columns=['_merge'])
                        
                    except gspread.WorksheetNotFound:
                        ws_energia = sh.add_worksheet("Registro_Energía", 1000, 20)
                        df_nuevos = df_mostrar
                        df_existente = pd.DataFrame()
                        
                    if df_nuevos.empty:
                        st.info("👍 Todo al día. Los datos extraídos ya existen en Google Sheets.")
                    else:
                        df_nuevos = df_nuevos.fillna("")
                        columnas_sheet = df_existente.columns.tolist() if not df_existente.empty else df_nuevos.columns.tolist()
                        for col in columnas_sheet:
                            if col not in df_nuevos.columns:
                                df_nuevos[col] = ""
                        df_nuevos = df_nuevos[columnas_sheet]
                        
                        if df_existente.empty:
                            set_with_dataframe(ws_energia, df_nuevos)
                        else:
                            ws_energia.append_rows(df_nuevos.values.tolist())
                            
                        st.success(f"🎉 ¡Éxito! Se agregaron {len(df_nuevos)} filas nuevas a Sheets.")
                except Exception as e:
                    st.error(f"Error de conexión con Sheets: {e}")
