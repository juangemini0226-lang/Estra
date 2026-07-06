import streamlit as st
import pandas as pd
import requests
import time
from datetime import datetime, timezone, timedelta
from requests.auth import HTTPBasicAuth
import gspread
from google.oauth2.service_account import Credentials
from gspread_dataframe import set_with_dataframe, get_as_dataframe
import plotly.express as px  # <-- LIBRERÍA PARA MEJORES GRÁFICOS

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Extracción Energía", page_icon="🔌", layout="wide")

st.title("🔌 KERN IoP - Panel de Extracción Unificado (InfluxDB)")
st.write("Extrae consumos históricos desde Grafana/InfluxDB, analízalos y envíalos a la base de datos maestra.")

# --- CREDENCIALES Y CONSTANTES ---
USERNAME = "ahenao_estra"
PASSWORD = "X490fDvd"
AUTH = HTTPBasicAuth(USERNAME, PASSWORD)
HEADERS = {
    "Accept": "application/json",
    "Content-Type": "application/json"
}
URL_GRAFANA = "https://kern-iop.tech/api/ds/query"
SHEET_URL = "https://docs.google.com/spreadsheets/d/1lRg2Fc1pk3HBfXkYwXhWnFlTAGxx9gvoZ4hRnJ1AhXY/edit#gid=0"
TZ_COLOMBIA = timezone(timedelta(hours=-5))

# 🛠️ DICCIONARIO DE MAPEO INFLUXDB
MAPEO_INFLUX = {
    43: {"nombre": "S02", "tag_energia": "energy22", "tag_potencia": "energy22"},
    44: {"nombre": "H73", "tag_energia": "energy20", "tag_potencia": "energy20"},
    45: {"nombre": "H72", "tag_energia": "energy23", "tag_potencia": "energy23"},
    46: {"nombre": "H71", "tag_energia": "energy21", "tag_potencia": "energy21"},
    47: {"nombre": "H75", "tag_energia": "energy18", "tag_potencia": "energy18"},
    48: {"nombre": "H69", "tag_energia": "energy4", "tag_potencia": "energy4"},
    49: {"nombre": "H80", "tag_energia": "energy5", "tag_potencia": "energy5"},
    50: {"nombre": "H81", "tag_energia": "energy3", "tag_potencia": "energy3"},
    51: {"nombre": "H83", "tag_energia": "energy8", "tag_potencia": "energy8"},
    53: {"nombre": "H79", "tag_energia": "energy1", "tag_potencia": "energy1"},
    42: {"nombre": "H85", "tag_energia": "energy24", "tag_potencia": "energy24"},
    39: {"nombre": "H82", "tag_energia": "energy7", "tag_potencia": "energy7"},
    38: {"nombre": "H84", "tag_energia": "energy16", "tag_potencia": "energy16"},
    37: {"nombre": "H86", "tag_energia": "energy2", "tag_potencia": "energy2"},
    36: {"nombre": "H64", "tag_energia": "energy13", "tag_potencia": "energy13"},
    60: {"nombre": "H76", "tag_energia": "energy14", "tag_potencia": "energy14"}, 
    41: {"nombre": "H74", "tag_energia": "energy9", "tag_potencia": "energy9"}
}

# Generamos las opciones del select dinámicamente desde el diccionario
MAQUINAS_DISPONIBLES = {f"{v['nombre']} ({k})": k for k, v in MAPEO_INFLUX.items()}

# --- FUNCIONES DE CONEXIÓN Y GOOGLE SHEETS ---
@st.cache_resource
def conectar_sheets():
    scope = ["https://www.googleapis.com/auth/spreadsheets"]
    creds_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
    return gspread.authorize(creds)

# --- FUNCIONES DE EXTRACCIÓN GRAFANA / INFLUXDB ---
def extraer_dataframe_json(json_resp, ref_id, nombre_columna):
    try:
        valores = json_resp['results'][ref_id]['frames'][0]['data']['values']
        return pd.DataFrame({'ts': valores[0], nombre_columna: valores[1]})
    except (KeyError, IndexError, TypeError):
        return pd.DataFrame(columns=['ts', nombre_columna])

def api_post_grafana_energia(dt_inicio_aware, dt_fin_aware, tag_energia):
    dt_inicio_utc = dt_inicio_aware.astimezone(timezone.utc)
    dt_fin_utc = dt_fin_aware.astimezone(timezone.utc)

    iso_inicio = dt_inicio_utc.strftime('%Y-%m-%dT%H:%M:%S.000Z')
    iso_fin = dt_fin_utc.strftime('%Y-%m-%dT%H:%M:%S.000Z')
    ts_inicio_ms = str(int(dt_inicio_aware.timestamp() * 1000))
    ts_fin_ms = str(int(dt_fin_aware.timestamp() * 1000))

    payload = {
        "queries": [{
            "refId": "A", "datasourceId": 5, "rawQuery": True, "resultFormat": "time_series",
            "query": f"SELECT last(\"value\")/1000 FROM \"NRG005\" WHERE (\"production_plant\" = 'medellin' AND \"variable_type\" = '{tag_energia}') AND $timeFilter GROUP BY time(1m) fill(null)"
        }],
        "range": {"from": iso_inicio, "to": iso_fin, "raw": {"from": iso_inicio, "to": iso_fin}},
        "from": ts_inicio_ms, "to": ts_fin_ms
    }
    try:
        resp = requests.post(URL_GRAFANA, json=payload, headers=HEADERS, auth=AUTH, verify=True, timeout=30)
        if resp.status_code == 200:
            return resp.json()
    except Exception as e:
        st.toast(f"Error HTTP en energía: {e}")
    return None

def parsear_energia(json_resp):
    df_A = extraer_dataframe_json(json_resp, 'A', 'Acumulador') if json_resp else pd.DataFrame(columns=['ts', 'Acumulador'])
    if df_A.empty:
        return pd.DataFrame(columns=['Fecha y hora', 'Energía [kWh]'])

    df_A['Energía [kWh]'] = df_A['Acumulador'].ffill().diff().round(4).fillna(0)
    df_A['Fecha y hora'] = pd.to_datetime(df_A['ts'], unit='ms').dt.tz_localize('UTC').dt.tz_convert('America/Bogota').dt.tz_localize(None).dt.floor('min')
    return df_A[['Fecha y hora', 'Energía [kWh]']]

def api_post_grafana_potencia(dt_inicio_aware, dt_fin_aware, tag_potencia):
    iso_inicio = dt_inicio_aware.astimezone(timezone.utc).strftime('%Y-%m-%dT%H:%M:%S.000Z')
    iso_fin = dt_fin_aware.astimezone(timezone.utc).strftime('%Y-%m-%dT%H:%M:%S.000Z')
    ts_inicio_ms = str(int(dt_inicio_aware.timestamp() * 1000))
    ts_fin_ms = str(int(dt_fin_aware.timestamp() * 1000))

    payload = {
        "queries": [{
            "refId": "A", "datasourceId": 5, "rawQuery": True, "resultFormat": "table",
            "query": f"SELECT last(\"value\")/1000 FROM \"NRG004\" WHERE (\"production_plant\" = 'medellin' AND \"variable_type\" = '{tag_potencia}') AND $timeFilter GROUP BY time(1s) fill(none)"
        }],
        "range": {"from": iso_inicio, "to": iso_fin, "raw": {"from": iso_inicio, "to": iso_fin}},
        "from": ts_inicio_ms, "to": ts_fin_ms
    }
    try:
        resp = requests.post(URL_GRAFANA, json=payload, headers=HEADERS, auth=AUTH, verify=True, timeout=30)
        if resp.status_code == 200:
            return resp.json()
    except Exception as e:
        st.toast(f"Error HTTP en potencia: {e}")
    return None

def parsear_potencia(json_resp):
    if not json_resp or 'A' not in json_resp.get('results', {}):
        return pd.DataFrame(columns=['Fecha y hora', 'Potencia [kW]'])
    try:
        tiempos = json_resp['results']['A']['frames'][0]['data']['values'][0]
        valores = json_resp['results']['A']['frames'][0]['data']['values'][1]
    except (KeyError, IndexError, TypeError):
        return pd.DataFrame(columns=['Fecha y hora', 'Potencia [kW]'])

    df = pd.DataFrame({'timestamp_ms': tiempos, 'Potencia [kW]': valores})
    if df.empty:
        return pd.DataFrame(columns=['Fecha y hora', 'Potencia [kW]'])

    df['Potencia [kW]'] = df['Potencia [kW]'].round(3)
    df['Fecha y hora'] = pd.to_datetime(df['timestamp_ms'], unit='ms').dt.tz_localize('UTC').dt.tz_convert('America/Bogota').dt.tz_localize(None).dt.floor('min')
    return df.groupby('Fecha y hora', as_index=False)['Potencia [kW]'].max()

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
    st.info("💡 La extracción trae **Energía y Potencia** automáticamente para cada máquina. Los valores negativos se corregirán a 0.")

dt_inicio_total = datetime.combine(d_inicio, t_inicio).replace(tzinfo=TZ_COLOMBIA)
dt_fin_total = datetime.combine(d_fin, t_fin).replace(tzinfo=TZ_COLOMBIA)

if st.button("🚀 Extraer Datos de InfluxDB", type="primary", use_container_width=True):
    if not maq_seleccionadas:
        st.error("⚠️ Debes seleccionar al menos una máquina.")
        st.stop()

    if dt_inicio_total >= dt_fin_total:
        st.error("⚠️ La fecha de inicio debe ser menor a la fecha de fin.")
        st.stop()

    ids_maquinas = [MAQUINAS_DISPONIBLES[m] for m in maq_seleccionadas]

    with st.status("Extrayendo datos desde InfluxDB...", expanded=True) as status:
        frames_encontrados = []
        
        total_dias = (dt_fin_total - dt_inicio_total).days + 1
        pasos_estimados = (total_dias // 5 + 1) * len(ids_maquinas)
        paso_actual = 0
        
        barra_progreso = st.progress(0)
        texto_progreso = st.empty()

        for id_maq in ids_maquinas:
            datos_maq = MAPEO_INFLUX[id_maq]
            nombre_maq = datos_maq["nombre"]
            tag_energia = datos_maq["tag_energia"]
            tag_potencia = datos_maq["tag_potencia"]

            ciclo_inicio = dt_inicio_total
            while ciclo_inicio < dt_fin_total:
                ciclo_fin = min(ciclo_inicio + timedelta(days=5), dt_fin_total)
                rango_str = f"{ciclo_inicio.strftime('%m-%d')} al {ciclo_fin.strftime('%m-%d')}"
                texto_progreso.write(f"🔄 Consultando: {nombre_maq} | {rango_str}")

                # 1. Extraemos Energía y Potencia
                json_energia = api_post_grafana_energia(ciclo_inicio, ciclo_fin, tag_energia)
                df_energia = parsear_energia(json_energia)

                json_potencia = api_post_grafana_potencia(ciclo_inicio, ciclo_fin, tag_potencia)
                df_potencia = parsear_potencia(json_potencia)

                # 2. Unimos ambas tablas por fecha y hora
                if not df_energia.empty or not df_potencia.empty:
                    df_tmp = df_energia.merge(df_potencia, on='Fecha y hora', how='outer')
                    
                    df_tmp['Energía [kWh]'] = df_tmp['Energía [kWh]'].fillna(0)
                    df_tmp['Potencia [kW]'] = df_tmp['Potencia [kW]'].fillna(0)
                    df_tmp['maquina_o_puesto'] = nombre_maq
                    df_tmp['id_maquina_api'] = id_maq
                    
                    frames_encontrados.append(df_tmp)

                ciclo_inicio = ciclo_fin + timedelta(minutes=1)
                paso_actual += 1
                barra_progreso.progress(min(paso_actual / pasos_estimados, 1.0))
                time.sleep(0.3)

        st.write("🗜️ Consolidando base de datos...")
        if frames_encontrados:
            df_final = pd.concat(frames_encontrados, ignore_index=True)
            df_final = df_final.sort_values(['Fecha y hora', 'maquina_o_puesto'])
            df_final = df_final.drop_duplicates(subset=["Fecha y hora", "maquina_o_puesto"], keep="first")

            if "Fecha y hora" in df_final.columns:
                df_final["Fecha y hora"] = pd.to_datetime(df_final["Fecha y hora"]).dt.strftime('%Y-%m-%d %H:%M:%S')

            # 3. Renombramos columnas para mantener compatibilidad
            renombres = {
                'Fecha y hora': 'Timestamp', 
                'maquina_o_puesto': 'ID_Maquina_Texto',
                'Energía [kWh]': 'Energia_kWh',
                'Potencia [kW]': 'Potencia_kW'
            }
            df_final.rename(columns={k: v for k, v in renombres.items() if k in df_final.columns}, inplace=True)

            # 4. APLICAMOS EL CLIP (convertir negativos en 0) a las columnas correctas
            if "Energia_kWh" in df_final.columns:
                df_final["Energia_kWh"] = pd.to_numeric(df_final["Energia_kWh"], errors="coerce").clip(lower=0)
            if "Potencia_kW" in df_final.columns:
                df_final["Potencia_kW"] = pd.to_numeric(df_final["Potencia_kW"], errors="coerce").clip(lower=0)

            # Reordenar las columnas para mejor estética
            cols_orden = ["Timestamp", "ID_Maquina_Texto", "id_maquina_api", "Energia_kWh", "Potencia_kW"]
            df_final = df_final[[c for c in cols_orden if c in df_final.columns]]

            st.session_state['df_energia_extraido'] = df_final
            status.update(label=f"¡Extracción Exitosa! {len(df_final)} registros procesados.", state="complete", expanded=False)
        else:
            status.update(label="No se encontraron datos", state="error")
            st.error("🚩 No se encontraron datos para los parámetros ingresados.")

# --- RESULTADOS Y ACCIONES POST-EXTRACCIÓN ---
if 'df_energia_extraido' in st.session_state:
    df_mostrar = st.session_state['df_energia_extraido']
    st.success("✅ Datos listos para descarga, análisis o sincronización.")

    with st.expander("👀 Vista Previa de los Datos", expanded=True):
        st.dataframe(df_mostrar.head(100), use_container_width=True)

    # --- ANÁLISIS Y VISUALIZACIÓN ---
    def bloque_analisis(df_base: pd.DataFrame, variable_col: str, titulo: str, unidad: str, key_prefix: str):
        if variable_col not in df_base.columns:
            st.info(f"La columna {variable_col} no está disponible.")
            return

        maquinas_disponibles = sorted(df_base["ID_Maquina_Texto"].dropna().unique().tolist())

        col_a, col_b, col_c = st.columns(3)
        with col_a:
            granularidad = st.selectbox("Agrupar por:", options=["Minuto (sin agrupar)", "Hora", "Día"], key=f"{key_prefix}_gran")
        with col_b:
            estadistico = st.selectbox("Estadístico:", options=["Suma", "Promedio", "Mediana", "Mínimo", "Máximo", "Desviación estándar"], key=f"{key_prefix}_stat")
        with col_c:
            maquinas_sel = st.multiselect("Máquinas:", options=maquinas_disponibles, default=maquinas_disponibles, key=f"{key_prefix}_maq")

        df_analisis = df_base[df_base["ID_Maquina_Texto"].isin(maquinas_sel)].copy()
        df_analisis["Timestamp"] = pd.to_datetime(df_analisis["Timestamp"], errors="coerce")
        df_analisis[variable_col] = pd.to_numeric(df_analisis[variable_col], errors="coerce")
        df_analisis = df_analisis.dropna(subset=["Timestamp", variable_col])

        if df_analisis.empty:
            st.info(f"No hay datos de {titulo.lower()} para graficar.")
            return

        freq_map = {"Minuto (sin agrupar)": None, "Hora": "h", "Día": "D"}
        freq = freq_map[granularidad]

        func_map = {"Suma": "sum", "Promedio": "mean", "Mediana": "median", "Mínimo": "min", "Máximo": "max", "Desviación estándar": "std"}
        func = func_map[estadistico]

        df_analisis["Periodo"] = df_analisis["Timestamp"].dt.floor(freq) if freq else df_analisis["Timestamp"]

        tabla_resumen = df_analisis.groupby(["Periodo", "ID_Maquina_Texto"])[variable_col].agg(func).reset_index()

        st.markdown(f"**{estadistico} de {titulo} ({unidad}) por {granularidad.lower()}**")
        
        # Gráfico interactivo con Plotly
        fig = px.line(
            tabla_resumen,
            x="Periodo",
            y=variable_col,
            color="ID_Maquina_Texto",
            markers=True,
            template="plotly_white"
        )
        fig.update_layout(
            xaxis_title="Tiempo",
            yaxis_title=f"{titulo} ({unidad})",
            legend_title="Máquina",
            hovermode="x unified",
            margin=dict(l=0, r=0, t=30, b=0)
        )
        st.plotly_chart(fig, use_container_width=True)

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Suma total", f"{df_analisis[variable_col].sum():,.2f} {unidad}")
        m2.metric("Promedio", f"{df_analisis[variable_col].mean():,.2f} {unidad}")
        m3.metric("Mediana", f"{df_analisis[variable_col].median():,.2f} {unidad}")
        m4.metric("Máximo", f"{df_analisis[variable_col].max():,.2f} {unidad}")

    with st.expander("📊 Análisis y Visualización", expanded=True):
        if df_mostrar.empty:
            st.info("No hay datos para analizar.")
        else:
            tab_energia, tab_potencia = st.tabs(["⚡ Energía Consumida (kWh)", "🔌 Potencia (kW)"])
            with tab_energia:
                # Aquí se invoca el bloque para la columna Energia_kWh
                bloque_analisis(df_mostrar, "Energia_kWh", "Energía Consumida", "kWh", "energia")
            with tab_potencia:
                # Aquí se invoca el bloque para la columna Potencia_kW
                bloque_analisis(df_mostrar, "Potencia_kW", "Potencia", "kW", "potencia")

    col_dl, col_up = st.columns(2)

    with col_dl:
        csv_data = df_mostrar.to_csv(index=False, encoding='utf-8-sig')
        st.download_button(
            label="⬇️ Descargar como CSV",
            data=csv_data,
            file_name=f"Kern_Influx_Energia_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
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
