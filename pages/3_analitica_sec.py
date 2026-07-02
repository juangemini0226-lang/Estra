import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import gspread
from google.oauth2.service_account import Credentials

# 1. Configuración de la página
st.set_page_config(page_title="Analítica SEC", page_icon="⚡", layout="wide")

st.title("⚡ Analítica de Consumo Específico de Energía (SEC)")
st.write("Cruza los datos de Producción, Masas y Energía para obtener el SEC por Orden de Trabajo.")

SHEET_URL = 'https://docs.google.com/spreadsheets/d/1lRg2Fc1pk3HBfXkYwXhWnFlTAGxx9gvoZ4hRnJ1AhXY/edit#gid=0'

# Columnas mínimas que DEBE tener cada hoja para ser considerada válida.
# IMPORTANTE: en el archivo real, la hoja de masas es 'Material_Data' y su
# llave YA se llama 'ID_Job' (no 'OT').
COLUMNAS_REQUERIDAS_PROD = {'Máquina', 'Trabajo / Orden', 'Tiempo Empezar', 'Tiempo Final'}
COLUMNAS_REQUERIDAS_MASAS = {'ID_Job', 'Total', 'Descripcion'}

HOJA_PROD_PREFERIDA = "produccion SEC"
HOJA_MASAS_PREFERIDA = "Material_Data"

TOLERANCIA_MINUTOS = 15  # margen de tolerancia temporal para emparejar energía

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
    ws = sh.worksheet(nombre_hoja)  # SIEMPRE por nombre, nunca por posición
    return pd.DataFrame(ws.get_all_records())


# ==========================================
# FUNCIONES DE CÁLCULO — MASAS
# ==========================================
def normalizar_maquina(id_maquina):
    if pd.isna(id_maquina): return id_maquina
    texto = str(id_maquina).strip().upper()
    if texto.endswith('MED'): return texto[:-3]
    return texto


def clasificar_unidad_masa(descripcion, total):
    """
    Extrae la unidad al final de la Descripcion (ej. '... ) - Kgs',
    '... ) - gr', '... ) - UN') y decide si esa fila debe contarse
    como masa del producto y con qué factor de conversión.

    Reglas:
    - 'kgs' / 'kg'  -> se usa tal cual (ya está en kilogramos)
    - 'gr'          -> se divide entre 1000 (masterbatch/aditivos en gramos)
    - 'un'          -> se EXCLUYE: es conteo de piezas de empaque
                       (bolsas, cajas, etiquetas), no es masa
    - cualquier otra cosa -> se EXCLUYE y se marca como 'desconocido'
                       para poder auditarla, en vez de sumarla a ciegas
    """
    texto = str(descripcion).strip()
    match = re.search(r'-\s*([A-Za-z]+)\s*$', texto)
    unidad = match.group(1).strip().lower() if match else 'desconocido'
    valor_total = float(total) if pd.notna(total) else 0.0

    if unidad in ('kgs', 'kg'):
        return unidad, valor_total, True
    elif unidad == 'gr':
        return unidad, valor_total / 1000.0, True
    elif unidad == 'un':
        return unidad, 0.0, False
    else:
        return unidad if unidad else 'desconocido', 0.0, False


# ==========================================
# FUNCIONES DE CÁLCULO — ENERGÍA
# ==========================================
def calcular_metricas_ot(energia_ot, inicio_real, fin_real):
    duracion_minutos = (fin_real - inicio_real).total_seconds() / 60.0
    duracion_minutos_efectiva = max(duracion_minutos, 1.0)
    n_lecturas = len(energia_ot)

    if n_lecturas == 0:
        return {
            'Energia_Total_kWh': 0.0, 'Potencia_Promedio_kW': np.nan, 'Potencia_Max_kW': np.nan,
            'Duracion_Min_OT': round(duracion_minutos_efectiva, 1),
            'Minutos_Cubiertos': 0, 'Cobertura_Pct': 0.0, 'Continuidad_Pct': 0.0,
            'Max_Gap_Min': np.nan, 'N_Lecturas_Energia': 0,
            'Score_Confiabilidad': 0.0, 'Confiabilidad': 'Sin datos'
        }

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
        'Duracion_Min_OT': round(duracion_minutos_efectiva, 1),
        'Minutos_Cubiertos': int(minutos_cubiertos),
        'Cobertura_Pct': round(cobertura_pct, 1),
        'Continuidad_Pct': round(continuidad_pct, 1),
        'Max_Gap_Min': round(float(max_gap), 1) if pd.notna(max_gap) else np.nan,
        'N_Lecturas_Energia': int(n_lecturas),
        'Score_Confiabilidad': round(score, 3),
        'Confiabilidad': label
    }


def buscar_energia_ot(df_energia, maquina, inicio_real, fin_real, tolerancia_min=TOLERANCIA_MINUTOS):
    inicio_span = inicio_real - pd.Timedelta(minutes=tolerancia_min)
    fin_span = fin_real + pd.Timedelta(minutes=tolerancia_min)
    mask = (df_energia['ID_Maquina_Texto'] == maquina) & (df_energia['Timestamp'] >= inicio_span) & (df_energia['Timestamp'] <= fin_span)
    return df_energia[mask].copy(), inicio_span, fin_span


# ==========================================
# 1. RESOLUCIÓN DE HOJAS
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
        st.info(f"ℹ️ Varias hojas califican como Producción: {candidatas_prod}. "
                f"Se preseleccionó '{HOJA_PROD_PREFERIDA}' por ser la hoja esperada; cambia si no aplica.")
        idx_default = candidatas_prod.index(HOJA_PROD_PREFERIDA) if HOJA_PROD_PREFERIDA in candidatas_prod else 0
        hoja_prod_elegida = st.selectbox("Confirma la hoja de Producción:", candidatas_prod, index=idx_default)

with col_b:
    if len(candidatas_masas) == 0:
        st.error(f"⚠️ Ninguna hoja tiene las columnas de Masas esperadas: {sorted(COLUMNAS_REQUERIDAS_MASAS)}")
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
                n_prod_antes = len(df_prod)
                df_prod = df_prod.dropna(subset=['Inicio_Limpio', 'Fin_Limpio'])
                n_prod_sin_fecha = n_prod_antes - len(df_prod)

                # Detectar ID_Job duplicados en Producción: si el mismo Job aparece
                # más de una vez (re-procesos, corridas repetidas), el merge con
                # masas lo va a multiplicar y las métricas de energía se
                # calcularán igual para ambas filas -> hay que avisarlo.
                duplicados_prod = df_prod['ID_Job'][df_prod['ID_Job'].duplicated(keep=False)].unique().tolist()

                df_prod['ID_Maquina_Normalizado'] = df_prod['ID_Maquina'].apply(normalizar_maquina)

                # --- PREPARAR MASAS: clasificación por unidad ---
                st.write("⚖️ Clasificando y consolidando Masas por ID_Job (filtrando por unidad)...")
                df_masas = df_masas_raw.copy()
                df_masas['ID_Job'] = df_masas['ID_Job'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                df_masas['Total'] = df_masas['Total'].astype(str).str.replace(',', '.', regex=False)
                df_masas['Total'] = pd.to_numeric(df_masas['Total'], errors='coerce')

                clasificacion = df_masas.apply(
                    lambda row: clasificar_unidad_masa(row['Descripcion'], row['Total']), axis=1
                )
                df_masas['Unidad_Detectada'] = [c[0] for c in clasificacion]
                df_masas['Valor_Kg_Equivalente'] = [c[1] for c in clasificacion]
                df_masas['Incluido_En_Calculo'] = [c[2] for c in clasificacion]

                n_excluidas = int((~df_masas['Incluido_En_Calculo']).sum())
                resumen_exclusion = df_masas.loc[~df_masas['Incluido_En_Calculo'], 'Unidad_Detectada'].value_counts()

                df_masas_incluidas = df_masas[df_masas['Incluido_En_Calculo']]
                df_masas_agg = df_masas_incluidas.groupby('ID_Job', as_index=False)['Valor_Kg_Equivalente'].sum()
                df_masas_agg.rename(columns={'Valor_Kg_Equivalente': 'Total_Masa_Kg'}, inplace=True)

                if n_excluidas > 0:
                    st.info(f"ℹ️ Se excluyeron {n_excluidas} filas de '{hoja_masas_elegida}' que no representan masa "
                            f"(unidades tipo UN/empaque o desconocidas). Detalle: {resumen_exclusion.to_dict()}. "
                            f"Puedes auditar cada fila en el Inspector de Orden más abajo.")
                if n_prod_sin_fecha > 0:
                    st.info(f"ℹ️ Se descartaron {n_prod_sin_fecha} filas de Producción sin fecha de inicio/fin válida.")
                if duplicados_prod:
                    st.warning(f"⚠️ {len(duplicados_prod)} ID_Job aparecen más de una vez en '{hoja_prod_elegida}' "
                               f"(ej. {duplicados_prod[:5]}...). El cruce va a generar una fila por cada aparición; "
                               f"revisa si son corridas legítimamente repetidas o duplicados de carga.")

                # --- MERGE PRODUCCIÓN + MASAS ---
                st.write("🔗 Uniendo Producción y Masas por ID_Job...")
                df_consolidado = pd.merge(df_prod, df_masas_agg, on='ID_Job', how='inner')

                n_prod_sin_masa = df_prod['ID_Job'].nunique() - df_consolidado['ID_Job'].nunique()
                if df_consolidado.empty:
                    st.warning("⚠️ El cruce Producción↔Masas dio 0 filas. Revisa que los ID_Job coincidan en formato.")
                    st.write("Ejemplos ID_Job Producción:", df_prod['ID_Job'].unique()[:5].tolist())
                    st.write("Ejemplos ID_Job Masas:", df_masas_agg['ID_Job'].unique()[:5].tolist())
                elif n_prod_sin_masa > 0:
                    st.info(f"ℹ️ {n_prod_sin_masa} órdenes de Producción no encontraron masa asociada y quedaron fuera del cruce.")

                # --- DETECCIÓN DE SOLAPES DE VENTANA DE ENERGÍA ---
                # Si dos OTs consecutivas en la misma máquina tienen sus ventanas
                # [Inicio-15min, Fin+15min] traslapadas, la misma lectura de
                # energía puede estar siendo contada en ambas órdenes.
                st.write("🕒 Revisando solapes de ventana temporal entre OTs de la misma máquina...")
                df_consolidado = df_consolidado.sort_values(['ID_Maquina_Normalizado', 'Inicio_Limpio']).reset_index(drop=True)
                df_consolidado['Inicio_Span'] = df_consolidado['Inicio_Limpio'] - pd.Timedelta(minutes=TOLERANCIA_MINUTOS)
                df_consolidado['Fin_Span'] = df_consolidado['Fin_Limpio'] + pd.Timedelta(minutes=TOLERANCIA_MINUTOS)
                fin_span_prev = df_consolidado.groupby('ID_Maquina_Normalizado')['Fin_Span'].shift(1)
                maquina_prev = df_consolidado.groupby('ID_Maquina_Normalizado')['ID_Maquina_Normalizado'].shift(1)
                df_consolidado['Solape_Ventana_Energia'] = (
                    (df_consolidado['ID_Maquina_Normalizado'] == maquina_prev) &
                    (df_consolidado['Inicio_Span'] < fin_span_prev)
                )
                n_solapes = int(df_consolidado['Solape_Ventana_Energia'].sum())
                if n_solapes > 0:
                    st.warning(f"⚠️ {n_solapes} OTs tienen su ventana de ±{TOLERANCIA_MINUTOS} min traslapada con la OT "
                               f"anterior de la misma máquina — posible doble conteo de energía. Se marcan en la tabla final "
                               f"con 'Solape_Ventana_Energia = True'.")

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
                df_energia['ID_Maquina_Texto'] = df_energia['ID_Maquina_Texto'].astype(str).str.strip().str.upper()

                # Chequeo de cobertura de máquinas: ¿las máquinas de Producción
                # existen tal cual en el CSV de energía?
                maquinas_prod = set(df_consolidado['ID_Maquina_Normalizado'].dropna().unique())
                maquinas_nrg = set(df_energia['ID_Maquina_Texto'].dropna().unique())
                maquinas_sin_energia = sorted(maquinas_prod - maquinas_nrg)
                if maquinas_sin_energia:
                    st.warning(f"⚠️ Estas máquinas de Producción no aparecen (con ese nombre exacto) en el CSV de energía: "
                               f"{maquinas_sin_energia}. Todas sus OTs quedarán con 'Sin datos'.")

                # --- CRUCE FINAL (MOTOR SEC) ---
                st.write(f"🧠 Ejecutando emparejamiento con tolerancia temporal (±{TOLERANCIA_MINUTOS} min)...")
                resultados_sec = []

                for _, ot in df_consolidado.iterrows():
                    maquina_ot = ot['ID_Maquina_Normalizado']
                    inicio_real = ot['Inicio_Limpio']
                    fin_real = ot['Fin_Limpio']

                    if pd.isna(inicio_real) or pd.isna(fin_real):
                        continue

                    energia_ot, _, _ = buscar_energia_ot(df_energia, maquina_ot, inicio_real, fin_real)
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
                # Guardamos los datos estandarizados (no solo el resultado final)
                # para poder auditar cualquier OT en el Inspector más abajo.
                st.session_state['df_prod_std'] = df_prod
                st.session_state['df_masas_std'] = df_masas
                st.session_state['df_energia_std'] = df_energia
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

    col1, col2, col3 = st.columns(3)
    with col1:
        maquinas = ['Todas'] + sorted(df_board['ID_Maquina_Normalizado'].dropna().unique().tolist())
        filtro_maq = st.selectbox("🏭 Máquina:", maquinas)
    with col2:
        confiabilidades = ['Todas'] + sorted(df_board['Confiabilidad'].dropna().unique().tolist())
        filtro_conf = st.selectbox("🎯 Confiabilidad del Dato:", confiabilidades)
    with col3:
        solo_solapes = st.checkbox("⚠️ Ver solo OTs con solape de ventana", value=False)

    if filtro_maq != 'Todas':
        df_board = df_board[df_board['ID_Maquina_Normalizado'] == filtro_maq]
    if filtro_conf != 'Todas':
        df_board = df_board[df_board['Confiabilidad'] == filtro_conf]
    if solo_solapes:
        df_board = df_board[df_board['Solape_Ventana_Energia'] == True]

    st.markdown("#### Indicadores Globales (Según filtro)")
    kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
    kpi1.metric("Órdenes Procesadas", f"{len(df_board)}")
    kpi2.metric("Masa Total (Kg)", f"{df_board['Total_Masa_Kg'].sum():,.2f}")
    kpi3.metric("Energía Total (kWh)", f"{df_board['Energia_Total_kWh'].sum():,.2f}")

    masa_total = df_board['Total_Masa_Kg'].sum()
    sec_promedio = df_board['Energia_Total_kWh'].sum() / masa_total if masa_total > 0 else 0
    kpi4.metric("SEC Promedio (kWh/kg)", f"{sec_promedio:.4f}")

    pct_alta = (df_board['Confiabilidad'] == 'Alta').mean() * 100 if len(df_board) > 0 else 0
    kpi5.metric("% OTs Confiabilidad Alta", f"{pct_alta:.0f}%")

    st.dataframe(
        df_board[['ID_Job', 'ID_Maquina', 'Inicio_Limpio', 'Fin_Limpio', 'Total_Masa_Kg',
                   'Energia_Total_kWh', 'SEC_kWh_kg', 'Cobertura_Pct', 'Continuidad_Pct',
                   'N_Lecturas_Energia', 'Confiabilidad', 'Solape_Ventana_Energia']],
        use_container_width=True
    )

    # ==========================================
    # 🔍 INSPECTOR DE ORDEN — trazabilidad completa por OT
    # ==========================================
    st.divider()
    st.markdown("## 🔍 Inspector de Orden")
    st.write("Elige un `ID_Job` para ver exactamente qué filas de masa se sumaron/excluyeron "
             "y qué lecturas de energía cayeron dentro de su ventana temporal.")

    ids_disponibles = sorted(df_board['ID_Job'].dropna().unique().tolist())
    if ids_disponibles:
        id_job_elegido = st.selectbox("ID_Job a inspeccionar:", ids_disponibles)

        fila_ot = df_board[df_board['ID_Job'] == id_job_elegido].iloc[0]

        st.markdown(f"### Orden `{id_job_elegido}` — Máquina `{fila_ot['ID_Maquina']}`")

        ic1, ic2, ic3, ic4 = st.columns(4)
        ic1.metric("Masa (Kg)", f"{fila_ot['Total_Masa_Kg']:.2f}")
        ic2.metric("Energía (kWh)", f"{fila_ot['Energia_Total_kWh']:.2f}")
        ic3.metric("SEC (kWh/kg)", f"{fila_ot['SEC_kWh_kg']:.4f}" if pd.notna(fila_ot['SEC_kWh_kg']) else "N/A")
        ic4.metric("Confiabilidad", fila_ot['Confiabilidad'])

        if fila_ot.get('Solape_Ventana_Energia', False):
            st.warning("⚠️ Esta orden tiene su ventana de energía traslapada con la OT anterior en la misma máquina — "
                       "la energía mostrada podría incluir consumo de la orden vecina.")

        tab_masa, tab_energia, tab_cobertura = st.tabs(["⚖️ Detalle de Masas", "⚡ Lecturas de Energía", "🎯 Cómo se calculó la Confiabilidad"])

        with tab_masa:
            df_masas_std = st.session_state.get('df_masas_std')
            if df_masas_std is not None:
                detalle_masa = df_masas_std[df_masas_std['ID_Job'] == str(id_job_elegido)].copy()
                if detalle_masa.empty:
                    st.info("No hay filas de masa para este ID_Job (verifica formato de llave).")
                else:
                    detalle_masa['¿Incluida en Total_Masa_Kg?'] = detalle_masa['Incluido_En_Calculo'].map({True: '✅ Sí', False: '❌ No'})
                    st.dataframe(
                        detalle_masa[['Descripcion', 'Total', 'Unidad_Detectada', 'Valor_Kg_Equivalente', '¿Incluida en Total_Masa_Kg?']],
                        use_container_width=True
                    )
                    total_incluido = detalle_masa.loc[detalle_masa['Incluido_En_Calculo'], 'Valor_Kg_Equivalente'].sum()
                    total_excluido_filas = (~detalle_masa['Incluido_En_Calculo']).sum()
                    st.caption(f"Suma de filas incluidas: **{total_incluido:.2f} Kg** "
                               f"(debe coincidir con la métrica de arriba). Filas excluidas: {total_excluido_filas}.")

        with tab_energia:
            df_energia_std = st.session_state.get('df_energia_std')
            if df_energia_std is not None:
                maquina_norm = fila_ot['ID_Maquina_Normalizado']
                inicio_real = pd.to_datetime(fila_ot['Inicio_Limpio'])
                fin_real = pd.to_datetime(fila_ot['Fin_Limpio'])
                energia_ot, inicio_span, fin_span = buscar_energia_ot(df_energia_std, maquina_norm, inicio_real, fin_real)

                st.caption(f"Ventana real de la OT: **{inicio_real} → {fin_real}**. "
                           f"Ventana con tolerancia (±{TOLERANCIA_MINUTOS} min): **{inicio_span} → {fin_span}**. "
                           f"Máquina buscada en energía: `{maquina_norm}`.")

                if energia_ot.empty:
                    st.info("No se encontró ninguna lectura de energía en esta ventana. Revisa si el nombre de máquina "
                            "coincide exactamente entre Producción y el CSV de energía.")
                else:
                    energia_ot_ordenada = energia_ot.sort_values('Timestamp')
                    energia_ot_ordenada['¿Dentro del rango real (sin tolerancia)?'] = (
                        (energia_ot_ordenada['Timestamp'] >= inicio_real) & (energia_ot_ordenada['Timestamp'] <= fin_real)
                    ).map({True: '✅ Sí', False: '🟡 Solo en margen ±15min'})
                    st.dataframe(
                        energia_ot_ordenada[['Timestamp', 'Energia_kWh', 'Potencia_kW', '¿Dentro del rango real (sin tolerancia)?']],
                        use_container_width=True
                    )
                    n_en_margen = (energia_ot_ordenada['¿Dentro del rango real (sin tolerancia)?'] == '🟡 Solo en margen ±15min').sum()
                    if n_en_margen > 0:
                        st.caption(f"🟡 {n_en_margen} de {len(energia_ot_ordenada)} lecturas están solo en el margen de tolerancia, "
                                   f"no en el rango exacto de la OT.")

        with tab_cobertura:
            duracion = fila_ot['Duracion_Min_OT']
            cubiertos = fila_ot['Minutos_Cubiertos']
            cobertura = fila_ot['Cobertura_Pct']
            continuidad = fila_ot['Continuidad_Pct']
            gap = fila_ot['Max_Gap_Min']
            score = fila_ot['Score_Confiabilidad']

            st.markdown(f"""
- **Duración de la OT:** {duracion:.1f} minutos
- **Minutos con al menos 1 lectura de energía:** {cubiertos} → **Cobertura = {cobertura:.1f}%**
  (`minutos_cubiertos / duración_OT`, tope 100%)
- **Mayor hueco (gap) sin lecturas:** {gap:.1f} min → **Continuidad = {continuidad:.1f}%**
  (100% si el gap más grande es ≤ 5 min; si no, penaliza proporcional al exceso sobre la duración)
- **Score final = Cobertura × 0.5 + Continuidad × 0.5 = {score:.3f}**
- **Etiqueta:** Alta (≥0.85) / Media (≥0.60) / Baja (resto) → **{fila_ot['Confiabilidad']}**
""")
    else:
        st.info("No hay órdenes disponibles con el filtro actual para inspeccionar.")
