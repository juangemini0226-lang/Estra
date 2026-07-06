import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import re
import sqlite3
import tempfile
import datetime as dt
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# 1. Configuración de la página
st.set_page_config(page_title="Analítica SEC", page_icon="⚡", layout="wide")

st.title("⚡ Analítica de Consumo Específico de Energía (SEC)")
st.write("Cruza Producción, Masas, Costos y Energía para obtener el SEC por Orden de Trabajo, "
         "incluyendo el impacto de la no conformidad.")

SHEET_URL = 'https://docs.google.com/spreadsheets/d/1lRg2Fc1pk3HBfXkYwXhWnFlTAGxx9gvoZ4hRnJ1AhXY/edit#gid=0'

# --- Ubicación de la base de datos de energía en Drive ---
CARPETA_ENERGIA_DRIVE_ID = "131X02ZCk-UyABfxFMHZD0Loidb1jfIw3"
NOMBRE_ARCHIVO_ENERGIA_DB = "energia.db"
# Nombre de tabla esperado dentro del .db (ajusta aquí si tu tabla se llama distinto).
TABLA_ENERGIA_PREFERIDA = "registros_energia"

# Columnas mínimas que DEBE tener cada hoja para ser considerada válida.
COLUMNAS_REQUERIDAS_PROD = {
    'Máquina', 'Trabajo / Orden', 'Tiempo Empezar', 'Tiempo Final', 'Número de Parte',
    'Producción Total', 'Producción Buena', 'Producción de Rechazo',
    'Tiempo de Actividad', 'Tiempo de Inactividad'
}
COLUMNAS_REQUERIDAS_MASAS = {'ID_Job', 'Total', 'Descripcion'}
COLUMNAS_REQUERIDAS_COSTO = {'Item', 'costo estandar'}

HOJA_PROD_PREFERIDA = "produccion detallada"
HOJA_MASAS_PREFERIDA = "Material_Data"
HOJA_COSTO_PREFERIDA = "Maestra Costo Estandar"

TOLERANCIA_MINUTOS = 15          # margen de tolerancia temporal para emparejar energía
EPOCH_EXCEL_DURACION = dt.date(1899, 12, 31)  # base para reconstruir duraciones > 24h mal formateadas por Excel
UMBRAL_VENTANA_MIN_HORAS = 2.0   # tolerancia mínima (horas) entre calendario y (activo+inactivo)
UMBRAL_VENTANA_PCT = 0.20        # tolerancia relativa (20% de la duración calendario)

# Posibles nombres de columnas dentro de energia.db, para tolerar variantes de esquema.
ALIAS_COLUMNAS_ENERGIA = {
    'Timestamp': ['Timestamp', 'Fecha y hora', 'fecha_hora', 'timestamp', 'Fecha_Hora'],
    'ID_Maquina_Texto': ['ID_Maquina_Texto', 'maquina_o_puesto', 'Maquina', 'ID_Maquina', 'maquina'],
    'Energia_kWh': ['Energia_kWh', 'Energía [kWh]', 'energia_kwh', 'Energia'],
    'Potencia_kW': ['Potencia_kW', 'Potencia [kW]', 'potencia_kw', 'Potencia'],
}

# ==========================================
# FUNCIONES DE CONEXIÓN Y DESCARGA — GOOGLE SHEETS
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
# FUNCIONES DE CONEXIÓN Y DESCARGA — GOOGLE DRIVE (energia.db)
# ==========================================
@st.cache_resource
def conectar_drive():
    # drive.readonly basta: solo necesitamos leer/descargar el archivo, no crearlo ni editarlo.
    scope = ["https://www.googleapis.com/auth/drive.readonly"]
    creds_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
    return build('drive', 'v3', credentials=creds)


def buscar_archivo_en_drive(nombre_archivo, carpeta_id):
    """Busca por nombre exacto dentro de una carpeta específica de Drive."""
    servicio = conectar_drive()
    query = (
        f"name = '{nombre_archivo}' and '{carpeta_id}' in parents and trashed = false"
    )
    resultado = servicio.files().list(
        q=query,
        fields="files(id, name, modifiedTime, size)",
        spaces='drive'
    ).execute()
    archivos = resultado.get('files', [])
    return archivos[0] if archivos else None


def _descargar_bytes_drive(file_id):
    servicio = conectar_drive()
    request = servicio.files().get_media(fileId=file_id)
    buffer = io.BytesIO()
    downloader = MediaIoBaseDownload(buffer, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    buffer.seek(0)
    return buffer.read()


@st.cache_data(ttl=300, show_spinner=False)
def descargar_energia_db(_cache_bust=None):
    """
    Busca energia.db en la carpeta de Drive configurada, lo descarga y devuelve
    (DataFrame de la tabla de energía, metadata del archivo) o (None, None) si no existe.
    El parámetro _cache_bust permite forzar una recarga manual (botón "Recargar").
    """
    archivo = buscar_archivo_en_drive(NOMBRE_ARCHIVO_ENERGIA_DB, CARPETA_ENERGIA_DRIVE_ID)
    if archivo is None:
        return None, None, "no_encontrado", []

    contenido = _descargar_bytes_drive(archivo['id'])

    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(suffix='.db', delete=False) as tmp:
            tmp.write(contenido)
            tmp_path = tmp.name

        conn = sqlite3.connect(tmp_path)
        try:
            tablas = pd.read_sql(
                "SELECT name FROM sqlite_master WHERE type='table'", conn
            )['name'].tolist()

            if not tablas:
                return None, archivo, "sin_tablas", []

            tabla_a_usar = TABLA_ENERGIA_PREFERIDA if TABLA_ENERGIA_PREFERIDA in tablas else tablas[0]
            df = pd.read_sql(f"SELECT * FROM {tabla_a_usar}", conn)
            return df, archivo, tabla_a_usar, tablas
        finally:
            conn.close()
    finally:
        if tmp_path and os.path.exists(tmp_path):
            os.remove(tmp_path)


def estandarizar_columnas_energia(df_crudo):
    """
    Renombra las columnas de energia.db a los nombres internos que usa el motor SEC
    (Timestamp, ID_Maquina_Texto, Energia_kWh, Potencia_kW), tolerando variantes de
    nombre según ALIAS_COLUMNAS_ENERGIA. Devuelve (df_renombrado, columnas_faltantes).
    """
    df = df_crudo.copy()
    df.columns = df.columns.str.strip()
    mapa_renombre = {}
    faltantes = []

    for nombre_interno, alias in ALIAS_COLUMNAS_ENERGIA.items():
        encontrado = next((a for a in alias if a in df.columns), None)
        if encontrado:
            mapa_renombre[encontrado] = nombre_interno
        else:
            faltantes.append(nombre_interno)

    df.rename(columns=mapa_renombre, inplace=True)
    return df, faltantes


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
    Extrae la unidad al final de la Descripcion y decide si esa fila debe
    contarse como masa del producto y con qué factor de conversión.
    'kgs'/'kg' -> tal cual. 'gr' -> /1000. 'un' o desconocido -> se excluye.
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
# FUNCIONES DE CÁLCULO — TIEMPOS DE ACTIVIDAD/INACTIVIDAD
# ==========================================
def parsear_duracion_minutos(valor):
    """
    Convierte el valor crudo de 'Tiempo de Actividad'/'Tiempo de Inactividad'
    a minutos, sin importar cómo Excel/Sheets lo haya tipado:
    - datetime.timedelta -> ya es una duración real, se usa tal cual.
    - datetime.datetime  -> Excel guardó una duración >24h como fecha+hora;
      se reconstruye usando el offset de días desde 1899-12-31 (el mismo
      epoch que usa Excel para sus números seriales de fecha/hora).
    - datetime.time      -> duración dentro de un mismo día (<24h).
    - cualquier otro caso (NaN, texto raro) -> np.nan (dato no disponible).
    """
    if isinstance(valor, dt.timedelta):
        return valor.total_seconds() / 60.0
    if isinstance(valor, dt.datetime):
        dias = (valor.date() - EPOCH_EXCEL_DURACION).days
        return dias * 24 * 60 + valor.hour * 60 + valor.minute + valor.second / 60.0
    if isinstance(valor, dt.time):
        return valor.hour * 60 + valor.minute + valor.second / 60.0
    return np.nan


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
candidatas_costo = detectar_hoja(info_hojas, COLUMNAS_REQUERIDAS_COSTO)

with st.expander("🔎 Ver hojas detectadas en el archivo (diagnóstico)"):
    for nombre, columnas in info_hojas.items():
        st.write(f"**{nombre}** → columnas: {columnas}")

col_a, col_b, col_c = st.columns(3)

with col_a:
    if len(candidatas_prod) == 0:
        st.error(f"⚠️ Ninguna hoja tiene las columnas de Producción esperadas (incluye calidad y tiempos de actividad).")
        hoja_prod_elegida = st.selectbox("Elige manualmente la hoja de Producción:", list(info_hojas.keys()))
    elif len(candidatas_prod) == 1:
        hoja_prod_elegida = candidatas_prod[0]
        st.success(f"✅ Hoja de Producción detectada: **{hoja_prod_elegida}**")
    else:
        st.info(f"ℹ️ Varias hojas califican como Producción: {candidatas_prod}. "
                f"Se preseleccionó '{HOJA_PROD_PREFERIDA}'.")
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
        idx_default = candidatas_masas.index(HOJA_MASAS_PREFERIDA) if HOJA_MASAS_PREFERIDA in candidatas_masas else 0
        hoja_masas_elegida = st.selectbox("Confirma la hoja de Masas:", candidatas_masas, index=idx_default)

with col_c:
    if len(candidatas_costo) == 0:
        st.error(f"⚠️ Ninguna hoja tiene las columnas de Costo esperadas: {sorted(COLUMNAS_REQUERIDAS_COSTO)}")
        hoja_costo_elegida = st.selectbox("Elige manualmente la hoja de Costo Estándar:", list(info_hojas.keys()))
    elif len(candidatas_costo) == 1:
        hoja_costo_elegida = candidatas_costo[0]
        st.success(f"✅ Hoja de Costo detectada: **{hoja_costo_elegida}**")
    else:
        idx_default = candidatas_costo.index(HOJA_COSTO_PREFERIDA) if HOJA_COSTO_PREFERIDA in candidatas_costo else 0
        hoja_costo_elegida = st.selectbox("Confirma la hoja de Costo Estándar:", candidatas_costo, index=idx_default)

with st.spinner(f"Descargando '{hoja_prod_elegida}', '{hoja_masas_elegida}' y '{hoja_costo_elegida}'..."):
    df_prod_raw = descargar_hoja(hoja_prod_elegida)
    df_masas_raw = descargar_hoja(hoja_masas_elegida)
    df_costo_raw = descargar_hoja(hoja_costo_elegida)

if df_prod_raw.empty or df_masas_raw.empty:
    st.warning("No se encontraron datos de Producción o Masas en las hojas seleccionadas.")
    st.stop()

faltantes_masas = COLUMNAS_REQUERIDAS_MASAS - set(df_masas_raw.columns)
if faltantes_masas:
    st.error(f"❌ La hoja '{hoja_masas_elegida}' no tiene las columnas requeridas: {faltantes_masas}.")
    st.stop()

faltantes_prod = COLUMNAS_REQUERIDAS_PROD - set(df_prod_raw.columns)
if faltantes_prod:
    st.error(f"❌ La hoja '{hoja_prod_elegida}' no tiene las columnas requeridas: {faltantes_prod}. "
             f"Columnas encontradas: {list(df_prod_raw.columns)}")
    st.stop()

faltantes_costo = COLUMNAS_REQUERIDAS_COSTO - set(df_costo_raw.columns)
if faltantes_costo:
    st.error(f"❌ La hoja '{hoja_costo_elegida}' no tiene las columnas requeridas: {faltantes_costo}.")
    st.stop()

st.success(f"Datos base descargados: {len(df_prod_raw)} OTs de Producción, {len(df_masas_raw)} registros de Materiales "
           f"y {len(df_costo_raw)} referencias de Costo Estándar.")

# ==========================================
# 2. CONSUMO ENERGÉTICO — AHORA DESDE energia.db EN DRIVE (sin CSV manual)
# ==========================================
st.markdown("### 2. Consumo Energético (energia.db en Drive)")

col_refresh, col_info = st.columns([1, 4])
with col_refresh:
    forzar_recarga = st.button("🔄 Recargar energia.db")

if forzar_recarga:
    descargar_energia_db.clear()

with st.spinner(f"Buscando '{NOMBRE_ARCHIVO_ENERGIA_DB}' en la carpeta de Drive..."):
    # el segundo argumento solo se usa para invalidar el cache cuando se pulsa "Recargar"
    df_energia_bruto, meta_archivo, tabla_usada, tablas_disponibles = descargar_energia_db(
        _cache_bust=dt.datetime.now().isoformat() if forzar_recarga else None
    )

if df_energia_bruto is None:
    if tabla_usada == "no_encontrado":
        st.error(f"❌ No se encontró un archivo llamado **{NOMBRE_ARCHIVO_ENERGIA_DB}** en la carpeta de Drive "
                 f"configurada (ID `{CARPETA_ENERGIA_DRIVE_ID}`). Verifica que exista y que la cuenta de servicio "
                 f"tenga acceso de Lector/Editor a esa carpeta.")
    else:
        st.error(f"❌ Se encontró y descargó **{NOMBRE_ARCHIVO_ENERGIA_DB}**, pero no tiene tablas legibles.")
    st.stop()

df_energia, columnas_faltantes = estandarizar_columnas_energia(df_energia_bruto)

if columnas_faltantes:
    st.error(f"❌ La tabla `{tabla_usada}` de energia.db no tiene (ni con alias conocidos) estas columnas: "
             f"{columnas_faltantes}. Columnas reales encontradas: {list(df_energia_bruto.columns)}. "
             f"Ajusta `ALIAS_COLUMNAS_ENERGIA` o `TABLA_ENERGIA_PREFERIDA` en el código con el nombre correcto.")
    if len(tablas_disponibles) > 1:
        st.info(f"Otras tablas disponibles en el archivo: {tablas_disponibles}")
    st.stop()

with col_info:
    fecha_mod = meta_archivo.get('modifiedTime', 'desconocida') if meta_archivo else 'desconocida'
    st.success(f"✅ **{NOMBRE_ARCHIVO_ENERGIA_DB}** cargado desde Drive (tabla `{tabla_usada}`, {len(df_energia):,} registros). "
               f"Última modificación en Drive: {fecha_mod}.")

if st.button("🚀 Iniciar Cruce y Cálculo SEC", type="primary"):
    with st.status("Procesando Motor SEC...", expanded=True) as status:
        try:
            # --- PREPARAR PRODUCCIÓN ---
            st.write("⚙️ Estandarizando Producción...")
            df_prod = df_prod_raw.copy()
            df_prod.rename(columns={
                'Máquina': 'ID_Maquina', 'Trabajo / Orden': 'ID_Job',
                'Tiempo Empezar': 'Inicio', 'Tiempo Final': 'Fin',
                'Número de Parte': 'ID_Parte'
            }, inplace=True)
            df_prod['ID_Job'] = df_prod['ID_Job'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            df_prod['ID_Parte'] = df_prod['ID_Parte'].astype(str).str.strip()

            df_prod['Inicio_Limpio'] = pd.to_datetime(df_prod['Inicio'], errors='coerce', format='mixed', dayfirst=True)
            # Esto también captura los 'Current' (OTs aún en curso, sin fecha de fin real)
            df_prod['Fin_Limpio'] = pd.to_datetime(df_prod['Fin'], errors='coerce', format='mixed', dayfirst=True)
            n_prod_antes = len(df_prod)
            n_en_curso = (df_prod['Fin'].astype(str).str.strip().str.lower() == 'current').sum()
            df_prod = df_prod.dropna(subset=['Inicio_Limpio', 'Fin_Limpio'])
            n_prod_sin_fecha = n_prod_antes - len(df_prod)

            duplicados_prod = df_prod['ID_Job'][df_prod['ID_Job'].duplicated(keep=False)].unique().tolist()
            df_prod['ID_Maquina_Normalizado'] = df_prod['ID_Maquina'].apply(normalizar_maquina)

            # --- VALIDACIÓN DE VENTANA DE TIEMPO CONFIABLE ---
            st.write("🕵️ Validando ventanas de tiempo (calendario vs. tiempo real de la OT)...")
            df_prod['Activo_Min'] = df_prod['Tiempo de Actividad'].apply(parsear_duracion_minutos)
            df_prod['Inactivo_Min'] = df_prod['Tiempo de Inactividad'].apply(parsear_duracion_minutos)
            df_prod['Duracion_Calendario_Min'] = (df_prod['Fin_Limpio'] - df_prod['Inicio_Limpio']).dt.total_seconds() / 60.0
            df_prod['Duracion_Real_Min'] = df_prod['Activo_Min'] + df_prod['Inactivo_Min']

            tiene_tiempos = df_prod['Duracion_Real_Min'].notna()
            diferencia_min = (df_prod['Duracion_Calendario_Min'] - df_prod['Duracion_Real_Min']).abs()
            umbral_min = np.maximum(UMBRAL_VENTANA_MIN_HORAS * 60, UMBRAL_VENTANA_PCT * df_prod['Duracion_Calendario_Min'])
            df_prod['Ventana_Confiable'] = tiene_tiempos & (diferencia_min <= umbral_min)
            df_prod['Diferencia_Ventana_Min'] = np.where(tiene_tiempos, diferencia_min, np.nan)

            n_ventana_no_confiable = int((~df_prod['Ventana_Confiable']).sum())

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

            # --- PREPARAR COSTO ESTÁNDAR ---
            st.write("💲 Preparando Costo Estándar por referencia (Item = Número de Parte)...")
            df_costo = df_costo_raw.copy()
            df_costo['ID_Parte'] = df_costo['Item'].astype(str).str.strip()
            df_costo['Costo_Unitario_Estandar'] = pd.to_numeric(df_costo['costo estandar'], errors='coerce')
            df_costo_dedup = df_costo.drop_duplicates(subset='ID_Parte', keep='first')
            cols_costo_extra = [c for c in ['Materiales', 'Recursos', 'CIF', 'Proceso Externo'] if c in df_costo_dedup.columns]
            df_costo_final = df_costo_dedup[['ID_Parte', 'Costo_Unitario_Estandar'] + cols_costo_extra]

            if n_excluidas > 0:
                st.info(f"ℹ️ Se excluyeron {n_excluidas} filas de '{hoja_masas_elegida}' que no representan masa "
                        f"(unidades tipo UN/empaque o desconocidas). Detalle: {resumen_exclusion.to_dict()}.")
            if n_prod_sin_fecha > 0:
                st.info(f"ℹ️ Se descartaron {n_prod_sin_fecha} filas de Producción sin fecha válida "
                        f"(incluye {n_en_curso} OTs aún en curso marcadas como 'Current').")
            if duplicados_prod:
                st.warning(f"⚠️ {len(duplicados_prod)} ID_Job aparecen más de una vez en '{hoja_prod_elegida}' "
                           f"(ej. {duplicados_prod[:5]}...).")
            if n_ventana_no_confiable > 0:
                st.warning(f"⚠️ {n_ventana_no_confiable} OTs tienen una diferencia grande entre el tiempo calendario "
                           f"(Fin−Inicio) y su tiempo real (Actividad+Inactividad) — probablemente la máquina corrió "
                           f"otras órdenes en el medio. **Se excluyen del cruce de energía** para no contaminar el "
                           f"cálculo de otras OTs (umbral: {UMBRAL_VENTANA_MIN_HORAS}h o {UMBRAL_VENTANA_PCT*100:.0f}% "
                           f"de la duración, lo que sea mayor).")

            # --- MERGE PRODUCCIÓN + MASAS + COSTO ---
            st.write("🔗 Uniendo Producción, Masas y Costo Estándar...")
            df_consolidado = pd.merge(df_prod, df_masas_agg, on='ID_Job', how='inner')
            df_consolidado = pd.merge(df_consolidado, df_costo_final, on='ID_Parte', how='left')

            n_prod_sin_masa = df_prod['ID_Job'].nunique() - df_consolidado['ID_Job'].nunique()
            n_sin_costo = df_consolidado['Costo_Unitario_Estandar'].isna().sum()

            if df_consolidado.empty:
                st.warning("⚠️ El cruce Producción↔Masas dio 0 filas. Revisa que los ID_Job coincidan en formato.")
            else:
                if n_prod_sin_masa > 0:
                    st.info(f"ℹ️ {n_prod_sin_masa} órdenes de Producción no encontraron masa asociada y quedaron fuera del cruce.")
                if n_sin_costo > 0:
                    st.warning(f"⚠️ {n_sin_costo} OTs no encontraron costo estándar para su Número de Parte — "
                               f"su costo de no conformidad quedará en blanco.")

            # --- VALIDACIÓN: Producción Total debe cuadrar con Buena + Rechazo ---
            for col in ['Producción Total', 'Producción Buena', 'Producción de Rechazo']:
                df_consolidado[col] = pd.to_numeric(df_consolidado[col], errors='coerce').fillna(0)
            diff_prod = (df_consolidado['Producción Total'] -
                         (df_consolidado['Producción Buena'] + df_consolidado['Producción de Rechazo'])).abs()
            tol_prod = np.maximum(1, 0.01 * df_consolidado['Producción Total'])
            df_consolidado['Produccion_Cuadra'] = diff_prod <= tol_prod
            n_prod_no_cuadra = int((~df_consolidado['Produccion_Cuadra']).sum())
            if n_prod_no_cuadra > 0:
                st.warning(f"⚠️ {n_prod_no_cuadra} OTs tienen 'Producción Total' que no coincide con "
                           f"'Buena + Rechazo' — revisar en el Inspector de Orden (columna 'Produccion_Cuadra').")

            # --- DETECCIÓN DE SOLAPES DE VENTANA DE ENERGÍA ---
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
                           f"anterior de la misma máquina — posible doble conteo de energía.")

            # --- PREPARAR ENERGÍA (ya viene de energia.db, descargado arriba) ---
            st.write(f"⚡ Preparando datos de Energía desde energia.db (tabla `{tabla_usada}`)...")
            df_energia_prep = df_energia.copy()

            df_energia_prep['Timestamp'] = pd.to_datetime(
                df_energia_prep['Timestamp'].astype(str), format='mixed', errors='coerce'
            ).dt.floor('min')
            df_energia_prep.dropna(subset=['Timestamp', 'Energia_kWh'], inplace=True)
            df_energia_prep['Energia_kWh'] = pd.to_numeric(df_energia_prep['Energia_kWh'], errors='coerce')
            df_energia_prep['Potencia_kW'] = pd.to_numeric(df_energia_prep['Potencia_kW'], errors='coerce')
            df_energia_prep['ID_Maquina_Texto'] = df_energia_prep['ID_Maquina_Texto'].astype(str).str.strip().str.upper()

            maquinas_prod = set(df_consolidado['ID_Maquina_Normalizado'].dropna().unique())
            maquinas_nrg = set(df_energia_prep['ID_Maquina_Texto'].dropna().unique())
            maquinas_sin_energia = sorted(maquinas_prod - maquinas_nrg)
            if maquinas_sin_energia:
                st.warning(f"⚠️ Estas máquinas de Producción no aparecen (con ese nombre exacto) en energia.db: "
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

                ventana_confiable = bool(ot['Ventana_Confiable'])

                if ventana_confiable:
                    energia_ot, _, _ = buscar_energia_ot(df_energia_prep, maquina_ot, inicio_real, fin_real)
                else:
                    # No confiamos en el rango calendario: no buscamos energía para no
                    # atribuirle a esta OT el consumo de otras órdenes intermedias.
                    energia_ot = df_energia_prep.iloc[0:0]

                metricas = calcular_metricas_ot(energia_ot, inicio_real, fin_real)
                if not ventana_confiable:
                    metricas['Confiabilidad'] = 'Sin Ventana Confiable'

                masa_total_kg = ot['Total_Masa_Kg'] if pd.notna(ot['Total_Masa_Kg']) else np.nan
                prod_total = ot['Producción Total']
                prod_buena = ot['Producción Buena']
                prod_rechazo = ot['Producción de Rechazo']
                costo_unitario = ot['Costo_Unitario_Estandar'] if pd.notna(ot.get('Costo_Unitario_Estandar', np.nan)) else np.nan

                masa_buena_kg = (masa_total_kg * (prod_buena / prod_total)
                                  if (pd.notna(masa_total_kg) and prod_total > 0) else np.nan)

                energia_total = metricas['Energia_Total_kWh']
                activo_min = ot['Activo_Min']
                inactivo_min = ot['Inactivo_Min']
                duracion_real_min = ot['Duracion_Real_Min']

                if ventana_confiable and pd.notna(duracion_real_min) and duracion_real_min > 0:
                    energia_activa_kwh = energia_total * (activo_min / duracion_real_min)
                    energia_parada_kwh = energia_total * (inactivo_min / duracion_real_min)
                    parada_pct = (inactivo_min / duracion_real_min) * 100
                else:
                    energia_activa_kwh = np.nan
                    energia_parada_kwh = np.nan
                    parada_pct = np.nan

                sec_total = (round(energia_total / masa_total_kg, 4)
                             if (ventana_confiable and pd.notna(masa_total_kg) and masa_total_kg > 0) else np.nan)
                sec_inyeccion = (round(energia_activa_kwh / masa_total_kg, 4)
                                  if (pd.notna(energia_activa_kwh) and pd.notna(masa_total_kg) and masa_total_kg > 0) else np.nan)
                sec_conforme = (round(energia_total / masa_buena_kg, 4)
                                 if (ventana_confiable and pd.notna(masa_buena_kg) and masa_buena_kg > 0) else np.nan)

                costo_no_conformidad = round(prod_rechazo * costo_unitario, 0) if pd.notna(costo_unitario) else np.nan
                energia_desperdiciada_rechazo = (round(energia_total * (prod_rechazo / prod_total), 4)
                                                   if (ventana_confiable and prod_total > 0) else np.nan)

                fila_resultado = ot.to_dict()
                fila_resultado.update(metricas)
                fila_resultado['Masa_Buena_Kg'] = round(masa_buena_kg, 3) if pd.notna(masa_buena_kg) else np.nan
                fila_resultado['Energia_Activa_Estimada_kWh'] = round(energia_activa_kwh, 3) if pd.notna(energia_activa_kwh) else np.nan
                fila_resultado['Energia_Parada_Estimada_kWh'] = round(energia_parada_kwh, 3) if pd.notna(energia_parada_kwh) else np.nan
                fila_resultado['Energia_Parada_Pct'] = round(parada_pct, 1) if pd.notna(parada_pct) else np.nan
                fila_resultado['SEC_Total_kWh_kg'] = sec_total
                fila_resultado['SEC_Inyeccion_kWh_kg'] = sec_inyeccion
                fila_resultado['SEC_Conforme_kWh_kg'] = sec_conforme
                fila_resultado['Costo_No_Conformidad'] = costo_no_conformidad
                fila_resultado['Energia_Desperdiciada_Rechazo_kWh'] = energia_desperdiciada_rechazo
                resultados_sec.append(fila_resultado)

            df_resultado_final = pd.DataFrame(resultados_sec)

            status.update(label="¡Cálculo SEC completado exitosamente!", state="complete", expanded=False)
            st.session_state['df_sec_calculado'] = df_resultado_final
            st.session_state['df_prod_std'] = df_prod
            st.session_state['df_masas_std'] = df_masas
            st.session_state['df_energia_std'] = df_energia_prep
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

    # Bandera clara de "esta OT sí tiene un dato de SEC utilizable" (energía encontrada + masa + ventana confiable).
    df_board['SEC_Viable'] = df_board['SEC_Total_kWh_kg'].notna()

    n_total_ots = len(df_board)
    n_viables = int(df_board['SEC_Viable'].sum())
    n_no_viables = n_total_ots - n_viables

    st.markdown(
        f"**{n_viables} de {n_total_ots} OTs** ({(n_viables/n_total_ots*100 if n_total_ots else 0):.0f}%) "
        f"tienen datos de energía viables y ya muestran su SEC calculado."
    )

    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        maquinas = ['Todas'] + sorted(df_board['ID_Maquina_Normalizado'].dropna().unique().tolist())
        filtro_maq = st.selectbox("🏭 Máquina:", maquinas)
    with col2:
        confiabilidades = ['Todas'] + sorted(df_board['Confiabilidad'].dropna().unique().tolist())
        filtro_conf = st.selectbox("🎯 Confiabilidad del Dato:", confiabilidades)
    with col3:
        solo_solapes = st.checkbox("⚠️ Solo con solape de ventana", value=False)
    with col4:
        solo_no_cuadra = st.checkbox("⚠️ Solo Producción sin cuadre", value=False)
    with col5:
        filtro_sec = st.selectbox(
            "⚡ SEC calculado:",
            ["Todas", "✅ Solo con SEC viable", "🚫 Solo sin SEC (sin datos viables)"]
        )

    if filtro_maq != 'Todas':
        df_board = df_board[df_board['ID_Maquina_Normalizado'] == filtro_maq]
    if filtro_conf != 'Todas':
        df_board = df_board[df_board['Confiabilidad'] == filtro_conf]
    if solo_solapes:
        df_board = df_board[df_board['Solape_Ventana_Energia'] == True]
    if solo_no_cuadra:
        df_board = df_board[df_board['Produccion_Cuadra'] == False]
    if filtro_sec == "✅ Solo con SEC viable":
        df_board = df_board[df_board['SEC_Viable']]
    elif filtro_sec == "🚫 Solo sin SEC (sin datos viables)":
        df_board = df_board[~df_board['SEC_Viable']]

    # Las OTs con SEC viable van primero, para que salten a la vista de inmediato.
    df_board = df_board.sort_values('SEC_Viable', ascending=False)

    st.markdown("#### Indicadores Globales (Según filtro)")
    kpi1, kpi2, kpi3, kpi4, kpi5, kpi6 = st.columns(6)
    kpi1.metric("Órdenes Procesadas", f"{len(df_board)}")
    kpi2.metric("✅ Con SEC viable", f"{int(df_board['SEC_Viable'].sum())}")
    kpi3.metric("Masa Total (Kg)", f"{df_board['Total_Masa_Kg'].sum():,.2f}")
    kpi4.metric("Energía Total (kWh)", f"{df_board['Energia_Total_kWh'].sum():,.2f}")
    masa_total = df_board['Total_Masa_Kg'].sum()
    sec_promedio = df_board['Energia_Total_kWh'].sum() / masa_total if masa_total > 0 else 0
    kpi5.metric("SEC Total Promedio (kWh/kg)", f"{sec_promedio:.4f}")
    pct_alta = (df_board['Confiabilidad'] == 'Alta').mean() * 100 if len(df_board) > 0 else 0
    kpi6.metric("% OTs Confiabilidad Alta", f"{pct_alta:.0f}%")

    st.markdown("#### SEC en sus 3 variantes")
    s1, s2, s3 = st.columns(3)
    masa_buena_total = df_board['Masa_Buena_Kg'].sum()
    energia_activa_total = df_board['Energia_Activa_Estimada_kWh'].sum()
    sec_iny_prom = energia_activa_total / masa_total if masa_total > 0 else 0
    sec_conf_prom = df_board['Energia_Total_kWh'].sum() / masa_buena_total if masa_buena_total > 0 else 0
    s1.metric("⚙️ SEC Inyección (activo/masa total)", f"{sec_iny_prom:.4f} kWh/kg")
    s2.metric("✅ SEC Conforme (toda energía/masa buena)", f"{sec_conf_prom:.4f} kWh/kg")
    energia_parada_total = df_board['Energia_Parada_Estimada_kWh'].sum()
    pct_parada_energia = (energia_parada_total / df_board['Energia_Total_kWh'].sum() * 100
                            if df_board['Energia_Total_kWh'].sum() > 0 else 0)
    s3.metric("🛑 Energía en Paros", f"{energia_parada_total:,.2f} kWh ({pct_parada_energia:.1f}%)")
    st.caption("La energía activa/parada es una **estimación proporcional** al tiempo (Activo vs Inactivo reportado "
               "en producción), ya que no existe una marca minuto a minuto del estado de la máquina en la señal de energía.")

    st.markdown("#### Impacto de la No Conformidad")
    n1, n2, n3 = st.columns(3)
    costo_nc_total = df_board['Costo_No_Conformidad'].sum()
    energia_nc_total = df_board['Energia_Desperdiciada_Rechazo_kWh'].sum()
    rechazo_total = df_board['Producción de Rechazo'].sum()
    n1.metric("💸 Costo No Conformidad", f"${costo_nc_total:,.0f}")
    n2.metric("⚡ Energía en Rechazo", f"{energia_nc_total:,.2f} kWh")
    n3.metric("📦 Unidades Rechazadas", f"{rechazo_total:,.0f}")

    df_tabla = df_board.copy()
    df_tabla['SEC_Viable'] = df_tabla['SEC_Viable'].map({True: '✅ Sí', False: '🚫 No'})

    st.dataframe(
        df_tabla[['SEC_Viable', 'ID_Job', 'ID_Maquina', 'Inicio_Limpio', 'Fin_Limpio', 'Total_Masa_Kg', 'Masa_Buena_Kg',
                   'Energia_Total_kWh', 'SEC_Total_kWh_kg', 'SEC_Inyeccion_kWh_kg', 'SEC_Conforme_kWh_kg',
                   'Producción de Rechazo', 'Costo_No_Conformidad', 'Energia_Desperdiciada_Rechazo_kWh',
                   'Cobertura_Pct', 'Continuidad_Pct', 'Confiabilidad', 'Ventana_Confiable',
                   'Solape_Ventana_Energia', 'Produccion_Cuadra']],
        use_container_width=True
    )

    if n_no_viables > 0:
        st.caption(
            f"ℹ️ Las {n_no_viables} OTs marcadas con 🚫 no tienen SEC calculado porque: no se encontró ninguna "
            f"lectura de energía en su ventana de tiempo, la ventana de tiempo no era confiable (columna "
            f"`Ventana_Confiable` = False), o no tienen masa asociada > 0. Usa el Inspector de Orden de abajo "
            f"para ver la razón exacta de cada caso."
        )

    # ==========================================
    # 🔍 INSPECTOR DE ORDEN — trazabilidad completa por OT
    # ==========================================
    st.divider()
    st.markdown("## 🔍 Inspector de Orden")
    st.write("Elige un `ID_Job` para ver la trazabilidad completa: masa, energía, calidad, paros y confiabilidad.")

    ids_disponibles = sorted(df_board['ID_Job'].dropna().unique().tolist())
    if ids_disponibles:
        id_job_elegido = st.selectbox("ID_Job a inspeccionar:", ids_disponibles)
        fila_ot = df_board[df_board['ID_Job'] == id_job_elegido].iloc[0]

        st.markdown(f"### Orden `{id_job_elegido}` — Máquina `{fila_ot['ID_Maquina']}`")

        ic1, ic2, ic3, ic4, ic5 = st.columns(5)
        ic1.metric("Masa Total (Kg)", f"{fila_ot['Total_Masa_Kg']:.2f}")
        ic2.metric("Masa Buena (Kg)", f"{fila_ot['Masa_Buena_Kg']:.2f}" if pd.notna(fila_ot['Masa_Buena_Kg']) else "N/A")
        ic3.metric("Energía (kWh)", f"{fila_ot['Energia_Total_kWh']:.2f}")
        ic4.metric("SEC Conforme", f"{fila_ot['SEC_Conforme_kWh_kg']:.4f}" if pd.notna(fila_ot['SEC_Conforme_kWh_kg']) else "N/A")
        ic5.metric("Confiabilidad", fila_ot['Confiabilidad'])

        if not fila_ot['Ventana_Confiable']:
            st.error(f"🚫 Esta OT tiene ventana **no confiable**: la diferencia entre tiempo calendario y tiempo real "
                     f"(Activo+Inactivo) es de {fila_ot['Diferencia_Ventana_Min']:.0f} minutos. No se buscó energía "
                     f"para evitar atribuirle consumo de otra orden.")
        if fila_ot.get('Solape_Ventana_Energia', False):
            st.warning("⚠️ Esta orden tiene su ventana de energía traslapada con la OT anterior en la misma máquina.")
        if not fila_ot.get('Produccion_Cuadra', True):
            st.warning(f"⚠️ Producción Total ({fila_ot['Producción Total']:.0f}) no coincide con "
                       f"Buena+Rechazo ({fila_ot['Producción Buena']:.0f}+{fila_ot['Producción de Rechazo']:.0f}).")

        tab_masa, tab_energia, tab_calidad, tab_paros, tab_cobertura = st.tabs(
            ["⚖️ Masas", "⚡ Energía", "🧯 No Conformidad", "⏱️ Paros", "🎯 Confiabilidad"]
        )

        with tab_masa:
            df_masas_std = st.session_state.get('df_masas_std')
            if df_masas_std is not None:
                detalle_masa = df_masas_std[df_masas_std['ID_Job'] == str(id_job_elegido)].copy()
                if detalle_masa.empty:
                    st.info("No hay filas de masa para este ID_Job.")
                else:
                    detalle_masa['¿Incluida?'] = detalle_masa['Incluido_En_Calculo'].map({True: '✅ Sí', False: '❌ No'})
                    st.dataframe(
                        detalle_masa[['Descripcion', 'Total', 'Unidad_Detectada', 'Valor_Kg_Equivalente', '¿Incluida?']],
                        use_container_width=True
                    )
                    total_incluido = detalle_masa.loc[detalle_masa['Incluido_En_Calculo'], 'Valor_Kg_Equivalente'].sum()
                    st.caption(f"Suma de filas incluidas: **{total_incluido:.2f} Kg**.")

        with tab_energia:
            df_energia_std = st.session_state.get('df_energia_std')
            if not fila_ot['Ventana_Confiable']:
                st.info("No se consultó energía para esta OT (ventana no confiable).")
            elif df_energia_std is not None:
                maquina_norm = fila_ot['ID_Maquina_Normalizado']
                inicio_real = pd.to_datetime(fila_ot['Inicio_Limpio'])
                fin_real = pd.to_datetime(fila_ot['Fin_Limpio'])
                energia_ot, inicio_span, fin_span = buscar_energia_ot(df_energia_std, maquina_norm, inicio_real, fin_real)

                st.caption(f"Ventana real: **{inicio_real} → {fin_real}**. Con tolerancia (±{TOLERANCIA_MINUTOS} min): "
                           f"**{inicio_span} → {fin_span}**.")

                if energia_ot.empty:
                    st.info("No se encontró ninguna lectura de energía en esta ventana.")
                else:
                    energia_ot_ordenada = energia_ot.sort_values('Timestamp')
                    st.dataframe(
                        energia_ot_ordenada[['Timestamp', 'Energia_kWh', 'Potencia_kW']],
                        use_container_width=True
                    )

                c1, c2 = st.columns(2)
                c1.metric("Energía activa estimada", f"{fila_ot['Energia_Activa_Estimada_kWh']:.2f} kWh"
                          if pd.notna(fila_ot['Energia_Activa_Estimada_kWh']) else "N/A")
                c2.metric("Energía en parada estimada", f"{fila_ot['Energia_Parada_Estimada_kWh']:.2f} kWh "
                          f"({fila_ot['Energia_Parada_Pct']:.1f}%)" if pd.notna(fila_ot['Energia_Parada_Pct']) else "N/A")

        with tab_calidad:
            st.markdown(f"**Producción Total:** {fila_ot['Producción Total']:.0f} | "
                        f"**Buena:** {fila_ot['Producción Buena']:.0f} | "
                        f"**Rechazo:** {fila_ot['Producción de Rechazo']:.0f}")
            costo_unit = fila_ot.get('Costo_Unitario_Estandar', np.nan)
            if pd.notna(costo_unit):
                st.metric("Costo No Conformidad", f"${fila_ot['Costo_No_Conformidad']:,.0f}",
                          help=f"Costo estándar unitario: ${costo_unit:,.0f} × {fila_ot['Producción de Rechazo']:.0f} unidades rechazadas")
            else:
                st.info("No se encontró costo estándar para el Número de Parte de esta OT.")

            cols_defecto = [c for c in fila_ot.index if c.startswith('Defecto_')]
            defectos = {c.replace('Defecto_', ''): fila_ot[c] for c in cols_defecto if pd.notna(fila_ot[c]) and fila_ot[c] > 0}
            if defectos:
                df_defectos = pd.DataFrame(list(defectos.items()), columns=['Tipo de Defecto', 'Cantidad']).sort_values('Cantidad', ascending=False)
                st.dataframe(df_defectos, use_container_width=True)
            else:
                st.caption("Sin defectos registrados por tipo para esta OT.")

        with tab_paros:
            cols_paro_tiempo = [c for c in fila_ot.index if c.startswith('Paro_') and c.endswith('_Tiempo')]
            filas_paro = []
            for c in cols_paro_tiempo:
                causa = c.replace('Paro_', '').replace('_Tiempo', '')
                col_cuenta = f'Paro_{causa}_Cuenta'
                valor_tiempo = fila_ot[c]
                valor_cuenta = fila_ot[col_cuenta] if col_cuenta in fila_ot.index else np.nan
                if pd.notna(valor_cuenta) and valor_cuenta and valor_cuenta > 0:
                    filas_paro.append({'Causa de Paro': causa, 'Tiempo': str(valor_tiempo), 'Cuenta': valor_cuenta})
            if filas_paro:
                st.dataframe(pd.DataFrame(filas_paro).sort_values('Cuenta', ascending=False), use_container_width=True)
            else:
                st.caption("Sin paros registrados por causa para esta OT.")

            st.markdown("---")
            p1, p2, p3 = st.columns(3)
            p1.metric("Tiempo Calendario", f"{fila_ot['Duracion_Calendario_Min']:.0f} min")
            p2.metric("Tiempo Real (Activo+Inactivo)", f"{fila_ot['Duracion_Real_Min']:.0f} min"
                      if pd.notna(fila_ot['Duracion_Real_Min']) else "N/A")
            p3.metric("Diferencia", f"{fila_ot['Diferencia_Ventana_Min']:.0f} min"
                      if pd.notna(fila_ot['Diferencia_Ventana_Min']) else "N/A")

        with tab_cobertura:
            duracion = fila_ot['Duracion_Min_OT']
            cubiertos = fila_ot['Minutos_Cubiertos']
            cobertura = fila_ot['Cobertura_Pct']
            continuidad = fila_ot['Continuidad_Pct']
            gap = fila_ot['Max_Gap_Min']
            score = fila_ot['Score_Confiabilidad']

            st.markdown(f"""
- **Duración de la OT (calendario):** {duracion:.1f} minutos
- **Minutos con al menos 1 lectura de energía:** {cubiertos} → **Cobertura = {cobertura:.1f}%**
- **Mayor hueco (gap) sin lecturas:** {gap:.1f} min → **Continuidad = {continuidad:.1f}%**
- **Score final = Cobertura × 0.5 + Continuidad × 0.5 = {score:.3f}**
- **Etiqueta:** Alta (≥0.85) / Media (≥0.60) / Baja (resto) / Sin Ventana Confiable → **{fila_ot['Confiabilidad']}**

**Cómo se calculan las 3 variantes de SEC:**
- `SEC_Total_kWh_kg` = Energía Total consumida en la ventana ÷ Masa Total transformada
- `SEC_Inyeccion_kWh_kg` = Energía estimada durante Tiempo de Actividad ÷ Masa Total transformada
- `SEC_Conforme_kWh_kg` = Energía Total consumida (activa + parada) ÷ Masa **Buena** (excluyendo el rechazo)
""")
    else:
        st.info("No hay órdenes disponibles con el filtro actual para inspeccionar.")
