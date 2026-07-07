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
UMBRAL_PCT_CERO_SOSPECHOSO = 40.0  # % de lecturas en 0 dentro de la ventana a partir del cual el SEC se marca como sospechoso

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


@st.cache_resource(show_spinner=False)
def obtener_conexion_energia_db(clave_cache):
    """
    Descarga energia.db UNA sola vez (materializándolo a un archivo temporal local) y
    devuelve una conexión sqlite3 abierta, reutilizable entre reruns de Streamlit mientras
    'clave_cache' (la fecha de modificación en Drive) no cambie. A diferencia del enfoque
    anterior, esta función NO carga la tabla completa en un DataFrame — eso es lo que
    agotaba la memoria del servidor con archivos de cientos de MB. Las consultas reales se
    hacen después, por ventana de tiempo, con SQL (ver buscar_energia_ot).
    Devuelve (conn, archivo_meta, tabla_usada, tablas_disponibles) o (None, archivo_meta/None, None, []).
    """
    archivo = buscar_archivo_en_drive(NOMBRE_ARCHIVO_ENERGIA_DB, CARPETA_ENERGIA_DRIVE_ID)
    if archivo is None:
        return None, None, None, []

    contenido = _descargar_bytes_drive(archivo['id'])
    tmp = tempfile.NamedTemporaryFile(suffix='.db', delete=False)
    tmp.write(contenido)
    tmp.close()

    conn = sqlite3.connect(tmp.name, check_same_thread=False)
    tablas = pd.read_sql("SELECT name FROM sqlite_master WHERE type='table'", conn)['name'].tolist()
    if not tablas:
        conn.close()
        return None, archivo, None, []

    tabla_a_usar = TABLA_ENERGIA_PREFERIDA if TABLA_ENERGIA_PREFERIDA in tablas else tablas[0]
    return conn, archivo, tabla_a_usar, tablas


def resolver_columnas_energia_db(conn, tabla):
    """
    Inspecciona las columnas reales de la tabla (via PRAGMA, sin leer filas) y arma el mapeo
    a los nombres internos (Timestamp, ID_Maquina_Texto, Energia_kWh, Potencia_kW),
    tolerando variantes de nombre según ALIAS_COLUMNAS_ENERGIA.
    """
    info_columnas = pd.read_sql(f'PRAGMA table_info("{tabla}")', conn)
    columnas_reales = info_columnas['name'].tolist()

    mapa = {}
    faltantes = []
    for nombre_interno, alias in ALIAS_COLUMNAS_ENERGIA.items():
        encontrado = next((a for a in alias if a in columnas_reales), None)
        if encontrado:
            mapa[nombre_interno] = encontrado
        else:
            faltantes.append(nombre_interno)
    return mapa, faltantes, columnas_reales


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
    Convierte el valor crudo de 'Tiempo de Actividad'/'Tiempo de Inactividad' a minutos.
    Soporta los distintos tipos que puede entregar cada fuente de datos:
    - Vía gspread (Google Sheets API): normalmente llega como TEXTO ya formateado,
      ej. '21:53:00' o '7:21:00'. El formato de duración de Sheets ('[h]:mm:ss') expresa
      el total de horas sin reiniciar cada 24h, así que un texto como '168:00:00' (una
      semana) se interpreta correctamente sin necesidad de reconstruir nada.
    - Vía lectura directa de un .xlsx con openpyxl/pandas: llega como
      datetime.timedelta, datetime.datetime (cuando Excel guardó una duración >24h como
      fecha+hora, usando el epoch 1899-12-31) o datetime.time (duración <24h).
    - Número (int/float): fracción de día en formato serial de Excel/Sheets.
    - Vacío, 'Current', texto no reconocible o None -> np.nan (dato no disponible).
    """
    if valor is None:
        return np.nan
    if isinstance(valor, dt.timedelta):
        return valor.total_seconds() / 60.0
    if isinstance(valor, dt.datetime):
        dias = (valor.date() - EPOCH_EXCEL_DURACION).days
        return dias * 24 * 60 + valor.hour * 60 + valor.minute + valor.second / 60.0
    if isinstance(valor, dt.time):
        return valor.hour * 60 + valor.minute + valor.second / 60.0
    if isinstance(valor, (int, float)):
        if pd.isna(valor):
            return np.nan
        return float(valor) * 24 * 60

    texto = str(valor).strip()
    if texto == '' or texto.lower() in ('nan', 'none'):
        return np.nan

    # 'HH:MM:SS' o 'H:MM:SS' — soporta horas >24 (ej. '168:00:00' = 1 semana)
    m = re.match(r'^(\d{1,5}):(\d{2}):(\d{2})$', texto)
    if m:
        h, mi, s = m.groups()
        return int(h) * 60 + int(mi) + int(s) / 60.0

    # 'H:MM' sin segundos
    m = re.match(r'^(\d{1,5}):(\d{2})$', texto)
    if m:
        h, mi = m.groups()
        return int(h) * 60 + int(mi)

    # Número como texto (serial de fecha, fracción de día), por si Sheets lo entrega así
    try:
        return float(texto.replace(',', '.')) * 24 * 60
    except ValueError:
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
            'Pct_Lecturas_Cero': np.nan,
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

    # % de minutos donde el medidor reportó exactamente 0 kWh — muchos ceros dentro de
    # una ventana con buena cobertura suele indicar falla de lectura/umbral de standby
    # del medidor, no que la máquina realmente no consumió energía. Esto hace que el
    # SEC salga artificialmente bajo (mismo denominador de masa, menos energía sumada).
    pct_lecturas_cero = round(float((energia_ot['Energia_kWh'] <= 1e-9).mean() * 100), 1)

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
        'Pct_Lecturas_Cero': pct_lecturas_cero,
        'Score_Confiabilidad': round(score, 3),
        'Confiabilidad': label
    }


def buscar_energia_ot(conn, tabla, mapa_col, maquina, inicio_real, fin_real, tolerancia_min=TOLERANCIA_MINUTOS):
    """
    Trae SOLO las lecturas de energía de la máquina y ventana de tiempo de esta OT,
    consultando directamente energia.db con SQL — nunca carga la tabla completa en memoria.
    """
    inicio_span = inicio_real - pd.Timedelta(minutes=tolerancia_min)
    fin_span = fin_real + pd.Timedelta(minutes=tolerancia_min)

    col_ts, col_maq = mapa_col['Timestamp'], mapa_col['ID_Maquina_Texto']
    col_e, col_p = mapa_col['Energia_kWh'], mapa_col['Potencia_kW']

    query = f'''
        SELECT "{col_ts}" AS Timestamp, "{col_maq}" AS ID_Maquina_Texto,
               "{col_e}" AS Energia_kWh, "{col_p}" AS Potencia_kW
        FROM "{tabla}"
        WHERE UPPER(TRIM("{col_maq}")) = UPPER(?)
          AND "{col_ts}" >= ? AND "{col_ts}" <= ?
    '''
    params = (
        maquina,
        inicio_span.strftime('%Y-%m-%d %H:%M:%S'),
        fin_span.strftime('%Y-%m-%d %H:%M:%S'),
    )
    df = pd.read_sql(query, conn, params=params)

    df['Timestamp'] = pd.to_datetime(df['Timestamp'], errors='coerce', format='mixed').dt.floor('min')
    df.dropna(subset=['Timestamp', 'Energia_kWh'], inplace=True)
    df['Energia_kWh'] = pd.to_numeric(df['Energia_kWh'], errors='coerce')
    df['Potencia_kW'] = pd.to_numeric(df['Potencia_kW'], errors='coerce')
    df['ID_Maquina_Texto'] = df['ID_Maquina_Texto'].astype(str).str.strip().str.upper()
    return df, inicio_span, fin_span


def listar_maquinas_energia_db(conn, tabla, mapa_col):
    """Máquinas distintas presentes en energia.db, sin traer ninguna otra columna ni fila completa."""
    col_maq = mapa_col['ID_Maquina_Texto']
    df = pd.read_sql(f'SELECT DISTINCT "{col_maq}" AS m FROM "{tabla}"', conn)
    return set(df['m'].dropna().astype(str).str.strip().str.upper())


def contar_registros_energia_db(conn, tabla):
    return int(pd.read_sql(f'SELECT COUNT(*) AS n FROM "{tabla}"', conn)['n'].iloc[0])


# ==========================================
# FUNCIONES DE CÁLCULO — DIAGNÓSTICO DE CONVERGENCIA
# ==========================================
def agregar_columnas_diagnostico(df):
    """
    Agrega, de forma vectorizada, las columnas que permiten diagnosticar por qué una OT
    no converge: si los defectos por tipo cuadran con la Producción de Rechazo reportada,
    si los paros por causa cuadran con el Tiempo de Inactividad reportado, y arma un texto
    con el motivo consolidado (o varios motivos) por el que la OT no tiene SEC viable o
    tiene datos inconsistentes.
    """
    df = df.copy()

    # --- Defectos por tipo vs Producción de Rechazo ---
    cols_defecto = [c for c in df.columns if c.startswith('Defecto_')]
    if cols_defecto:
        df['Suma_Defectos'] = df[cols_defecto].apply(pd.to_numeric, errors='coerce').fillna(0).sum(axis=1)
    else:
        df['Suma_Defectos'] = np.nan
    df['Diferencia_Defectos_vs_Rechazo'] = df['Producción de Rechazo'] - df['Suma_Defectos']
    tol_defectos = np.maximum(1, 0.05 * df['Producción de Rechazo'])
    df['Defectos_Cuadran'] = df['Suma_Defectos'].notna() & (df['Diferencia_Defectos_vs_Rechazo'].abs() <= tol_defectos)

    # --- Paros por causa vs Tiempo de Inactividad ---
    cols_paro_tiempo = [c for c in df.columns if c.startswith('Paro_') and c.endswith('_Tiempo')]
    cols_paro_min = []
    for c in cols_paro_tiempo:
        col_min = f'_{c}_Min_tmp'
        df[col_min] = df[c].apply(parsear_duracion_minutos)
        cols_paro_min.append(col_min)
    if cols_paro_min:
        df['Suma_Paros_Min'] = df[cols_paro_min].sum(axis=1, skipna=True)
        df.drop(columns=cols_paro_min, inplace=True)
    else:
        df['Suma_Paros_Min'] = np.nan
    df['Diferencia_Paros_vs_Inactivo'] = df['Inactivo_Min'] - df['Suma_Paros_Min']
    tol_paros = np.maximum(5, 0.10 * df['Inactivo_Min'].fillna(0))
    df['Paros_Cuadran'] = df['Suma_Paros_Min'].notna() & (df['Diferencia_Paros_vs_Inactivo'].abs() <= tol_paros)

    # --- SEC sospechoso: se calculó, pero muchos minutos de la ventana tienen el medidor en 0 ---
    df['SEC_Sospechoso'] = (
        df['SEC_Total_kWh_kg'].notna() &
        df['Pct_Lecturas_Cero'].notna() &
        (df['Pct_Lecturas_Cero'] > UMBRAL_PCT_CERO_SOSPECHOSO)
    )

    # --- Motivo consolidado (texto legible por OT) ---
    def _motivo(row):
        motivos = []
        if not row['Ventana_Confiable']:
            dif = row.get('Diferencia_Ventana_Min', np.nan)
            motivos.append(
                f"Ventana de tiempo no confiable: el calendario (Fin−Inicio) difiere "
                f"{dif:.0f} min del tiempo real (Actividad+Inactividad)" if pd.notna(dif)
                else "Ventana de tiempo no confiable (faltan Tiempo de Actividad/Inactividad)"
            )
        elif row.get('N_Lecturas_Energia', 0) == 0:
            motivos.append("No se encontró ninguna lectura de energía para esa máquina en esa ventana de tiempo")
        elif row.get('SEC_Sospechoso', False):
            motivos.append(
                f"SEC calculado, pero {row['Pct_Lecturas_Cero']:.0f}% de los minutos en la ventana tienen el "
                f"medidor en 0 kWh — revisar la gráfica de energía de esta OT, el SEC podría estar subestimado"
            )
        if pd.isna(row.get('Total_Masa_Kg', np.nan)) or row.get('Total_Masa_Kg', 0) <= 0:
            motivos.append("Sin masa (kg) asociada a esta OT")
        if not row.get('Defectos_Cuadran', True):
            motivos.append(
                f"Defectos por tipo suman {row['Suma_Defectos']:.0f} vs. Producción de Rechazo "
                f"reportada de {row['Producción de Rechazo']:.0f} (dif. {row['Diferencia_Defectos_vs_Rechazo']:.0f})"
            )
        if not row.get('Paros_Cuadran', True):
            motivos.append(
                f"Paros por causa suman {row['Suma_Paros_Min']:.0f} min vs. Tiempo de Inactividad "
                f"reportado de {row.get('Inactivo_Min', float('nan')):.0f} min "
                f"(dif. {row['Diferencia_Paros_vs_Inactivo']:.0f} min)"
            )
        if not row.get('Produccion_Cuadra', True):
            motivos.append(
                f"Producción Total ({row['Producción Total']:.0f}) no coincide con Buena+Rechazo "
                f"({row['Producción Buena']:.0f}+{row['Producción de Rechazo']:.0f})"
            )
        if row.get('Solape_Ventana_Energia', False):
            motivos.append("Ventana de energía traslapada con la OT anterior de la misma máquina (riesgo de doble conteo)")
        return " | ".join(motivos) if motivos else "Sin inconsistencias detectadas"

    df['Diagnostico'] = df.apply(_motivo, axis=1)
    df['N_Problemas_Detectados'] = df['Diagnostico'].apply(
        lambda t: 0 if t == "Sin inconsistencias detectadas" else t.count('|') + 1
    )
    return df


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
# 2. CONSUMO ENERGÉTICO — AHORA DESDE energia.db EN DRIVE (consulta SQL, sin cargar todo a memoria)
# ==========================================
st.markdown("### 2. Consumo Energético (energia.db en Drive)")

col_refresh, col_info = st.columns([1, 4])
with col_refresh:
    forzar_recarga = st.button("🔄 Recargar energia.db")

with st.spinner(f"Buscando '{NOMBRE_ARCHIVO_ENERGIA_DB}' en la carpeta de Drive..."):
    archivo_preview = buscar_archivo_en_drive(NOMBRE_ARCHIVO_ENERGIA_DB, CARPETA_ENERGIA_DRIVE_ID)

if archivo_preview is None:
    st.error(f"❌ No se encontró un archivo llamado **{NOMBRE_ARCHIVO_ENERGIA_DB}** en la carpeta de Drive "
             f"configurada (ID `{CARPETA_ENERGIA_DRIVE_ID}`). Verifica que exista y que la cuenta de servicio "
             f"tenga acceso de Lector/Editor a esa carpeta.")
    st.stop()

clave_cache_energia = archivo_preview.get('modifiedTime', 'sin_fecha')
if forzar_recarga:
    obtener_conexion_energia_db.clear()

with st.spinner("Abriendo energia.db (solo se descarga completo si cambió en Drive; las consultas después son por ventana de tiempo)..."):
    conn_energia, meta_archivo, tabla_usada, tablas_disponibles = obtener_conexion_energia_db(clave_cache_energia)

if conn_energia is None:
    st.error(f"❌ Se encontró y descargó **{NOMBRE_ARCHIVO_ENERGIA_DB}**, pero no tiene tablas legibles.")
    st.stop()

mapa_col_energia, columnas_faltantes, columnas_reales_energia = resolver_columnas_energia_db(conn_energia, tabla_usada)

if columnas_faltantes:
    st.error(f"❌ La tabla `{tabla_usada}` de energia.db no tiene (ni con alias conocidos) estas columnas: "
             f"{columnas_faltantes}. Columnas reales encontradas: {columnas_reales_energia}. "
             f"Ajusta `ALIAS_COLUMNAS_ENERGIA` o `TABLA_ENERGIA_PREFERIDA` en el código con el nombre correcto.")
    if len(tablas_disponibles) > 1:
        st.info(f"Otras tablas disponibles en el archivo: {tablas_disponibles}")
    st.stop()

with col_info:
    fecha_mod = meta_archivo.get('modifiedTime', 'desconocida') if meta_archivo else 'desconocida'
    n_registros_energia = contar_registros_energia_db(conn_energia, tabla_usada)
    st.success(f"✅ **{NOMBRE_ARCHIVO_ENERGIA_DB}** conectado (tabla `{tabla_usada}`, {n_registros_energia:,} registros totales). "
               f"Última modificación en Drive: {fecha_mod}. Las consultas se hacen por ventana de tiempo de cada OT, "
               f"nunca se carga la tabla completa a memoria.")

# Se guarda para que el Inspector de Orden pueda reabrir la misma conexión (cacheada) más adelante.
st.session_state['clave_cache_energia'] = clave_cache_energia
st.session_state['tabla_energia_usada'] = tabla_usada
st.session_state['mapa_col_energia'] = mapa_col_energia

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

            # --- Diagnóstico del parseo: cuántas filas SÍ se pudieron leer y muestra de las que no ---
            n_activo_nan = int(df_prod['Activo_Min'].isna().sum())
            n_inactivo_nan = int(df_prod['Inactivo_Min'].isna().sum())
            with st.expander(
                f"🔬 Diagnóstico de lectura de 'Tiempo de Actividad'/'Tiempo de Inactividad' "
                f"({len(df_prod) - n_activo_nan}/{len(df_prod)} de Actividad y "
                f"{len(df_prod) - n_inactivo_nan}/{len(df_prod)} de Inactividad se leyeron con éxito)"
            ):
                st.write("Muestra de valores crudos vs. minutos ya interpretados (primeras 8 filas):")
                st.dataframe(
                    df_prod[['ID_Job', 'Tiempo de Actividad', 'Activo_Min', 'Tiempo de Inactividad', 'Inactivo_Min']].head(8),
                    use_container_width=True
                )
                if n_activo_nan > 0 or n_inactivo_nan > 0:
                    filas_no_leidas = df_prod[df_prod['Activo_Min'].isna() | df_prod['Inactivo_Min'].isna()]
                    st.write(f"Muestra de filas que **no** se pudieron interpretar ({len(filas_no_leidas)} en total):")
                    st.dataframe(
                        filas_no_leidas[['ID_Job', 'Tiempo de Actividad', 'Tiempo de Inactividad']].head(10),
                        use_container_width=True
                    )
                    st.caption("Si ves aquí un formato de texto distinto a 'HH:MM:SS' (ej. con la palabra 'día', "
                               "o un formato con coma decimal), avísame con un par de ejemplos exactos de esta "
                               "tabla para ajustar el parser a ese formato puntual.")

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

            # --- ENERGÍA: no se carga la tabla completa, solo se listan las máquinas presentes ---
            st.write(f"⚡ Consultando máquinas disponibles en energia.db (tabla `{tabla_usada}`)...")
            maquinas_prod = set(df_consolidado['ID_Maquina_Normalizado'].dropna().unique())
            maquinas_nrg = listar_maquinas_energia_db(conn_energia, tabla_usada, mapa_col_energia)
            maquinas_sin_energia = sorted(maquinas_prod - maquinas_nrg)
            if maquinas_sin_energia:
                st.warning(f"⚠️ Estas máquinas de Producción no aparecen (con ese nombre exacto) en energia.db: "
                           f"{maquinas_sin_energia}. Todas sus OTs quedarán con 'Sin datos'.")

            # --- CRUCE FINAL (MOTOR SEC) ---
            st.write(f"🧠 Ejecutando emparejamiento con tolerancia temporal (±{TOLERANCIA_MINUTOS} min), "
                     f"consultando energia.db OT por OT (sin cargarlo completo)...")
            resultados_sec = []

            for _, ot in df_consolidado.iterrows():
                maquina_ot = ot['ID_Maquina_Normalizado']
                inicio_real = ot['Inicio_Limpio']
                fin_real = ot['Fin_Limpio']

                if pd.isna(inicio_real) or pd.isna(fin_real):
                    continue

                ventana_confiable = bool(ot['Ventana_Confiable'])

                if ventana_confiable:
                    energia_ot, _, _ = buscar_energia_ot(
                        conn_energia, tabla_usada, mapa_col_energia, maquina_ot, inicio_real, fin_real
                    )
                else:
                    # No confiamos en el rango calendario: no buscamos energía para no
                    # atribuirle a esta OT el consumo de otras órdenes intermedias.
                    energia_ot = pd.DataFrame(columns=['Timestamp', 'ID_Maquina_Texto', 'Energia_kWh', 'Potencia_kW'])

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

            st.write("🩺 Calculando diagnóstico de convergencia (defectos, paros, ventanas)...")
            df_resultado_final = agregar_columnas_diagnostico(df_resultado_final)

            status.update(label="¡Cálculo SEC completado exitosamente!", state="complete", expanded=False)
            st.session_state['df_sec_calculado'] = df_resultado_final
            st.session_state['df_prod_std'] = df_prod
            st.session_state['df_masas_std'] = df_masas
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

    col1, col2, col3, col4, col5, col6 = st.columns(6)
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
    with col6:
        solo_sospechoso = st.checkbox(f"🕵️ Solo SEC sospechoso (>{UMBRAL_PCT_CERO_SOSPECHOSO:.0f}% ceros)", value=False)

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
    if solo_sospechoso:
        df_board = df_board[df_board['SEC_Sospechoso']]

    # Las OTs con SEC viable van primero, para que salten a la vista de inmediato.
    df_board = df_board.sort_values('SEC_Viable', ascending=False)

    st.markdown("#### Indicadores Globales (Según filtro)")
    kpi1, kpi2, kpi3, kpi4, kpi5, kpi6, kpi7 = st.columns(7)
    kpi1.metric("Órdenes Procesadas", f"{len(df_board)}")
    kpi2.metric("✅ Con SEC viable", f"{int(df_board['SEC_Viable'].sum())}")
    kpi3.metric("🕵️ SEC sospechoso", f"{int(df_board['SEC_Sospechoso'].sum())}")
    kpi4.metric("Masa Total (Kg)", f"{df_board['Total_Masa_Kg'].sum():,.2f}")
    kpi5.metric("Energía Total (kWh)", f"{df_board['Energia_Total_kWh'].sum():,.2f}")
    masa_total = df_board['Total_Masa_Kg'].sum()
    sec_promedio = df_board['Energia_Total_kWh'].sum() / masa_total if masa_total > 0 else 0
    kpi6.metric("SEC Total Promedio (kWh/kg)", f"{sec_promedio:.4f}")
    pct_alta = (df_board['Confiabilidad'] == 'Alta').mean() * 100 if len(df_board) > 0 else 0
    kpi7.metric("% OTs Confiabilidad Alta", f"{pct_alta:.0f}%")

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
    df_tabla['SEC_Sospechoso'] = df_tabla['SEC_Sospechoso'].map({True: '🕵️ Sí', False: '—'})

    st.dataframe(
        df_tabla[['SEC_Viable', 'SEC_Sospechoso', 'ID_Job', 'ID_Maquina', 'Inicio_Limpio', 'Fin_Limpio',
                   'Total_Masa_Kg', 'Masa_Buena_Kg', 'Energia_Total_kWh', 'Pct_Lecturas_Cero',
                   'SEC_Total_kWh_kg', 'SEC_Inyeccion_kWh_kg', 'SEC_Conforme_kWh_kg',
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
    if int(df_board['SEC_Sospechoso'].sum()) > 0:
        st.caption(
            f"🕵️ Las OTs marcadas como **SEC sospechoso** sí tienen SEC calculado, pero más del "
            f"{UMBRAL_PCT_CERO_SOSPECHOSO:.0f}% de los minutos en su ventana tienen el medidor reportando "
            f"exactamente 0 kWh (columna `Pct_Lecturas_Cero`). Eso puede ser real (la máquina realmente estuvo "
            f"parada) o una falla del medidor — revisa la gráfica de energía de esa OT en el Inspector de Orden "
            f"para diferenciar un caso del otro antes de confiar en ese SEC."
        )

    # ==========================================
    # 🩺 DIAGNÓSTICO DE CONVERGENCIA — muestra acotada de OTs
    # ==========================================
    st.divider()
    st.markdown("## 🩺 Diagnóstico de Convergencia (muestra)")
    st.write("Revisa una porción manejable de OTs para entender por qué no convergen: ventanas de tiempo que no "
             "cuadran, defectos por tipo que no suman lo mismo que la Producción de Rechazo, o paros por causa "
             "que no suman el Tiempo de Inactividad reportado.")

    dcol1, dcol2, dcol3 = st.columns([1.2, 1.5, 1])
    with dcol1:
        tamano_muestra = st.number_input("Tamaño de la muestra:", min_value=5, max_value=200, value=15, step=5)
    with dcol2:
        estrategia_muestra = st.selectbox(
            "¿Qué OTs incluir en la muestra?",
            [
                "🔥 Las que más problemas tienen (peor caso primero)",
                "🚫 Solo sin SEC viable",
                "🕵️ Solo SEC sospechoso (muchos ceros en energía)",
                "⚠️ Solo con defectos que no cuadran",
                "⚠️ Solo con paros que no cuadran",
                "⚠️ Solo con ventana no confiable",
                "🎲 Aleatoria",
                "📋 Las primeras N (tal como vienen)",
            ]
        )
    with dcol3:
        st.metric("OTs con ≥1 problema", f"{int((df_board['N_Problemas_Detectados'] > 0).sum())} / {len(df_board)}")

    df_diag = df_board.copy()
    if estrategia_muestra == "🔥 Las que más problemas tienen (peor caso primero)":
        df_muestra = df_diag.sort_values('N_Problemas_Detectados', ascending=False).head(int(tamano_muestra))
    elif estrategia_muestra == "🚫 Solo sin SEC viable":
        df_muestra = df_diag[~df_diag['SEC_Viable']].head(int(tamano_muestra))
    elif estrategia_muestra == "🕵️ Solo SEC sospechoso (muchos ceros en energía)":
        df_muestra = df_diag[df_diag['SEC_Sospechoso']].sort_values('Pct_Lecturas_Cero', ascending=False).head(int(tamano_muestra))
    elif estrategia_muestra == "⚠️ Solo con defectos que no cuadran":
        df_muestra = df_diag[~df_diag['Defectos_Cuadran']].head(int(tamano_muestra))
    elif estrategia_muestra == "⚠️ Solo con paros que no cuadran":
        df_muestra = df_diag[~df_diag['Paros_Cuadran']].head(int(tamano_muestra))
    elif estrategia_muestra == "⚠️ Solo con ventana no confiable":
        df_muestra = df_diag[~df_diag['Ventana_Confiable']].head(int(tamano_muestra))
    elif estrategia_muestra == "🎲 Aleatoria":
        df_muestra = df_diag.sample(n=min(int(tamano_muestra), len(df_diag)), random_state=None) if len(df_diag) > 0 else df_diag
    else:  # primeras N tal como vienen
        df_muestra = df_diag.head(int(tamano_muestra))

    if df_muestra.empty:
        st.info("No hay OTs que cumplan ese criterio con el filtro actual del tablero.")
    else:
        st.markdown(f"#### Resumen de la muestra ({len(df_muestra)} OTs)")
        st.dataframe(
            df_muestra[['ID_Job', 'ID_Maquina', 'SEC_Viable', 'Ventana_Confiable', 'Diferencia_Ventana_Min',
                        'Produccion_Cuadra', 'Defectos_Cuadran', 'Diferencia_Defectos_vs_Rechazo',
                        'Paros_Cuadran', 'Diferencia_Paros_vs_Inactivo', 'N_Problemas_Detectados', 'Diagnostico']],
            use_container_width=True
        )

        st.markdown("#### Detalle OT por OT (dentro de la muestra)")
        for _, ot_diag in df_muestra.iterrows():
            etiqueta_sec = "✅" if ot_diag['SEC_Viable'] else "🚫"
            with st.expander(f"{etiqueta_sec} OT `{ot_diag['ID_Job']}` — Máquina `{ot_diag['ID_Maquina']}` "
                              f"— {ot_diag['N_Problemas_Detectados']} problema(s)"):
                d1, d2, d3 = st.columns(3)
                d1.metric("Tiempo Calendario", f"{ot_diag['Duracion_Calendario_Min']:.0f} min")
                d2.metric("Tiempo Real (Activo+Inactivo)", f"{ot_diag['Duracion_Real_Min']:.0f} min"
                          if pd.notna(ot_diag['Duracion_Real_Min']) else "N/A")
                d3.metric("Diferencia de ventana", f"{ot_diag['Diferencia_Ventana_Min']:.0f} min"
                          if pd.notna(ot_diag['Diferencia_Ventana_Min']) else "N/A")

                e1, e2, e3 = st.columns(3)
                e1.metric("Producción Rechazo", f"{ot_diag['Producción de Rechazo']:.0f}")
                e2.metric("Suma Defectos por tipo", f"{ot_diag['Suma_Defectos']:.0f}" if pd.notna(ot_diag['Suma_Defectos']) else "N/A")
                e3.metric("Diferencia", f"{ot_diag['Diferencia_Defectos_vs_Rechazo']:.0f}" if pd.notna(ot_diag['Diferencia_Defectos_vs_Rechazo']) else "N/A")

                f1, f2, f3 = st.columns(3)
                f1.metric("Tiempo de Inactividad", f"{ot_diag['Inactivo_Min']:.0f} min" if pd.notna(ot_diag['Inactivo_Min']) else "N/A")
                f2.metric("Suma Paros por causa", f"{ot_diag['Suma_Paros_Min']:.0f} min" if pd.notna(ot_diag['Suma_Paros_Min']) else "N/A")
                f3.metric("Diferencia", f"{ot_diag['Diferencia_Paros_vs_Inactivo']:.0f} min" if pd.notna(ot_diag['Diferencia_Paros_vs_Inactivo']) else "N/A")

                st.markdown("**Motivo(s) detectado(s):**")
                if ot_diag['Diagnostico'] == "Sin inconsistencias detectadas":
                    st.success("Sin inconsistencias detectadas para esta OT.")
                else:
                    for motivo in ot_diag['Diagnostico'].split(" | "):
                        st.warning(f"• {motivo}")

                if ot_diag['Ventana_Confiable']:
                    energia_ot_diag, inicio_span_diag, fin_span_diag = buscar_energia_ot(
                        conn_energia, tabla_usada, mapa_col_energia,
                        ot_diag['ID_Maquina_Normalizado'],
                        pd.to_datetime(ot_diag['Inicio_Limpio']), pd.to_datetime(ot_diag['Fin_Limpio'])
                    )
                    if energia_ot_diag.empty:
                        st.info("No hay lecturas de energía en esta ventana para graficar.")
                    else:
                        g1, g2 = st.columns(2)
                        g1.metric("SEC Total", f"{ot_diag['SEC_Total_kWh_kg']:.4f} kWh/kg" if pd.notna(ot_diag['SEC_Total_kWh_kg']) else "N/A")
                        g2.metric("% minutos en 0 kWh", f"{ot_diag['Pct_Lecturas_Cero']:.0f}%" if pd.notna(ot_diag['Pct_Lecturas_Cero']) else "N/A")
                        st.line_chart(
                            energia_ot_diag.sort_values('Timestamp').set_index('Timestamp')[['Potencia_kW', 'Energia_kWh']],
                            use_container_width=True
                        )
                else:
                    st.caption("Ventana no confiable: no se consultó energía, por lo tanto no hay gráfica.")

        st.download_button(
            "⬇️ Descargar esta muestra en CSV",
            data=df_muestra.to_csv(index=False).encode('utf-8'),
            file_name="muestra_diagnostico_sec.csv",
            mime="text/csv"
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
            clave_cache_insp = st.session_state.get('clave_cache_energia')
            tabla_insp = st.session_state.get('tabla_energia_usada')
            mapa_col_insp = st.session_state.get('mapa_col_energia')

            if not fila_ot['Ventana_Confiable']:
                st.info("No se consultó energía para esta OT (ventana no confiable).")
            elif clave_cache_insp and tabla_insp and mapa_col_insp:
                # Misma conexión cacheada de la sección 2 (no vuelve a descargar el archivo).
                conn_insp, _, _, _ = obtener_conexion_energia_db(clave_cache_insp)
                maquina_norm = fila_ot['ID_Maquina_Normalizado']
                inicio_real = pd.to_datetime(fila_ot['Inicio_Limpio'])
                fin_real = pd.to_datetime(fila_ot['Fin_Limpio'])
                energia_ot, inicio_span, fin_span = buscar_energia_ot(
                    conn_insp, tabla_insp, mapa_col_insp, maquina_norm, inicio_real, fin_real
                )

                st.caption(f"Ventana real: **{inicio_real} → {fin_real}**. Con tolerancia (±{TOLERANCIA_MINUTOS} min): "
                           f"**{inicio_span} → {fin_span}**.")

                if energia_ot.empty:
                    st.info("No se encontró ninguna lectura de energía en esta ventana.")
                else:
                    if fila_ot.get('SEC_Sospechoso', False):
                        st.warning(
                            f"🕵️ SEC sospechoso: {fila_ot['Pct_Lecturas_Cero']:.0f}% de los minutos en esta ventana "
                            f"tienen el medidor en 0 kWh. Mira la gráfica de abajo para ver si son ceros dispersos "
                            f"(probable falla del medidor) o un tramo continuo (probable parada real de la máquina)."
                        )

                    st.markdown("**Comportamiento de energía durante la OT:**")
                    energia_ot_ordenada = energia_ot.sort_values('Timestamp')
                    st.line_chart(
                        energia_ot_ordenada.set_index('Timestamp')[['Potencia_kW', 'Energia_kWh']],
                        use_container_width=True
                    )
                    g1, g2, g3 = st.columns(3)
                    g1.metric("Lecturas totales", f"{fila_ot['N_Lecturas_Energia']:.0f}")
                    g2.metric("% minutos en 0 kWh", f"{fila_ot['Pct_Lecturas_Cero']:.0f}%" if pd.notna(fila_ot['Pct_Lecturas_Cero']) else "N/A")
                    g3.metric("Potencia máx.", f"{fila_ot['Potencia_Max_kW']:.2f} kW" if pd.notna(fila_ot['Potencia_Max_kW']) else "N/A")

                    st.markdown("**Detalle minuto a minuto:**")
                    st.dataframe(
                        energia_ot_ordenada[['Timestamp', 'Energia_kWh', 'Potencia_kW']],
                        use_container_width=True
                    )

                c1, c2 = st.columns(2)
                c1.metric("Energía activa estimada", f"{fila_ot['Energia_Activa_Estimada_kWh']:.2f} kWh"
                          if pd.notna(fila_ot['Energia_Activa_Estimada_kWh']) else "N/A")
                c2.metric("Energía en parada estimada", f"{fila_ot['Energia_Parada_Estimada_kWh']:.2f} kWh "
                          f"({fila_ot['Energia_Parada_Pct']:.1f}%)" if pd.notna(fila_ot['Energia_Parada_Pct']) else "N/A")
            else:
                st.info("Vuelve a correr el cruce en esta sesión para poder consultar el detalle de energía aquí.")

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
