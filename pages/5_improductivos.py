import streamlit as st
import pandas as pd
import csv
import io
import re
import time
import gspread
from google.oauth2.service_account import Credentials
from gspread_dataframe import set_with_dataframe, get_as_dataframe

# 1. Configuración de la página
st.set_page_config(page_title="Improductivos | Extractor", page_icon="🛑", layout="wide")

st.title("🛑 Extractor de Tiempos Improductivos")
st.write(
    "Sube el CSV del **Informe de Tiempo de Inactividad de Trabajo** (reporte de paros por OT/máquina) "
    "para extraer las causas de paro y cargarlas a la pestaña **Improductivos** del tablero maestro."
)

# Misma hoja de cálculo que usan los demás módulos del tablero
SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1lRg2Fc1pk3HBfXkYwXhWnFlTAGxx9gvoZ4hRnJ1AhXY/edit#gid=0'
NOMBRE_HOJA_DESTINO = "Improductivos"

# Etiqueta ancla: todas las filas del reporte traen este literal justo antes
# del bloque de datos reales (Departamento, Máquina, Trabajo, Parte, tiempos y causas).
ETIQUETA_ANCLA = "Departmento"


@st.cache_resource
def conectar_sheets():
    scope = ["https://www.googleapis.com/auth/spreadsheets"]
    creds_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
    return gspread.authorize(creds)


def hhmm_a_minutos(valor):
    """Convierte 'H:MM' o 'HH:MM' a minutos totales (float). Devuelve None si no se puede parsear."""
    if valor is None:
        return None
    valor = str(valor).strip()
    if not re.match(r'^\d{1,4}:\d{2}$', valor):
        return None
    horas, minutos = valor.split(":")
    try:
        return int(horas) * 60 + int(minutos)
    except ValueError:
        return None


def procesar_fila(row):
    """
    El reporte no usa el patrón 'Etiqueta : valor' (como el de Producción),
    sino que repite un bloque de encabezados literales en cada fila y luego
    los datos reales en posiciones fijas relativas a la etiqueta 'Departmento'.

    Estructura relativa (offset desde el índice de 'Departmento'):
      +0  'Departmento' (etiqueta)      +1 Departamento        +2 '-'   +3 Descripción Depto
      +4  Máquina                       +5 Trabajo / Orden     +6 Número de Parte
      +7  Tiempo de Actividad           +8 Tiempo de Inactividad   +9 % de Inactividad
      +10 Código Causa 1  +11 Cuenta 1  +12 Tiempo 1
      +13 Código Causa 2  +14 Cuenta 2  +15 Tiempo 2
      +16 Código Causa 3  +17 Cuenta 3  +18 Tiempo 3
      ... más adelante en la fila: fecha/hora de generación del reporte
    """
    fila_limpia = [str(c).strip() for c in row]
    try:
        i = fila_limpia.index(ETIQUETA_ANCLA)
    except ValueError:
        return None

    if len(fila_limpia) < i + 19:
        return None

    base = {
        "Departamento": fila_limpia[i + 1],
        "Máquina": fila_limpia[i + 4],
        "Trabajo / Orden": fila_limpia[i + 5],
        "Número de Parte": fila_limpia[i + 6],
        "Tiempo de Actividad": fila_limpia[i + 7],
        "Tiempo de Inactividad": fila_limpia[i + 8],
        "% de Inactividad": fila_limpia[i + 9],
    }
    base["Tiempo de Actividad (min)"] = hhmm_a_minutos(base["Tiempo de Actividad"])
    base["Tiempo de Inactividad (min)"] = hhmm_a_minutos(base["Tiempo de Inactividad"])

    # Fecha de generación del reporte: es el primer campo con formato dd/mm/aaaa, hh:mm
    # que aparece después del bloque de causas.
    fecha_reporte = None
    for celda in fila_limpia[i + 19:]:
        if re.match(r'^\d{2}/\d{2}/\d{4},\s*\d{1,2}:\d{2}$', celda):
            fecha_reporte = celda
            break
    base["Fecha Reporte"] = fecha_reporte

    causas = []
    offsets_causa = [(i + 10, i + 11, i + 12), (i + 13, i + 14, i + 15), (i + 16, i + 17, i + 18)]
    for orden, (idx_codigo, idx_cuenta, idx_tiempo) in enumerate(offsets_causa, start=1):
        codigo = fila_limpia[idx_codigo]
        if not codigo or codigo == "N/A":
            continue
        fila_causa = dict(base)
        fila_causa["Orden Causa"] = orden
        fila_causa["Código Causa"] = codigo
        fila_causa["Cuenta Paros"] = fila_limpia[idx_cuenta]
        fila_causa["Tiempo Causa"] = fila_limpia[idx_tiempo]
        fila_causa["Tiempo Causa (min)"] = hhmm_a_minutos(fila_causa["Tiempo Causa"])
        causas.append(fila_causa)

    return causas


# Componente web para cargar el archivo
uploaded_file = st.file_uploader("Por favor, sube tu archivo CSV de improductivos:", type=["csv"])

if uploaded_file is not None:
    raw_data = uploaded_file.read()
    try:
        decoded_file = raw_data.decode('utf-8')
    except UnicodeDecodeError:
        decoded_file = raw_data.decode('latin1')

    st.info("Procesando archivo... Por favor espera.")

    registros = []
    csv_reader = csv.reader(io.StringIO(decoded_file), delimiter=',')
    filas_ignoradas = 0

    for row in csv_reader:
        if not row:
            continue
        causas_fila = procesar_fila(row)
        if causas_fila:
            registros.extend(causas_fila)
        else:
            filas_ignoradas += 1

    columnas_orden = [
        "Fecha Reporte", "Departamento", "Máquina", "Trabajo / Orden", "Número de Parte",
        "Tiempo de Actividad", "Tiempo de Actividad (min)",
        "Tiempo de Inactividad", "Tiempo de Inactividad (min)", "% de Inactividad",
        "Orden Causa", "Código Causa", "Cuenta Paros", "Tiempo Causa", "Tiempo Causa (min)",
    ]

    if not registros:
        st.warning(
            "No se pudo extraer ningún registro. Verifica que el CSV corresponda al "
            "Informe de Tiempo de Inactividad de Trabajo (debe contener la etiqueta "
            f"'{ETIQUETA_ANCLA}' en cada fila)."
        )
    else:
        df_final = pd.DataFrame(registros)[columnas_orden]

        st.success(
            f"¡Extracción completada! Se procesaron {df_final['Trabajo / Orden'].nunique()} "
            f"OTs y {df_final.shape[0]} causas de paro."
        )
        if filas_ignoradas:
            st.caption(f"⚠️ {filas_ignoradas} fila(s) no coincidían con el formato esperado y fueron omitidas.")

        with st.expander("👀 Ver vista previa de los datos extraídos"):
            st.dataframe(df_final)

        st.markdown("### 📊 Top causas de improductividad (minutos totales, este archivo)")
        resumen_causas = (
            df_final.groupby("Código Causa")["Tiempo Causa (min)"]
            .sum()
            .sort_values(ascending=False)
        )
        st.bar_chart(resumen_causas)

        st.markdown("### 📤 Carga a la Nube")
        if st.button("🚀 Actualizar 'Improductivos' en Google Sheets", type="primary"):
            with st.status("Actualizando Google Sheets...", expanded=True) as status:
                try:
                    gc = conectar_sheets()
                    sh = gc.open_by_url(SPREADSHEET_URL)

                    llave_primaria = ["Trabajo / Orden", "Orden Causa"]
                    for col in llave_primaria:
                        df_final[col] = df_final[col].astype(str).str.strip()

                    st.write(f"📝 Actualizando pestaña '{NOMBRE_HOJA_DESTINO}'...")
                    try:
                        worksheet = sh.worksheet(NOMBRE_HOJA_DESTINO)
                        # dtype=str es crítico: sin esto, columnas como 'Trabajo / Orden'
                        # (todo dígitos) se leen como int64/float64, pierden ceros a la
                        # izquierda y dejan de coincidir con el string del CSV nuevo →
                        # en vez de actualizar la fila, la duplica.
                        df_existente = get_as_dataframe(worksheet, dtype=str).dropna(how='all').dropna(axis=1, how='all')
                    except gspread.WorksheetNotFound:
                        worksheet = sh.add_worksheet(
                            title=NOMBRE_HOJA_DESTINO, rows="2000", cols=str(len(df_final.columns))
                        )
                        df_existente = pd.DataFrame()

                    if df_existente.empty or not all(k in df_existente.columns for k in llave_primaria):
                        df_combinado = df_final
                    else:
                        for col in llave_primaria:
                            df_existente[col] = df_existente[col].astype(str).str.strip()
                        df_combinado = pd.concat([df_existente, df_final], ignore_index=True)
                        df_combinado.drop_duplicates(subset=llave_primaria, keep='last', inplace=True)

                    df_combinado = df_combinado[columnas_orden]
                    time.sleep(5)
                    worksheet.clear()
                    set_with_dataframe(worksheet, df_combinado.fillna(""))

                    status.update(label="Proceso completado exitosamente", state="complete", expanded=False)
                    st.balloons()
                    st.success(f"🎉 ¡Pestaña '{NOMBRE_HOJA_DESTINO}' actualizada con éxito en Google Sheets!")

                except Exception as e:
                    status.update(label="Ocurrió un error", state="error")
                    st.error(f"Ocurrió un error durante la actualización: {e}")
