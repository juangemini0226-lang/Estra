import streamlit as st
import pandas as pd
import io
import re
import gspread
from google.oauth2.service_account import Credentials

# 1. Configuración de la página
st.set_page_config(page_title="Masas | Extractor", page_icon="⚖️", layout="wide")

st.title("⚖️ Masa Transformada - Panel Unificado")
st.write("Carga el CSV de consumos de material, extrae OT/Material/Masa, visualiza el comportamiento y sincroniza con Google Sheets.")

# URL de Google Sheets — apunta al documento y a la pestaña específica de destino.
SHEET_URL = 'https://docs.google.com/spreadsheets/d/1lRg2Fc1pk3HBfXkYwXhWnFlTAGxx9gvoZ4hRnJ1AhXY/edit?gid=2109835940#gid=2109835940'
NOMBRE_HOJA_DESTINO = "Material_Data"
GID_ESPERADO = 2109835940  # extraído de la URL de arriba, para validar que no cambió


@st.cache_resource
def conectar_sheets():
    scope = ["https://www.googleapis.com/auth/spreadsheets"]
    creds_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
    return gspread.authorize(creds)


def conectar_hoja_destino():
    """
    Capa de seguridad: abre el spreadsheet y localiza la pestaña por NOMBRE
    (no por posición como .sheet1, que podía apuntar a la pestaña equivocada
    si 'Material_Data' no era la primera hoja). Además valida que el gid de
    la pestaña encontrada coincida con el gid esperado, para detectar si la
    hoja fue renombrada, movida o recreada.

    Devuelve: (worksheet, spreadsheet_titulo, lista_hojas_disponibles, gid_coincide)
    """
    gc = conectar_sheets()
    spreadsheet = gc.open_by_url(SHEET_URL)
    hojas_disponibles = [ws.title for ws in spreadsheet.worksheets()]

    if NOMBRE_HOJA_DESTINO not in hojas_disponibles:
        raise ValueError(
            f"No se encontró la pestaña '{NOMBRE_HOJA_DESTINO}' en el documento. "
            f"Pestañas disponibles: {', '.join(hojas_disponibles)}"
        )

    worksheet = spreadsheet.worksheet(NOMBRE_HOJA_DESTINO)
    gid_coincide = (worksheet.id == GID_ESPERADO)

    return worksheet, spreadsheet.title, hojas_disponibles, gid_coincide


def extraer_masas(content: bytes) -> pd.DataFrame:
    """
    Misma lógica del notebook original: recorre el CSV crudo buscando el
    patrón fecha (dd/mm/yyyy) que marca el inicio de cada bloque de datos,
    y a partir de ahí lee Material, Máquina, OT y Total.
    """
    df_raw = pd.read_csv(io.BytesIO(content), header=None, on_bad_lines='skip', dtype=str, encoding='latin1')
    datos_extraidos = []

    for index, row in df_raw.iterrows():
        for col_idx, value in enumerate(row):
            val_str = str(value).strip()

            # Si encontramos una fecha dd/mm/yyyy, sabemos que ahí empiezan los datos
            if re.match(r'^\d{2}/\d{2}/\d{4}$', val_str):
                if col_idx >= 2 and col_idx + 3 < len(row):
                    ot_val = str(row[col_idx + 2]).strip()

                    if ot_val.isdigit():
                        raw_material = str(row[col_idx - 2]).strip()
                        id_mat = raw_material.split()[0] if pd.notna(row[col_idx - 2]) else "Desc"
                        nombre_mat = raw_material[len(id_mat):].strip(" -()") if len(raw_material) > len(id_mat) else "Sin nombre"
                        maquina = str(row[col_idx - 1]).strip()
                        fecha = val_str
                        ot = ot_val
                        total_str = str(row[col_idx + 3]).replace(',', '.')

                        try:
                            total_float = float(total_str)
                        except ValueError:
                            total_float = 0.0

                        if id_mat not in ["Totals", "0,00", "0.00"]:
                            datos_extraidos.append({
                                'Fecha': fecha,
                                'Maquina': maquina,
                                'OT': ot,
                                'Id_Material': id_mat,
                                'Nombre_Material': nombre_mat,
                                'Total': total_float
                            })
                break

    if not datos_extraidos:
        return pd.DataFrame()

    df_nuevos = pd.DataFrame(datos_extraidos)
    df_agrupado = df_nuevos.groupby(
        ['Fecha', 'Maquina', 'OT', 'Id_Material', 'Nombre_Material'], dropna=False, as_index=False
    )['Total'].sum()
    return df_agrupado


def obtener_keys_existentes(sheet) -> tuple[set, list]:
    """Lee la hoja y arma el set de llaves ya existentes (OT_Maquina_Fecha_Material)."""
    datos_existentes = sheet.get_all_values()
    keys_existentes = set()

    if not datos_existentes:
        sheet.append_row(['Fecha', 'Maquina', 'OT', 'Id_Material', 'Nombre_Material', 'Total'])
        return keys_existentes, datos_existentes

    headers = [str(h).strip().lower() for h in datos_existentes[0]]
    try:
        idx_fecha = next(i for i, h in enumerate(headers) if 'fecha' in h)
        idx_maq = next(i for i, h in enumerate(headers) if 'maquina' in h or 'máquina' in h)
        idx_ot = next(i for i, h in enumerate(headers) if 'ot' in h or 'numero' in h or 'trabajo' in h)
        idx_id = next(i for i, h in enumerate(headers) if 'id' in h or 'material' in h)

        for row_vals in datos_existentes[1:]:
            if len(row_vals) > max(idx_fecha, idx_maq, idx_ot, idx_id):
                f_ot = str(row_vals[idx_ot]).strip()
                f_maq = str(row_vals[idx_maq]).strip().upper()
                f_fec = str(row_vals[idx_fecha]).strip()
                f_id = str(row_vals[idx_id]).strip().upper()
                keys_existentes.add(f"{f_ot}_{f_maq}_{f_fec}_{f_id}")

    except StopIteration:
        st.write("⚠️ No se reconocieron los encabezados. Se usará el mapeo de columnas por defecto (Fecha, Maquina, OT, Id_Material).")
        for row_vals in datos_existentes[1:]:
            if len(row_vals) >= 4:
                f_ot = str(row_vals[2]).strip()
                f_maq = str(row_vals[1]).strip().upper()
                f_fec = str(row_vals[0]).strip()
                f_id = str(row_vals[3]).strip().upper()
                keys_existentes.add(f"{f_ot}_{f_maq}_{f_fec}_{f_id}")

    return keys_existentes, datos_existentes


# ─────────────────────────────────────────────────────────────
# PASO 1: CARGAR
# ─────────────────────────────────────────────────────────────
st.divider()
st.markdown("### 1️⃣ Cargar archivo")
uploaded_file = st.file_uploader("Sube tu archivo CSV (ej. materialesjunio.csv):", type=['csv'])

# ─────────────────────────────────────────────────────────────
# PASO 2: EXTRAER
# ─────────────────────────────────────────────────────────────
if uploaded_file is not None:
    st.markdown("### 2️⃣ Extraer datos")
    if st.button("🔍 Extraer datos del CSV", type="primary", use_container_width=True):
        with st.status("Extrayendo datos de materiales...", expanded=True) as status:
            st.write("📄 Leyendo archivo...")
            content = uploaded_file.read()

            st.write("🔍 Buscando bloques de datos (fecha, máquina, OT, material, masa)...")
            df_agrupado = extraer_masas(content)

            if df_agrupado.empty:
                status.update(label="No se encontraron datos", state="error")
                st.error("⚠️ No se encontraron datos válidos. Verifica que el archivo corresponde al formato esperado.")
            else:
                st.session_state['df_masa_extraida'] = df_agrupado
                st.session_state['nombre_archivo_masa'] = uploaded_file.name
                status.update(
                    label=f"¡Extracción exitosa! {len(df_agrupado)} registros consolidados.",
                    state="complete", expanded=False
                )

# ─────────────────────────────────────────────────────────────
# PASO 3: VISUALIZAR
# ─────────────────────────────────────────────────────────────
if 'df_masa_extraida' in st.session_state:
    df_masa = st.session_state['df_masa_extraida']

    st.markdown("### 3️⃣ Visualizar")
    st.success(f"✅ {len(df_masa)} registros extraídos de **{st.session_state.get('nombre_archivo_masa', 'archivo')}**, listos para revisar.")

    with st.expander("👀 Vista previa de los datos extraídos", expanded=True):
        st.dataframe(df_masa, use_container_width=True)

    with st.expander("📊 Análisis y Visualización", expanded=True):
        col_a, col_b = st.columns(2)
        with col_a:
            dimension = st.selectbox("Agrupar por:", options=["Máquina", "Material", "Fecha", "OT"])
        with col_b:
            estadistico = st.selectbox(
                "Estadístico:",
                options=["Suma", "Promedio", "Mediana", "Mínimo", "Máximo", "Conteo de registros"]
            )

        col_map = {"Máquina": "Maquina", "Material": "Id_Material", "Fecha": "Fecha", "OT": "OT"}
        func_map = {
            "Suma": "sum", "Promedio": "mean", "Mediana": "median",
            "Mínimo": "min", "Máximo": "max", "Conteo de registros": "count"
        }
        col_dim = col_map[dimension]
        func = func_map[estadistico]

        tabla = df_masa.groupby(col_dim, as_index=False)['Total'].agg(func)
        tabla = tabla.rename(columns={'Total': estadistico})

        st.markdown(f"**{estadistico} de masa (Total) por {dimension}**")

        if dimension == "Fecha":
            tabla_grafico = tabla.copy()
            tabla_grafico['Fecha_dt'] = pd.to_datetime(tabla_grafico['Fecha'], format='%d/%m/%Y', errors='coerce')
            tabla_grafico = tabla_grafico.sort_values('Fecha_dt').set_index('Fecha')
            st.line_chart(tabla_grafico[estadistico])
        else:
            tabla_grafico = tabla.sort_values(estadistico, ascending=False).head(20).set_index(col_dim)
            st.bar_chart(tabla_grafico[estadistico])

        st.markdown("###### Tabla resumen")
        st.dataframe(tabla.sort_values(estadistico, ascending=False), use_container_width=True)

        st.markdown("###### Métricas generales")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Masa total (Total)", f"{df_masa['Total'].sum():,.2f}")
        m2.metric("OTs únicas", f"{df_masa['OT'].nunique()}")
        m3.metric("Materiales únicos", f"{df_masa['Id_Material'].nunique()}")
        m4.metric("Máquinas", f"{df_masa['Maquina'].nunique()}")

    # ─────────────────────────────────────────────────────────
    # PASO 4: CARGAR A GOOGLE SHEETS
    # ─────────────────────────────────────────────────────────
    st.markdown("### 4️⃣ Cargar a Google Sheets")

    col_dl, col_up = st.columns(2)

    with col_dl:
        csv_data = df_masa.to_csv(index=False, encoding='utf-8-sig')
        st.download_button(
            label="⬇️ Descargar como CSV",
            data=csv_data,
            file_name="Masa_Transformada_extraido.csv",
            mime="text/csv",
            use_container_width=True
        )

    with col_up:
        # --- CAPA DE SEGURIDAD: validar destino ANTES de mostrar el botón de subida ---
        try:
            worksheet, titulo_doc, hojas_disponibles, gid_coincide = conectar_hoja_destino()
            destino_valido = True
        except Exception as e:
            worksheet, titulo_doc, hojas_disponibles, gid_coincide = None, None, [], False
            destino_valido = False
            st.error(f"❌ No se pudo validar el destino en Google Sheets: {e}")

        if destino_valido:
            st.markdown("**📍 Destino verificado:**")
            st.write(f"📄 Documento: `{titulo_doc}`")
            st.write(f"📑 Hoja: `{worksheet.title}`")

            if gid_coincide:
                st.success(f"✅ El gid de la pestaña coincide con el esperado ({GID_ESPERADO}).")
            else:
                st.warning(
                    f"⚠️ El gid de la pestaña '{NOMBRE_HOJA_DESTINO}' es {worksheet.id}, "
                    f"distinto al gid esperado ({GID_ESPERADO}). Es posible que la hoja haya sido "
                    f"movida o recreada. El nombre coincide, así que se puede continuar, pero revísalo."
                )

            confirmado = st.checkbox(
                f"Confirmo que quiero subir los datos a la hoja **{NOMBRE_HOJA_DESTINO}** de arriba.",
                value=False
            )

            if st.button(
                "☁️ Sincronizar a Google Sheets (sin duplicar)",
                type="primary", use_container_width=True, disabled=not confirmado
            ):
                with st.spinner("Subiendo registros nuevos..."):
                    try:
                        st.write("☁️ Validando duplicados contra la hoja destino...")
                        keys_existentes, _ = obtener_keys_existentes(worksheet)

                        filas_a_subir = []
                        for _, row in df_masa.iterrows():
                            r_ot = str(row['OT']).strip()
                            r_maq = str(row['Maquina']).strip().upper()
                            r_fec = str(row['Fecha']).strip()
                            r_id = str(row['Id_Material']).strip().upper()

                            llave_nueva = f"{r_ot}_{r_maq}_{r_fec}_{r_id}"

                            if llave_nueva not in keys_existentes:
                                filas_a_subir.append([
                                    row['Fecha'], row['Maquina'], row['OT'],
                                    row['Id_Material'], row['Nombre_Material'], row['Total']
                                ])

                        if filas_a_subir:
                            worksheet.append_rows(filas_a_subir)
                            st.success(f"🎉 ¡Éxito! Se subieron {len(filas_a_subir)} registros nuevos a '{NOMBRE_HOJA_DESTINO}'.")
                            with st.expander("👀 Ver registros subidos"):
                                st.dataframe(
                                    pd.DataFrame(filas_a_subir, columns=['Fecha', 'Maquina', 'OT', 'Id_Material', 'Nombre_Material', 'Total']),
                                    use_container_width=True
                                )
                        else:
                            st.info("👍 Todo al día. La información de este archivo ya se encuentra en Google Sheets.")

                    except Exception as e:
                        st.error(f"❌ Error de conexión con Google Sheets: {e}")
