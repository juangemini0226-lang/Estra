import streamlit as st
import pandas as pd
import io
import re
import gspread
from google.oauth2.service_account import Credentials

# 1. Configuración de la página
st.set_page_config(page_title="Masas | Extractor", page_icon="⚖️", layout="wide")

st.title("⚖️ Carga de Materiales a Google Sheets")
st.write("Sube el archivo CSV de consumos de material para extraer las OTs, Id_Material y sus masas.")

# URL de Google Sheets
SHEET_URL = 'https://docs.google.com/spreadsheets/d/1lRg2Fc1pk3HBfXkYwXhWnFlTAGxx9gvoZ4hRnJ1AhXY/edit#gid=0'

@st.cache_resource
def conectar_sheets():
    scope = ["https://www.googleapis.com/auth/spreadsheets"]
    creds_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
    return gspread.authorize(creds)

# 2. Componente de Carga
uploaded_file = st.file_uploader("Por favor, sube tu archivo CSV (ej. materialesjunio.csv):", type=['csv'])

if uploaded_file is not None:
    if st.button("🚀 Procesar y Subir", type="primary"):
        
        with st.status("Procesando datos de materiales...", expanded=True) as status:
            try:
                st.write("📄 Leyendo archivo...")
                content = uploaded_file.read()
                
                # Leemos el CSV crudo con encoding latin1
                df_raw = pd.read_csv(io.BytesIO(content), header=None, on_bad_lines='skip', dtype=str, encoding='latin1')
                datos_extraidos = []

                st.write("🔍 Extrayendo información...")
                
                # Iterar sobre las filas para encontrar el bloque de datos usando Regex
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
                    status.update(label="No se encontraron datos", state="error")
                    st.error("⚠️ No se encontraron datos válidos. Verifica que el archivo corresponde al formato.")
                    st.stop()

                df_nuevos = pd.DataFrame(datos_extraidos)

                # Agrupar sumando el total
                df_agrupado = df_nuevos.groupby(['Fecha', 'Maquina', 'OT', 'Id_Material', 'Nombre_Material'], dropna=False, as_index=False)['Total'].sum()
                st.write(f"📊 Se consolidaron {len(df_agrupado)} registros únicos del CSV.")

                st.write("☁️ Conectando a Google Sheets para validación de duplicados...")
                gc = conectar_sheets()
                # Conectamos a la primera hoja (Sheet1 o Material_Data)
                sheet = gc.open_by_url(SHEET_URL).sheet1

                datos_existentes = sheet.get_all_values()
                keys_existentes = set()

                if not datos_existentes:
                    sheet.append_row(['Fecha', 'Maquina', 'OT', 'Id_Material', 'Nombre_Material', 'Total'])
                else:
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
                        st.write("⚠️ No se reconocieron los encabezados. Se usará el mapeo de columnas por defecto.")
                        for row_vals in datos_existentes[1:]:
                            if len(row_vals) >= 4:
                                f_ot = str(row_vals[2]).strip()
                                f_maq = str(row_vals[1]).strip().upper()
                                f_fec = str(row_vals[0]).strip()
                                f_id = str(row_vals[3]).strip().upper()
                                keys_existentes.add(f"{f_ot}_{f_maq}_{f_fec}_{f_id}")

                filas_a_subir = []
                st.write("⚖️ Cruzando datos para omitir los ya existentes...")
                
                for _, row in df_agrupado.iterrows():
                    r_ot = str(row['OT']).strip()
                    r_maq = str(row['Maquina']).strip().upper()
                    r_fec = str(row['Fecha']).strip()
                    r_id = str(row['Id_Material']).strip().upper()

                    llave_nueva = f"{r_ot}_{r_maq}_{r_fec}_{r_id}"

                    if llave_nueva not in keys_existentes:
                        filas_a_subir.append([row['Fecha'], row['Maquina'], row['OT'], row['Id_Material'], row['Nombre_Material'], row['Total']])

                if filas_a_subir:
                    st.write(f"🚀 Subiendo {len(filas_a_subir)} registros NUEVOS a la nube...")
                    sheet.append_rows(filas_a_subir)
                    
                    status.update(label="¡Sincronización exitosa!", state="complete", expanded=False)
                    st.success(f"✅ ¡Sincronización exitosa! Se subieron {len(filas_a_subir)} registros.")
                    
                    with st.expander("👀 Ver registros subidos"):
                        st.dataframe(pd.DataFrame(filas_a_subir, columns=['Fecha', 'Maquina', 'OT', 'Id_Material', 'Nombre_Material', 'Total']))
                else:
                    status.update(label="Datos al día", state="complete", expanded=False)
                    st.info("✅ Todo está al día. La información de este archivo ya se encuentra en Google Sheets.")

            except Exception as e:
                status.update(label="Ocurrió un error", state="error")
                st.error(f"❌ Ocurrió un error inesperado: {e}")
