import streamlit as st
import pandas as pd
import csv
import io
import time
import gspread
from google.oauth2.service_account import Credentials
from gspread_dataframe import set_with_dataframe, get_as_dataframe

# Configuración de la página web
st.set_page_config(page_title="Extractor de Producción", page_icon="⚙️", layout="wide")
st.title("⚙️ Extractor de Datos de Producción & Carga a Google Sheets")
st.write("Sube el archivo CSV de producción para procesar los datos y actualizar el tablero maestro.")

# URL del Google Sheet fija
SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1lRg2Fc1pk3HBfXkYwXhWnFlTAGxx9gvoZ4hRnJ1AhXY/edit#gid=0'

# 1. Función para conectar con Google Sheets usando Secrets de Streamlit
def conectar_sheets():
    scope = ["https://www.googleapis.com/auth/spreadsheets"]
    # Se obtienen las credenciales de forma segura desde la configuración de Streamlit
    creds_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
    return gspread.authorize(creds)

# 2. Función de extracción inteligente (Tu lógica original)
def extraer_valor(fila, etiqueta, num_valores=1):
    fila_limpia = [str(celda).strip() for celda in fila]
    try:
        idx = fila_limpia.index(etiqueta)
        if num_valores == 1:
            return fila_limpia[idx + 1]
        else:
            return [fila_limpia[idx + i] for i in range(1, num_valores + 1)]
    except ValueError:
        return None if num_valores == 1 else [None] * num_valores

# Componente web para cargar el archivo
uploaded_file = st.file_uploader("Por favor, sube tu archivo CSV:", type=["csv"])

if uploaded_file is not None:
    # Leer y decodificar el archivo cargado
    raw_data = uploaded_file.read()
    try:
        decoded_file = raw_data.decode('utf-8')
    except UnicodeDecodeError:
        decoded_file = raw_data.decode('latin1')

    st.info("Procesando archivo... Por favor espera.")
    
    # --- PROCESAMIENTO PRINCIPAL ---
    datos_extraidos = []
    csv_reader = csv.reader(io.StringIO(decoded_file), delimiter=',')

    for row in csv_reader:
        if len(row) < 50:
            continue

        fila_dict = {}
        # Datos Generales
        maquina_datos = extraer_valor(row, "Máquina :", 2)
        fila_dict["Máquina"] = maquina_datos[0] if maquina_datos else None
        fila_dict["Descripción Máquina"] = maquina_datos[1] if maquina_datos else None
        fila_dict["Tiempo Empezar"] = extraer_valor(row, "Tiempo Empezar:")
        fila_dict["Tiempo Final"] = extraer_valor(row, "Tiempo Final:")
        fila_dict["Trabajo / Orden"] = extraer_valor(row, "Trabajo :")

        parte_datos = extraer_valor(row, "Número de Parte :", 2)
        fila_dict["Número de Parte"] = parte_datos[0] if parte_datos else None
        fila_dict["Descripción Parte"] = parte_datos[1] if parte_datos else None

        molde_datos = extraer_valor(row, "Instrumento :", 2)
        fila_dict["Molde"] = molde_datos[0] if molde_datos else None
        fila_dict["Descripción Molde"] = molde_datos[1] if molde_datos else None

        # KPIs
        fila_dict["Producción Total"] = extraer_valor(row, "Producción Total :Total Parts:")
        fila_dict["Producción Buena"] = extraer_valor(row, "Producción Buena :")
        fila_dict["Producción de Rechazo"] = extraer_valor(row, "Producción de Rechazo :")
        fila_dict["Producción Empacada"] = extraer_valor(row, "Producción Empacada :", 2)[1] if extraer_valor(row, \"Producción Empacada :\", 2) else None
        fila_dict["Ciclos de Máquina"] = extraer_valor(row, "Ciclos de Máquina :Machine Cycles:")
        ciclos_fuera = extraer_valor(row, "Ciclos Fuera Especificación :", 2)
        fila_dict["Ciclos Fuera de Especificación"] = ciclos_fuera[0] if ciclos_fuera else None
        fila_dict["% Fuera de Especificación"] = extraer_valor(row, "% Fuera de Especificación :")
        fila_dict["Eficiencia de Ciclo (%)"] = extraer_valor(row, "Eficiencia de Ciclo :")
        fila_dict["Eficiencia de % (OEE)"] = extraer_valor(row, "Eficiencia de % :")
        fila_dict["Cavidades Medias"] = extraer_valor(row, "Cavidades Medias :")
        fila_dict["Cavidades Totales"] = extraer_valor(row, "Cavidades Totales :")

        # Tiempos
        fila_dict["Tiempo de Actividad"] = extraer_valor(row, "Tiempo de Actividad :")
        fila_dict["Tiempo de Inactividad"] = extraer_valor(row, "Tiempo de Inactividad :")
        fila_dict["% de Inactividad"] = extraer_valor(row, "% de Inactividad :")
        fila_dict["Velocidad Media"] = extraer_valor(row, "Velocidad Media :")
        fila_dict["Velocidad Uniforme"] = extraer_valor(row, "Velocidad Uniforme :")

        # Paros
        paros = ["DESCONO/PROD", "CAMB_REF", "MTTO", "TALLER M", "PARO POR PROCESO", "PERSONAL", "PLANEACIÓN", "INI_FIN PROCESO"]
        for paro in paros:
            datos_paro = extraer_valor(row, paro, 2)
            fila_dict[f"Paro_{paro}_Tiempo"] = datos_paro[0] if datos_paro else "0:00"
            fila_dict[f"Paro_{paro}_Cuenta"] = datos_paro[1] if datos_paro else "0"

        # Defectos
        defectos = ["S LLENAR", "QUEMADOS", "C.COLOR", "MANCHAS", "VETAS", "BURBUJAS", "DEFORMAC", "MAL TONO", "JASPEO", "REVIENTE", "OTRAS C", "ARRANQUE PROCESO"]
        for defecto in defectos:
            fila_dict[f"Defecto_{defecto}"] = extraer_valor(row, defecto)

        datos_extraidos.append(fila_dict)

    df_final = pd.DataFrame(datos_extraidos)
    
    st.success(f"¡Extracción completada! Se procesaron {df_final.shape[0]} registros de órdenes.")
    st.subheader("Vista previa de los datos extraídos")
    st.dataframe(df_final.head())

    # Botón web para iniciar la actualización en Google Sheets
    if st.button("🚀 Actualizar Base de Datos en Google Sheets"):
        try:
            gc = conectar_sheets()
            sh = gc.open_by_url(SPREADSHEET_URL)
            llave_primaria = 'Trabajo / Orden'
            
            # --- PESTAÑA 1: PRODUCCIÓN DETALLADA ---
            st.write("1. Actualizando 'produccion detallada'...")
            try:
                worksheet = sh.worksheet("produccion detallada")
                df_existente = get_as_dataframe(worksheet).dropna(how='all').dropna(axis=1, how='all')
            except gspread.WorksheetNotFound:
                worksheet = sh.add_worksheet(title="produccion detallada", rows="1000", cols=str(len(df_final.columns)))
                df_existente = pd.DataFrame()

            df_final[llave_primaria] = df_final[llave_primaria].astype(str).str.strip()
            
            if df_existente.empty or llave_primaria not in df_existente.columns:
                df_combinado = df_final
            else:
                df_existente[llave_primaria] = df_existente[llave_primaria].astype(str).str.strip()
                df_combinado = pd.concat([df_existente, df_final], ignore_index=True)
                df_combinado.drop_duplicates(subset=[llave_primaria], keep='last', inplace=True)

            time.sleep(5) # Evitar cuotas de la API
            worksheet.clear()
            set_with_dataframe(worksheet, df_combinado.fillna(""))
            st.toast("Pestaña 'produccion detallada' al día.")

            # --- PESTAÑA 2: PRODUCCIÓN SEC ---
            st.write("2. Actualizando 'produccion SEC'...")
            columnas_sec = ["Máquina", "Tiempo Empezar", "Tiempo Final", "Trabajo / Orden", "Número de Parte", "Molde", "Producción Total"]
            df_sec = df_final[columnas_sec].copy()
            
            # Formateos
            for col in ["Máquina", "Trabajo / Orden", "Número de Parte", "Molde"]:
                df_sec[col] = df_sec[col].astype(str).str.strip().replace('None', '')
            
            for col_fecha in ["Tiempo Empezar", "Tiempo Final"]:
                mask_current = df_sec[col_fecha].astype(str).str.strip().str.lower() == 'current'
                df_sec[col_fecha] = pd.to_datetime(df_sec[col_fecha], format='%d/%m/%Y, %H:%M', errors='coerce')
                df_sec[col_fecha] = df_sec[col_fecha].dt.strftime('%Y-%m-%d %H:%M:%S')
                df_sec.loc[mask_current, col_fecha] = 'En proceso'
                
            df_sec['Producción Total'] = df_sec['Producción Total'].astype(str).str.replace('.', '', regex=False)
            df_sec['Producción Total'] = pd.to_numeric(df_sec['Producción Total'], errors='coerce').astype('Int64')

            try:
                worksheet_sec = sh.worksheet("produccion SEC")
                df_existente_sec = get_as_dataframe(worksheet_sec).dropna(how='all').dropna(axis=1, how='all')
            except gspread.WorksheetNotFound:
                worksheet_sec = sh.add_worksheet(title="produccion SEC", rows="1000", cols=str(len(columnas_sec)))
                df_existente_sec = pd.DataFrame()

            if df_existente_sec.empty or llave_primaria not in df_existente_sec.columns:
                df_combinado_sec = df_sec
            else:
                df_existente_sec[llave_primaria] = df_existente_sec[llave_primaria].astype(str).str.strip()
                df_combinado_sec = pd.concat([df_existente_sec, df_sec], ignore_index=True)
                df_combinado_sec.drop_duplicates(subset=[llave_primaria], keep='last', inplace=True)

            df_combinado_sec = df_combinado_sec[columnas_sec]
            time.sleep(5)
            worksheet_sec.clear()
            df_export = df_combinado_sec.copy()
            df_export['Producción Total'] = df_export['Producción Total'].fillna("")
            set_with_dataframe(worksheet_sec, df_export.fillna(""))
            st.toast("Pestaña 'produccion SEC' al día.")

            # --- PESTAÑA 3: NO CONFORMIDADES ---
            st.write("3. Actualizando 'No_conformidad_produccion'...")
            columnas_req_nc = [
                'Máquina', 'Descripción Máquina', 'Trabajo / Orden', 'Número de Parte',
                'Descripción Parte', 'Molde', 'Producción Total', 'Producción Buena', 'Producción de Rechazo',
                'Defecto_S LLENAR', 'Defecto_QUEMADOS', 'Defecto_C.COLOR', 'Defecto_MANCHAS',
                'Defecto_VETAS', 'Defecto_BURBUJAS', 'Defecto_DEFORMAC', 'Defecto_MAL TONO',
                'Defecto_JASPEO', 'Defecto_REVIENTE', 'Defecto_OTRAS C', 'Defecto_ARRANQUE PROCESO'
            ]
            columnas_validas_nc = [col for col in columnas_req_nc if col in df_combinado.columns]
            df_no_conformidades = df_combinado[columnas_validas_nc].copy()

            try:
                worksheet_destino = sh.worksheet("No_conformidad_produccion")
            except gspread.WorksheetNotFound:
                worksheet_destino = sh.add_worksheet(title="No_conformidad_produccion", rows="1000", cols=str(len(df_no_conformidades.columns)))

            time.sleep(5)
            worksheet_destino.clear()
            set_with_dataframe(worksheet_destino, df_no_conformidades.fillna(""))
            
            st.success("🎉 ¡BASE DE DATOS ACTUALIZADA CON ÉXITO EN GOOGLE SHEETS!")
            
        except Exception as e:
            st.error(f"Ocurrió un error durante la actualización: {e}")
