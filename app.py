import streamlit as st

# Configuración global de la aplicación
st.set_page_config(page_title="Tablero de Operaciones", page_icon="🏭", layout="wide")

st.title("🏭 Panel de Operaciones - Industrias Estra")
st.write("Bienvenido al sistema integrado. Selecciona un módulo para comenzar a trabajar:")

st.markdown("<br>", unsafe_allow_html=True)

# Diseño de 3 columnas para las tarjetas
col1, col2, col3 = st.columns(3)

with col1:
    with st.container(border=True):
        st.markdown("### ⚙️ Producción")
        st.write("Extracción de OTs, tiempos, defectos, paros y carga a Google Sheets.")
        st.markdown("<br>", unsafe_allow_html=True)
        # Este botón redirige automáticamente al archivo dentro de la carpeta 'pages'
        if st.button("Abrir Producción", use_container_width=True, type="primary"):
            st.switch_page("pages/1_produccion.py")

with col2:
    with st.container(border=True):
        st.markdown("### ⚖️ Masas y Materiales")
        st.write("Carga de materiales consumidos y sincronización de kilogramos por OT.")
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Abrir Masas", use_container_width=True, type="primary"):
            st.switch_page("pages/2_masas.py")

with col3:
    with st.container(border=True):
        st.markdown("### ⚡ Analítica SEC")
        st.write("Cruce de consumo de energía con producción para indicadores SEC.")
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Abrir Analítica", use_container_width=True, type="primary"):
            st.switch_page("pages/3_analitica_sec.py")

st.divider()
st.caption("Módulos futuros como Mantenimiento (CMMS) o Calidad pueden integrarse fácilmente en este panel.")
