import streamlit as st

st.set_page_config(page_title="Tablero de Operaciones", page_icon="🏭", layout="wide")

st.title("🏭 Panel de Operaciones - Industrias Estra")
st.write("Bienvenido al sistema integrado. Selecciona un módulo para comenzar a trabajar:")

st.markdown("<br>", unsafe_allow_html=True)

# Diseño de 4 columnas (Cambio aquí para acomodar la nueva tarjeta)
col1, col2, col3, col4 = st.columns(4)

with col1:
    with st.container(border=True):
        st.markdown("### ⚙️ Producción")
        st.write("Extracción de OTs, tiempos y paros.")
        if st.button("Abrir Producción", use_container_width=True, type="primary"):
            st.switch_page("pages/1_produccion.py")

with col2:
    with st.container(border=True):
        st.markdown("### ⚖️ Materiales")
        st.write("Sincronización de masas y consumo de kg.")
        if st.button("Abrir Masas", use_container_width=True, type="primary"):
            st.switch_page("pages/2_masas.py")

with col3:
    with st.container(border=True):
        st.markdown("### 🔌 API Kern IoP")
        st.write("Extracción directa de energía desde la API.")
        if st.button("Abrir API Energía", use_container_width=True, type="primary"):
            st.switch_page("pages/4_energia.py")

with col4:
    with st.container(border=True):
        st.markdown("### ⚡ Análisis SEC")
        st.write("Cruce de energía y tiempo de máquina.")
        if st.button("Abrir Analítica", use_container_width=True, type="primary"):
            st.switch_page("pages/3_analitica_sec.py")

st.divider()
