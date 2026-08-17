"""
Gestión y Logística de Productos Farmacéuticos y Vencimientos
Hospital Penco Lirquén
Punto de entrada de la aplicación Streamlit.
"""

import streamlit as st

from src.catalogo_base import get_catalogo
from src.database import init_db, obtener_registros
from src.ui import render_step1_form
from src.importer import render_carga_masiva

st.set_page_config(
    page_title="Gestión de Vencimientos — Hospital Penco Lirquén",
    page_icon="💊",
    layout="wide",
)

st.title("💊 Gestión y Logística de Productos Farmacéuticos y Vencimientos")
st.caption("Hospital Penco Lirquén")

# Conexión a base de datos (persistente durante la sesión del contenedor)
conn = init_db()

# Catálogo con fallback a modo manual si falla
catalogo = get_catalogo()

tab1, tab2, tab3 = st.tabs(
    ["📋 Informe Bodega (individual)", "📥 Carga masiva", "📊 Registros"]
)

with tab1:
    render_step1_form(catalogo, conn)

with tab2:
    render_carga_masiva(conn)

with tab3:
    st.subheader("📊 Registros ingresados")
    df = obtener_registros(conn)
    if df.empty:
        st.info("Aún no hay registros ingresados.")
    else:
        st.dataframe(df, use_container_width=True)
        st.metric("Total de registros", len(df))