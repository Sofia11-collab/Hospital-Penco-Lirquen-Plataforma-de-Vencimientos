"""
Punto de entrada principal de la aplicación Streamlit.
"""
import streamlit as st
from src.database import init_db, get_connection
from src.ui import render_ui

# Configuración de página
st.set_page_config(
    page_title="Gestión de Vencimientos - Hospital Penco Lirquén",
    page_icon="💊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Inicializar Base de Datos
init_db()

# Estado de Sesión para Autenticación
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False
if "user_info" not in st.session_state:
    st.session_state["user_info"] = None

def login_form():
    st.title("💊 Gestión y Logística de Vencimientos")
    st.subheader("Hospital Penco Lirquén — Iniciar Sesión")
    
    with st.form("form_login"):
        col1, col2 = st.columns(2)
        with col1:
            usuario = st.text_input("Usuario")
        with col2:
            password = st.text_input("Contraseña", type="password")
            
        submitted = st.form_submit_button("Ingresar")
        
        if submitted:
            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute(
                "SELECT usuario, rol, nombre_completo FROM usuarios WHERE usuario=? AND password=? AND estado='Activo'",
                (usuario, password)
            )
            user = cursor.fetchone()
            conn.close()
            
            if user:
                st.session_state["logged_in"] = True
                st.session_state["user_info"] = {
                    "usuario": user["usuario"],
                    "rol": user["rol"],
                    "nombre_completo": user["nombre_completo"]
                }
                st.success("Inicio de sesión exitoso.")
                st.rerun()
            else:
                st.error("Usuario o contraseña incorrectos.")

def main():
    if not st.session_state["logged_in"]:
        login_form()
    else:
        render_ui(st.session_state["user_info"])

if __name__ == "__main__":
    main()