import streamlit as st
import os
import base64
from src.database import init_db, get_connection
from src.ui import render_ui

# Configuración inicial de la página
st.set_page_config(page_title="Gestión Vencimientos", page_icon="💊", layout="wide")

def get_base64_image(image_path):
    """Convierte la imagen a base64 para poder aplicarle CSS"""
    if os.path.exists(image_path):
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    return ""

def authenticate_user(username, password):
    """Verifica las credenciales directamente en la base de datos"""
    conn = get_connection()
    user = conn.cursor().execute(
        "SELECT usuario, rol, nombre_completo FROM usuarios WHERE usuario=? AND password=? AND estado='Activo'", 
        (username, password)
    ).fetchone()
    conn.close()
    
    if user:
        return {
            "usuario": user["usuario"],
            "rol": user["rol"],
            "nombre_completo": user["nombre_completo"]
        }
    return None

def main():
    init_db()

    if "logged_in" not in st.session_state:
        st.session_state["logged_in"] = False

    if not st.session_state["logged_in"]:
        # =======================================================
        # DISEÑO DE LA PANTALLA DE INICIO DE SESIÓN
        # =======================================================
        
        # Reducir el enorme espacio en blanco superior por defecto de Streamlit
        st.markdown("""
            <style>
                .block-container { padding-top: 3rem !important; padding-bottom: 1rem !important; }
                header { visibility: hidden !important; }
            </style>
        """, unsafe_allow_html=True)
        
        # 1. Logo circular (El truco border-radius: 50% corta las esquinas blancas)
        logo_base64 = get_base64_image("assets/hospital-penco-lirquen.png")
        if not logo_base64:
            logo_base64 = get_base64_image("assets/logo.png")
            
        if logo_base64:
            st.markdown(
                f"""
                <div style="display: flex; justify-content: center; margin-bottom: 10px;">
                    <img src="data:image/png;base64,{logo_base64}" style="width: 120px; border-radius: 50%; opacity: 0.9;">
                </div>
                """,
                unsafe_allow_html=True
            )
            
        # 2. Título Centrado y en Mayúsculas (con márgenes reducidos)
        st.markdown("<h1 style='text-align: center; margin-bottom: 0.2rem; font-weight: 800; font-size: 2.2rem;'>GESTIÓN Y LOGÍSTICA DE VENCIMIENTOS</h1>", unsafe_allow_html=True)
        
        # 3. Formulario centrado y más estrecho
        col_vacia1, col_centro, col_vacia2 = st.columns([1, 1.2, 1])
        
        with col_centro:
            st.markdown("<h4 style='text-align: center; color: #888; margin-bottom: 1rem; font-size: 1.1rem;'>Hospital Penco Lirquén — Iniciar Sesión</h4>", unsafe_allow_html=True)
            
            with st.form("login_form"):
                usuario = st.text_input("Usuario")
                password = st.text_input("Contraseña", type="password")
                
                # El botón ahora está inmediatamente debajo de la contraseña
                submit_button = st.form_submit_button("Ingresar", use_container_width=True)
                
                if submit_button:
                    user_info = authenticate_user(usuario, password)
                    if user_info:
                        st.session_state["logged_in"] = True
                        st.session_state["user_info"] = user_info
                        st.rerun()
                    else:
                        st.error("Usuario o contraseña incorrectos. Por favor, intente nuevamente.")
    else:
        # Si ya inició sesión, carga el sistema principal
        render_ui(st.session_state["user_info"])

if __name__ == "__main__":
    main()
