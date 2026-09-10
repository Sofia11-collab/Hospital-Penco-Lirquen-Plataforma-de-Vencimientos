import streamlit as st
import psycopg2
from src.database import get_connection
from src.ui import render_ui

# Configuración principal de la página
st.set_page_config(page_title="Gestión Vencimientos", page_icon="💊", layout="wide")

# Inicializar variables de sesión para el Login
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False
if "user_info" not in st.session_state:
    st.session_state["user_info"] = None

def login():
    st.markdown("<h1 style='text-align: center;'>💊 Plataforma de Gestión de Vencimientos</h1>", unsafe_allow_html=True)
    st.markdown("<h3 style='text-align: center; color: gray;'>Hospital Penco Lirquén</h3>", unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        with st.form("login_form"):
            username = st.text_input("Usuario")
            password = st.text_input("Contraseña", type="password")
            submit = st.form_submit_button("Ingresar", use_container_width=True)

            if submit:
                if not username or not password:
                    st.error("Por favor, ingrese sus credenciales.")
                else:
                    try:
                        conn = get_connection()
                        cursor = conn.cursor()
                        # Buscar usuario en PostgreSQL
                        cursor.execute("SELECT id, usuario, rol, nombre_completo, estado FROM usuarios WHERE usuario=%s AND password=%s", (username, password))
                        user = cursor.fetchone()
                        conn.close()

                        if user:
                            # user[4] corresponde a la columna 'estado'
                            if user[4] == 'Inactivo':
                                st.error("🚫 Esta cuenta está inactiva. Contacte al administrador.")
                            else:
                                st.session_state["logged_in"] = True
                                st.session_state["user_info"] = {
                                    "id": user[0],
                                    "usuario": user[1],
                                    "rol": user[2],
                                    "nombre_completo": user[3]
                                }
                                st.rerun()
                        else:
                            st.error("⚠️ Usuario o contraseña incorrectos")
                    except Exception as e:
                        st.error(f"Error de conexión a la base de datos: {e}")

if not st.session_state["logged_in"]:
    login()
else:
    render_ui(st.session_state["user_info"])
