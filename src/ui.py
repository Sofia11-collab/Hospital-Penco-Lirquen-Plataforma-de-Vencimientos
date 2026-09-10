"""
Módulo de Interfaz de Usuario para los 5 Pasos Operativos, Carga Masiva,
Gestión de Usuarios, Dashboard, Consolidado General, Interfaz Segura y Alertas Sanitarias.
"""
import io
import os
import time
import base64
import psycopg2
import streamlit as st
import pandas as pd
import altair as alt
from datetime import datetime, timedelta
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openpyxl.worksheet.datavalidation import DataValidation
from src.database import get_connection
from src.catalogo import (
    get_catalogo, get_opciones_selectbox, buscar_por_etiqueta,
    BODEGAS_OFICIALES, PROVEEDORES_OFICIALES, TIPOS_DOCUMENTO, OPCION_MANUAL
)
from src.importer import procesar_carga_masiva

ESTADOS_TRAMITE_PROVEEDOR = [
    "Correo enviado - En espera de respuesta de proveedor",
    "Correo enviado - Canje aceptado",
    "Correo enviado - Canje rechazado",
    "Correo enviado - En espera de retiro por proveedor",
    "Bulto retirado por proveedor",
    "Canje aceptado - En espera de Nota de credito",
    "Canje aceptado - En espera de reposición del producto",
    "Nota de credito recibida",
    "Producto recibido"
]

BODEGAS_PASO4 = BODEGAS_OFICIALES + ["Bodega de Excluidos"]
OPCIONES_FISICA_P4 = BODEGAS_PASO4 + ["Despachado / Retirado de Bodega"]

OPCIONES_DIFUSION_RED = [
    "No iniciada la difusión",
    "Ofrecido a la Red",
    "Aceptado por Establecimiento",
    "Producto donado"
]

OPCIONES_REDISTRIBUCION_STOCK = [
    "Baja mensual CENABAST",
    "Reprogramación anual CENABAST",
    "Pedido especial CENABAST",
    "Consulta de compra"
]

TEMAS_COLOR_CLAROS = {
    "Claro (Predeterminado)": {"bg": "#f8f9fa", "card": "#ffffff", "text": "#1a1d20", "border": "#dee2e6"},
    "Azul Pastel": {"bg": "#edf2f7", "card": "#e2e8f0", "text": "#0f172a", "border": "#cbd5e1"},
    "Rojo Suave": {"bg": "#fef2f2", "card": "#fee2e2", "text": "#450a0a", "border": "#fca5a5"},
    "Amarillo / Crema": {"bg": "#fefce8", "card": "#fef08a", "text": "#422006", "border": "#fde047"},
    "Violeta Lavanda": {"bg": "#faf5ff", "card": "#f3e8ff", "text": "#3b0764", "border": "#d8b4fe"},
    "Verde Menta": {"bg": "#f0fdf4", "card": "#dcfce7", "text": "#052e16", "border": "#86efac"},
    "Rosado Pastel": {"bg": "#fdf2f8", "card": "#fce7f3", "text": "#500724", "border": "#fbcfe8"}
}

def aplicar_estilo_tema(nombre_tema):
    tema = TEMAS_COLOR_CLAROS.get(nombre_tema, TEMAS_COLOR_CLAROS["Claro (Predeterminado)"])
    css = f"""
    <style>
        [data-testid="stHeader"] {{ display: none !important; }}
        #MainMenu {{ visibility: hidden !important; }}
        footer {{ visibility: hidden !important; }}
        .stDeployButton, .stAppDeployButton, [data-testid="stToolbar"] {{ display: none !important; visibility: hidden !important; }}
        .block-container {{ padding-top: 1.5rem !important; padding-bottom: 1.5rem !important; }}
        .stApp {{ background-color: {tema['bg']} !important; color: {tema['text']} !important; }}
        [data-testid="stSidebar"] > div:first-child {{ padding-top: 1.5rem !important; }}
        [data-testid="stSidebar"] {{ background-color: {tema['card']} !important; border-right: 1px solid {tema['border']} !important; }}
        [data-testid="stSidebar"] *, [data-testid="stSidebar"] label, [data-testid="stSidebar"] p, [data-testid="stSidebar"] span {{ color: {tema['text']} !important; font-weight: 600 !important; }}
        [data-testid="stSidebar"] div[role="radiogroup"] > label {{ padding: 10px 12px; background-color: transparent; border-radius: 8px; margin-bottom: 4px; transition: all 0.2s ease; }}
        [data-testid="stSidebar"] div[role="radiogroup"] > label:hover {{ background-color: rgba(0, 0, 0, 0.04); transform: translateX(4px); }}
        [data-testid="stSidebar"] div[role="radiogroup"] > label > div:first-child {{ transform: scale(0.85); }}
        .stMarkdown, .stText, h1, h2, h3, h4, h5, h6, label, p, span {{ color: {tema['text']} !important; }}
        [data-testid="stMetricValue"] {{ color: {tema['text']} !important; }}
        .stMarkdown p {{ margin-bottom: 0 !important; }}
        code {{ background-color: #dcfce7 !important; color: #15803d !important; border: 1px solid #86efac !important; font-weight: bold !important; padding: 2px 8px !important; border-radius: 4px !important; }}
        
        /* Proteccion contra modo oscuro en cajas de texto y menus */
        div[data-baseweb="select"] > div, .stTextInput div[data-baseweb="input"], .stNumberInput div[data-baseweb="input"], .stDateInput div[data-baseweb="input"], .stTextArea div[data-baseweb="textarea"], div[data-baseweb="popover"] {{ background-color: #ffffff !important; border: 1px solid #cbd5e1 !important; color: #0f172a !important; }}
        div[data-baseweb="select"] * {{ color: #0f172a !important; background-color: transparent !important; }}
        .stTextInput input, .stNumberInput input, .stDateInput input, .stTextArea textarea {{ background-color: #ffffff !important; color: #0f172a !important; -webkit-text-fill-color: #0f172a !important; }}
        ul[data-testid="stSelectboxVirtualDropdown"] {{ background-color: #ffffff !important; border: 1px solid #cbd5e1 !important; }}
        ul[data-testid="stSelectboxVirtualDropdown"] li {{ background-color: #ffffff !important; color: #0f172a !important; }}
        ul[data-testid="stSelectboxVirtualDropdown"] li:hover {{ background-color: #e2e8f0 !important; color: #0f172a !important; }}
        
        .stButton button, .stDownloadButton button, [data-testid="stFileUploader"] button {{ background-color: #ffffff !important; color: #0f172a !important; border: 1px solid #cbd5e1 !important; font-weight: bold !important; box-shadow: 0 1px 2px rgba(0,0,0,0.05) !important; }}
        .stButton button *, .stDownloadButton button *, [data-testid="stFileUploader"] button * {{ color: #0f172a !important; }}
        div[data-testid="stForm"] button {{ background-color: #0284c7 !important; border: 1px solid #0369a1 !important; }}
        div[data-testid="stForm"] button *, div[data-testid="stForm"] button p, div[data-testid="stForm"] button span {{ color: #ffffff !important; font-weight: bold !important; font-size: 15px !important; }}
        div[data-testid="stForm"] button:hover {{ background-color: #0369a1 !important; }}
        [data-testid="stFileUploader"] section {{ background-color: #ffffff !important; border: 2px dashed #cbd5e1 !important; }}
        [data-testid="stFileUploader"] section * {{ color: #334155 !important; }}
        
        /* Ajuste de Logo Centrado Perfecto */
        [data-testid="stSidebar"] [data-testid="stImage"] {{ display: flex !important; justify-content: center !important; width: 100% !important; margin-top: -15px !important; margin-bottom: 20px !important; background: transparent !important; border: none !important; box-shadow: none !important; padding: 0 !important; }}
        [data-testid="stSidebar"] [data-testid="stImage"] > div {{ display: flex !important; justify-content: center !important; width: 100% !important; }}
        [data-testid="stSidebar"] [data-testid="stImage"] img {{ margin: 0 auto !important; display: block !important; max-width: 170px !important; width: 100% !important; height: auto !important; object-fit: contain !important; mix-blend-mode: multiply !important; }}
    </style>
    """
    st.markdown(css, unsafe_allow_html=True)

def calcular_semaforo_vencimiento(fecha_str, motivo):
    if motivo == "Alerta Sanitaria":
        return "🚨 ALERTA (ISP)"
    if not fecha_str or pd.isna(fecha_str): 
        return "⚪ N/A"
    try:
        venc = datetime.strptime(str(fecha_str).strip()[:10], "%Y-%m-%d").date()
        dias = (venc - datetime.now().date()).days
        if dias <= 0:
            return "⚫ Vencido (≤ 0d)"
        elif dias <= 60:
            return "🔴 Roja (1-60d)"
        elif dias <= 120:
            return "🟡 Amarilla (61-120d)"
        else:
            return "🟢 Verde (> 120d)"
    except:
        return "⚪ Error Fecha"

def generar_anexo_ii_docx(datos):
    doc = Document()
    p_membrete = doc.add_paragraph()
    run_mem1 = p_membrete.add_run("DEPARTAMENTO AGENCIA NACIONAL DE MEDICAMENTOS\n")
    run_mem2 = p_membrete.add_run("SUBDEPARTAMENTO DE FISCALIZACIÓN")
    run_mem1.font.size = Pt(8); run_mem1.font.bold = True; run_mem1.font.color.rgb = RGBColor(128, 128, 128)
    run_mem2.font.size = Pt(8); run_mem2.font.bold = True; run_mem2.font.color.rgb = RGBColor(128, 128, 128)
    p_titulo = doc.add_paragraph(); p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_t1 = p_titulo.add_run("ANEXO II\n"); run_t2 = p_titulo.add_run("FORMULARIO DE REPORTE DE EXISTENCIAS DE PRODUCTO RETIRADO DEL MERCADO\n"); run_t3 = p_titulo.add_run("PRODUCTOS FARMACÉUTICOS (ARTS. 60° y 71° 3, D.S. N° 3/2010)\n")
    for r in [run_t1, run_t2, run_t3]: r.font.bold = True; r.font.size = Pt(11); r.font.name = 'Arial'
    p_intro = doc.add_paragraph(); p_intro.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run_intro = p_intro.add_run("Documento emitido por cada establecimiento que ha recibido producto afecto(s) a retiro del mercado, que da cuenta de la cantidad recibida, stock y cantidad distribuida de los lotes en cuestión, que debe ser informado al distribuidor de quien obtuvo el producto, mediante el presente formato o por otro documento que contenga la misma información, para ser reportado finalmente al titular del producto sujeto a retiro del mercado.")
    run_intro.font.italic = True; run_intro.font.size = Pt(11); doc.add_paragraph() 
    p_h1 = doc.add_paragraph(); r_h1 = p_h1.add_run("ANTECEDENTES DEL PRODUCTO EN PROCESO DE RETIRO DEL MERCADO\n"); r_h1.font.bold = True
    r_h1_sub = p_h1.add_run("(Datos aportados por el titular o distribuidor al establecimiento receptor que suscribe la información del presente formulario)")
    r_h1_sub.font.bold = True; r_h1_sub.font.italic = True
    t1 = doc.add_table(rows=5, cols=2); t1.style = 'Table Grid'
    t1.cell(0,0).text = "PRODUCTO"; t1.cell(0,1).text = datos.get('descripcion', '')
    t1.cell(1,0).text = "PRINCIPIO ACTIVO"; t1.cell(1,1).text = datos.get('principio_activo', '')
    t1.cell(2,0).text = "TITULAR"; t1.cell(2,1).text = datos.get('titular', '')
    t1.cell(3,0).text = "N° DE REGISTRO SANITARIO"; t1.cell(3,1).text = datos.get('registro_sanitario', '')
    t1.cell(4,0).text = "NUMERO DE SERIE(S) / LOTE(S)"; t1.cell(4,1).text = datos.get('lote', '')
    doc.add_paragraph() 
    p_h2 = doc.add_paragraph(); r_h2 = p_h2.add_run("ANTECEDENTES DEL ESTABLECIMIENTO RECEPTOR QUE SUSCRIBE LA INFORMACIÓN"); r_h2.font.bold = True
    t2 = doc.add_table(rows=5, cols=2); t2.style = 'Table Grid'
    t2.cell(0,0).text = "NOMBRE y DIRECCIÓN"; t2.cell(0,1).text = "Hospital Penco Lirquén"
    t2.cell(1,0).text = "TIPO DE ESTABLECIMIENTO/LÍNEAS DE ACTIVIDAD"; t2.cell(1,1).text = "Establecimiento de Salud Pública"
    t2.cell(2,0).text = "REPRESENTANTE LEGAL"; t2.cell(2,1).text = datos.get('representante_legal', '')
    t2.cell(3,0).text = "DIRECTOR TÉCNICO/RESPONSABLE TÉCNICO"; t2.cell(3,1).text = datos.get('director_tecnico', '')
    t2.cell(4,0).text = "FECHA DE REPORTE"; t2.cell(4,1).text = datetime.now().strftime('%d-%m-%Y')
    doc.add_paragraph() 
    p_h3 = doc.add_paragraph(); r_h3 = p_h3.add_run("ANTECEDENTES MOVIMIENTO DEL PRODUCTO"); r_h3.font.bold = True
    t3 = doc.add_table(rows=8, cols=2); t3.style = 'Table Grid'
    t3.cell(0,0).text = "PROVEEDOR"; t3.cell(0,1).text = datos.get('proveedor', '')
    t3.cell(1,0).text = "CANTIDAD RECIBIDA EN EL ESTABLECIMIENTO RECEPTOR"; t3.cell(1,1).text = "N/A"
    t3.cell(2,0).text = "CANTIDAD EN EXISTENCIA (STOCK) NO DISTRIBUIDA"; t3.cell(2,1).text = f"{datos.get('cantidad', '')} {datos.get('unidad', '')}"
    t3.cell(3,0).text = "CANTIDAD DISTRIBUIDA A OTROS ESTABLECIMIENTOS RECEPTORES"; t3.cell(3,1).text = "0"
    t3.cell(4,0).text = "REGISTRO DE DISTRIBUCIÓN"; t3.cell(4,1).text = "N/A"
    t3.cell(5,0).text = "OTRAS OBSERVACIONES"; t3.cell(5,1).text = datos.get('observaciones', '')
    t3.cell(6,0).text = "IDENTIFICACIÓN DE NOTIFICANTE y TELÉFONO DIRECTO"; t3.cell(6,1).text = f"{datos.get('director_tecnico', '')} - "
    t3.cell(7,0).text = "FIRMA DE NOTIFICANTE"; t3.cell(7,1).text = "" 
    bio = io.BytesIO(); doc.save(bio)
    return bio.getvalue()

def actualizar_bd_alertas(conn):
    cursor = conn.cursor()
    columnas = ["alerta_numero VARCHAR(255)", "alerta_fecha VARCHAR(255)", "titular_registro VARCHAR(255)", "registro_sanitario VARCHAR(255)", "principio_activo VARCHAR(255)", "archivo_canje TEXT", "nombre_archivo_canje TEXT"]
    for col in columnas:
        try: 
            cursor.execute(f"ALTER TABLE productos ADD COLUMN {col}")
            conn.commit()
        except Exception: 
            conn.rollback() 

def parche_arreglar_carga_masiva(conn):
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE productos 
        SET paso_actual = 2, 
            estado_global = 'En trámite',
            motivo_informe = COALESCE(motivo_informe, 'Gestión pronto vencimiento')
        WHERE paso_actual = 1 OR estado_global IS NULL OR estado_global = ''
    """)
    cursor.execute("UPDATE productos SET cantidad = 0 WHERE cantidad IS NULL")
    conn.commit()

def render_ui(user_info: dict):
    conn = get_connection()
    actualizar_bd_alertas(conn)
    parche_arreglar_carga_masiva(conn)
    
    opciones_temas = list(TEMAS_COLOR_CLAROS.keys())
    tema_actual = st.session_state.get("tema_seleccionado", "Claro (Predeterminado)")
    if tema_actual not in opciones_temas: tema_actual = "Claro (Predeterminado)"
    idx_tema = opciones_temas.index(tema_actual)
    aplicar_estilo_tema(tema_actual)

    cfg_columnas = {
        "ID": st.column_config.NumberColumn("ID", width="small"),
        "Nivel Alerta": st.column_config.TextColumn("Alerta", width="small"),
        "Código": st.column_config.TextColumn("Código", width="small"),
        "Lote": st.column_config.TextColumn("Lote", width="small"),
        "Cant": st.column_config.NumberColumn("Cant.", width="small"),
        "Cant.": st.column_config.NumberColumn("Cant.", width="small"),
        "Compra": st.column_config.TextColumn("Compra", width="small"),
        "Bodega": st.column_config.TextColumn("Bodega", width="medium"),
        "Física": st.column_config.TextColumn("Ubic. Física", width="medium"),
        "Bulto": st.column_config.TextColumn("Bulto", width="small"),
        "Proveedor": st.column_config.TextColumn("Proveedor", width="medium"),
        "Trámite": st.column_config.TextColumn("Trámite", width="medium"),
        "Doc": st.column_config.TextColumn("N° Doc", width="small"),
        "Canje": st.column_config.TextColumn("Canje", width="small"),
        "Vencimiento": st.column_config.TextColumn("Vencimiento", width="small"),
        "Descripción": st.column_config.TextColumn("Descripción del Producto", width="large")
    }

    col_user, col_rol, col_btn = st.columns([5, 4, 2], vertical_alignment="center")
    with col_user: st.markdown(f"👤 **Usuario:** {user_info['nombre_completo']}")
    with col_rol: st.markdown(f"🛡️ **Rol:** `{user_info['rol'].upper()}`")
    with col_btn:
        if st.button("🚪 Cerrar Sesión", use_container_width=True): st.session_state["logged_in"] = False; st.rerun()

    borde_color = TEMAS_COLOR_CLAROS[tema_actual]["border"]
    st.markdown(f"<hr style='margin-top: 0.5rem; margin-bottom: 1rem; border: none; border-top: 1px solid {borde_color};' />", unsafe_allow_html=True)

    # Inserción del logo centrada usando columnas
    path1, path2 = "assets/hospital-penco-lirquen.png", "assets/logo.png"
    img_path = path1 if os.path.exists(path1) else (path2 if os.path.exists(path2) else None)
    if img_path:
        _, col_logo, _ = st.sidebar.columns([1, 6, 1])
        with col_logo:
            st.image(img_path, use_container_width=True)

    rol = user_info['rol']
    
    st.sidebar.markdown("### 📂 SELECCIONE MÓDULO")
    opciones_modulos = ["💊 Gestión de Vencimientos", "🚨 Alertas Sanitarias", "⚙️ Reportes y Adm."]
    modulo_sel = st.sidebar.selectbox("Módulos Principales", opciones_modulos, label_visibility="collapsed")
    st.sidebar.markdown(f"<hr style='margin: 0.5rem 0; border: none; border-top: 1px solid {borde_color};' />", unsafe_allow_html=True)
    st.sidebar.markdown(f"<p style='font-size: 14px; margin-bottom: 8px;'><b>Estás en:</b> {modulo_sel.split(' ', 1)[1]}</p>", unsafe_allow_html=True)
    
    tabs_disponibles = []
    
    if modulo_sel == "💊 Gestión de Vencimientos":
        tabs_disponibles = ["📋 1. Informe Bodega", "⚖️ 2. Canjes (Jefatura)", "🚚 3. Registro/Prov.", "📦 4. Bulto/Ubicación", "📜 5. Resolución/Cierre"]
        if rol in ["admin", "jefatura_admin", "bodega"]: 
            tabs_disponibles.append("📤 Carga Masiva")
            
    elif modulo_sel == "🚨 Alertas Sanitarias":
        tabs_disponibles = ["📋 1. Ingresar Nueva Alerta", "🚨 2. Gestión Anexo II", "📦 3. Bulto/Ubicación (Alerta)"]
            
    elif modulo_sel == "⚙️ Reportes y Adm.":
        tabs_disponibles = ["🔍 Consolidado General", "📊 Dashboard / Análisis"]
        if rol in ["admin", "jefatura_admin"]: 
            tabs_disponibles.append("👥 Gestión de Usuarios")

    tab_seleccionada = st.sidebar.radio("Pasos", tabs_disponibles, label_visibility="collapsed")

    st.sidebar.markdown(f"<hr style='margin: 1.5rem 0 0.5rem 0; border: none; border-top: 1px dashed {borde_color};' />", unsafe_allow_html=True)
    st.sidebar.markdown("<p style='font-size: 13px; margin-bottom: 4px;'>🎨 Tema Visual</p>", unsafe_allow_html=True)
    tema_sel = st.sidebar.selectbox("Tema Visual", opciones_temas, index=idx_tema, label_visibility="collapsed", key="tema_sidebar")
    if tema_sel != tema_actual: st.session_state["tema_seleccionado"] = tema_sel; st.rerun()

    # --- PASO 1 (VENCIMIENTOS Y ALERTAS) ---
    if tab_seleccionada in ["📋 1. Informe Bodega", "📋 1. Ingresar Nueva Alerta"]:
        es_modulo_alerta = (tab_seleccionada == "📋 1. Ingresar Nueva Alerta")
        if es_modulo_alerta: 
            st.markdown("## 🚨 Ingreso Rápido de Alerta Sanitaria")
            st.caption("Los productos pasarán a Cuarentena e irán al Anexo II.")
        else: 
            st.markdown("## 📋 Paso 1 — Informe de Bodega")
            
        catalogo = get_catalogo()
        opciones = get_opciones_selectbox(catalogo)
        
        tab_p1_ingresar, tab_p1_editar = st.tabs(["➕ Nuevo Registro", "✏️ Editar / Modificar Registros Existentes"])

        with tab_p1_ingresar:
            if rol in ["admin", "jefatura_admin", "bodega"]:
                c_busq1, c_busq2 = st.columns(2)
                with c_busq1:
                    bodega = st.selectbox("Bodega/Farmacia Origen *", BODEGAS_OFICIALES, key="b_ingreso")
                    tipo_prod = st.selectbox("Tipo de producto", ["Fármaco", "Insumo"], key="tp_ingreso")
                with c_busq2:
                    prod_sel = st.selectbox("Buscar Código Reyimen / Producto *", opciones, key="busqueda_producto")
                    
                cod_auto, desc_auto, unidad_auto = "", "", ""
                item = buscar_por_etiqueta(catalogo, prod_sel)
                if item: cod_auto = item.get("codigo", ""); desc_auto = item.get("descripcion", ""); unidad_auto = item.get("unidad", "")
                st.divider()

                if es_modulo_alerta:
                    motivo = "Alerta Sanitaria"
                    st.error("🚨 MODO ALERTA SANITARIA: Se han activado los campos de urgencia.")
                else:
                    motivo = "Gestión pronto vencimiento"

                with st.form("form_paso1"):
                    c1, c2 = st.columns(2)
                    codigo = c1.text_input("Código Reyimen *", value=cod_auto)
                    unidad = c2.text_input("Unidad *", value=unidad_auto)
                    descripcion = c1.text_input("Descripción *", value=desc_auto)
                    cantidad = c2.number_input("Cantidad *", min_value=0.0, step=1.0)
                    
                    if es_modulo_alerta:
                        tipo_compra = "No Aplica (Alerta)"
                        alerta_numero = c1.text_input("N° Alerta Sanitaria ISP *")
                        alerta_fecha = c2.date_input("Fecha de Alerta ISP *", format="DD/MM/YYYY")
                        vencimiento = c1.date_input("Fecha de Vencimiento *", format="DD/MM/YYYY")
                        lote = c2.text_input("Lote *")
                    else:
                        if "farmacia" in bodega.lower():
                            tipo_compra = c1.selectbox("Tipo de compra", ["Desconocido (Farmacia)", "CENABAST", "Compra propia"])
                        else:
                            tipo_compra = c1.selectbox("Tipo de compra *", ["CENABAST", "Compra propia"])
                        
                        vencimiento = c2.date_input("Fecha de Vencimiento *", format="DD/MM/YYYY")
                        lote = c1.text_input("Lote *")
                    
                    if st.form_submit_button("Guardar Paso 1"):
                        if not codigo or not descripcion or not lote or (es_modulo_alerta and not alerta_numero):
                            st.error("Complete todos los campos obligatorios (*)")
                        else:
                            conn.cursor().execute("SELECT id FROM productos WHERE codigo_reyimen=%s AND lote=%s AND bodega_origen=%s AND estado_global IN ('En trámite', 'CUARENTENA')", (codigo, lote, bodega))
                            if conn.cursor().fetchone(): st.warning("⚠️ Este producto ya fue ingresado y está activo.")
                            else:
                                estado_inicial = 'CUARENTENA' if es_modulo_alerta else 'En trámite'
                                ub_fisica = "Bodega de Excluidos" if es_modulo_alerta else ""
                                bulto = ""
                                a_num = alerta_numero if es_modulo_alerta else ""
                                a_fec = str(alerta_fecha) if es_modulo_alerta else ""
                                conn.cursor().execute("""
                                INSERT INTO productos (bodega_origen, tipo_producto, codigo_reyimen, descripcion, unidad, cantidad, vencimiento, lote, motivo_informe, tipo_documento, usuario_registro, paso_actual, estado_global, ubicacion_fisica, numero_bulto, alerta_numero, alerta_fecha)
                                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                                """, (bodega, tipo_prod, codigo, descripcion, unidad, cantidad, str(vencimiento), lote, motivo, tipo_compra, user_info['usuario'], 2, estado_inicial, ub_fisica, bulto, a_num, a_fec))
                                conn.commit()
                                st.success("✅ Producto registrado."); time.sleep(1.5); st.rerun()
            else:
                st.info("🔒 **Modo de solo lectura:** Solo los usuarios con rol de Bodega pueden ingresar nuevos registros al sistema.")

        with tab_p1_editar:
            df_p1 = pd.read_sql_query("""SELECT id AS "ID", bodega_origen AS "Bodega", codigo_reyimen AS "Código", descripcion AS "Descripción", lote AS "Lote", cantidad AS "Cant", vencimiento AS "Vencimiento", motivo_informe AS "Motivo", estado_global AS "Estado" FROM productos WHERE paso_actual IN (1, 2) OR estado_global = 'CUARENTENA'""", conn)
            
            if df_p1.empty: 
                st.info("No hay registros modificables.")
            else:
                df_p1["Nivel Alerta"] = df_p1.apply(lambda row: calcular_semaforo_vencimiento(row["Vencimiento"], row["Motivo"]), axis=1)
                df_p1["Vencimiento"] = pd.to_datetime(df_p1["Vencimiento"], errors='coerce').dt.strftime('%d/%m/%Y')
                
                st.markdown("#### 🔍 Filtros de Búsqueda")
                col_f1, col_f2 = st.columns(2)
                filtro_semaforo = col_f1.multiselect("Filtrar por Semáforo de Riesgo", ["⚫ Vencido (≤ 0d)", "🔴 Roja (1-60d)", "🟡 Amarilla (61-120d)", "🟢 Verde (> 120d)", "🚨 ALERTA (ISP)"], placeholder="Seleccione colores...", key="sem_p1_e")
                filtro_codigo = col_f2.text_input("Filtrar por Código Reyimen", placeholder="Ej: 543...", key="cod_p1_e")
                
                df_filtrado = df_p1.copy()
                if filtro_semaforo: df_filtrado = df_filtrado[df_filtrado["Nivel Alerta"].isin(filtro_semaforo)]
                if filtro_codigo: df_filtrado = df_filtrado[df_filtrado["Código"].astype(str).str.contains(filtro_codigo, case=False, na=False)]
                
                st.markdown("<br>", unsafe_allow_html=True)
                if df_filtrado.empty:
                    st.warning("No hay registros que coincidan con los filtros.")
                else:
                    cols_mostrar = ["ID", "Nivel Alerta", "Código", "Descripción", "Bodega", "Lote", "Cant", "Vencimiento"]
                    st.dataframe(df_filtrado[cols_mostrar], hide_index=True, use_container_width=True, height=180, column_config=cfg_columnas)
                    
                    if rol in ["admin", "jefatura_admin", "bodega"]:
                        opciones_id = df_filtrado['ID'].tolist()
                        formato_opciones = {row['ID']: f"{row['Código']} - {row['Descripción']} (Lote: {row['Lote']})" for idx, row in df_filtrado.iterrows()}
                        
                        st.markdown("#### ⚙️ Edición de Registro Seleccionado")
                        id_mod = st.selectbox("Seleccione el registro a modificar:", opciones_id, format_func=lambda x: formato_opciones[x], key="sel_p1_e")
                        
                        cursor_mod = conn.cursor()
                        cursor_mod.execute("SELECT * FROM productos WHERE id=%s", (id_mod,))
                        prod_data = cursor_mod.fetchone()
                        
                        if prod_data:
                            column_names = [desc[0] for desc in cursor_mod.description]
                            prod_data = dict(zip(column_names, prod_data))
                            
                            with st.form("form_editar_p1"):
                                c_e1, c_e2 = st.columns(2)
                                idx_bodega = BODEGAS_OFICIALES.index(prod_data['bodega_origen']) if prod_data['bodega_origen'] in BODEGAS_OFICIALES else 0
                                new_bodega = c_e1.selectbox("Bodega Origen *", BODEGAS_OFICIALES, index=idx_bodega)
                                
                                valor_cantidad = float(prod_data['cantidad']) if prod_data['cantidad'] is not None and str(prod_data['cantidad']).strip() != '' else 0.0
                                new_cantidad = c_e2.number_input("Cantidad *", value=valor_cantidad, min_value=0.0, step=1.0)
                                
                                valor_lote = str(prod_data['lote']) if prod_data['lote'] is not None else ""
                                new_lote = c_e1.text_input("Lote *", value=valor_lote)
                                
                                try: 
                                    if prod_data['vencimiento']:
                                        fecha_init = datetime.strptime(str(prod_data['vencimiento']).strip()[:10], "%Y-%m-%d").date()
                                    else:
                                        fecha_init = datetime.now().date()
                                except: 
                                    fecha_init = datetime.now().date()
                                    
                                new_vencimiento = c_e2.date_input("Fecha de Vencimiento *", value=fecha_init, format="DD/MM/YYYY")
                                b1, b2 = st.columns(2)
                                if b1.form_submit_button("💾 Guardar Cambios"):
                                    conn.cursor().execute("UPDATE productos SET bodega_origen=%s, cantidad=%s, lote=%s, vencimiento=%s WHERE id=%s", (new_bodega, new_cantidad, new_lote, str(new_vencimiento), id_mod))
                                    conn.commit(); st.success("Actualizado."); time.sleep(1); st.rerun()
                                if b2.form_submit_button("🗑️ Eliminar Registro"):
                                    conn.cursor().execute("DELETE FROM productos WHERE id=%s", (id_mod,))
                                    conn.commit(); st.warning("Eliminado."); time.sleep(1); st.rerun()

    # --- PASO 2 ---
    elif tab_seleccionada == "⚖️ 2. Canjes (Jefatura)":
        st.markdown("## ⚖️ Paso 2 — Gestión de Canjes")
        df = pd.read_sql_query("""SELECT id AS "ID", bodega_origen AS "Bodega", codigo_reyimen AS "Código", descripcion AS "Descripción", tipo_documento AS "Compra", cantidad AS "Cant", lote AS "Lote", vencimiento AS "Vencimiento", motivo_informe AS "Motivo" FROM productos WHERE paso_actual = 2 AND estado_global = 'En trámite' AND motivo_informe != 'Alerta Sanitaria'""", conn)
        
        if df.empty: 
            st.info("No hay productos pendientes de canje comercial.")
        else:
            df["Nivel Alerta"] = df.apply(lambda row: calcular_semaforo_vencimiento(row["Vencimiento"], row["Motivo"]), axis=1)
            df["Vencimiento"] = pd.to_datetime(df["Vencimiento"], errors='coerce').dt.strftime('%d/%m/%Y')
            columnas_orden = ["ID", "Nivel Alerta", "Código", "Descripción", "Bodega", "Compra", "Lote", "Cant", "Vencimiento"]
            df = df[columnas_orden]
            
            st.markdown("#### 🔍 Filtros de Búsqueda")
            col_f1, col_f2 = st.columns(2)
            filtro_semaforo = col_f1.multiselect("Filtrar por Semáforo de Riesgo", ["⚫ Vencido (≤ 0d)", "🔴 Roja (1-60d)", "🟡 Amarilla (61-120d)", "🟢 Verde (> 120d)"], placeholder="Seleccione colores...", key="sem_p2")
            filtro_codigo = col_f2.text_input("Filtrar por Código Reyimen", placeholder="Ej: 543...", key="cod_p2")
            
            df_filtrado = df.copy()
            if filtro_semaforo: df_filtrado = df_filtrado[df_filtrado["Nivel Alerta"].isin(filtro_semaforo)]
            if filtro_codigo: df_filtrado = df_filtrado[df_filtrado["Código"].astype(str).str.contains(filtro_codigo, case=False, na=False)]
            
            st.markdown("<br>", unsafe_allow_html=True)
            if df_filtrado.empty:
                st.warning("No hay productos pendientes que coincidan con los filtros seleccionados.")
            else:
                cols_mostrar = ["ID", "Nivel Alerta", "Código", "Descripción", "Bodega", "Compra", "Cant", "Vencimiento"]
                st.dataframe(df_filtrado[cols_mostrar], hide_index=True, use_container_width=True, height=180, column_config=cfg_columnas)
                
                if rol in ["admin", "jefatura_admin", "jefatura"]:
                    opciones_id = df_filtrado['ID'].tolist()
                    formato_opciones = {row['ID']: f"{row['Código']} - {row['Descripción']} (Lote: {row['Lote']})" for idx, row in df_filtrado.iterrows()}
                    
                    st.markdown("#### ⚙️ Decisión de Canje")
                    prod_id = st.selectbox("Seleccione el producto para evaluar canje:", opciones_id, format_func=lambda x: formato_opciones[x], key="sel_p2")
                    
                    with st.form("form_paso2"):
                        aplica_canje = st.selectbox("¿Aplica Canje? *", ["Aplica", "No aplica"])
                        obs_jefatura = st.text_area("Observaciones (Normas del proveedor, exigencias, etc.)")
                        archivo_adjunto = st.file_uploader("📎 Adjuntar Carta de Canje o Respaldo (Opcional)", type=["pdf", "png", "jpg", "jpeg", "docx", "doc"])
                        
                        if st.form_submit_button("Guardar y Avanzar"):
                            siguiente_paso = 4 if aplica_canje == "No aplica" else 3
                            
                            archivo_b64 = ""
                            nombre_archivo = ""
                            if archivo_adjunto is not None:
                                archivo_b64 = base64.b64encode(archivo_adjunto.read()).decode()
                                nombre_archivo = archivo_adjunto.name
                                conn.cursor().execute("UPDATE productos SET estado_canje=%s, observacion_paso2=%s, fecha_paso2=%s, paso_actual=%s, archivo_canje=%s, nombre_archivo_canje=%s WHERE id=%s", (aplica_canje, obs_jefatura, str(datetime.now().date()), siguiente_paso, archivo_b64, nombre_archivo, prod_id))
                            else:
                                conn.cursor().execute("UPDATE productos SET estado_canje=%s, observacion_paso2=%s, fecha_paso2=%s, paso_actual=%s WHERE id=%s", (aplica_canje, obs_jefatura, str(datetime.now().date()), siguiente_paso, prod_id))
                            
                            conn.commit()
                            if siguiente_paso == 4:
                                st.success("Al no aplicar canje, el producto avanzó directamente al Paso 4.")
                            else:
                                st.success("Avanzado al Paso 3 con éxito.")
                            time.sleep(2)
                            st.rerun()
                else:
                    st.info("🔒 **Modo de solo lectura:** Tu rol de usuario te permite visualizar estos datos para trazabilidad, pero no tienes permisos para gestionar decisiones de canje.")

    # --- PASO 3 ---
    elif tab_seleccionada == "🚚 3. Registro/Prov.":
        st.markdown("## 🚚 Paso 3 — Registro y Proveedor")
        tab_p3_nuevo, tab_p3_seguimiento = st.tabs(["➕ Pendientes de Ingreso", "🔄 Seguimiento de Trámites"])
        
        with tab_p3_nuevo:
            df_p3_nuevo = pd.read_sql_query("""SELECT id AS "ID", codigo_reyimen AS "Código", descripcion AS "Descripción", cantidad AS "Cant", lote AS "Lote", vencimiento AS "Vencimiento", motivo_informe AS "Motivo", estado_canje AS "Canje", observacion_paso2 AS "Obs_P2" FROM productos WHERE paso_actual = 3 AND estado_global = 'En trámite'""", conn)
            
            if df_p3_nuevo.empty: 
                st.info("No hay nuevos productos.")
            else:
                df_p3_nuevo["Nivel Alerta"] = df_p3_nuevo.apply(lambda row: calcular_semaforo_vencimiento(row["Vencimiento"], row["Motivo"]), axis=1)
                df_p3_nuevo["Vencimiento"] = pd.to_datetime(df_p3_nuevo["Vencimiento"], errors='coerce').dt.strftime('%d/%m/%Y')
                
                st.markdown("#### 🔍 Filtros de Búsqueda")
                col_f1, col_f2 = st.columns(2)
                filtro_semaforo = col_f1.multiselect("Filtrar por Semáforo de Riesgo", ["⚫ Vencido (≤ 0d)", "🔴 Roja (1-60d)", "🟡 Amarilla (61-120d)", "🟢 Verde (> 120d)"], placeholder="Seleccione colores...", key="sem_p3_n")
                filtro_codigo = col_f2.text_input("Filtrar por Código Reyimen", placeholder="Ej: 543...", key="cod_p3_n")
                
                df_filtrado = df_p3_nuevo.copy()
                if filtro_semaforo: df_filtrado = df_filtrado[df_filtrado["Nivel Alerta"].isin(filtro_semaforo)]
                if filtro_codigo: df_filtrado = df_filtrado[df_filtrado["Código"].astype(str).str.contains(filtro_codigo, case=False, na=False)]
                
                st.markdown("<br>", unsafe_allow_html=True)
                if df_filtrado.empty:
                    st.warning("No hay productos pendientes que coincidan con los filtros.")
                else:
                    cols_mostrar = ["ID", "Nivel Alerta", "Código", "Descripción", "Canje", "Lote", "Cant", "Vencimiento"]
                    st.dataframe(df_filtrado[cols_mostrar], hide_index=True, use_container_width=True, height=180, column_config=cfg_columnas)
                    
                    if rol in ["admin", "jefatura_admin", "registro"]:
                        opciones_id = df_filtrado['ID'].tolist()
                        formato_opciones = {row['ID']: f"{row['Código']} - {row['Descripción']} (Lote: {row['Lote']})" for idx, row in df_filtrado.iterrows()}
                        
                        st.markdown("#### ⚙️ Ingreso de Trámite Comercial")
                        prod_id = st.selectbox("Seleccione el producto para ingresar trámite:", opciones_id, format_func=lambda x: formato_opciones[x], key="sel_p3_n")
                        
                        obs_previa = df_filtrado.loc[df_filtrado['ID'] == prod_id, 'Obs_P2'].values[0]
                        if obs_previa: st.info(f"📝 **Instrucciones de Jefatura:** {obs_previa}")
                        
                        cursor_p3 = conn.cursor()
                        cursor_p3.execute("SELECT archivo_canje, nombre_archivo_canje FROM productos WHERE id=%s", (prod_id,))
                        prod_data_file = cursor_p3.fetchone()
                        
                        if prod_data_file and prod_data_file[0]:
                            st.info("📎 Jefatura ha adjuntado un documento de respaldo para este canje:")
                            st.download_button("Descargar Respaldo de Jefatura", base64.b64decode(prod_data_file[0]), file_name=prod_data_file[1])
                        
                        with st.form("form_paso3_ingreso"):
                            proveedor, tipo_doc = st.selectbox("Proveedor *", PROVEEDORES_OFICIALES), st.selectbox("Tipo Doc *", TIPOS_DOCUMENTO)
                            num_doc, tramite = st.text_input("N° Documento / OC *"), st.selectbox("Estado del trámite *", ESTADOS_TRAMITE_PROVEEDOR)
                            obs = st.text_area("Observaciones del Área de Registro")
                            if st.form_submit_button("Avanzar a Paso 4"):
                                conn.cursor().execute("UPDATE productos SET proveedor=%s, tipo_documento=%s, numero_documento_oc=%s, tramite_proveedor=%s, fecha_paso3=%s, observacion_paso3=%s, paso_actual=4 WHERE id=%s", (proveedor, tipo_doc, num_doc, tramite, str(datetime.now().date()), obs, prod_id))
                                conn.commit(); st.success("Avanzado."); time.sleep(1.5); st.rerun()
                    else:
                        st.info("🔒 **Modo de solo lectura:** Solo el área de Registro y Abastecimiento puede ingresar trámites comerciales.")

        with tab_p3_seguimiento:
            df_p3_seg = pd.read_sql_query("""SELECT id AS "ID", codigo_reyimen AS "Código", descripcion AS "Descripción", cantidad AS "Cant", lote AS "Lote", vencimiento AS "Vencimiento", motivo_informe AS "Motivo", proveedor AS "Proveedor", tramite_proveedor AS "Trámite", numero_documento_oc AS "Doc" FROM productos WHERE paso_actual >= 3 AND estado_global = 'En trámite' AND motivo_informe != 'Alerta Sanitaria'""", conn)
            
            if df_p3_seg.empty: 
                st.info("No hay trámites en seguimiento.")
            else:
                df_p3_seg["Nivel Alerta"] = df_p3_seg.apply(lambda row: calcular_semaforo_vencimiento(row["Vencimiento"], row["Motivo"]), axis=1)
                df_p3_seg["Vencimiento"] = pd.to_datetime(df_p3_seg["Vencimiento"], errors='coerce').dt.strftime('%d/%m/%Y')
                
                st.markdown("#### 🔍 Filtros de Búsqueda")
                col_f1, col_f2 = st.columns(2)
                filtro_semaforo_seg = col_f1.multiselect("Filtrar por Semáforo de Riesgo", ["⚫ Vencido (≤ 0d)", "🔴 Roja (1-60d)", "🟡 Amarilla (61-120d)", "🟢 Verde (> 120d)"], placeholder="Seleccione colores...", key="sem_p3_s")
                filtro_codigo_seg = col_f2.text_input("Filtrar por Código Reyimen", placeholder="Ej: 543...", key="cod_p3_s")
                
                df_filtrado_seg = df_p3_seg.copy()
                if filtro_semaforo_seg: df_filtrado_seg = df_filtrado_seg[df_filtrado_seg["Nivel Alerta"].isin(filtro_semaforo_seg)]
                if filtro_codigo_seg: df_filtrado_seg = df_filtrado_seg[df_filtrado_seg["Código"].astype(str).str.contains(filtro_codigo_seg, case=False, na=False)]
                
                st.markdown("<br>", unsafe_allow_html=True)
                if df_filtrado_seg.empty:
                    st.warning("No hay trámites que coincidan con los filtros.")
                else:
                    cols_mostrar_seg = ["ID", "Nivel Alerta", "Código", "Descripción", "Proveedor", "Trámite", "Doc", "Cant", "Vencimiento"]
                    st.dataframe(df_filtrado_seg[cols_mostrar_seg], hide_index=True, use_container_width=True, height=180, column_config=cfg_columnas)
                    
                    if rol in ["admin", "jefatura_admin", "registro"]:
                        opciones_id_seg = df_filtrado_seg['ID'].tolist()
                        formato_opciones_seg = {row['ID']: f"{row['Código']} - {row['Descripción']} (Lote: {row['Lote']})" for idx, row in df_filtrado_seg.iterrows()}
                        
                        st.markdown("#### ⚙️ Actualización de Estado")
                        id_seg = st.selectbox("Seleccione el trámite a actualizar:", opciones_id_seg, format_func=lambda x: formato_opciones_seg[x], key="sel_p3_s")
                        
                        cursor_p3_seg = conn.cursor()
                        cursor_p3_seg.execute("SELECT archivo_canje, nombre_archivo_canje FROM productos WHERE id=%s", (id_seg,))
                        prod_seg_file = cursor_p3_seg.fetchone()
                        
                        if prod_seg_file and prod_seg_file[0]:
                            st.info("📎 Jefatura ha adjuntado un documento de respaldo para este canje:")
                            st.download_button("Descargar Respaldo de Jefatura", base64.b64decode(prod_seg_file[0]), file_name=prod_seg_file[1], key="dl_seg_p3")

                        cursor_p3_upd = conn.cursor()
                        cursor_p3_upd.execute("SELECT * FROM productos WHERE id=%s", (id_seg,))
                        prod_seg = cursor_p3_upd.fetchone()
                        
                        if prod_seg:
                            col_names = [desc[0] for desc in cursor_p3_upd.description]
                            prod_seg = dict(zip(col_names, prod_seg))
                            
                            with st.form("form_paso3_actualizar"):
                                u_prov = st.selectbox("Proveedor", PROVEEDORES_OFICIALES, index=PROVEEDORES_OFICIALES.index(prod_seg['proveedor']) if prod_seg['proveedor'] in PROVEEDORES_OFICIALES else 0)
                                u_doc = st.text_input("N° Doc", value=prod_seg['numero_documento_oc'] or "")
                                u_tram = st.selectbox("Estado", ESTADOS_TRAMITE_PROVEEDOR, index=ESTADOS_TRAMITE_PROVEEDOR.index(prod_seg['tramite_proveedor']) if prod_seg['tramite_proveedor'] in ESTADOS_TRAMITE_PROVEEDOR else 0)
                                u_obs = st.text_area("Obs.", value=prod_seg['observacion_paso3'] or "")
                                if st.form_submit_button("Guardar Cambios"):
                                    conn.cursor().execute("UPDATE productos SET proveedor=%s, numero_documento_oc=%s, tramite_proveedor=%s, observacion_paso3=%s WHERE id=%s", (u_prov, u_doc, u_tram, u_obs, id_seg))
                                    conn.commit(); st.success("Actualizado."); time.sleep(1); st.rerun()

    # --- PASO 4 ---
    elif tab_seleccionada == "📦 4. Bulto/Ubicación":
        st.markdown("## 📦 Paso 4 — Bulto y Ubicaciones")
        tab_p4_nuevo, tab_p4_seg = st.tabs(["➕ Asignar Nueva Ubicación", "🔄 Seguimiento de Bultos"])
        
        with tab_p4_nuevo:
            df = pd.read_sql_query("""SELECT id AS "ID", codigo_reyimen AS "Código", descripcion AS "Descripción", cantidad AS "Cant", lote AS "Lote", proveedor AS "Proveedor", vencimiento AS "Vencimiento", motivo_informe AS "Motivo" FROM productos WHERE paso_actual = 4 AND estado_global = 'En trámite' AND motivo_informe != 'Alerta Sanitaria'""", conn)
            
            if df.empty: 
                st.info("No hay productos pendientes de asignación de bulto.")
            else:
                df["Nivel Alerta"] = df.apply(lambda row: calcular_semaforo_vencimiento(row["Vencimiento"], row["Motivo"]), axis=1)
                df["Vencimiento"] = pd.to_datetime(df["Vencimiento"], errors='coerce').dt.strftime('%d/%m/%Y')
                
                st.markdown("#### 🔍 Filtros de Búsqueda")
                col_f1, col_f2 = st.columns(2)
                filtro_semaforo = col_f1.multiselect("Filtrar por Semáforo de Riesgo", ["⚫ Vencido (≤ 0d)", "🔴 Roja (1-60d)", "🟡 Amarilla (61-120d)", "🟢 Verde (> 120d)"], placeholder="Seleccione colores...", key="sem_p4_n")
                filtro_codigo = col_f2.text_input("Filtrar por Código Reyimen", placeholder="Ej: 543...", key="cod_p4_n")
                
                df_filtrado = df.copy()
                if filtro_semaforo: df_filtrado = df_filtrado[df_filtrado["Nivel Alerta"].isin(filtro_semaforo)]
                if filtro_codigo: df_filtrado = df_filtrado[df_filtrado["Código"].astype(str).str.contains(filtro_codigo, case=False, na=False)]
                
                st.markdown("<br>", unsafe_allow_html=True)
                if df_filtrado.empty:
                    st.warning("No hay productos que coincidan con los filtros.")
                else:
                    cols_mostrar = ["ID", "Nivel Alerta", "Código", "Descripción", "Lote", "Proveedor", "Cant", "Vencimiento"]
                    st.dataframe(df_filtrado[cols_mostrar], hide_index=True, use_container_width=True, height=180, column_config=cfg_columnas)
                    
                    if rol in ["admin", "jefatura_admin", "bodega"]:
                        opciones_id = df_filtrado['ID'].tolist()
                        formato_opciones = {row['ID']: f"{row['Código']} - {row['Descripción']} (Lote: {row['Lote']})" for idx, row in df_filtrado.iterrows()}
                        
                        st.markdown("#### ⚙️ Asignación de Ubicación Física")
                        prod_id = st.selectbox("Seleccione el producto a gestionar:", opciones_id, format_func=lambda x: formato_opciones[x], key="sel_p4_n")
                        
                        with st.form("form_p4"):
                            ub_fisica, ub_comp = st.selectbox("Ubicación Física *", OPCIONES_FISICA_P4), st.selectbox("Ubicación Computacional *", BODEGAS_PASO4)
                            bulto, obs = st.text_input("N° Bulto *"), st.text_area("Observaciones Paso 4")
                            if st.form_submit_button("Avanzar a Paso 5"):
                                conn.cursor().execute("UPDATE productos SET ubicacion_fisica=%s, ubicacion_computacional=%s, numero_bulto=%s, observacion_paso4=%s, paso_actual=5 WHERE id=%s", (ub_fisica, ub_comp, bulto, obs, prod_id))
                                conn.commit(); st.success("Avanzado a Paso 5."); time.sleep(1); st.rerun()
                    else:
                        st.info("🔒 **Modo de solo lectura:** Solo el área de Bodega tiene permisos para asignar ubicaciones físicas y computacionales.")
                        
        with tab_p4_seg:
            df_seg4 = pd.read_sql_query("""SELECT id AS "ID", codigo_reyimen AS "Código", descripcion AS "Descripción", cantidad AS "Cant", lote AS "Lote", ubicacion_fisica AS "Física", numero_bulto AS "Bulto", vencimiento AS "Vencimiento", motivo_informe AS "Motivo" FROM productos WHERE paso_actual >= 4 AND estado_global = 'En trámite' AND motivo_informe != 'Alerta Sanitaria'""", conn)
            
            if df_seg4.empty: 
                st.info("No hay bultos activos en seguimiento.")
            else:
                df_seg4["Nivel Alerta"] = df_seg4.apply(lambda row: calcular_semaforo_vencimiento(row["Vencimiento"], row["Motivo"]), axis=1)
                df_seg4["Vencimiento"] = pd.to_datetime(df_seg4["Vencimiento"], errors='coerce').dt.strftime('%d/%m/%Y')
                
                st.markdown("#### 🔍 Filtros de Búsqueda")
                col_f1, col_f2 = st.columns(2)
                filtro_semaforo_seg = col_f1.multiselect("Filtrar por Semáforo de Riesgo", ["⚫ Vencido (≤ 0d)", "🔴 Roja (1-60d)", "🟡 Amarilla (61-120d)", "🟢 Verde (> 120d)"], placeholder="Seleccione colores...", key="sem_p4_s")
                filtro_codigo_seg = col_f2.text_input("Filtrar por Código Reyimen", placeholder="Ej: 543...", key="cod_p4_s")
                
                df_filtrado_seg = df_seg4.copy()
                if filtro_semaforo_seg: df_filtrado_seg = df_filtrado_seg[df_filtrado_seg["Nivel Alerta"].isin(filtro_semaforo_seg)]
                if filtro_codigo_seg: df_filtrado_seg = df_filtrado_seg[df_filtrado_seg["Código"].astype(str).str.contains(filtro_codigo_seg, case=False, na=False)]
                
                st.markdown("<br>", unsafe_allow_html=True)
                if df_filtrado_seg.empty:
                    st.warning("No hay bultos que coincidan con los filtros.")
                else:
                    cols_mostrar_seg = ["ID", "Nivel Alerta", "Código", "Descripción", "Física", "Bulto", "Cant", "Vencimiento"]
                    st.dataframe(df_filtrado_seg[cols_mostrar_seg], hide_index=True, use_container_width=True, height=180, column_config=cfg_columnas)
                    
                    if rol in ["admin", "jefatura_admin", "bodega"]:
                        opciones_id_seg = df_filtrado_seg['ID'].tolist()
                        formato_opciones_seg = {row['ID']: f"{row['Código']} - {row['Descripción']} (Lote: {row['Lote']})" for idx, row in df_filtrado_seg.iterrows()}
                        
                        st.markdown("#### ⚙️ Modificación de Bulto Existente")
                        id_seg4 = st.selectbox("Seleccione el bulto a actualizar:", opciones_id_seg, format_func=lambda x: formato_opciones_seg[x], key="sel_p4_s")
                        
                        cursor_p4 = conn.cursor()
                        cursor_p4.execute("SELECT * FROM productos WHERE id=%s", (id_seg4,))
                        prod_seg4 = cursor_p4.fetchone()
                        
                        if prod_seg4:
                            col_names = [desc[0] for desc in cursor_p4.description]
                            prod_seg4 = dict(zip(col_names, prod_seg4))
                            
                            with st.form("form_p4_actualizar"):
                                idx_fis = OPCIONES_FISICA_P4.index(prod_seg4['ubicacion_fisica']) if prod_seg4['ubicacion_fisica'] in OPCIONES_FISICA_P4 else 0
                                idx_comp = BODEGAS_PASO4.index(prod_seg4['ubicacion_computacional']) if prod_seg4['ubicacion_computacional'] in BODEGAS_PASO4 else 0
                                
                                u_fisica = st.selectbox("Ubicación Física", OPCIONES_FISICA_P4, index=idx_fis)
                                u_comp = st.selectbox("Ubicación Computacional", BODEGAS_PASO4, index=idx_comp)
                                u_bulto = st.text_input("N° Bulto", value=prod_seg4['numero_bulto'] or "")
                                u_obs4 = st.text_area("Observaciones", value=prod_seg4['observacion_paso4'] or "")
                                
                                if st.form_submit_button("Actualizar Bulto"):
                                    conn.cursor().execute("UPDATE productos SET ubicacion_fisica=%s, ubicacion_computacional=%s, numero_bulto=%s, observacion_paso4=%s WHERE id=%s", (u_fisica, u_comp, u_bulto, u_obs4, id_seg4))
                                    conn.commit(); st.success("Ubicación actualizada."); time.sleep(1); st.rerun()

    # --- PASO 5 ---
    elif tab_seleccionada == "📜 5. Resolución/Cierre":
        st.markdown("## 📜 Paso 5 — Resolución y Cierre")
        tab_p5_cierre, tab_p5_sin_canje = st.tabs(["🔒 Cierre con Carta de Canje", "🟢 Cierre sin Carta de Canje"])
        
        with tab_p5_cierre:
            # Solo productos CON canje ("Aplica")
            df_p5 = pd.read_sql_query("""SELECT id AS "ID", codigo_reyimen AS "Código", descripcion AS "Descripción", cantidad AS "Cant", lote AS "Lote", proveedor AS "Proveedor", numero_bulto AS "Bulto", estado_global AS "Estado", vencimiento AS "Vencimiento", motivo_informe AS "Motivo" FROM productos WHERE paso_actual = 5 AND estado_global != 'Concluido' AND estado_canje = 'Aplica' AND motivo_informe != 'Alerta Sanitaria'""", conn)
            
            if df_p5.empty: 
                st.info("No hay productos con canje pendientes de cierre.")
            else:
                df_p5["Nivel Alerta"] = df_p5.apply(lambda row: calcular_semaforo_vencimiento(row["Vencimiento"], row["Motivo"]), axis=1)
                df_p5["Vencimiento"] = pd.to_datetime(df_p5["Vencimiento"], errors='coerce').dt.strftime('%d/%m/%Y')
                
                st.markdown("#### 🔍 Filtros de Búsqueda")
                col_f1, col_f2 = st.columns(2)
                filtro_semaforo = col_f1.multiselect("Filtrar por Semáforo de Riesgo", ["⚫ Vencido (≤ 0d)", "🔴 Roja (1-60d)", "🟡 Amarilla (61-120d)", "🟢 Verde (> 120d)"], placeholder="Seleccione colores...", key="sem_p5_c")
                filtro_codigo = col_f2.text_input("Filtrar por Código Reyimen", placeholder="Ej: 543...", key="cod_p5_c")
                
                df_filtrado = df_p5.copy()
                if filtro_semaforo: df_filtrado = df_filtrado[df_filtrado["Nivel Alerta"].isin(filtro_semaforo)]
                if filtro_codigo: df_filtrado = df_filtrado[df_filtrado["Código"].astype(str).str.contains(filtro_codigo, case=False, na=False)]
                
                st.markdown("<br>", unsafe_allow_html=True)
                if df_filtrado.empty:
                    st.warning("No hay productos que coincidan con los filtros.")
                else:
                    cols_mostrar = ["ID", "Nivel Alerta", "Código", "Descripción", "Lote", "Proveedor", "Bulto", "Cant", "Vencimiento"]
                    st.dataframe(df_filtrado[cols_mostrar], hide_index=True, use_container_width=True, height=180, column_config=cfg_columnas)
                    
                    if rol in ["admin", "jefatura_admin", "jefatura"]:
                        opciones_id = df_filtrado['ID'].tolist()
                        formato_opciones = {row['ID']: f"{row['Código']} - {row['Descripción']} (Lote: {row['Lote']})" for idx, row in df_filtrado.iterrows()}
                        
                        st.markdown("#### ⚙️ Cierre Definitivo de Trámite")
                        prod_id = st.selectbox("Seleccione el producto a CERRAR:", opciones_id, format_func=lambda x: formato_opciones[x], key="sel_p5_c")
                        
                        with st.form("form_p5"):
                            num_res = st.text_input("N° Resolución o Documento *")
                            estado_fin = st.selectbox("Estado Final *", [
                                "Canjeado - reposición del producto", 
                                "Canjeado - nota de credito recibida", 
                                "Producto no canjeado"
                            ])
                            obs = st.text_area("Resolución / Comentarios finales")
                            if st.form_submit_button("Finalizar y Archivar"):
                                conn.cursor().execute("UPDATE productos SET resolucion_numero=%s, estado_final=%s, observacion_paso5=%s, estado_global='Concluido' WHERE id=%s", (num_res, estado_fin, obs, prod_id))
                                conn.commit(); st.success("Archivado."); time.sleep(1.5); st.rerun()
                    else:
                        st.info("🔒 **Modo de solo lectura:** Solo Jefatura puede aplicar el cierre definitivo de los trámites.")
                            
        with tab_p5_sin_canje:
            # Solo productos SIN canje ("No aplica" o compras propias)
            df_p5_sc = pd.read_sql_query("""SELECT id AS "ID", codigo_reyimen AS "Código", descripcion AS "Descripción", cantidad AS "Cant", lote AS "Lote", estado_canje AS "Canje", numero_bulto AS "Bulto", vencimiento AS "Vencimiento", motivo_informe AS "Motivo" FROM productos WHERE paso_actual = 5 AND estado_global != 'Concluido' AND (estado_canje = 'No aplica' OR estado_canje IS NULL OR estado_canje != 'Aplica') AND motivo_informe != 'Alerta Sanitaria'""", conn)
            
            if df_p5_sc.empty:
                st.info("No hay productos sin canje pendientes de cierre.")
            else:
                df_p5_sc["Nivel Alerta"] = df_p5_sc.apply(lambda row: calcular_semaforo_vencimiento(row["Vencimiento"], row["Motivo"]), axis=1)
                df_p5_sc["Vencimiento"] = pd.to_datetime(df_p5_sc["Vencimiento"], errors='coerce').dt.strftime('%d/%m/%Y')
                
                st.markdown("#### 🔍 Filtros de Búsqueda")
                col_f1, col_f2 = st.columns(2)
                filtro_semaforo_sc = col_f1.multiselect("Filtrar por Semáforo de Riesgo", ["⚫ Vencido (≤ 0d)", "🔴 Roja (1-60d)", "🟡 Amarilla (61-120d)", "🟢 Verde (> 120d)"], placeholder="Seleccione colores...", key="sem_p5_sc")
                filtro_codigo_sc = col_f2.text_input("Filtrar por Código Reyimen", placeholder="Ej: 543...", key="cod_p5_sc")
                
                df_filtrado_sc = df_p5_sc.copy()
                if filtro_semaforo_sc: df_filtrado_sc = df_filtrado_sc[df_filtrado_sc["Nivel Alerta"].isin(filtro_semaforo_sc)]
                if filtro_codigo_sc: df_filtrado_sc = df_filtrado_sc[df_filtrado_sc["Código"].astype(str).str.contains(filtro_codigo_sc, case=False, na=False)]
                
                st.markdown("<br>", unsafe_allow_html=True)
                if df_filtrado_sc.empty:
                    st.warning("No hay productos sin canje que coincidan con los filtros.")
                else:
                    cols_mostrar_sc = ["ID", "Nivel Alerta", "Código", "Descripción", "Lote", "Canje", "Bulto", "Cant", "Vencimiento"]
                    st.dataframe(df_filtrado_sc[cols_mostrar_sc], hide_index=True, use_container_width=True, height=180, column_config=cfg_columnas)
                    
                    if rol in ["admin", "jefatura_admin", "jefatura"]:
                        opciones_id_sc = df_filtrado_sc['ID'].tolist()
                        formato_opciones_sc = {row['ID']: f"{row['Código']} - {row['Descripción']} (Lote: {row['Lote']})" for idx, row in df_filtrado_sc.iterrows()}
                        
                        st.markdown("#### ⚙️ Cierre Definitivo (Sin Canje)")
                        id_sc = st.selectbox("Seleccione el producto para CERRAR:", opciones_id_sc, format_func=lambda x: formato_opciones_sc[x], key="sel_p5_sc")
                        
                        cursor_sc = conn.cursor()
                        cursor_sc.execute("SELECT * FROM productos WHERE id=%s", (id_sc,))
                        prod_sc = cursor_sc.fetchone()
                        
                        if prod_sc:
                            col_names = [desc[0] for desc in cursor_sc.description]
                            prod_sc = dict(zip(col_names, prod_sc))
                            
                            with st.form("form_paso5_sin_canje"):
                                st.markdown("<p style='color: #475569; font-size: 14px; font-weight: 600; margin-bottom: 0;'>Opciones de Gestión de Red (Opcional)</p>", unsafe_allow_html=True)
                                c1, c2 = st.columns(2)
                                with c1: difusion_sel = st.selectbox("DIFUSIÓN A LA RED", OPCIONES_DIFUSION_RED, label_visibility="collapsed")
                                with c2: redistribucion_sel = st.selectbox("REDISTRIBUCIÓN STOCK", OPCIONES_REDISTRIBUCION_STOCK, label_visibility="collapsed")
                                
                                st.markdown("<hr style='margin: 10px 0px; border-top: 1px solid #cbd5e1;'>", unsafe_allow_html=True)
                                st.markdown("<p style='color: #475569; font-size: 14px; font-weight: 600; margin-bottom: 0;'>Resolución Final</p>", unsafe_allow_html=True)
                                
                                num_res_sc = st.text_input("N° Resolución o Documento de Baja *")
                                estado_fin_sc = st.selectbox("Estado Final *", ["Producto dado de baja", "Producto donado"], label_visibility="collapsed")
                                obs_sc = st.text_area("Resolución / Comentarios finales", value=prod_sc['observacion_paso5'] or "", label_visibility="collapsed")
                                
                                if st.form_submit_button("Finalizar y Archivar"):
                                    conn.cursor().execute("UPDATE productos SET tipo_gestion_canje=%s, observacion_paso2=%s, observacion_paso5=%s, resolucion_numero=%s, estado_final=%s, estado_global='Concluido' WHERE id=%s", (difusion_sel, redistribucion_sel, obs_sc, num_res_sc, estado_fin_sc, id_sc))
                                    conn.commit(); st.success("Archivado con éxito."); time.sleep(1.5); st.rerun()

    # --- ALERTAS, CARGA MASIVA Y ADMIN ---
    elif tab_seleccionada == "🚨 2. Gestión Anexo II":
        st.markdown("## 🚨 Paso 2 — Anexo II (Alertas Sanitarias)")
        df_alertas = pd.read_sql_query("""SELECT id AS "ID", alerta_numero AS "N° Alerta", codigo_reyimen AS "Código", descripcion AS "Descripción", cantidad AS "Cant", lote AS "Lote", proveedor AS "Proveedor Asignado", estado_global AS "Estado", vencimiento AS "Vencimiento", motivo_informe AS "Motivo" FROM productos WHERE motivo_informe = 'Alerta Sanitaria' AND paso_actual = 2""", conn)
        if df_alertas.empty: st.info("No hay Alertas Sanitarias pendientes de Anexo II.")
        else:
            df_alertas["Nivel Alerta"] = df_alertas.apply(lambda row: calcular_semaforo_vencimiento(row["Vencimiento"], row["Motivo"]), axis=1)
            df_alertas["Vencimiento"] = pd.to_datetime(df_alertas["Vencimiento"], errors='coerce').dt.strftime('%d/%m/%Y')
            st.dataframe(df_alertas, hide_index=True, use_container_width=True, height=180, column_config=cfg_columnas)
            
            if rol in ["admin", "jefatura_admin", "jefatura"]:
                id_alerta = st.selectbox("Seleccione Alerta a gestionar", df_alertas['ID'].tolist())
                
                cursor_alerta = conn.cursor()
                cursor_alerta.execute("SELECT * FROM productos WHERE id=%s", (id_alerta,))
                prod_alerta = cursor_alerta.fetchone()
                
                if prod_alerta:
                    col_names = [desc[0] for desc in cursor_alerta.description]
                    prod_alerta = dict(zip(col_names, prod_alerta))
                    
                    st.markdown(f"#### 📝 Redacción de Anexo II: **{prod_alerta['descripcion']}**")
                    with st.form("form_anexo_ii"):
                        col1, col2 = st.columns(2)
                        principio_activo = col1.text_input("Principio Activo", value=prod_alerta['principio_activo'] or "")
                        titular = col1.text_input("Titular del Producto *", value=prod_alerta['titular_registro'] or "")
                        reg_sanitario = col1.text_input("N° de Registro Sanitario *", value=prod_alerta['registro_sanitario'] or "")
                        proveedor = col2.selectbox("Proveedor *", PROVEEDORES_OFICIALES, index=PROVEEDORES_OFICIALES.index(prod_alerta['proveedor']) if prod_alerta['proveedor'] in PROVEEDORES_OFICIALES else 0)
                        rep_legal = col2.text_input("Representante Legal Hospital", value="Falta Rellenar")
                        dir_tecnico = col2.text_input("Director Técnico / QF Responsable", value=user_info['nombre_completo'])
                        obs_alerta = st.text_area("Otras Observaciones", value=prod_alerta['observacion_paso2'] or "")
                        if st.form_submit_button("💾 Guardar Datos para Anexo II"):
                            conn.cursor().execute("UPDATE productos SET principio_activo=%s, titular_registro=%s, registro_sanitario=%s, proveedor=%s, observacion_paso2=%s WHERE id=%s", (principio_activo, titular, reg_sanitario, proveedor, obs_alerta, id_alerta))
                            conn.commit(); st.success("Datos guardados."); time.sleep(1); st.rerun()
                    
                    if prod_alerta['titular_registro'] and prod_alerta['registro_sanitario']:
                        datos_docx = { "descripcion": prod_alerta['descripcion'], "principio_activo": prod_alerta['principio_activo'], "titular": prod_alerta['titular_registro'], "registro_sanitario": prod_alerta['registro_sanitario'], "lote": prod_alerta['lote'], "representante_legal": "Representante Hospital", "director_tecnico": user_info['nombre_completo'], "proveedor": prod_alerta['proveedor'], "cantidad": prod_alerta['cantidad'], "unidad": prod_alerta['unidad'], "observaciones": prod_alerta['observacion_paso2'] }
                        st.download_button("📄 Descargar Documento Anexo II (.docx)", data=generar_anexo_ii_docx(datos_docx), file_name=f"ANEXO_II_{prod_alerta['lote']}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                        if st.button("✅ Marcar como 'Notificado al Proveedor' y Enviar a Bulto"):
                            conn.cursor().execute("UPDATE productos SET estado_global='Alerta Notificada al Proveedor', paso_actual=3 WHERE id=%s", (id_alerta,))
                            conn.commit(); st.success("Avanzado al Paso 3."); time.sleep(1.5); st.rerun()
            else:
                st.info("🔒 **Modo de solo lectura:** Solo Jefatura puede completar y generar el Anexo II.")

    elif tab_seleccionada == "📦 3. Bulto/Ubicación (Alerta)":
        st.markdown("## 📦 Paso 3 — Bulto y Ubicaciones (Alertas Sanitarias)")
        tab_p3a_nuevo, tab_p3a_seg = st.tabs(["➕ Asignar Bulto a Alerta", "🔄 Seguimiento de Alertas Físicas"])
        
        with tab_p3a_nuevo:
            df_p3a = pd.read_sql_query("""SELECT id AS "ID", alerta_numero AS "N° Alerta", codigo_reyimen AS "Código", descripcion AS "Descripción", cantidad AS "Cant", lote AS "Lote", proveedor AS "Proveedor", vencimiento AS "Vencimiento", motivo_informe AS "Motivo" FROM productos WHERE motivo_informe = 'Alerta Sanitaria' AND paso_actual = 3 AND estado_global != 'Concluido'""", conn)
            
            if df_p3a.empty: 
                st.info("No hay alertas pendientes de asignación de bulto.")
            else:
                df_p3a["Nivel Alerta"] = df_p3a.apply(lambda row: calcular_semaforo_vencimiento(row["Vencimiento"], row["Motivo"]), axis=1)
                df_p3a["Vencimiento"] = pd.to_datetime(df_p3a["Vencimiento"], errors='coerce').dt.strftime('%d/%m/%Y')
                
                st.markdown("#### 🔍 Filtros de Búsqueda")
                filtro_codigo = st.text_input("Filtrar por Código Reyimen", placeholder="Ej: 543...", key="cod_p3a_n")
                
                df_filtrado = df_p3a.copy()
                if filtro_codigo: df_filtrado = df_filtrado[df_filtrado["Código"].astype(str).str.contains(filtro_codigo, case=False, na=False)]
                
                st.markdown("<br>", unsafe_allow_html=True)
                if df_filtrado.empty:
                    st.warning("No hay productos que coincidan con los filtros.")
                else:
                    st.dataframe(df_filtrado, hide_index=True, use_container_width=True, height=180, column_config=cfg_columnas)
                    
                    if rol in ["admin", "jefatura_admin", "bodega"]:
                        opciones_id = df_filtrado['ID'].tolist()
                        formato_opciones = {row['ID']: f"{row['Código']} - {row['Descripción']} (Lote: {row['Lote']})" for idx, row in df_filtrado.iterrows()}
                        
                        st.markdown("#### ⚙️ Asignación de Ubicación y Cierre de Alerta")
                        prod_id = st.selectbox("Seleccione la alerta a gestionar:", opciones_id, format_func=lambda x: formato_opciones[x], key="sel_p3a_n")
                        
                        with st.form("form_p3a"):
                            ub_fisica, ub_comp = st.selectbox("Ubicación Física *", OPCIONES_FISICA_P4), st.selectbox("Ubicación Computacional *", BODEGAS_PASO4)
                            bulto = st.text_input("N° Bulto *")
                            obs = st.text_area("Observaciones Finales (Opcional)")
                            if st.form_submit_button("Guardar Ubicación y Finalizar Alerta"):
                                conn.cursor().execute("UPDATE productos SET ubicacion_fisica=%s, ubicacion_computacional=%s, numero_bulto=%s, observacion_paso4=%s, paso_actual=4, estado_global='Concluido', estado_final='Retirado por Alerta Sanitaria' WHERE id=%s", (ub_fisica, ub_comp, bulto, obs, prod_id))
                                conn.commit(); st.success("Alerta Sanitaria finalizada y archivada."); time.sleep(1.5); st.rerun()
                    else:
                        st.info("🔒 **Modo de solo lectura:** Solo Bodega puede ubicar y finalizar los bultos de Alertas Sanitarias.")

        with tab_p3a_seg:
            df_seg = pd.read_sql_query("""SELECT id AS "ID", alerta_numero AS "N° Alerta", codigo_reyimen AS "Código", descripcion AS "Descripción", cantidad AS "Cant", lote AS "Lote", ubicacion_fisica AS "Física", numero_bulto AS "Bulto", vencimiento AS "Vencimiento", motivo_informe AS "Motivo" FROM productos WHERE motivo_informe = 'Alerta Sanitaria' AND paso_actual >= 3 AND ubicacion_fisica != '' AND estado_global = 'Concluido'""", conn)
            
            if df_seg.empty: 
                st.info("No hay bultos de alertas para seguimiento.")
            else:
                df_seg["Nivel Alerta"] = df_seg.apply(lambda row: calcular_semaforo_vencimiento(row["Vencimiento"], row["Motivo"]), axis=1)
                df_seg["Vencimiento"] = pd.to_datetime(df_seg["Vencimiento"], errors='coerce').dt.strftime('%d/%m/%Y')
                
                st.markdown("#### 🔍 Filtros de Búsqueda")
                filtro_codigo_seg = st.text_input("Filtrar por Código Reyimen", placeholder="Ej: 543...", key="cod_p3a_s")
                
                df_filtrado_seg = df_seg.copy()
                if filtro_codigo_seg: df_filtrado_seg = df_filtrado_seg[df_filtrado_seg["Código"].astype(str).str.contains(filtro_codigo_seg, case=False, na=False)]
                
                st.markdown("<br>", unsafe_allow_html=True)
                if df_filtrado_seg.empty:
                    st.warning("No hay bultos que coincidan con los filtros.")
                else:
                    st.dataframe(df_filtrado_seg, hide_index=True, use_container_width=True, height=180, column_config=cfg_columnas)
                    
                    if rol in ["admin", "jefatura_admin", "bodega"]:
                        opciones_id_seg = df_filtrado_seg['ID'].tolist()
                        formato_opciones_seg = {row['ID']: f"{row['Código']} - {row['Descripción']} (Lote: {row['Lote']})" for idx, row in df_filtrado_seg.iterrows()}
                        
                        st.markdown("#### ⚙️ Modificación de Bulto Existente")
                        id_seg = st.selectbox("Seleccione el bulto a actualizar:", opciones_id_seg, format_func=lambda x: formato_opciones_seg[x], key="sel_p3a_s")
                        
                        cursor_p3a = conn.cursor()
                        cursor_p3a.execute("SELECT * FROM productos WHERE id=%s", (id_seg,))
                        prod_seg = cursor_p3a.fetchone()
                        
                        if prod_seg:
                            col_names = [desc[0] for desc in cursor_p3a.description]
                            prod_seg = dict(zip(col_names, prod_seg))
                            
                            with st.form("form_p3a_actualizar"):
                                idx_fis = OPCIONES_FISICA_P4.index(prod_seg['ubicacion_fisica']) if prod_seg['ubicacion_fisica'] in OPCIONES_FISICA_P4 else 0
                                idx_comp = BODEGAS_PASO4.index(prod_seg['ubicacion_computacional']) if prod_seg['ubicacion_computacional'] in BODEGAS_PASO4 else 0
                                
                                u_fisica = st.selectbox("Ubicación Física", OPCIONES_FISICA_P4, index=idx_fis)
                                u_comp = st.selectbox("Ubicación Computacional", BODEGAS_PASO4, index=idx_comp)
                                u_bulto = st.text_input("N° Bulto", value=prod_seg['numero_bulto'] or "")
                                u_obs = st.text_area("Observaciones", value=prod_seg['observacion_paso4'] or "")
                                
                                if st.form_submit_button("Actualizar Bulto"):
                                    conn.cursor().execute("UPDATE productos SET ubicacion_fisica=%s, ubicacion_computacional=%s, numero_bulto=%s, observacion_paso4=%s WHERE id=%s", (u_fisica, u_comp, u_bulto, u_obs, id_seg))
                                    conn.commit(); st.success("Ubicación actualizada."); time.sleep(1); st.rerun()

    # --- CARGA MASIVA Y REPORTES ---
    elif tab_seleccionada == "📤 Carga Masiva":
        st.markdown("## 📤 Carga Masiva de Productos")
        
        df_plantilla = pd.DataFrame([{"BODEGA ORIGEN": "Bodega AZ09 (Fármacos)", "TIPO PRODUCTO": "Fármaco", "CÓDIGO REYIMEN": "1365", "DESCRIPCIÓN": "BUPIVACAINA", "TIPO COMPRA": "CENABAST", "UNIDAD": "FRASCO", "CANTIDAD": 100, "FECHA VENCIMIENTO": "31/10/2026", "LOTE": "L12345"}])
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_plantilla.to_excel(writer, index=False, sheet_name='Plantilla_Carga')
            worksheet = writer.sheets['Plantilla_Carga']
            
            dv_bodega = DataValidation(type="list", formula1=f'"{",".join(BODEGAS_OFICIALES)}"', allow_blank=False)
            dv_tipo = DataValidation(type="list", formula1='"Fármaco,Insumo"', allow_blank=False)
            dv_compra = DataValidation(type="list", formula1='"CENABAST,Compra propia,Desconocido (Farmacia)"', allow_blank=False)
            
            worksheet.add_data_validation(dv_bodega)
            worksheet.add_data_validation(dv_tipo)
            worksheet.add_data_validation(dv_compra)
            
            dv_bodega.add('A2:A1000') 
            dv_tipo.add('B2:B1000')   
            dv_compra.add('E2:E1000') 
            
            anchos = {'A': 28, 'B': 18, 'C': 18, 'D': 45, 'E': 25, 'F': 15, 'G': 12, 'H': 22, 'I': 15}
            for col, ancho in anchos.items():
                worksheet.column_dimensions[col].width = ancho

        st.markdown("#### 1. Descargar Plantilla Modelo")
        st.download_button(label="📥 Descargar Plantilla Excel", data=output.getvalue(), file_name="Plantilla_Carga_Masiva_Validada.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        st.markdown("#### 2. Subir Archivo Completado")
        uploaded = st.file_uploader("Subir archivo Excel o CSV (Asegúrate de usar formato DD/MM/YYYY)", type=["xlsx", "xls", "csv"])
        if uploaded and st.button("🚀 Procesar e Ingresar Productos"):
            ok, msg = procesar_carga_masiva(uploaded, user_info['usuario'])
            if ok: st.success(msg); time.sleep(1.5); st.rerun()
            else: st.error(msg)
            
    elif tab_seleccionada == "🔍 Consolidado General":
        st.markdown("## 🔍 Consolidado General")
        df_cons = pd.read_sql_query("""SELECT id AS "ID", codigo_reyimen AS "Código", descripcion AS "Descripción", bodega_origen AS "Bodega", cantidad AS "Cant", lote AS "Lote", vencimiento AS "Venc", motivo_informe AS "Motivo", estado_global AS "Estado", proveedor AS "Proveedor" FROM productos""", conn)
        
        df_cons["Nivel Alerta"] = df_cons.apply(lambda row: calcular_semaforo_vencimiento(row["Venc"], row["Motivo"]), axis=1)
        df_cons["Venc"] = pd.to_datetime(df_cons["Venc"], errors='coerce').dt.strftime('%d/%m/%Y')
        cols_cons = ["ID", "Nivel Alerta", "Código", "Descripción", "Bodega", "Cant", "Lote", "Venc", "Estado", "Proveedor"]
        df_cons = df_cons[cols_cons]
        st.dataframe(df_cons, hide_index=True, column_config=cfg_columnas)
        
    elif tab_seleccionada == "📊 Dashboard / Análisis":
        st.markdown("## 📊 Dashboard y Estadísticas")
        df_all = pd.read_sql_query("SELECT * FROM productos", conn)
        if not df_all.empty:
            df_all["Nivel Alerta"] = df_all.apply(lambda row: calcular_semaforo_vencimiento(row["vencimiento"], row["motivo_informe"]), axis=1)
            
            vencidos = len(df_all[df_all["Nivel Alerta"] == "⚫ Vencido (≤ 0d)"])
            rojas = len(df_all[df_all["Nivel Alerta"] == "🔴 Roja (1-60d)"])
            amarillas = len(df_all[df_all["Nivel Alerta"] == "🟡 Amarilla (61-120d)"])
            verdes = len(df_all[df_all["Nivel Alerta"] == "🟢 Verde (> 120d)"])
            
            df_all["vencimiento"] = pd.to_datetime(df_all["vencimiento"], errors='coerce').dt.strftime('%d/%m/%Y')
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_all.to_excel(writer, index=False, sheet_name='Consolidado')
            
            st.download_button(
                label="📥 Descargar Reporte Completo en Excel",
                data=output.getvalue(),
                file_name=f"Reporte_Vencimientos_{datetime.now().strftime('%d_%m_%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.markdown("<br>", unsafe_allow_html=True)
            
            st.markdown("#### 🚦 Semáforo de Riesgo (Vencimientos)")
            
            source = pd.DataFrame({
                "Nivel": ["Vencido (≤ 0d)", "Crítico (Roja)", "Tramitación (Amarilla)", "Seguro (Verde)"],
                "Cantidad": [vencidos, rojas, amarillas, verdes]
            })
            
            if source["Cantidad"].sum() > 0:
                source = source[source["Cantidad"] > 0]
                
                grafico_pastel = alt.Chart(source).mark_arc(innerRadius=60).encode(
                    theta=alt.Theta(field="Cantidad", type="quantitative"),
                    color=alt.Color(
                        field="Nivel", 
                        type="nominal", 
                        scale=alt.Scale(
                            domain=["Vencido (≤ 0d)", "Crítico (Roja)", "Tramitación (Amarilla)", "Seguro (Verde)"], 
                            range=["#1e293b", "#ef4444", "#f59e0b", "#10b981"]
                        ),
                        legend=alt.Legend(title="Estado", orient="right")
                    ),
                    tooltip=['Nivel', 'Cantidad']
                ).properties(height=350)
                
                st.altair_chart(grafico_pastel, use_container_width=True)
            else:
                st.info("Aún no hay productos registrados para generar el gráfico.")
            
            sem1, sem2, sem3, sem4 = st.columns(4)
            sem1.metric("⚫ Vencido (Mermas)", vencidos)
            sem2.metric("🔴 Alerta Roja (Crítico)", rojas)
            sem3.metric("🟡 Alerta Amarilla (Tramitación)", amarillas)
            sem4.metric("🟢 Alerta Verde (Monitoreo)", verdes)
            
            st.markdown("---")
            
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Trámites Totales", len(df_all))
            m2.metric("En Trámite Activo", len(df_all[df_all['estado_global'].isin(['En trámite', 'CUARENTENA', 'Alerta Notificada al Proveedor'])]))
            m3.metric("Concluidos", len(df_all[df_all['estado_global'] == 'Concluido']))
            
            estados_canje = ['Canjeado - reposición del producto', 'Canjeado - nota de credito recibida']
            canjeadas_df = df_all[df_all['estado_final'].isin(estados_canje)]
            unidades_canjeadas = int(canjeadas_df['cantidad'].sum()) if not canjeadas_df.empty else 0
            m4.metric("Unid. Canjeadas", unidades_canjeadas)
            
    elif tab_seleccionada == "👥 Gestión de Usuarios":
        st.markdown("## 👥 Gestión de Usuarios")
        tab_crear, tab_editar = st.tabs(["➕ Crear Usuario", "✏️ Editar / Eliminar"])
        with tab_crear:
            with st.form("form_nuevo_usuario"):
                u_user, u_pass = st.text_input("Usuario *"), st.text_input("Contraseña *", type="password")
                u_nombre, u_rol = st.text_input("Nombre *"), st.selectbox("Rol", ["admin", "jefatura_admin", "jefatura", "registro", "bodega"])
                if st.form_submit_button("Crear"):
                    if not u_user or not u_pass or not u_nombre:
                        st.error("⚠️ Todos los campos con asterisco (*) son obligatorios.")
                    else:
                        try:
                            conn.cursor().execute("INSERT INTO usuarios (usuario, password, rol, nombre_completo) VALUES (%s, %s, %s, %s)", (u_user, u_pass, u_rol, u_nombre))
                            conn.commit()
                            st.success("Usuario creado exitosamente."); time.sleep(1); st.rerun()
                        except psycopg2.IntegrityError:
                            st.error(f"⚠️ El nombre de usuario '{u_user}' ya está en uso. Por favor, elige otro distinto (ej: {u_user}2).")
                            conn.rollback()
            st.dataframe(pd.read_sql_query("SELECT id, usuario, rol, nombre_completo, estado FROM usuarios", conn), hide_index=True, use_container_width=True, height=180)
        with tab_editar:
            df_users_edit = pd.read_sql_query("SELECT id, usuario, rol, nombre_completo, estado FROM usuarios", conn)
            if not df_users_edit.empty:
                st.markdown("#### 📋 Listado Actual de Usuarios")
                st.dataframe(df_users_edit, hide_index=True, use_container_width=True)
                
                opciones_user_id = df_users_edit['id'].tolist()
                formato_opciones_user = {row['id']: f"ID {row['id']} - {row['nombre_completo']} ({row['usuario']}) - Rol: {row['rol'].upper()}" for idx, row in df_users_edit.iterrows()}
                
                st.markdown("#### ⚙️ Editar o Eliminar Usuario")
                id_mod = st.selectbox("Seleccione el usuario a Modificar/Eliminar:", opciones_user_id, format_func=lambda x: formato_opciones_user[x])
                
                cursor_usr = conn.cursor()
                cursor_usr.execute("SELECT * FROM usuarios WHERE id=%s", (id_mod,))
                user_data = cursor_usr.fetchone()
                
                if user_data:
                    col_names = [desc[0] for desc in cursor_usr.description]
                    user_data = dict(zip(col_names, user_data))
                    
                    with st.form("form_editar_usuario"):
                        new_u_nombre, new_u_user = st.text_input("Nombre", value=user_data['nombre_completo']), st.text_input("Usuario", value=user_data['usuario'])
                        
                        lista_roles = ["admin", "jefatura_admin", "jefatura", "registro", "bodega"]
                        idx_rol = lista_roles.index(user_data['rol']) if user_data['rol'] in lista_roles else 0
                        
                        new_u_rol, new_u_estado = st.selectbox("Rol", lista_roles, index=idx_rol), st.selectbox("Estado", ["Activo", "Inactivo"], index=["Activo", "Inactivo"].index(user_data['estado']))
                        new_u_pass = st.text_input("Nueva Contraseña (Opcional)", type="password")
                        b1, b2 = st.columns(2)
                        if b1.form_submit_button("Guardar Cambios"):
                            if new_u_pass.strip(): conn.cursor().execute("UPDATE usuarios SET usuario=%s, password=%s, rol=%s, nombre_completo=%s, estado=%s WHERE id=%s", (new_u_user, new_u_pass, new_u_rol, new_u_nombre, new_u_estado, id_mod))
                            else: conn.cursor().execute("UPDATE usuarios SET usuario=%s, rol=%s, nombre_completo=%s, estado=%s WHERE id=%s", (new_u_user, new_u_rol, new_u_nombre, new_u_estado, id_mod))
                            conn.commit(); st.success("Actualizado."); time.sleep(1); st.rerun()
                        if b2.form_submit_button("Eliminar Usuario"):
                            conn.cursor().execute("DELETE FROM usuarios WHERE id=%s", (id_mod,))
                            conn.commit(); st.warning("Eliminado."); time.sleep(1); st.rerun()

    conn.close()
