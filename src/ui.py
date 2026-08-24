"""
Módulo de Interfaz de Usuario para los 5 Pasos Operativos, Carga Masiva,
Gestión de Usuarios, Dashboard, Consolidado General, Interfaz Segura y Alertas Sanitarias.
"""
import io
import os
import time
import sqlite3
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from docx import Document
from docx.shared import Pt
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
        .stDeployButton {{ display: none !important; }}
        .block-container {{ padding-top: 1.5rem !important; padding-bottom: 1.5rem !important; }}
        .stApp {{ background-color: {tema['bg']} !important; color: {tema['text']} !important; }}
        [data-testid="stSidebar"] {{ background-color: {tema['card']} !important; border-right: 1px solid {tema['border']} !important; }}
        [data-testid="stSidebar"] *, [data-testid="stSidebar"] label, [data-testid="stSidebar"] p, [data-testid="stSidebar"] span {{ color: {tema['text']} !important; font-weight: 600 !important; }}
        [data-testid="stSidebar"] div[role="radiogroup"] > label {{ padding: 10px 12px; background-color: transparent; border-radius: 8px; margin-bottom: 4px; transition: all 0.2s ease; }}
        [data-testid="stSidebar"] div[role="radiogroup"] > label:hover {{ background-color: rgba(0, 0, 0, 0.04); transform: translateX(4px); }}
        [data-testid="stSidebar"] div[role="radiogroup"] > label > div:first-child {{ transform: scale(0.85); }}
        .stMarkdown, .stText, h1, h2, h3, h4, h5, h6, label, p, span {{ color: {tema['text']} !important; }}
        [data-testid="stMetricValue"] {{ color: {tema['text']} !important; }}
        .stMarkdown p {{ margin-bottom: 0 !important; }}
        code {{ background-color: #dcfce7 !important; color: #15803d !important; border: 1px solid #86efac !important; font-weight: bold !important; padding: 2px 8px !important; border-radius: 4px !important; }}
        
        [data-testid="collapsedControl"], [data-testid="stSidebarCollapsedControl"] {{
            opacity: 1 !important; background-color: #ffffff !important; border: 1px solid #cbd5e1 !important; border-radius: 8px !important; box-shadow: 0 2px 5px rgba(0,0,0,0.08) !important; color: #0f172a !important; transition: all 0.3s ease; z-index: 999999 !important;
        }}
        [data-testid="collapsedControl"]:hover, [data-testid="stSidebarCollapsedControl"]:hover {{ background-color: #f1f5f9 !important; border-color: #94a3b8 !important; }}
        [data-testid="collapsedControl"] svg, [data-testid="stSidebarCollapsedControl"] svg {{ fill: #0f172a !important; color: #0f172a !important; }}

        div[data-baseweb="select"] > div, .stTextInput div[data-baseweb="input"], .stNumberInput div[data-baseweb="input"], .stDateInput div[data-baseweb="input"], .stTextArea div[data-baseweb="textarea"] {{ background-color: #ffffff !important; border: 1px solid #cbd5e1 !important; }}
        .stTextInput input, .stNumberInput input, .stDateInput input, .stTextArea textarea {{ background-color: #ffffff !important; color: #0f172a !important; }}
        div[data-baseweb="select"] span {{ color: #0f172a !important; }}
        ul[data-testid="stSelectboxVirtualDropdown"] {{ background-color: #ffffff !important; border: 1px solid #cbd5e1 !important; }}
        ul[data-testid="stSelectboxVirtualDropdown"] li {{ color: #0f172a !important; }}
        .stButton button, .stDownloadButton button, [data-testid="stFileUploader"] button {{ background-color: #ffffff !important; color: #0f172a !important; border: 1px solid #cbd5e1 !important; font-weight: bold !important; box-shadow: 0 1px 2px rgba(0,0,0,0.05) !important; }}
        .stButton button *, .stDownloadButton button *, [data-testid="stFileUploader"] button * {{ color: #0f172a !important; }}
        div[data-testid="stForm"] button {{ background-color: #0284c7 !important; border: 1px solid #0369a1 !important; }}
        div[data-testid="stForm"] button *, div[data-testid="stForm"] button p, div[data-testid="stForm"] button span {{ color: #ffffff !important; font-weight: bold !important; font-size: 15px !important; }}
        div[data-testid="stForm"] button:hover {{ background-color: #0369a1 !important; }}
        [data-testid="stFileUploader"] section {{ background-color: #ffffff !important; border: 2px dashed #cbd5e1 !important; }}
        [data-testid="stFileUploader"] section * {{ color: #334155 !important; }}
        
        [data-testid="stSidebar"] [data-testid="stImage"] {{ display: flex !important; justify-content: center !important; margin-bottom: 20px !important; background: transparent !important; border: none !important; box-shadow: none !important; padding: 0 !important; }}
        [data-testid="stSidebar"] [data-testid="stImage"] img {{ display: block !important; margin: 0 auto !important; image-rendering: high-quality !important; -webkit-font-smoothing: antialiased !important; mix-blend-mode: multiply !important; }}
    </style>
    """
    st.markdown(css, unsafe_allow_html=True)

def generar_anexo_ii_docx(datos):
    """Genera el documento Word (Anexo II) con la información de la Alerta Sanitaria."""
    doc = Document()
    
    # Título Principal
    titulo = doc.add_heading('ANEXO II', level=1)
    titulo.alignment = 1 # Centrado
    
    sub = doc.add_heading('FORMULARIO DE REPORTE DE EXISTENCIAS DE PRODUCTO RETIRADO DEL MERCADO', level=2)
    sub.alignment = 1
    
    doc.add_paragraph('PRODUCTOS FARMACÉUTICOS (ARTS. 60° y 71° 3, D.S. N° 3/2010)').alignment = 1
    doc.add_paragraph('Documento emitido por cada establecimiento que ha recibido producto afecto(s) a retiro del mercado, que da cuenta de la cantidad recibida, stock y cantidad distribuida de los lotes en cuestión, que debe ser informado al distribuidor de quien obtuvo el producto.')
    
    doc.add_heading('ANTECEDENTES DEL PRODUCTO EN PROCESO DE RETIRO DEL MERCADO', level=3)
    doc.add_paragraph(f"PRODUCTO: {datos.get('descripcion', '')}")
    doc.add_paragraph(f"PRINCIPIO ACTIVO: {datos.get('principio_activo', '')}")
    doc.add_paragraph(f"TITULAR: {datos.get('titular', '')}")
    doc.add_paragraph(f"N° DE REGISTRO SANITARIO: {datos.get('registro_sanitario', '')}")
    doc.add_paragraph(f"NUMERO DE SERIE(S) / LOTE(S): {datos.get('lote', '')}")
    
    doc.add_heading('ANTECEDENTES DEL ESTABLECIMIENTO RECEPTOR', level=3)
    doc.add_paragraph("NOMBRE y DIRECCIÓN: Hospital Penco Lirquén")
    doc.add_paragraph("TIPO DE ESTABLECIMIENTO/LÍNEAS DE ACTIVIDAD: Establecimiento de Salud Pública")
    doc.add_paragraph(f"REPRESENTANTE LEGAL: {datos.get('representante_legal', '')}")
    doc.add_paragraph(f"DIRECTOR TÉCNICO/RESPONSABLE TÉCNICO: {datos.get('director_tecnico', '')}")
    doc.add_paragraph(f"FECHA DE REPORTE: {datetime.now().strftime('%d-%m-%Y')}")
    
    doc.add_heading('ANTECEDENTES MOVIMIENTO DEL PRODUCTO', level=3)
    doc.add_paragraph(f"PROVEEDOR: {datos.get('proveedor', '')}")
    doc.add_paragraph(f"CANTIDAD EN EXISTENCIA (STOCK) NO DISTRIBUIDA: {datos.get('cantidad', '')} {datos.get('unidad', '')}")
    doc.add_paragraph(f"CANTIDAD DISTRIBUIDA A OTROS ESTABLECIMIENTOS: 0")
    doc.add_paragraph("REGISTRO DE DISTRIBUCIÓN: N/A")
    doc.add_paragraph(f"OTRAS OBSERVACIONES: {datos.get('observaciones', '')}")
    
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def actualizar_bd_alertas(conn):
    """Actualiza la estructura de la base de datos de forma segura para soportar Alertas."""
    cursor = conn.cursor()
    columnas = [
        "alerta_numero TEXT", "alerta_fecha TEXT", "titular_registro TEXT", 
        "registro_sanitario TEXT", "principio_activo TEXT"
    ]
    for col in columnas:
        try:
            cursor.execute(f"ALTER TABLE productos ADD COLUMN {col}")
        except sqlite3.OperationalError:
            pass # Si la columna ya existe, ignora el error y continua
    conn.commit()

def render_ui(user_info: dict):
    conn = get_connection()
    actualizar_bd_alertas(conn)

    opciones_temas = list(TEMAS_COLOR_CLAROS.keys())
    tema_actual = st.session_state.get("tema_seleccionado", "Claro (Predeterminado)")
    if tema_actual not in opciones_temas:
        tema_actual = "Claro (Predeterminado)"
        st.session_state["tema_seleccionado"] = tema_actual
    idx_tema = opciones_temas.index(tema_actual)

    col_user, col_rol, col_tema, col_btn = st.columns([3, 2, 3, 2], vertical_alignment="center")
    with col_user: st.markdown(f"👤 **Usuario:** {user_info['nombre_completo']}")
    with col_rol: st.markdown(f"🛡️ **Rol:** `{user_info['rol'].upper()}`")
    with col_tema:
        tema_sel = st.selectbox("Color", opciones_temas, index=idx_tema, label_visibility="collapsed")
        st.session_state["tema_seleccionado"] = tema_sel
    with col_btn:
        if st.button("🚪 Cerrar Sesión", use_container_width=True):
            st.session_state["logged_in"] = False
            st.rerun()

    borde_color = TEMAS_COLOR_CLAROS[tema_sel]["border"]
    st.markdown(f"<hr style='margin-top: 0.5rem; margin-bottom: 1rem; border: none; border-top: 1px solid {borde_color};' />", unsafe_allow_html=True)
    aplicar_estilo_tema(tema_sel)

    path1, path2 = "assets/hospital-penco-lirquen.png", "assets/logo.png"
    if os.path.exists(path1): st.sidebar.image(path1, width=150)
    elif os.path.exists(path2): st.sidebar.image(path2, width=150)

    rol = user_info['rol']
    
    # --- MENÚ DE NAVEGACIÓN ESTILIZADO CON EMOJIS ---
    tabs_disponibles = []
    if rol in ["admin", "bodega"]:
        tabs_disponibles.append("📋 1. Informe Bodega")
        tabs_disponibles.append("📤 Carga Masiva")
    if rol in ["admin", "jefatura"]:
        tabs_disponibles.append("🚨 Alertas Sanitarias")
        tabs_disponibles.append("⚖️ 2. Canjes (Jefatura)")
    if rol in ["admin", "registro"]:
        tabs_disponibles.append("🚚 3. Registro/Prov.")
    if rol in ["admin", "bodega"]:
        tabs_disponibles.append("📦 4. Bulto/Ubicación")
    if rol in ["admin", "jefatura"]:
        tabs_disponibles.append("📜 5. Resolución/Cierre")
    
    tabs_disponibles.append("🔍 Consolidado General")
    tabs_disponibles.append("📊 Dashboard / Análisis")
    
    if rol == "admin":
        tabs_disponibles.append("👥 Gestión de Usuarios")

    # label_visibility="collapsed" oculta la palabra "Navegación"
    tab_seleccionada = st.sidebar.radio("Navegación", tabs_disponibles, label_visibility="collapsed")

    # --- PASO 1 ---
    if tab_seleccionada == "📋 1. Informe Bodega":
        st.header("📋 Paso 1 — Informe de Bodega")
        catalogo = get_catalogo()
        opciones = get_opciones_selectbox(catalogo)
        
        tab_p1_ingresar, tab_p1_editar = st.tabs(["➕ Nuevo Registro", "✏️ Editar / Modificar Registros Existentes"])

        with tab_p1_ingresar:
            c_busq1, c_busq2 = st.columns(2)
            with c_busq1:
                bodega = st.selectbox("Bodega/Farmacia Origen *", BODEGAS_OFICIALES, key="b_ingreso")
                tipo_prod = st.selectbox("Tipo de producto", ["Fármaco", "Insumo"], key="tp_ingreso")
            
            with c_busq2:
                prod_sel = st.selectbox("Buscar Código Reyimen / Producto *", opciones, key="busqueda_producto")
                
            cod_auto, desc_auto, unidad_auto = "", "", ""
            item = buscar_por_etiqueta(catalogo, prod_sel)
            if item:
                cod_auto = item.get("codigo", "")
                desc_auto = item.get("descripcion", "")
                unidad_auto = item.get("unidad", "")

            with st.form("form_paso1"):
                motivo = st.selectbox(
                    "Motivo de informe *",
                    ["Gestión pronto vencimiento", "Alerta Sanitaria", "Falla de calidad"]
                )
                
                es_alerta = (motivo == "Alerta Sanitaria")
                if es_alerta:
                    st.error("🚨 MODO ALERTA SANITARIA: El producto será bloqueado y pasará a Cuarentena Inmediata.")
                
                col1, col2 = st.columns(2)
                with col1:
                    codigo = st.text_input("Código Reyimen *", value=cod_auto)
                    descripcion = st.text_input("Descripción *", value=desc_auto)
                    if not es_alerta:
                        tipo_compra = st.selectbox("Tipo de compra *", ["CENABAST", "Compra propia"])
                    else:
                        tipo_compra = "No Aplica (Alerta)"
                
                with col2:
                    unidad = st.text_input("Unidad *", value=unidad_auto)
                    cantidad = st.number_input("Cantidad *", min_value=0.0, step=1.0)
                    
                    if es_alerta:
                        alerta_numero = st.text_input("N° Alerta Sanitaria ISP *")
                        alerta_fecha = st.date_input("Fecha de Alerta ISP *")
                        num_bulto_alerta = st.text_input("Número de Bulto Físico (Bodega Excluidos) *")

                c3, c4 = st.columns(2)
                with c3:
                    vencimiento = st.date_input("Fecha de Vencimiento *")
                with c4:
                    lote = st.text_input("Lote *")
                
                if st.form_submit_button("Guardar Paso 1"):
                    if not codigo or not descripcion or not lote or (es_alerta and (not alerta_numero or not num_bulto_alerta)):
                        st.error("Complete todos los campos obligatorios (*)")
                    else:
                        cursor = conn.cursor()
                        cursor.execute("""
                            SELECT id FROM productos 
                            WHERE codigo_reyimen=? AND lote=? AND bodega_origen=? AND estado_global IN ('En trámite', 'CUARENTENA')
                        """, (codigo, lote, bodega))
                        
                        if cursor.fetchone():
                            st.warning("⚠️ ¡Atención! Este producto (mismo código, lote y bodega) ya fue ingresado y está activo.")
                        else:
                            estado_inicial = 'CUARENTENA' if es_alerta else 'En trámite'
                            paso_inicial = 2
                            ub_fisica = "Bodega de Excluidos" if es_alerta else ""
                            bulto = num_bulto_alerta if es_alerta else ""
                            a_num = alerta_numero if es_alerta else ""
                            a_fec = str(alerta_fecha) if es_alerta else ""

                            cursor.execute("""
                            INSERT INTO productos (bodega_origen, tipo_producto, codigo_reyimen, descripcion, unidad, cantidad, vencimiento, lote, motivo_informe, tipo_documento, usuario_registro, paso_actual, estado_global, ubicacion_fisica, numero_bulto, alerta_numero, alerta_fecha)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                            """, (bodega, tipo_prod, codigo, descripcion, unidad, cantidad, str(vencimiento), lote, motivo, tipo_compra, user_info['usuario'], paso_inicial, estado_inicial, ub_fisica, bulto, a_num, a_fec))
                            conn.commit()
                            
                            st.success("✅ Producto registrado exitosamente en el sistema.")
                            time.sleep(1.5)
                            st.rerun()

        with tab_p1_editar:
            st.subheader("Registros Ingresados en Paso 1 (Modificables)")
            df_p1 = pd.read_sql_query("""
                SELECT id AS ID, bodega_origen AS "Bodega", codigo_reyimen AS "Código", descripcion AS "Descripción", lote AS "Lote", motivo_informe AS "Motivo", estado_global AS "Estado" 
                FROM productos WHERE paso_actual = 1 OR estado_global = 'CUARENTENA'
            """, conn)

            if df_p1.empty:
                st.info("No hay registros pendientes para modificar.")
            else:
                st.dataframe(df_p1, hide_index=True, use_container_width=True)
                id_mod = st.selectbox("Seleccione ID del registro a Modificar/Eliminar", df_p1['ID'].tolist())
                prod_data = conn.cursor().execute("SELECT * FROM productos WHERE id=?", (id_mod,)).fetchone()
                
                if prod_data:
                    with st.form("form_editar_p1"):
                        st.markdown(f"**Modificando Registro ID #{id_mod} — {prod_data['descripcion']}**")
                        col_e1, col_e2 = st.columns(2)
                        with col_e1:
                            idx_bodega = BODEGAS_OFICIALES.index(prod_data['bodega_origen']) if prod_data['bodega_origen'] in BODEGAS_OFICIALES else 0
                            new_bodega = st.selectbox("Bodega Origen *", BODEGAS_OFICIALES, index=idx_bodega)
                            new_cantidad = st.number_input("Cantidad *", value=float(prod_data['cantidad']), min_value=0.0, step=1.0)
                        
                        with col_e2:
                            new_lote = st.text_input("Lote *", value=prod_data['lote'])
                            try:
                                fecha_init = datetime.strptime(prod_data['vencimiento'], "%Y-%m-%d").date()
                            except Exception:
                                fecha_init = datetime.now().date()
                                
                            new_vencimiento = st.date_input("Fecha de Vencimiento *", value=fecha_init)

                        btn_col1, btn_col2 = st.columns(2)
                        with btn_col1:
                            guardar_mod = st.form_submit_button("💾 Guardar Cambios")
                        with btn_col2:
                            eliminar_mod = st.form_submit_button("🗑️ Eliminar Registro")

                        if guardar_mod:
                            cursor = conn.cursor()
                            cursor.execute("UPDATE productos SET bodega_origen=?, cantidad=?, lote=?, vencimiento=? WHERE id=?", (new_bodega, new_cantidad, new_lote, str(new_vencimiento), id_mod))
                            conn.commit()
                            st.success("Registro actualizado.")
                            time.sleep(1)
                            st.rerun()

                        if eliminar_mod:
                            cursor = conn.cursor()
                            cursor.execute("DELETE FROM productos WHERE id=?", (id_mod,))
                            conn.commit()
                            st.warning("Registro eliminado.")
                            time.sleep(1)
                            st.rerun()

    # --- NUEVO: ALERTAS SANITARIAS (JEFATURA) ---
    elif tab_seleccionada == "🚨 Alertas Sanitarias":
        st.header("🚨 Gestión Exclusiva de Alertas Sanitarias")
        st.markdown("Generación de **Anexo II** y notificación a proveedores para productos en Cuarentena.")
        
        df_alertas = pd.read_sql_query("""
            SELECT id AS ID, alerta_numero AS "N° Alerta", codigo_reyimen AS "Código", descripcion AS "Descripción", lote AS "Lote", cantidad AS "Cant.", proveedor AS "Proveedor Asignado", estado_global AS "Estado"
            FROM productos 
            WHERE motivo_informe = 'Alerta Sanitaria' AND estado_global IN ('CUARENTENA', 'Alerta Notificada al Proveedor')
        """, conn)

        if df_alertas.empty:
            st.info("No hay Alertas Sanitarias activas pendientes de gestión.")
        else:
            st.dataframe(df_alertas, hide_index=True, use_container_width=True)
            id_alerta = st.selectbox("Seleccione ID de Alerta a gestionar", df_alertas['ID'].tolist())
            prod_alerta = conn.cursor().execute("SELECT * FROM productos WHERE id=?", (id_alerta,)).fetchone()
            
            if prod_alerta:
                st.markdown(f"### Redacción de Anexo II para: {prod_alerta['descripcion']}")
                with st.form("form_anexo_ii"):
                    col1, col2 = st.columns(2)
                    with col1:
                        principio_activo = st.text_input("Principio Activo", value=prod_alerta['principio_activo'] or "")
                        titular = st.text_input("Titular del Producto *", value=prod_alerta['titular_registro'] or "")
                        reg_sanitario = st.text_input("N° de Registro Sanitario *", value=prod_alerta['registro_sanitario'] or "")
                    
                    with col2:
                        idx_prov = PROVEEDORES_OFICIALES.index(prod_alerta['proveedor']) if prod_alerta['proveedor'] in PROVEEDORES_OFICIALES else 0
                        proveedor = st.selectbox("Proveedor *", PROVEEDORES_OFICIALES, index=idx_prov)
                        rep_legal = st.text_input("Representante Legal Hospital", value="Falta Rellenar")
                        dir_tecnico = st.text_input("Director Técnico / QF Responsable", value=user_info['nombre_completo'])
                    
                    obs_alerta = st.text_area("Otras Observaciones", value=prod_alerta['observacion_paso2'] or "")
                    
                    if st.form_submit_button("💾 Guardar Datos para Anexo II"):
                        cursor = conn.cursor()
                        cursor.execute("""
                            UPDATE productos 
                            SET principio_activo=?, titular_registro=?, registro_sanitario=?, proveedor=?, observacion_paso2=?
                            WHERE id=?
                        """, (principio_activo, titular, reg_sanitario, proveedor, obs_alerta, id_alerta))
                        conn.commit()
                        st.success("Datos guardados.")
                        time.sleep(1)
                        st.rerun()

                if prod_alerta['titular_registro'] and prod_alerta['registro_sanitario']:
                    datos_docx = {
                        "descripcion": prod_alerta['descripcion'],
                        "principio_activo": prod_alerta['principio_activo'],
                        "titular": prod_alerta['titular_registro'],
                        "registro_sanitario": prod_alerta['registro_sanitario'],
                        "lote": prod_alerta['lote'],
                        "representante_legal": "Representante Hospital", 
                        "director_tecnico": user_info['nombre_completo'],
                        "proveedor": prod_alerta['proveedor'],
                        "cantidad": prod_alerta['cantidad'],
                        "unidad": prod_alerta['unidad'],
                        "observaciones": prod_alerta['observacion_paso2']
                    }
                    archivo_docx = generar_anexo_ii_docx(datos_docx)
                    
                    st.download_button(
                        label="📄 Descargar Documento Anexo II (.docx)",
                        data=archivo_docx,
                        file_name=f"ANEXO_II_{prod_alerta['lote']}_{prod_alerta['codigo_reyimen']}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                    
                    if prod_alerta['estado_global'] != 'Alerta Notificada al Proveedor':
                        if st.button("✅ Marcar como 'Notificado al Proveedor'"):
                            conn.cursor().execute("UPDATE productos SET estado_global='Alerta Notificada al Proveedor', paso_actual=5 WHERE id=?", (id_alerta,))
                            conn.commit()
                            st.success("Trámite avanzado. Ahora está listo para su Cierre Final en el Paso 5.")
                            time.sleep(1.5)
                            st.rerun()
                else:
                    st.info("Complete los datos requeridos (Titular, Registro Sanitario) y guarde para habilitar la descarga del Anexo II.")

    # --- CARGA MASIVA ---
    elif tab_seleccionada == "📤 Carga Masiva":
        st.header("📤 Carga Masiva de Productos (Paso 1)")
        st.markdown("### 1. Descargar Plantilla Modelo")
        st.caption("Descarga la planilla oficial formateada para que el bodeguero complete los datos de pronto vencimiento.")
        
        df_plantilla = pd.DataFrame([{
            "BODEGA ORIGEN": "Bodega AZ09 (Fármacos)",
            "TIPO PRODUCTO": "Fármaco",
            "CÓDIGO REYIMEN": "1365",
            "DESCRIPCIÓN": "BUPIVACAINA 0,50 % SOLUCION INYECTABLE 10 ML",
            "TIPO COMPRA": "CENABAST",
            "UNIDAD": "FRASCO",
            "CANTIDAD": 100,
            "MOTIVO INFORME": "Gestión pronto vencimiento",
            "FECHA VENCIMIENTO": "2026-10-31",
            "LOTE": "L12345"
        }])
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_plantilla.to_excel(writer, index=False, sheet_name='Plantilla_Carga')
            ws = writer.sheets['Plantilla_Carga']

            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
            from openpyxl.utils import get_column_letter
            from openpyxl.worksheet.datavalidation import DataValidation

            header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
            header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            thin_border = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))

            ws.row_dimensions[1].height = 28
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = header_alignment

            for col in ws.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    cell.border = thin_border
                    val_str = str(cell.value or '')
                    if len(val_str) > max_len: max_len = len(val_str)
                ws.column_dimensions[col_letter].width = max(max_len + 5, 18)

            dv_tipo = DataValidation(type="list", formula1='"Fármaco,Insumo"', allow_blank=True, showDropDown=True)
            ws.add_data_validation(dv_tipo)
            dv_tipo.add("B2:B2000")

            dv_compra = DataValidation(type="list", formula1='"CENABAST,Compra propia"', allow_blank=True, showDropDown=True)
            ws.add_data_validation(dv_compra)
            dv_compra.add("E2:E2000")

            dv_motivo = DataValidation(type="list", formula1='"Gestión pronto vencimiento,Alerta Sanitaria,Falla de calidad"', allow_blank=True, showDropDown=True)
            ws.add_data_validation(dv_motivo)
            dv_motivo.add("H2:H2000")

        excel_data = output.getvalue()
        st.download_button(label="📥 Descargar Plantilla Excel", data=excel_data, file_name="Plantilla_Carga_Masiva_Bodega.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.markdown(f"<hr style='margin-top: 1rem; margin-bottom: 0.5rem; border: none; border-top: 1px solid {borde_color};' />", unsafe_allow_html=True)

        st.markdown("### 2. Subir Archivo Completado")
        uploaded = st.file_uploader("Subir archivo Excel o CSV", type=["xlsx", "xls", "csv"])
        if uploaded and st.button("🚀 Procesar e Ingresar Productos"):
            ok, msg = procesar_carga_masiva(uploaded, user_info['usuario'])
            if ok:
                st.success(msg)
                time.sleep(1.5)
                st.rerun()
            else:
                st.error(msg)

    # --- PASO 2 ---
    elif tab_seleccionada == "⚖️ 2. Canjes (Jefatura)":
        st.header("⚖️ Paso 2 — Gestión de Canjes (Jefatura)")
        df = pd.read_sql_query("""
            SELECT id AS ID, bodega_origen AS Bodega, codigo_reyimen AS Código, descripcion AS Descripción, tipo_documento AS Compra, cantidad AS Cant, lote AS Lote, vencimiento AS Vencimiento 
            FROM productos WHERE paso_actual = 2 AND estado_global = 'En trámite' AND motivo_informe != 'Alerta Sanitaria'
        """, conn)
        
        if df.empty:
            st.info("No hay productos pendientes de canje comercial.")
        else:
            hoy = datetime.now().date()
            def calcular_meses(venc_str):
                try:
                    venc = datetime.strptime(str(venc_str), "%Y-%m-%d").date()
                    return max(0.0, round((venc - hoy).days / 30.44, 1))
                except: return 0.0

            df["Meses Vencer"] = df["Vencimiento"].apply(calcular_meses)
            st.dataframe(df, hide_index=True, use_container_width=True)
            prod_id = st.selectbox("Seleccione ID de Producto a gestionar", df['ID'].tolist())
            
            with st.form("form_paso2"):
                aplica_canje = st.selectbox("¿Aplica Canje? *", ["Aplica", "No aplica", "Revisión área de registro - producto compra propia"])
                if st.form_submit_button("Avanzar a Paso 3"):
                    conn.cursor().execute("UPDATE productos SET estado_canje=?, fecha_paso2=?, paso_actual=3 WHERE id=?", (aplica_canje, str(datetime.now().date()), prod_id))
                    conn.commit()
                    st.success("Producto avanzado al Paso 3.")
                    time.sleep(1.5)
                    st.rerun()

    # --- PASO 3 ---
    elif tab_seleccionada == "🚚 3. Registro/Prov.":
        st.header("🚚 Paso 3 — Área de Registro / Proveedor")
        
        tab_p3_nuevo, tab_p3_seguimiento = st.tabs(["➕ Pendientes de Ingreso", "🔄 Seguimiento de Trámites"])

        with tab_p3_nuevo:
            df_p3_nuevo = pd.read_sql_query("SELECT id AS ID, codigo_reyimen AS Código, descripcion AS Descripción, cantidad AS Cant, lote AS Lote, estado_canje AS Canje FROM productos WHERE paso_actual = 3 AND estado_global = 'En trámite'", conn)
            if df_p3_nuevo.empty:
                st.info("No hay nuevos productos pendientes de ingresar.")
            else:
                st.dataframe(df_p3_nuevo, hide_index=True, use_container_width=True)
                prod_id = st.selectbox("Seleccione ID a ingresar trámite", df_p3_nuevo['ID'].tolist())
                with st.form("form_paso3_ingreso"):
                    proveedor = st.selectbox("Proveedor Oficial *", PROVEEDORES_OFICIALES)
                    tipo_doc = st.selectbox("Tipo de Documento *", TIPOS_DOCUMENTO)
                    num_doc = st.text_input("N° Documento / OC *")
                    tramite = st.selectbox("Estado del trámite *", ESTADOS_TRAMITE_PROVEEDOR)
                    obs = st.text_area("Observaciones")
                    if st.form_submit_button("Avanzar a Paso 4"):
                        conn.cursor().execute("UPDATE productos SET proveedor=?, tipo_documento=?, numero_documento_oc=?, tramite_proveedor=?, fecha_paso3=?, observacion_paso3=?, paso_actual=4 WHERE id=?", (proveedor, tipo_doc, num_doc, tramite, str(datetime.now().date()), obs, prod_id))
                        conn.commit()
                        st.success("Avanzado al Paso 4.")
                        time.sleep(1.5)
                        st.rerun()

        with tab_p3_seguimiento:
            df_p3_seg = pd.read_sql_query("SELECT id AS ID, descripcion AS Descripción, proveedor AS Proveedor, tramite_proveedor AS Trámite, numero_documento_oc AS Doc FROM productos WHERE paso_actual >= 3 AND estado_global = 'En trámite' AND motivo_informe != 'Alerta Sanitaria'", conn)
            if df_p3_seg.empty:
                st.info("No hay trámites activos en seguimiento.")
            else:
                st.dataframe(df_p3_seg, hide_index=True, use_container_width=True)
                id_seg = st.selectbox("Seleccione ID a actualizar", df_p3_seg['ID'].tolist())
                prod_seg_data = conn.cursor().execute("SELECT * FROM productos WHERE id=?", (id_seg,)).fetchone()
                if prod_seg_data:
                    with st.form("form_paso3_actualizar"):
                        idx_prov = PROVEEDORES_OFICIALES.index(prod_seg_data['proveedor']) if prod_seg_data['proveedor'] in PROVEEDORES_OFICIALES else 0
                        idx_tram = ESTADOS_TRAMITE_PROVEEDOR.index(prod_seg_data['tramite_proveedor']) if prod_seg_data['tramite_proveedor'] in ESTADOS_TRAMITE_PROVEEDOR else 0
                        u_prov = st.selectbox("Proveedor", PROVEEDORES_OFICIALES, index=idx_prov)
                        u_doc = st.text_input("N° Doc", value=prod_seg_data['numero_documento_oc'] or "")
                        u_tram = st.selectbox("Estado Trámite", ESTADOS_TRAMITE_PROVEEDOR, index=idx_tram)
                        u_obs = st.text_area("Observación", value=prod_seg_data['observacion_paso3'] or "")
                        if st.form_submit_button("Guardar Actualización"):
                            conn.cursor().execute("UPDATE productos SET proveedor=?, numero_documento_oc=?, tramite_proveedor=?, observacion_paso3=? WHERE id=?", (u_prov, u_doc, u_tram, u_obs, id_seg))
                            conn.commit()
                            st.success("Actualizado.")
                            time.sleep(1)
                            st.rerun()

    # --- PASO 4 ---
    elif tab_seleccionada == "📦 4. Bulto/Ubicación":
        st.header("📦 Paso 4 — Gestión de Bulto y Ubicaciones")
        df = pd.read_sql_query("SELECT id AS ID, descripcion AS Descripción, lote AS Lote, proveedor AS Proveedor FROM productos WHERE paso_actual = 4 AND estado_global = 'En trámite' AND motivo_informe != 'Alerta Sanitaria'", conn)
        if df.empty:
            st.info("No hay productos pendientes para asignar bulto.")
        else:
            st.dataframe(df, hide_index=True, use_container_width=True)
            prod_id = st.selectbox("ID de Producto", df['ID'].tolist())
            with st.form("form_p4"):
                ub_fisica = st.selectbox("Ubicación Física *", BODEGAS_PASO4)
                ub_comp = st.selectbox("Ubicación Computacional *", BODEGAS_PASO4)
                bulto = st.text_input("N° Bulto")
                obs = st.text_area("Observaciones Paso 4")
                if st.form_submit_button("Avanzar a Paso 5"):
                    conn.cursor().execute("UPDATE productos SET ubicacion_fisica=?, ubicacion_computacional=?, numero_bulto=?, observacion_paso4=?, paso_actual=5 WHERE id=?", (ub_fisica, ub_comp, bulto, obs, prod_id))
                    conn.commit()
                    st.success("Avanzado a Paso 5.")
                    time.sleep(1)
                    st.rerun()

    # --- PASO 5 ---
    elif tab_seleccionada == "📜 5. Resolución/Cierre":
        st.header("📜 Paso 5 — Resolución y Cierre")
        tab_p5_cierre, tab_p5_sin_canje = st.tabs(["🔒 Cierre General de Proceso", "🟢 Productos Sin Carta de Canje"])

        with tab_p5_cierre:
            df_p5 = pd.read_sql_query("SELECT id AS ID, descripcion AS Descripción, lote AS Lote, proveedor AS Proveedor, numero_bulto AS Bulto, estado_global AS Estado FROM productos WHERE paso_actual = 5 AND estado_global != 'Concluido'", conn)
            if df_p5.empty:
                st.info("No hay productos pendientes para cierre.")
            else:
                st.dataframe(df_p5, hide_index=True, use_container_width=True)
                prod_id = st.selectbox("ID de Producto a CERRAR", df_p5['ID'].tolist())
                with st.form("form_p5"):
                    num_res = st.text_input("N° de Resolución ISP o Acta Destrucción *")
                    estado_fin = st.selectbox("Estado Final *", ["Retirado por Alerta Sanitaria", "Destruido", "Concluido", "Dado de baja", "Canjeado"])
                    obs = st.text_area("Resolución del caso")
                    if st.form_submit_button("Finalizar y Archivar"):
                        conn.cursor().execute("UPDATE productos SET resolucion_numero=?, estado_final=?, observacion_paso5=?, estado_global='Concluido' WHERE id=?", (num_res, estado_fin, obs, prod_id))
                        conn.commit()
                        st.success("Trámite finalizado e histórico.")
                        time.sleep(1.5)
                        st.rerun()

        with tab_p5_sin_canje:
            df_p5_sc = pd.read_sql_query("SELECT id AS ID, descripcion AS Descripción, lote AS Lote, estado_canje AS Canje FROM productos WHERE (estado_canje = 'No aplica' OR estado_canje LIKE '%compra propia%') AND estado_global = 'En trámite' AND motivo_informe != 'Alerta Sanitaria'", conn)
            if df_p5_sc.empty:
                st.info("No hay productos marcados sin carta de canje pendientes.")
            else:
                st.dataframe(df_p5_sc, hide_index=True, use_container_width=True)
                id_sc = st.selectbox("Seleccione ID para Difusión", df_p5_sc['ID'].tolist())
                prod_sc_data = conn.cursor().execute("SELECT * FROM productos WHERE id=?", (id_sc,)).fetchone()
                if prod_sc_data:
                    with st.form("form_paso5_sin_canje"):
                        c1, c2 = st.columns(2)
                        with c1: difusion_sel = st.selectbox("DIFUSIÓN A LA RED", OPCIONES_DIFUSION_RED)
                        with c2: redistribucion_sel = st.selectbox("REDISTRIBUCIÓN STOCK", OPCIONES_REDISTRIBUCION_STOCK)
                        obs_sc = st.text_area("Observaciones", value=prod_sc_data['observacion_paso5'] or "")
                        if st.form_submit_button("Guardar Gestión"):
                            conn.cursor().execute("UPDATE productos SET tipo_gestion_canje=?, observacion_paso2=?, observacion_paso5=? WHERE id=?", (difusion_sel, redistribucion_sel, obs_sc, id_sc))
                            conn.commit()
                            st.success("Gestión registrada.")
                            time.sleep(1)
                            st.rerun()

    # --- CONSOLIDADO & DASHBOARD ---
    elif tab_seleccionada == "🔍 Consolidado General":
        st.header("🔍 Consolidado")
        tab_cons_todos, tab_cons_historico = st.tabs(["🌐 Todos", "📁 Histórico Concluidos"])
        with tab_cons_todos:
            df_cons = pd.read_sql_query("SELECT id AS ID, codigo_reyimen AS Código, descripcion AS Descripción, bodega_origen AS Bodega, cantidad AS Cant, lote AS Lote, vencimiento AS Venc, estado_global AS Estado, proveedor AS Proveedor FROM productos", conn)
            st.dataframe(df_cons, hide_index=True, use_container_width=True)
        with tab_cons_historico:
            df_hist = pd.read_sql_query("SELECT id AS ID, codigo_reyimen AS Código, descripcion AS Descripción, lote AS Lote, estado_final AS Resolución, fecha_resolucion AS Fecha_Cierre FROM productos WHERE estado_global = 'Concluido'", conn)
            st.dataframe(df_hist, hide_index=True, use_container_width=True)

    elif tab_seleccionada == "📊 Dashboard / Análisis":
        st.header("📊 Dashboard")
        df_all = pd.read_sql_query("SELECT * FROM productos", conn)
        if df_all.empty:
            st.info("No hay datos suficientes para generar análisis.")
        else:
            tot_reg = len(df_all)
            concluidos = len(df_all[df_all['estado_global'] == 'Concluido'])
            tramite = len(df_all[df_all['estado_global'].isin(['En trámite', 'CUARENTENA', 'Alerta Notificada al Proveedor'])])
            canjeados_df = df_all[df_all['estado_final'] == 'Canjeado']
            unidades_canjeadas = canjeados_df['cantidad'].sum() if not canjeados_df.empty else 0
            
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total Trámites Registrados", tot_reg)
            m2.metric("En Trámite Activo", tramite)
            m3.metric("Trámites Concluidos", concluidos)
            m4.metric("Unidades Canjeadas/Rescatadas", int(unidades_canjeadas))
            
            st.markdown(f"<hr style='margin-top: 0.5rem; margin-bottom: 1.5rem; border: none; border-top: 1px solid {TEMAS_COLOR_CLAROS[st.session_state['tema_seleccionado']]['border']};' />", unsafe_allow_html=True)
            
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                st.subheader("📦 Trámites por Bodega")
                st.dataframe(df_all.groupby('bodega_origen').size().reset_index(name='TOTAL'), hide_index=True, use_container_width=True)
                st.subheader("⚖️ Estado de Canjes")
                st.dataframe(df_all.groupby('estado_canje').size().reset_index(name='TOTAL'), hide_index=True, use_container_width=True)
            with col_d2:
                st.subheader("⚠️ Motivos de Informe")
                st.dataframe(df_all.groupby('motivo_informe').size().reset_index(name='TOTAL'), hide_index=True, use_container_width=True)
                st.subheader("🚚 Avance de Trámite")
                st.dataframe(df_all[df_all['tramite_proveedor'].notnull()].groupby('tramite_proveedor').size().reset_index(name='TOTAL'), hide_index=True, use_container_width=True)

    elif tab_seleccionada == "👥 Gestión de Usuarios":
        st.header("👥 Gestión de Usuarios")
        tab_crear, tab_editar = st.tabs(["➕ Crear Usuario", "✏️ Editar / Eliminar"])
        with tab_crear:
            with st.form("form_nuevo_usuario"):
                u_user, u_pass = st.text_input("Usuario"), st.text_input("Contraseña", type="password")
                u_nombre, u_rol = st.text_input("Nombre"), st.selectbox("Rol", ["admin", "jefatura", "registro", "bodega"])
                if st.form_submit_button("Crear"):
                    conn.cursor().execute("INSERT INTO usuarios (usuario, password, rol, nombre_completo) VALUES (?, ?, ?, ?)", (u_user, u_pass, u_rol, u_nombre))
                    conn.commit()
                    st.success("Usuario creado.")
                    time.sleep(1)
                    st.rerun()
            st.dataframe(pd.read_sql_query("SELECT id, usuario, rol, nombre_completo, estado FROM usuarios", conn), hide_index=True, use_container_width=True)
        with tab_editar:
            df_users_edit = pd.read_sql_query("SELECT * FROM usuarios", conn)
            if not df_users_edit.empty:
                id_mod = st.selectbox("ID a Modificar", df_users_edit['id'].tolist())
                user_data = conn.cursor().execute("SELECT * FROM usuarios WHERE id=?", (id_mod,)).fetchone()
                if user_data:
                    with st.form("form_editar_usuario"):
                        new_u_nombre, new_u_user = st.text_input("Nombre", value=user_data['nombre_completo']), st.text_input("Usuario", value=user_data['usuario'])
                        new_u_rol, new_u_estado = st.selectbox("Rol", ["admin", "jefatura", "registro", "bodega"]), st.selectbox("Estado", ["Activo", "Inactivo"])
                        new_u_pass = st.text_input("Nueva Contraseña (Opcional)", type="password")
                        b1, b2 = st.columns(2)
                        if b1.form_submit_button("Guardar"):
                            if new_u_pass.strip(): conn.cursor().execute("UPDATE usuarios SET usuario=?, password=?, rol=?, nombre_completo=?, estado=? WHERE id=?", (new_u_user, new_u_pass, new_u_rol, new_u_nombre, new_u_estado, id_mod))
                            else: conn.cursor().execute("UPDATE usuarios SET usuario=?, rol=?, nombre_completo=?, estado=? WHERE id=?", (new_u_user, new_u_rol, new_u_nombre, new_u_estado, id_mod))
                            conn.commit()
                            st.success("Usuario actualizado.")
                            time.sleep(1)
                            st.rerun()
                        if b2.form_submit_button("Eliminar"):
                            conn.cursor().execute("DELETE FROM usuarios WHERE id=?", (id_mod,))
                            conn.commit()
                            st.warning("Usuario eliminado.")
                            time.sleep(1)
                            st.rerun()

    conn.close()
