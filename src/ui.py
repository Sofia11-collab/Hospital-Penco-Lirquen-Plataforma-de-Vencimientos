"""
Módulo de Interfaz de Usuario para los 5 Pasos Operativos y Gestión de Usuarios.
"""
import streamlit as st
import pandas as pd
from datetime import datetime
from src.database import get_connection
from src.catalogo import (
    get_catalogo, get_opciones_selectbox, buscar_por_etiqueta,
    BODEGAS_OFICIALES, PROVEEDORES_OFICIALES, TIPOS_DOCUMENTO, OPCION_MANUAL
)
from src.importer import procesar_carga_masiva

def render_ui(user_info: dict):
    st.sidebar.markdown(f"**Usuario:** {user_info['nombre_completo']}")
    st.sidebar.markdown(f"**Rol:** `{user_info['rol'].upper()}`")
    
    if st.sidebar.button("Cerrar Sesión"):
        st.session_state["logged_in"] = False
        st.rerun()

    rol = user_info['rol']
    
    # Menú de pestañas según rol
    tabs_disponibles = []
    if rol in ["admin", "bodega"]:
        tabs_disponibles.append("Paso 1: Informe Bodega")
        tabs_disponibles.append("Carga Masiva")
    if rol in ["admin", "jefatura"]:
        tabs_disponibles.append("Paso 2: Canjes (Jefatura)")
    if rol in ["admin", "registro"]:
        tabs_disponibles.append("Paso 3: Registro/Proveedor")
    if rol in ["admin", "bodega"]:
        tabs_disponibles.append("Paso 4: Bulto y Ubicación")
    if rol in ["admin", "jefatura"]:
        tabs_disponibles.append("Paso 5: Resolución/Cierre")
    
    tabs_disponibles.append("Histórico Concluidos")
    
    if rol == "admin":
        tabs_disponibles.append("Gestión de Usuarios")

    tab_seleccionada = st.sidebar.radio("Navegación", tabs_disponibles)
    
    conn = get_connection()

    # --- PASO 1 (REACTIVO + MOTIVO + TIPO DE COMPRA) ---
    if tab_seleccionada == "Paso 1: Informe Bodega":
        st.header("📋 Paso 1 — Informe de Bodega")
        catalogo = get_catalogo()
        opciones = get_opciones_selectbox(catalogo)
        
        c_busq1, c_busq2 = st.columns(2)
        with c_busq1:
            bodega = st.selectbox("Bodega/Farmacia Origen *", BODEGAS_OFICIALES)
            tipo_prod = st.selectbox("Tipo de producto", ["Fármaco", "Insumo"])
        
        with c_busq2:
            prod_sel = st.selectbox("Buscar Código Reyimen / Producto *", opciones, key="busqueda_producto")
            
        cod_auto, desc_auto, unidad_auto = "", "", ""
        item = buscar_por_etiqueta(catalogo, prod_sel)
        if item:
            cod_auto = item.get("codigo", "")
            desc_auto = item.get("descripcion", "")
            unidad_auto = item.get("unidad", "")

        with st.form("form_paso1"):
            col1, col2 = st.columns(2)
            with col1:
                codigo = st.text_input("Código Reyimen *", value=cod_auto)
                descripcion = st.text_input("Descripción *", value=desc_auto)
                tipo_compra = st.selectbox("Tipo de compra *", ["CENABAST", "Compra propia"])
            
            with col2:
                unidad = st.text_input("Unidad *", value=unidad_auto)
                cantidad = st.number_input("Cantidad *", min_value=0.0, step=1.0)
                motivo = st.selectbox(
                    "Motivo de informe *",
                    ["Gestión pronto vencimiento", "Alerta Sanitaria", "Falla de calidad"]
                )

            c3, c4 = st.columns(2)
            with c3:
                vencimiento = st.date_input("Fecha de Vencimiento *")
            with c4:
                lote = st.text_input("Lote *")
            
            if st.form_submit_button("Guardar Paso 1"):
                if not codigo or not descripcion or not lote:
                    st.error("Complete todos los campos obligatorios (*)")
                else:
                    cursor = conn.cursor()
                    cursor.execute("""
                    INSERT INTO productos (bodega_origen, tipo_producto, codigo_reyimen, descripcion, unidad, cantidad, vencimiento, lote, motivo_informe, tipo_documento, usuario_registro, paso_actual)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1)
                    """, (bodega, tipo_prod, codigo, descripcion, unidad, cantidad, str(vencimiento), lote, motivo, tipo_compra, user_info['usuario']))
                    conn.commit()
                    st.success("Producto registrado exitosamente en el Paso 1.")

    # --- CARGA MASIVA ---
    elif tab_seleccionada == "Carga Masiva":
        st.header("📤 Carga Masiva de Productos (Paso 1)")
        uploaded = st.file_uploader("Subir archivo Excel o CSV", type=["xlsx", "xls", "csv"])
        if uploaded and st.button("Procesar Archivo"):
            ok, msg = procesar_carga_masiva(uploaded, user_info['usuario'])
            if ok:
                st.success(msg)
            else:
                st.error(msg)

    # --- PASO 2 (FORMATO LIMPIO Y MAYÚSCULAS EN ENCABEZADOS) ---
    elif tab_seleccionada == "Paso 2: Canjes (Jefatura)":
        st.header("⚖️ Paso 2 — Gestión de Canjes (Jefatura)")
        df = pd.read_sql_query("""
            SELECT 
                id AS ID, 
                bodega_origen AS "Bodega Origen", 
                codigo_reyimen AS "Código Reyimen", 
                descripcion AS "Descripción", 
                tipo_documento AS "Tipo de Compra", 
                cantidad AS "Cantidad", 
                lote AS "Lote", 
                vencimiento AS "Vencimiento" 
            FROM productos 
            WHERE paso_actual = 1 AND estado_global = 'En trámite'
        """, conn)
        
        if df.empty:
            st.info("No hay productos pendientes en Paso 1.")
        else:
            st.dataframe(df, hide_index=True)
            prod_id = st.selectbox("Seleccione ID de Producto a gestionar", df['ID'].tolist())
            
            with st.form("form_paso2"):
                aplica_canje = st.selectbox("¿Aplica Canje? *", ["Aplica", "No aplica"])
                tipo_gestion = st.text_input("Tipo de gestión de canje")
                obs = st.text_area("Observaciones Paso 2")
                
                if st.form_submit_button("Avanzar a Paso 3"):
                    cursor = conn.cursor()
                    cursor.execute("""
                    UPDATE productos SET estado_canje=?, tipo_gestion_canje=?, fecha_paso2=?, observacion_paso2=?, paso_actual=2 WHERE id=?
                    """, (aplica_canje, tipo_gestion, str(datetime.now().date()), obs, prod_id))
                    conn.commit()
                    st.success("Producto avanzado al Paso 3.")
                    st.rerun()

    # --- PASO 3 ---
    elif tab_seleccionada == "Paso 3: Registro/Proveedor":
        st.header("🚚 Paso 3 — Área de Registro / Proveedor")
        df = pd.read_sql_query("""
            SELECT 
                id AS ID, 
                codigo_reyimen AS "Código Reyimen", 
                descripcion AS "Descripción", 
                lote AS "Lote", 
                estado_canje AS "Estado Canje" 
            FROM productos 
            WHERE paso_actual = 2 AND estado_global = 'En trámite'
        """, conn)
        
        if df.empty:
            st.info("No hay productos pendientes para gestionar con proveedor.")
        else:
            st.dataframe(df, hide_index=True)
            prod_id = st.selectbox("Seleccione ID de Producto a gestionar", df['ID'].tolist())
            
            with st.form("form_paso3"):
                proveedor = st.selectbox("Proveedor Oficial *", PROVEEDORES_OFICIALES)
                tipo_doc = st.selectbox("Tipo de Documento *", TIPOS_DOCUMENTO)
                num_doc = st.text_input("Número de Documento / Orden de Compra *")
                tramite = st.text_input("Estado del trámite con proveedor")
                obs = st.text_area("Observaciones Paso 3")
                
                if st.form_submit_button("Avanzar a Paso 4"):
                    cursor = conn.cursor()
                    cursor.execute("""
                    UPDATE productos SET proveedor=?, tipo_documento=?, numero_documento_oc=?, tramite_proveedor=?, fecha_paso3=?, observacion_paso3=?, paso_actual=3 WHERE id=?
                    """, (proveedor, tipo_doc, num_doc, tramite, str(datetime.now().date()), obs, prod_id))
                    conn.commit()
                    st.success("Producto avanzado al Paso 4.")
                    st.rerun()

    # --- PASO 4 ---
    elif tab_seleccionada == "Paso 4: Bulto y Ubicación":
        st.header("📦 Paso 4 — Gestión de Bulto y Ubicaciones")
        df = pd.read_sql_query("""
            SELECT 
                id AS ID, 
                codigo_reyimen AS "Código Reyimen", 
                descripcion AS "Descripción", 
                lote AS "Lote", 
                proveedor AS "Proveedor" 
            FROM productos 
            WHERE paso_actual = 3 AND estado_global = 'En trámite'
        """, conn)
        
        if df.empty:
            st.info("No hay productos pendientes para asignar bulto.")
        else:
            st.dataframe(df, hide_index=True)
            prod_id = st.selectbox("Seleccione ID de Producto a gestionar", df['ID'].tolist())
            
            with st.form("form_paso4"):
                ub_fisica = st.selectbox("Ubicación Física *", BODEGAS_OFICIALES)
                ub_comp = st.selectbox("Ubicación Computacional *", BODEGAS_OFICIALES)
                num_bulto = st.text_input("Número de Bulto *")
                obs = st.text_area("Observaciones Paso 4")
                
                if st.form_submit_button("Avanzar a Paso 5"):
                    cursor = conn.cursor()
                    cursor.execute("""
                    UPDATE productos SET ubicacion_fisica=?, ubicacion_computacional=?, numero_bulto=?, fecha_paso4=?, observacion_paso4=?, paso_actual=4 WHERE id=?
                    """, (ub_fisica, ub_comp, num_bulto, str(datetime.now().date()), obs, prod_id))
                    conn.commit()
                    st.success("Producto avanzado al Paso 5.")
                    st.rerun()

    # --- PASO 5 ---
    elif tab_seleccionada == "Paso 5: Resolución/Cierre":
        st.header("📜 Paso 5 — Resolución y Cierre")
        df = pd.read_sql_query("""
            SELECT 
                id AS ID, 
                codigo_reyimen AS "Código Reyimen", 
                descripcion AS "Descripción", 
                lote AS "Lote", 
                proveedor AS "Proveedor", 
                numero_bulto AS "Número Bulto" 
            FROM productos 
            WHERE paso_actual = 4 AND estado_global = 'En trámite'
        """, conn)
        
        if df.empty:
            st.info("No hay productos pendientes para cierre.")
        else:
            st.dataframe(df, hide_index=True)
            prod_id = st.selectbox("Seleccione ID de Producto a CERRAR", df['ID'].tolist())
            
            with st.form("form_paso5"):
                num_res = st.text_input("Número de Resolución *")
                estado_fin = st.selectbox("Estado Final *", ["Concluido", "Dado de baja", "Canjeado"])
                obs = st.text_area("Observaciones Paso 5")
                
                if st.form_submit_button("Finalizar Proceso"):
                    cursor = conn.cursor()
                    cursor.execute("""
                    UPDATE productos SET resolucion_numero=?, estado_final=?, fecha_resolucion=?, observacion_paso5=?, paso_actual=5, estado_global='Concluido' WHERE id=?
                    """, (num_res, estado_fin, str(datetime.now().date()), obs, prod_id))
                    conn.commit()
                    st.success("Producto CONCLUIDO con éxito. Se ha archivado al histórico.")
                    st.rerun()

    # --- HISTÓRICO ---
    elif tab_seleccionada == "Histórico Concluidos":
        st.header("📁 Registros Históricos Concluidos")
        df = pd.read_sql_query("SELECT * FROM productos WHERE estado_global = 'Concluido'", conn)
        st.dataframe(df, hide_index=True)

    # --- GESTIÓN DE USUARIOS (ADMIN) ---
    elif tab_seleccionada == "Gestión de Usuarios":
        st.header("👥 Módulo de Gestión de Usuarios")
        
        with st.form("form_nuevo_usuario"):
            st.subheader("Crear Nuevo Usuario")
            u_user = st.text_input("Usuario")
            u_pass = st.text_input("Contraseña", type="password")
            u_nombre = st.text_input("Nombre Completo")
            u_rol = st.selectbox("Rol de Acceso", ["admin", "jefatura", "registro", "bodega"])
            
            if st.form_submit_button("Crear Usuario"):
                if u_user and u_pass:
                    try:
                        cursor = conn.cursor()
                        cursor.execute("INSERT INTO usuarios (usuario, password, rol, nombre_completo) VALUES (?, ?, ?, ?)", (u_user, u_pass, u_rol, u_nombre))
                        conn.commit()
                        st.success("Usuario creado exitosamente.")
                    except Exception as e:
                        st.error(f"Error al crear usuario: {e}")

        st.subheader("Usuarios Registrados")
        df_users = pd.read_sql_query("""
            SELECT 
                id AS ID, 
                usuario AS "Usuario", 
                rol AS "Rol", 
                nombre_completo AS "Nombre Completo", 
                estado AS "Estado" 
            FROM usuarios
        """, conn)
        st.dataframe(df_users, hide_index=True)

    conn.close()