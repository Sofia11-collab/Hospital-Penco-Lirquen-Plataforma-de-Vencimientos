"""
Módulo de Interfaz de Usuario para los 5 Pasos Operativos, Carga Masiva,
Gestión de Usuarios, Dashboard y Consolidado General.
"""
import io
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
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
    tabs_disponibles.append("🔍 Consolidado General")
    tabs_disponibles.append("📊 Dashboard / Análisis")
    
    if rol == "admin":
        tabs_disponibles.append("Gestión de Usuarios")

    tab_seleccionada = st.sidebar.radio("Navegación", tabs_disponibles)
    
    conn = get_connection()

    # --- PASO 1 ---
    if tab_seleccionada == "Paso 1: Informe Bodega":
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
                        st.rerun()

        with tab_p1_editar:
            st.subheader("Registros Ingresados en Paso 1 (Modificables)")
            df_p1 = pd.read_sql_query("""
                SELECT 
                    id AS ID, 
                    bodega_origen AS "Bodega Origen", 
                    codigo_reyimen AS "Código Reyimen", 
                    descripcion AS "Descripción", 
                    tipo_documento AS "Tipo de Compra", 
                    cantidad AS "Cantidad", 
                    lote AS "Lote", 
                    vencimiento AS "Vencimiento",
                    motivo_informe AS "Motivo Informe"
                FROM productos 
                WHERE paso_actual = 1 AND estado_global = 'En trámite'
            """, conn)

            if df_p1.empty:
                st.info("No hay registros pendientes en Paso 1 para modificar.")
            else:
                st.dataframe(df_p1, hide_index=True, use_container_width=True)
                id_mod = st.selectbox("Seleccione ID del registro a Modificar o Eliminar", df_p1['ID'].tolist())
                prod_data = conn.cursor().execute("SELECT * FROM productos WHERE id=?", (id_mod,)).fetchone()
                
                if prod_data:
                    with st.form("form_editar_p1"):
                        st.markdown(f"**Modificando Registro ID #{id_mod} — {prod_data['descripcion']}**")
                        col_e1, col_e2 = st.columns(2)
                        with col_e1:
                            idx_bodega = BODEGAS_OFICIALES.index(prod_data['bodega_origen']) if prod_data['bodega_origen'] in BODEGAS_OFICIALES else 0
                            new_bodega = st.selectbox("Bodega Origen *", BODEGAS_OFICIALES, index=idx_bodega)
                            new_tipo_compra = st.selectbox("Tipo de compra *", ["CENABAST", "Compra propia"], index=0 if prod_data['tipo_documento']=="CENABAST" else 1)
                            new_cantidad = st.number_input("Cantidad *", value=float(prod_data['cantidad']), min_value=0.0, step=1.0)
                        
                        with col_e2:
                            new_lote = st.text_input("Lote *", value=prod_data['lote'])
                            try:
                                fecha_init = datetime.strptime(prod_data['vencimiento'], "%Y-%m-%d").date()
                            except Exception:
                                fecha_init = datetime.now().date()
                                
                            new_vencimiento = st.date_input("Fecha de Vencimiento *", value=fecha_init)
                            motivos_opt = ["Gestión pronto vencimiento", "Alerta Sanitaria", "Falla de calidad"]
                            idx_mot = motivos_opt.index(prod_data['motivo_informe']) if prod_data['motivo_informe'] in motivos_opt else 0
                            new_motivo = st.selectbox("Motivo de informe *", motivos_opt, index=idx_mot)

                        btn_col1, btn_col2 = st.columns(2)
                        with btn_col1:
                            guardar_mod = st.form_submit_button("💾 Guardar Cambios")
                        with btn_col2:
                            eliminar_mod = st.form_submit_button("🗑️ Eliminar Registro")

                        if guardar_mod:
                            cursor = conn.cursor()
                            cursor.execute("""
                                UPDATE productos 
                                SET bodega_origen=?, tipo_documento=?, cantidad=?, lote=?, vencimiento=?, motivo_informe=?
                                WHERE id=?
                            """, (new_bodega, new_tipo_compra, new_cantidad, new_lote, str(new_vencimiento), new_motivo, id_mod))
                            conn.commit()
                            st.success(f"Registro #{id_mod} actualizado correctamente.")
                            st.rerun()

                        if eliminar_mod:
                            cursor = conn.cursor()
                            cursor.execute("DELETE FROM productos WHERE id=?", (id_mod,))
                            conn.commit()
                            st.warning(f"Registro #{id_mod} eliminado.")
                            st.rerun()

    # --- CARGA MASIVA ---
    elif tab_seleccionada == "Carga Masiva":
        st.header("📤 Carga Masiva de Productos (Paso 1)")
        
        st.markdown("### 1. Descargar Plantilla Modelo")
        st.caption("Descarga la planilla oficial para que el bodeguero complete los datos de pronto vencimiento.")
        
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
        excel_data = output.getvalue()

        st.download_button(
            label="📥 Descargar Plantilla Excel (.xlsx)",
            data=excel_data,
            file_name="Plantilla_Carga_Masiva_Bodega.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.divider()

        st.markdown("### 2. Subir Archivo Completado")
        uploaded = st.file_uploader("Subir archivo Excel o CSV", type=["xlsx", "xls", "csv"])
        if uploaded and st.button("🚀 Procesar e Ingresar Productos"):
            ok, msg = procesar_carga_masiva(uploaded, user_info['usuario'])
            if ok:
                st.success(msg)
                st.rerun()
            else:
                st.error(msg)

    # --- PASO 2 ---
    elif tab_seleccionada == "Paso 2: Canjes (Jefatura)":
        st.header("⚖️ Paso 2 — Gestión de Canjes (Jefatura)")
        df = pd.read_sql_query("""
            SELECT 
                id AS ID, 
                bodega_origen AS Bodega, 
                codigo_reyimen AS Código, 
                descripcion AS Descripción, 
                tipo_documento AS Compra, 
                cantidad AS Cant, 
                lote AS Lote, 
                vencimiento AS Vencimiento 
            FROM productos 
            WHERE paso_actual = 1 AND estado_global = 'En trámite'
        """, conn)
        
        if df.empty:
            st.info("No hay productos pendientes en Paso 1.")
        else:
            hoy = datetime.now().date()
            
            def calcular_meses(venc_str):
                try:
                    venc = datetime.strptime(str(venc_str), "%Y-%m-%d").date()
                    dias = (venc - hoy).days
                    meses = round(dias / 30.44, 1)
                    return max(0.0, meses)
                except Exception:
                    return 0.0

            def calcular_fecha_limite(venc_str):
                try:
                    venc = datetime.strptime(str(venc_str), "%Y-%m-%d").date()
                    limite = venc - timedelta(days=60)
                    return str(limite)
                except Exception:
                    return ""

            df["Meses Vencer"] = df["Vencimiento"].apply(calcular_meses)
            df["Límite Retiro (60d)"] = df["Vencimiento"].apply(calcular_fecha_limite)

            st.dataframe(
                df, 
                hide_index=True, 
                use_container_width=True,
                column_config={
                    "ID": st.column_config.NumberColumn(width="small"),
                    "Bodega": st.column_config.TextColumn(width="small"),
                    "Código": st.column_config.TextColumn(width="small"),
                    "Descripción": st.column_config.TextColumn(width="large"),
                    "Compra": st.column_config.TextColumn(width="small"),
                    "Cant": st.column_config.NumberColumn(width="small"),
                    "Lote": st.column_config.TextColumn(width="small"),
                    "Vencimiento": st.column_config.TextColumn(width="small"),
                    "Meses Vencer": st.column_config.NumberColumn(width="small"),
                    "Límite Retiro (60d)": st.column_config.TextColumn(width="small"),
                }
            )
            prod_id = st.selectbox("Seleccione ID de Producto a gestionar", df['ID'].tolist())
            
            with st.form("form_paso2"):
                aplica_canje = st.selectbox(
                    "¿Aplica Canje? *", 
                    [
                        "Aplica", 
                        "No aplica", 
                        "Revisión área de registro - producto compra propia"
                    ]
                )
                
                if st.form_submit_button("Avanzar a Paso 3"):
                    cursor = conn.cursor()
                    cursor.execute("""
                    UPDATE productos SET estado_canje=?, fecha_paso2=?, paso_actual=2 WHERE id=?
                    """, (aplica_canje, str(datetime.now().date()), prod_id))
                    conn.commit()
                    st.success("Producto avanzado al Paso 3.")
                    st.rerun()

    # --- PASO 3 ---
    elif tab_seleccionada == "Paso 3: Registro/Proveedor":
        st.header("🚚 Paso 3 — Área de Registro / Proveedor")
        
        tab_p3_nuevo, tab_p3_seguimiento = st.tabs(["➕ Pendientes de Ingreso (Paso 2)", "🔄 Seguimiento y Actualización de Trámites"])

        with tab_p3_nuevo:
            df_p3_nuevo = pd.read_sql_query("""
                SELECT 
                    id AS ID, 
                    codigo_reyimen AS "Código Reyimen", 
                    descripcion AS "Descripción", 
                    cantidad AS "Cantidad",
                    lote AS "Lote", 
                    estado_canje AS "Estado Canje" 
                FROM productos 
                WHERE paso_actual = 2 AND estado_global = 'En trámite'
            """, conn)
            
            if df_p3_nuevo.empty:
                st.info("No hay nuevos productos pendientes de ingresar en Paso 3.")
            else:
                st.dataframe(df_p3_nuevo, hide_index=True, use_container_width=True)
                prod_id = st.selectbox("Seleccione ID de Producto a ingresar trámite", df_p3_nuevo['ID'].tolist())
                
                with st.form("form_paso3_ingreso"):
                    proveedor = st.selectbox("Proveedor Oficial *", PROVEEDORES_OFICIALES)
                    tipo_doc = st.selectbox("Tipo de Documento *", TIPOS_DOCUMENTO)
                    num_doc = st.text_input("Número de Documento / Orden de Compra *")
                    tramite = st.selectbox("Estado del trámite con proveedor *", ESTADOS_TRAMITE_PROVEEDOR)
                    obs = st.text_area("Observaciones Paso 3")
                    
                    if st.form_submit_button("Avanzar a Paso 4"):
                        cursor = conn.cursor()
                        cursor.execute("""
                        UPDATE productos SET proveedor=?, tipo_documento=?, numero_documento_oc=?, tramite_proveedor=?, fecha_paso3=?, observacion_paso3=?, paso_actual=3 WHERE id=?
                        """, (proveedor, tipo_doc, num_doc, tramite, str(datetime.now().date()), obs, prod_id))
                        conn.commit()
                        st.success("Producto registrado y avanzado al Paso 4.")
                        st.rerun()

        with tab_p3_seguimiento:
            st.subheader("Gestión Continua de Productos en Trámite (No Concluidos)")
            df_p3_seg = pd.read_sql_query("""
                SELECT 
                    id AS ID, 
                    codigo_reyimen AS "Código", 
                    descripcion AS "Descripción", 
                    proveedor AS "Proveedor", 
                    tramite_proveedor AS "Estado Trámite Actual",
                    numero_documento_oc AS "N° Doc/OC",
                    paso_actual AS "Paso Actual"
                FROM productos 
                WHERE paso_actual >= 3 AND estado_global = 'En trámite'
            """, conn)

            if df_p3_seg.empty:
                st.info("No hay productos con trámites activos en seguimiento.")
            else:
                st.dataframe(df_p3_seg, hide_index=True, use_container_width=True)
                
                id_seg = st.selectbox("Seleccione ID del Producto para actualizar su Estado de Trámite", df_p3_seg['ID'].tolist())
                prod_seg_data = conn.cursor().execute("SELECT * FROM productos WHERE id=?", (id_seg,)).fetchone()
                
                if prod_seg_data:
                    with st.form("form_paso3_actualizar"):
                        st.markdown(f"**Actualizando Estado para ID #{id_seg} — {prod_seg_data['descripcion']}**")
                        
                        idx_prov = PROVEEDORES_OFICIALES.index(prod_seg_data['proveedor']) if prod_seg_data['proveedor'] in PROVEEDORES_OFICIALES else 0
                        idx_tram = ESTADOS_TRAMITE_PROVEEDOR.index(prod_seg_data['tramite_proveedor']) if prod_seg_data['tramite_proveedor'] in ESTADOS_TRAMITE_PROVEEDOR else 0
                        
                        u_proveedor = st.selectbox("Proveedor Oficial *", PROVEEDORES_OFICIALES, index=idx_prov)
                        u_num_doc = st.text_input("Número de Documento / Orden de Compra *", value=prod_seg_data['numero_documento_oc'] or "")
                        u_tramite = st.selectbox("Actualizar Estado del Trámite *", ESTADOS_TRAMITE_PROVEEDOR, index=idx_tram)
                        u_obs = st.text_area("Nueva Observación / Historial de Gestión", value=prod_seg_data['observacion_paso3'] or "")
                        
                        if st.form_submit_button("💾 Guardar Actualización de Trámite"):
                            cursor = conn.cursor()
                            cursor.execute("""
                                UPDATE productos 
                                SET proveedor=?, numero_documento_oc=?, tramite_proveedor=?, observacion_paso3=?
                                WHERE id=?
                            """, (u_proveedor, u_num_doc, u_tramite, u_obs, id_seg))
                            conn.commit()
                            st.success(f"Estado del trámite para ID #{id_seg} actualizado correctamente.")
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
            st.dataframe(df, hide_index=True, use_container_width=True)
            prod_id = st.selectbox("Seleccione ID de Producto a gestionar", df['ID'].tolist())
            
            with st.form("form_paso4"):
                ub_fisica = st.selectbox("Ubicación Física *", BODEGAS_PASO4)
                ub_comp = st.selectbox("Ubicación Computacional *", BODEGAS_PASO4)
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
        
        tab_p5_cierre, tab_p5_sin_canje = st.tabs(["🔒 Cierre General de Proceso", "🟢 Productos Sin Carta de Canje"])

        with tab_p5_cierre:
            df_p5 = pd.read_sql_query("""
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
            
            if df_p5.empty:
                st.info("No hay productos pendientes para cierre.")
            else:
                st.dataframe(df_p5, hide_index=True, use_container_width=True)
                prod_id = st.selectbox("Seleccione ID de Producto a CERRAR", df_p5['ID'].tolist())
                
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

        with tab_p5_sin_canje:
            st.subheader("🟢 Gestión de Productos Sin Carta de Canje")
            df_p5_sc = pd.read_sql_query("""
                SELECT 
                    id AS ID, 
                    codigo_reyimen AS "Código", 
                    descripcion AS "Descripción", 
                    lote AS "Lote", 
                    estado_canje AS "Estado Canje",
                    tipo_gestion_canje AS "Difusión a la Red",
                    observacion_paso2 AS "Redistribución Stock"
                FROM productos 
                WHERE (estado_canje = 'No aplica' OR estado_canje LIKE '%compra propia%') AND estado_global = 'En trámite'
            """, conn)

            if df_p5_sc.empty:
                st.info("No hay productos marcados sin carta de canje pendientes de gestión.")
            else:
                st.dataframe(df_p5_sc, hide_index=True, use_container_width=True)
                
                id_sc = st.selectbox("Seleccione ID de Producto para registrar Difusión o Redistribución", df_p5_sc['ID'].tolist())
                prod_sc_data = conn.cursor().execute("SELECT * FROM productos WHERE id=?", (id_sc,)).fetchone()
                
                if prod_sc_data:
                    with st.form("form_paso5_sin_canje"):
                        st.markdown(f"**Gestión de Producto ID #{id_sc} — {prod_sc_data['descripcion']}**")
                        
                        col_sc1, col_sc2 = st.columns(2)
                        with col_sc1:
                            difusion_sel = st.selectbox("DIFUSIÓN A LA RED *", OPCIONES_DIFUSION_RED)
                        with col_sc2:
                            redistribucion_sel = st.selectbox("REDISTRIBUCIÓN STOCK *", OPCIONES_REDISTRIBUCION_STOCK)
                        
                        obs_sc = st.text_area("Observaciones / Detalles de la Difusión", value=prod_sc_data['observacion_paso5'] or "")

                        if st.form_submit_button("💾 Guardar Gestión Sin Canje"):
                            cursor = conn.cursor()
                            cursor.execute("""
                                UPDATE productos 
                                SET tipo_gestion_canje=?, observacion_paso2=?, observacion_paso5=?
                                WHERE id=?
                            """, (difusion_sel, redistribucion_sel, obs_sc, id_sc))
                            conn.commit()
                            st.success(f"Gestión registrada para el Producto #{id_sc}.")
                            st.rerun()

    # --- HISTÓRICO ---
    elif tab_seleccionada == "Histórico Concluidos":
        st.header("📁 Registros Históricos Concluidos")
        df = pd.read_sql_query("SELECT * FROM productos WHERE estado_global = 'Concluido'", conn)
        st.dataframe(df, hide_index=True, use_container_width=True)

    # --- CONSOLIDADO GENERAL ---
    elif tab_seleccionada == "🔍 Consolidado General":
        st.header("🔍 Consolidado General de Todos los Productos")
        
        df_cons = pd.read_sql_query("SELECT * FROM productos", conn)
        
        if df_cons.empty:
            st.info("No hay productos registrados en el sistema.")
        else:
            def determinar_estado_general(row):
                if row['estado_global'] == 'Concluido':
                    return f"Concluido ({row['estado_final'] or 'Archivado'})"
                
                paso = row['paso_actual']
                if paso == 1:
                    return "Paso 1: Nuevo / Pendiente Jefatura"
                elif paso == 2:
                    return f"Paso 2: Canje Jefatura ({row['estado_canje'] or 'En evaluación'})"
                elif paso == 3:
                    return f"Paso 3: Trámite Proveedor ({row['tramite_proveedor'] or 'En gestión'})"
                elif paso == 4:
                    return "Paso 4: Bulto Asignado / Pendiente Cierre"
                else:
                    return "En trámite"

            df_cons['Estado Actual del Producto'] = df_cons.apply(determinar_estado_general, axis=1)

            col_f1, col_f2 = st.columns(2)
            with col_f1:
                search_text = st.text_input("🔎 Buscar por Descripción, Código Reyimen o Lote")
            with col_f2:
                estado_filtro = st.selectbox(
                    "Filtrar por Estado Global", 
                    ["Todos", "En trámite", "Concluido"]
                )

            if estado_filtro != "Todos":
                df_cons = df_cons[df_cons['estado_global'] == estado_filtro]
                
            if search_text:
                mask = (
                    df_cons['descripcion'].str.contains(search_text, case=False, na=False) |
                    df_cons['codigo_reyimen'].str.contains(search_text, case=False, na=False) |
                    df_cons['lote'].str.contains(search_text, case=False, na=False)
                )
                df_cons = df_cons[mask]

            df_mostrar = df_cons[[
                'id', 'codigo_reyimen', 'descripcion', 'bodega_origen', 
                'cantidad', 'lote', 'vencimiento', 'Estado Actual del Producto', 'proveedor'
            ]].copy()

            df_mostrar.columns = [
                'ID', 'CÓDIGO REYIMEN', 'DESCRIPCIÓN', 'BODEGA ORIGEN', 
                'CANTIDAD', 'LOTE', 'VENCIMIENTO', 'ESTADO ACTUAL', 'PROVEEDOR'
            ]

            st.dataframe(df_mostrar, hide_index=True, use_container_width=True)

    # --- DASHBOARD / ANÁLISIS ---
    elif tab_seleccionada == "📊 Dashboard / Análisis":
        st.header("📊 ANÁLISIS DE DATOS")
        
        df_all = pd.read_sql_query("SELECT * FROM productos", conn)
        
        if df_all.empty:
            st.info("No hay datos suficientes para generar análisis.")
        else:
            tot_reg = len(df_all)
            concluidos = len(df_all[df_all['estado_global'] == 'Concluido'])
            tramite = len(df_all[df_all['estado_global'] == 'En trámite'])
            
            canjeados_df = df_all[df_all['estado_final'] == 'Canjeado']
            unidades_canjeadas = canjeados_df['cantidad'].sum() if not canjeados_df.empty else 0
            
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total Trámites Registrados", tot_reg)
            m2.metric("En Trámite Activo", tramite)
            m3.metric("Trámites Concluidos", concluidos)
            m4.metric("Unidades Canjeadas/Rescatadas", int(unidades_canjeadas))
            
            st.divider()

            col_d1, col_d2 = st.columns(2)
            
            with col_d1:
                st.subheader("📦 Trámites por Bodega de Origen")
                df_bod = df_all.groupby('bodega_origen').size().reset_index(name='CANTIDAD TRÁMITES')
                df_bod.columns = ['BODEGA ORIGEN', 'CANTIDAD TRÁMITES']
                st.dataframe(df_bod, hide_index=True, use_container_width=True)

                st.subheader("⚖️ Estado de Canjes (Jefatura)")
                df_canje = df_all.groupby('estado_canje').size().reset_index(name='TOTAL REGISTROS')
                df_canje.columns = ['ESTADO CANJE', 'TOTAL REGISTROS']
                st.dataframe(df_canje, hide_index=True, use_container_width=True)

            with col_d2:
                st.subheader("⚠️ Motivos de Informe")
                df_mot = df_all.groupby('motivo_informe').size().reset_index(name='TOTAL REGISTROS')
                df_mot.columns = ['MOTIVO DE INFORME', 'TOTAL REGISTROS']
                st.dataframe(df_mot, hide_index=True, use_container_width=True)

                st.subheader("🚚 Avance por Estado de Trámite")
                df_tram = df_all[df_all['tramite_proveedor'].notnull()].groupby('tramite_proveedor').size().reset_index(name='TOTAL')
                df_tram.columns = ['ESTADO TRÁMITE', 'TOTAL']
                st.dataframe(df_tram, hide_index=True, use_container_width=True)

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
                usuario AS "USUARIO", 
                rol AS "ROL", 
                nombre_completo AS "NOMBRE COMPLETO", 
                estado AS "ESTADO" 
            FROM usuarios
        """, conn)
        st.dataframe(df_users, hide_index=True, use_container_width=True)

    conn.close()