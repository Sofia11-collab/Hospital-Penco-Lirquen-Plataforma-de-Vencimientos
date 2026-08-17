"""
Interfaz del Paso 1: Informe Bodega.
Firma de las funciones públicas:
    render_step1_form(catalogo: list[dict], conn) -> None
"""

import streamlit as st
from datetime import date
import sqlite3
from typing import List, Dict

from src.catalogo_base import (
    OPCION_MANUAL,
    get_opciones_selectbox,
    buscar_por_etiqueta,
)

BODEGAS_OFICIALES = [
    "Bodega BZ02 (Insumos clínicos)",
    "Bodega CS08 (Insumos clínicos)",
    "Bodega CZ69 (Insumos clínicos)",
    "Bodega BZ28 (Insumos clínicos)",
    "Bodega AZ10 (Fármacos)",
    "Bodega AZ09 (Fármacos)",
    "Bodega BZ03 (Sueros)",
    "Farmacia",
]


def _inicializar_estado():
    """Inicializa las claves de session_state usadas para el autocompletado."""
    defaults = {
        "sel_codigo_reyimen": "",
        "sel_descripcion": "",
        "sel_unidad": "",
        "sel_costo": 0.0,
        "modo_manual": False,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def _on_change_catalogo(catalogo: List[Dict]):
    """Callback: al elegir un código en el selectbox, autocompleta los campos."""
    etiqueta = st.session_state.get("select_reyimen", "")
    if etiqueta == OPCION_MANUAL:
        st.session_state["modo_manual"] = True
        st.session_state["sel_codigo_reyimen"] = ""
        st.session_state["sel_descripcion"] = ""
        st.session_state["sel_unidad"] = ""
        st.session_state["sel_costo"] = 0.0
        return

    producto = buscar_por_etiqueta(catalogo, etiqueta)
    if producto:
        st.session_state["modo_manual"] = False
        st.session_state["sel_codigo_reyimen"] = producto["codigo"]
        st.session_state["sel_descripcion"] = producto["descripcion"]
        st.session_state["sel_unidad"] = producto["unidad"]
        st.session_state["sel_costo"] = float(producto["costo"])


def render_step1_form(catalogo: List[Dict], conn: sqlite3.Connection) -> None:
    """
    Renderiza el formulario completo del Paso 1 (Informe Bodega) y guarda
    el registro en la base de datos al enviar.

    catalogo: lista de productos entregada por catalogo_base.get_catalogo().
              Si llega vacía (falla de catálogo), se activa modo manual,
              pero el selector de Bodega NUNCA se deshabilita.
    conn: conexión sqlite3 ya inicializada (src.database.init_db()).
    """
    _inicializar_estado()

    st.subheader("📋 Paso 1 — Informe Bodega")

    catalogo_disponible = bool(catalogo)
    if not catalogo_disponible:
        st.warning(
            "⚠️ No fue posible cargar el catálogo Reyimen. "
            "Se activó el modo de ingreso manual. El registro de Bodega "
            "y el resto del formulario siguen funcionando con normalidad."
        )
        st.session_state["modo_manual"] = True

    col1, col2 = st.columns(2)
    with col1:
        estado_producto = st.selectbox(
            "Estado del producto", ["En revisión", "Concluido"], key="estado_producto"
        )
    with col2:
        mes_informe = st.date_input(
            "Mes de informe", value=date.today(), key="mes_informe"
        )

    # --- Selector de Bodega: SIEMPRE activo, pase lo que pase con el catálogo ---
    bodega_origen = st.selectbox(
        "Bodega/Farmacia origen *",
        BODEGAS_OFICIALES,
        key="bodega_origen",
        help="Campo obligatorio. Se mantiene activo aunque el catálogo falle.",
    )

    tipo_producto = st.selectbox(
        "Tipo de producto", ["Fármaco", "Insumo"], key="tipo_producto"
    )

    st.markdown("**Producto**")

    if catalogo_disponible and not st.session_state["modo_manual"]:
        opciones = get_opciones_selectbox(catalogo)
        st.selectbox(
            "Código Reyimen *",
            opciones,
            key="select_reyimen",
            on_change=_on_change_catalogo,
            args=(catalogo,),
            index=None,
            placeholder="Escribe para buscar por código o descripción...",
        )
        if st.button("✏️ Prefiero ingresar manualmente"):
            st.session_state["modo_manual"] = True
            st.rerun()
    else:
        st.caption(
            "Modo manual activo — completa código, unidad y descripción abajo."
        )
        if catalogo_disponible and st.button("🔍 Volver a usar el catálogo"):
            st.session_state["modo_manual"] = False
            st.rerun()

    colc1, colc2, colc3 = st.columns(3)
    with colc1:
        codigo_reyimen = st.text_input(
            "Código Reyimen *", value=st.session_state["sel_codigo_reyimen"]
        )
    with colc2:
        unidad = st.text_input(
            "Unidad *", value=st.session_state["sel_unidad"]
        )
    with colc3:
        costo_unitario = st.number_input(
            "Costo unitario estimado ($)",
            min_value=0.0,
            value=st.session_state["sel_costo"],
            step=10.0,
        )

    descripcion = st.text_input(
        "Descripción *", value=st.session_state["sel_descripcion"]
    )

    col3, col4 = st.columns(2)
    with col3:
        cantidad = st.number_input(
            "Cantidad", min_value=0.0, step=1.0, format="%.2f"
        )
    with col4:
        vencimiento = st.date_input("Vencimiento")

    lote = st.text_input("Lote *")
    motivo_informe = st.text_area("Motivo de informe *")

    col5, col6 = st.columns(2)
    with col5:
        tipo_compra = st.selectbox(
            "Tipo de compra", ["CENABAST", "COMPRA PROPIA"]
        )

    st.markdown("---")

    if st.button("💾 Guardar informe", type="primary"):
        errores = []
        if not bodega_origen:
            errores.append("Bodega/Farmacia origen es obligatorio.")
        if not codigo_reyimen:
            errores.append("Código Reyimen es obligatorio.")
        if not unidad:
            errores.append("Unidad es obligatoria.")
        if not descripcion:
            errores.append("Descripción es obligatoria.")
        if not lote:
            errores.append("Lote es obligatorio.")
        if not motivo_informe:
            errores.append("Motivo de informe es obligatorio.")

        if errores:
            for e in errores:
                st.error(e)
        else:
            registro = {
                "estado_producto": estado_producto,
                "mes_informe": str(mes_informe),
                "bodega_origen": bodega_origen,
                "tipo_producto": tipo_producto,
                "codigo_reyimen": codigo_reyimen,
                "unidad": unidad,
                "descripcion": descripcion,
                "cantidad": cantidad,
                "vencimiento": str(vencimiento),
                "lote": lote,
                "motivo_informe": motivo_informe,
                "tipo_compra": tipo_compra,
                "costo_unitario": costo_unitario,
            }
            from src.database import insertar_registro
            insertar_registro(conn, registro)
            st.success("✅ Informe guardado correctamente.")
            st.balloons()