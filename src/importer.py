"""
Carga masiva del Paso 1: plantilla Excel descargable + validación e
inserción automática de registros subidos por el usuario.

Firma pública:
    render_carga_masiva(conn) -> None
"""

import streamlit as st
import pandas as pd
import io
import sqlite3

from src.database import insertar_registros_bulk

# Encabezados esperados en el archivo del usuario (formato de la planilla actual)
COLUMNAS_ESPERADAS = [
    "BODEGA", "CODIGO REYIMEN", "DESCRIPCIÓN", "UNIDAD", "CANTIDAD",
    "VENCIMIENTO", "LOTE", "MOTIVO DE INFORME", "TIPO DE COMPRA", "PRECIO UNITARIO",
]

# Mapeo de encabezado de la planilla -> columna interna de la base de datos
MAPEO_COLUMNAS = {
    "BODEGA": "bodega_origen",
    "CODIGO REYIMEN": "codigo_reyimen",
    "DESCRIPCIÓN": "descripcion",
    "UNIDAD": "unidad",
    "CANTIDAD": "cantidad",
    "VENCIMIENTO": "vencimiento",
    "LOTE": "lote",
    "MOTIVO DE INFORME": "motivo_informe",
    "TIPO DE COMPRA": "tipo_compra",
    "PRECIO UNITARIO": "costo_unitario",
}


def _generar_plantilla_excel() -> io.BytesIO:
    """Genera un Excel vacío (solo encabezados) para que el usuario lo complete."""
    df_vacio = pd.DataFrame(columns=COLUMNAS_ESPERADAS)
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_vacio.to_excel(writer, index=False, sheet_name="Plantilla")
    buffer.seek(0)
    return buffer


def _leer_archivo(uploaded_file) -> pd.DataFrame:
    """Lee .xlsx, .xls o .csv según la extensión del archivo subido."""
    nombre = uploaded_file.name.lower()
    if nombre.endswith(".csv"):
        return pd.read_csv(uploaded_file)
    elif nombre.endswith(".xls"):
        return pd.read_excel(uploaded_file, engine="xlrd")
    else:  # .xlsx
        return pd.read_excel(uploaded_file, engine="openpyxl")


def _validar_columnas(df: pd.DataFrame) -> list:
    """Retorna la lista de columnas esperadas que faltan en el archivo (vacía = válido)."""
    columnas_archivo = [str(c).strip().upper() for c in df.columns]
    faltantes = [c for c in COLUMNAS_ESPERADAS if c not in columnas_archivo]
    return faltantes


def _normalizar_df(df: pd.DataFrame) -> pd.DataFrame:
    """Normaliza encabezados y renombra a las columnas internas de la base de datos."""
    df = df.copy()
    df.columns = [str(c).strip().upper() for c in df.columns]
    df = df.rename(columns=MAPEO_COLUMNAS)
    columnas_internas = list(MAPEO_COLUMNAS.values())
    for c in columnas_internas:
        if c not in df.columns:
            df[c] = ""
    df = df[columnas_internas].fillna("")
    df["estado_producto"] = "En revisión"
    df["mes_informe"] = ""
    df["tipo_producto"] = ""
    return df


def render_carga_masiva(conn: sqlite3.Connection) -> None:
    """
    Renderiza la sección de carga masiva: descarga de plantilla, subida de
    archivo, validación de columnas e inserción automática en la base de datos.
    """
    st.subheader("📥 Carga masiva — Paso 1")

    st.download_button(
        label="⬇️ Descargar plantilla Excel",
        data=_generar_plantilla_excel(),
        file_name="plantilla_informe_bodega.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.caption(
        "Encabezados requeridos: " + ", ".join(COLUMNAS_ESPERADAS)
    )

    archivo = st.file_uploader(
        "Sube el archivo completado (.xlsx, .xls o .csv)",
        type=["xlsx", "xls", "csv"],
    )

    if archivo is not None:
        try:
            df = _leer_archivo(archivo)
        except Exception as e:
            st.error(f"No fue posible leer el archivo: {e}")
            return

        faltantes = _validar_columnas(df)
        if faltantes:
            st.error(
                "El archivo no coincide con el formato esperado. "
                f"Faltan las columnas: {', '.join(faltantes)}"
            )
            return

        st.success(f"Archivo válido. Se detectaron {len(df)} filas.")
        st.dataframe(df, use_container_width=True)

        if st.button("📤 Insertar registros en el sistema", type="primary"):
            df_normalizado = _normalizar_df(df)
            insertados = insertar_registros_bulk(conn, df_normalizado)
            st.success(f"✅ Se insertaron {insertados} registros correctamente.")
            st.rerun()