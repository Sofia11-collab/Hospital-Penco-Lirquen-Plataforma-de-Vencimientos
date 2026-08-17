"""
Capa de persistencia (SQLite) para los informes de Bodega - Paso 1.

Nota: en Streamlit Cloud el disco es efímero (se reinicia en cada
redeploy o reinicio del contenedor). Para uso real y continuo del
personal de bodega, reemplazar esta capa por una fuente persistente
externa (Google Sheets, Supabase, Postgres, etc.).
"""

import sqlite3
import pandas as pd
from typing import Dict
from pathlib import Path

DB_PATH = Path("vencimientos_penco_lirquen.db")

COLUMNAS = [
    "estado_producto", "mes_informe", "bodega_origen", "tipo_producto",
    "codigo_reyimen", "unidad", "descripcion", "cantidad", "vencimiento",
    "lote", "motivo_informe", "tipo_compra", "costo_unitario",
]


def init_db(db_path: Path = DB_PATH) -> sqlite3.Connection:
    """Inicializa (o abre) la base de datos y crea la tabla si no existe."""
    conn = sqlite3.connect(db_path, check_same_thread=False)
    columnas_sql = ", ".join([f'"{c}" TEXT' for c in COLUMNAS])
    conn.execute(f"""
        CREATE TABLE IF NOT EXISTS informes_bodega (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            {columnas_sql}
        )
    """)
    conn.commit()
    return conn


def insertar_registro(conn: sqlite3.Connection, registro: Dict) -> None:
    """Inserta un único registro (dict) validando que existan todas las columnas."""
    fila = {c: registro.get(c, "") for c in COLUMNAS}
    placeholders = ", ".join(["?"] * len(COLUMNAS))
    columnas_sql = ", ".join([f'"{c}"' for c in COLUMNAS])
    conn.execute(
        f'INSERT INTO informes_bodega ({columnas_sql}) VALUES ({placeholders})',
        [str(fila[c]) for c in COLUMNAS],
    )
    conn.commit()


def insertar_registros_bulk(conn: sqlite3.Connection, df: pd.DataFrame) -> int:
    """Inserta múltiples registros desde un DataFrame ya normalizado. Retorna cuántos insertó."""
    count = 0
    for _, fila in df.iterrows():
        registro = {c: fila.get(c, "") for c in COLUMNAS}
        insertar_registro(conn, registro)
        count += 1
    return count


def obtener_registros(conn: sqlite3.Connection) -> pd.DataFrame:
    """Retorna todos los registros como DataFrame, más recientes primero."""
    try:
        return pd.read_sql_query(
            "SELECT * FROM informes_bodega ORDER BY id DESC", conn
        )
    except Exception:
        return pd.DataFrame(columns=["id"] + COLUMNAS)