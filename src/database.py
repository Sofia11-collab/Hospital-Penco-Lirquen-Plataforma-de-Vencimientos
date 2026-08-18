"""
Módulo de Base de Datos SQLite para la gestión de productos y usuarios.
"""
import sqlite3
from typing import List, Dict, Optional

DB_NAME = "gestion_vencimientos.db"

def get_connection():
    conn = sqlite3.connect(DB_NAME, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_connection()
    cursor = conn.cursor()
    
    # Tabla de Usuarios
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS usuarios (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        usuario TEXT UNIQUE NOT NULL,
        password TEXT NOT NULL,
        rol TEXT NOT NULL,
        nombre_completo TEXT,
        estado TEXT DEFAULT 'Activo'
    )
    """)
    
    # Insertar usuarios por defecto si no existen
    usuarios_base = [
        ("admin", "admin123", "admin", "Administrador General"),
        ("jefatura", "jefatura123", "jefatura", "Jefatura de Farmacia/Bodega"),
        ("registro", "registro123", "registro", "Encargado de Registro y Proveedores"),
        ("bodega", "bodega123", "bodega", "Personal de Bodega"),
    ]
    
    for u, p, r, n in usuarios_base:
        cursor.execute("INSERT OR IGNORE INTO usuarios (usuario, password, rol, nombre_completo) VALUES (?, ?, ?, ?)", (u, p, r, n))

    # Tabla de Productos (Flujo de 5 Pasos)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS productos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        fecha_registro TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        bodega_origen TEXT,
        tipo_producto TEXT,
        codigo_reyimen TEXT,
        descripcion TEXT,
        unidad TEXT,
        cantidad REAL,
        vencimiento TEXT,
        lote TEXT,
        motivo_informe TEXT,
        tipo_compra TEXT,
        precio_unitario REAL,
        usuario_registro TEXT,
        
        -- Paso 2
        estado_canje TEXT,
        tipo_gestion_canje TEXT,
        fecha_paso2 TEXT,
        observacion_paso2 TEXT,
        
        -- Paso 3
        proveedor TEXT,
        tipo_documento TEXT,
        numero_documento_oc TEXT,
        tramite_proveedor TEXT,
        fecha_paso3 TEXT,
        observacion_paso3 TEXT,
        
        -- Paso 4
        ubicacion_fisica TEXT,
        ubicacion_computacional TEXT,
        numero_bulto TEXT,
        fecha_paso4 TEXT,
        observacion_paso4 TEXT,
        
        -- Paso 5
        resolucion_numero TEXT,
        estado_final TEXT,
        fecha_resolucion TEXT,
        observacion_paso5 TEXT,
        
        -- Control
        paso_actual INTEGER DEFAULT 1,
        estado_global TEXT DEFAULT 'En trámite'
    )
    """)
    
    conn.commit()
    conn.close()