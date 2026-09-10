import psycopg2
import streamlit as st

def get_connection():
    # Se conecta a la base de datos en la nube usando el secreto
    conn = psycopg2.connect(st.secrets["DB_URL"])
    cursor = conn.cursor()
    
    # Crear tabla de usuarios
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS usuarios (
        id SERIAL PRIMARY KEY,
        usuario VARCHAR(255) UNIQUE,
        password VARCHAR(255),
        rol VARCHAR(50),
        nombre_completo VARCHAR(255),
        estado VARCHAR(50) DEFAULT 'Activo'
    )
    """)
    
    # Crear tabla de productos
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS productos (
        id SERIAL PRIMARY KEY,
        bodega_origen VARCHAR(255),
        tipo_producto VARCHAR(255),
        codigo_reyimen VARCHAR(255),
        descripcion TEXT,
        unidad VARCHAR(100),
        cantidad FLOAT,
        vencimiento VARCHAR(100),
        lote VARCHAR(255),
        motivo_informe VARCHAR(255),
        tipo_documento VARCHAR(255),
        usuario_registro VARCHAR(255),
        paso_actual INTEGER,
        estado_global VARCHAR(255),
        ubicacion_fisica VARCHAR(255),
        numero_bulto VARCHAR(255),
        alerta_numero VARCHAR(255),
        alerta_fecha VARCHAR(255),
        estado_canje VARCHAR(255),
        observacion_paso2 TEXT,
        fecha_paso2 VARCHAR(100),
        archivo_canje TEXT,
        nombre_archivo_canje TEXT,
        proveedor VARCHAR(255),
        numero_documento_oc VARCHAR(255),
        tramite_proveedor VARCHAR(255),
        observacion_paso3 TEXT,
        fecha_paso3 VARCHAR(100),
        ubicacion_computacional VARCHAR(255),
        observacion_paso4 TEXT,
        resolucion_numero VARCHAR(255),
        estado_final VARCHAR(255),
        observacion_paso5 TEXT,
        tipo_gestion_canje VARCHAR(255),
        principio_activo VARCHAR(255),
        titular_registro VARCHAR(255),
        registro_sanitario VARCHAR(255)
    )
    """)
    
    # Insertar admin por defecto si la tabla está vacía
    cursor.execute("SELECT * FROM usuarios WHERE usuario='admin'")
    if not cursor.fetchone():
        cursor.execute("INSERT INTO usuarios (usuario, password, rol, nombre_completo) VALUES ('admin', 'admin', 'admin', 'Administrador General')")
        
    conn.commit()
    return conn
