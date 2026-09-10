import pandas as pd
from src.database import get_connection

def procesar_carga_masiva(file, usuario_registro):
    try:
        # Leer el archivo dependiendo de su extensión
        if file.name.endswith('.csv'):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file)

        conn = get_connection()
        cursor = conn.cursor()
        count = 0

        for index, row in df.iterrows():
            # Extraer datos de las columnas asegurando que sean texto o números válidos
            bodega = str(row.get('BODEGA ORIGEN', '')).strip()
            tipo_prod = str(row.get('TIPO PRODUCTO', '')).strip()
            codigo = str(row.get('CÓDIGO REYIMEN', '')).strip()
            desc = str(row.get('DESCRIPCIÓN', '')).strip()
            compra = str(row.get('TIPO COMPRA', '')).strip()
            unidad = str(row.get('UNIDAD', '')).strip()
            cant = row.get('CANTIDAD', 0)
            venc = row.get('FECHA VENCIMIENTO', '')
            lote = str(row.get('LOTE', '')).strip()

            # Evitar procesar filas que estén completamente en blanco
            if not codigo or codigo == 'nan':
                continue

            # Blindaje para cantidades vacías
            if pd.isna(cant) or str(cant).strip() == '':
                cant = 0.0
            else:
                cant = float(cant)
            
            # Blindaje y conversión de fecha para que la BD pueda calcular el semáforo (YYYY-MM-DD)
            if pd.isna(venc) or str(venc).strip() == '':
                venc_str = ''
            else:
                try:
                    # Convierte desde DD/MM/YYYY del Excel al formato interno de la BD
                    venc_str = pd.to_datetime(venc, dayfirst=True).strftime('%Y-%m-%d')
                except:
                    venc_str = str(venc).strip()[:10]

            # Inyección en PostgreSQL usando %s (El nuevo idioma de la nube)
            cursor.execute("""
                INSERT INTO productos (
                    bodega_origen, tipo_producto, codigo_reyimen, descripcion, 
                    tipo_documento, unidad, cantidad, vencimiento, lote, 
                    motivo_informe, usuario_registro, paso_actual, estado_global
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """, (
                bodega, tipo_prod, codigo, desc, compra, unidad, cant, 
                venc_str, lote, 'Gestión pronto vencimiento', 
                usuario_registro, 2, 'En trámite'
            ))
            count += 1

        conn.commit()
        conn.close()
        return True, f"✅ Carga masiva exitosa: {count} productos ingresados directamente al Paso 2."
        
    except Exception as e:
        return False, f"Error al procesar el archivo: {e}"
