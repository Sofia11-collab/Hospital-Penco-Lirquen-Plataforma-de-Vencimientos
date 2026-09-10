import pandas as pd
from src.database import get_connection

def procesar_carga_masiva(file, usuario_registro):
    try:
        # Leer el archivo dependiendo de su extensión
        if file.name.endswith('.csv'):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file)

        # Blindaje 1: Estandarizar títulos (quitar espacios extra y poner todo en mayúsculas)
        df.columns = df.columns.str.strip().str.upper()

        conn = get_connection()
        cursor = conn.cursor()
        count = 0

        for index, row in df.iterrows():
            # Blindaje 2: Búsqueda flexible (con o sin tildes)
            codigo = str(row.get('CÓDIGO REYIMEN', row.get('CODIGO REYIMEN', ''))).strip()

            # Evitar procesar filas que estén completamente en blanco
            if not codigo or codigo.lower() == 'nan' or codigo.lower() == 'nat':
                continue

            bodega = str(row.get('BODEGA ORIGEN', '')).strip()
            tipo_prod = str(row.get('TIPO PRODUCTO', '')).strip()
            desc = str(row.get('DESCRIPCIÓN', row.get('DESCRIPCION', ''))).strip()
            compra = str(row.get('TIPO COMPRA', '')).strip()
            unidad = str(row.get('UNIDAD', '')).strip()
            cant = row.get('CANTIDAD', 0)
            venc = row.get('FECHA VENCIMIENTO', '')
            lote = str(row.get('LOTE', '')).strip()

            # Forzar cantidad a número
            try:
                cant = float(cant)
            except:
                cant = 0.0
            
            # Forzar fecha al formato correcto (YYYY-MM-DD)
            if pd.isna(venc) or str(venc).strip() == '':
                venc_str = ''
            else:
                try:
                    venc_str = pd.to_datetime(venc, dayfirst=True).strftime('%Y-%m-%d')
                except:
                    venc_str = str(venc).strip()[:10]

            # Inyección segura a PostgreSQL
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

        # Blindaje 3: Alerta clara si leyó el Excel pero no encontró datos
        if count == 0:
            return False, "⚠️ El archivo fue leído, pero se encontraron 0 productos. Asegúrate de usar la plantilla oficial y no alterar los títulos."

        return True, f"✅ Carga masiva exitosa: {count} productos ingresados correctamente al Paso 2."
        
    except Exception as e:
        return False, f"❌ Error interno al procesar el archivo: {e}"
