import pandas as pd
from src.database import get_connection

def procesar_carga_masiva(file, usuario_registro):
    try:
        # Leer el archivo dependiendo de su extensión
        if file.name.endswith('.csv'):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file)

        # BLINDAJE 1: Estandarizar títulos y eliminar objetos "NaN" de Pandas
        df.columns = df.columns.str.strip().str.upper()
        df = df.fillna('')

        conn = get_connection()
        cursor = conn.cursor()
        count = 0

        for index, row in df.iterrows():
            codigo = str(row.get('CÓDIGO REYIMEN', row.get('CODIGO REYIMEN', ''))).strip()

            # Saltar filas vacías
            if not codigo:
                continue

            bodega = str(row.get('BODEGA ORIGEN', '')).strip()
            tipo_prod = str(row.get('TIPO PRODUCTO', '')).strip()
            desc = str(row.get('DESCRIPCIÓN', row.get('DESCRIPCION', ''))).strip()
            compra = str(row.get('TIPO COMPRA', '')).strip()
            unidad = str(row.get('UNIDAD', '')).strip()
            cant_raw = row.get('CANTIDAD', 0)
            venc = row.get('FECHA VENCIMIENTO', '')
            lote = str(row.get('LOTE', '')).strip()

            # BLINDAJE 2: Forzar cantidad a número puro de Python
            try:
                cant = float(cant_raw) if cant_raw != '' else 0.0
            except:
                cant = 0.0
            
            # BLINDAJE 3: Forzar fecha al formato estricto (YYYY-MM-DD)
            if venc == '':
                venc_str = ''
            else:
                try:
                    venc_str = pd.to_datetime(venc, dayfirst=True).strftime('%Y-%m-%d')
                except:
                    venc_str = str(venc).strip()[:10]

            # Inyección a PostgreSQL
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

        if count == 0:
            return False, "⚠️ Se leyó el Excel, pero no se detectaron filas válidas. Revisa la primera columna."

        return True, f"✅ Carga masiva exitosa: {count} productos ingresados correctamente. (Búscalos directo en el Paso 2)"
        
    except Exception as e:
        return False, f"❌ Error de lectura en la nube: {e}"
