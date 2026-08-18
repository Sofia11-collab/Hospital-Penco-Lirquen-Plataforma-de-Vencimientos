"""
Módulo para la importación masiva de productos desde archivos Excel (.xlsx, .xls) o CSV.
"""
import pandas as pd
from datetime import datetime
from src.database import get_connection

def procesar_carga_masiva(file_uploaded, usuario_registro: str):
    try:
        filename = file_uploaded.name.lower()
        if filename.endswith('.csv'):
            df = pd.read_csv(file_uploaded)
        else:
            df = pd.read_excel(file_uploaded)

        if df.empty:
            return False, "El archivo subido se encuentra vacío."

        # Normalizar nombres de columnas (eliminar espacios y pasar a mayúsculas)
        df.columns = [str(col).strip().upper() for col in df.columns]

        # Mapeo flexible de encabezados
        mapeo_columnas = {
            'BODEGA ORIGEN': 'bodega_origen',
            'BODEGA': 'bodega_origen',
            'TIPO PRODUCTO': 'tipo_producto',
            'TIPO': 'tipo_producto',
            'CÓDIGO REYIMEN': 'codigo_reyimen',
            'CODIGO REYIMEN': 'codigo_reyimen',
            'CODIGO': 'codigo_reyimen',
            'DESCRIPCIÓN': 'descripcion',
            'DESCRIPCION': 'descripcion',
            'TIPO COMPRA': 'tipo_documento',
            'TIPO DE COMPRA': 'tipo_documento',
            'UNIDAD': 'unidad',
            'CANTIDAD': 'cantidad',
            'CANT': 'cantidad',
            'MOTIVO INFORME': 'motivo_informe',
            'MOTIVO': 'motivo_informe',
            'FECHA VENCIMIENTO': 'vencimiento',
            'VENCIMIENTO': 'vencimiento',
            'LOTE': 'lote'
        }

        df.rename(columns=mapeo_columnas, inplace=True)

        cols_requeridas = ['codigo_reyimen', 'descripcion', 'lote', 'vencimiento']
        for col in cols_requeridas:
            if col not in df.columns:
                return False, f"Falta la columna obligatoria '{col.upper()}' en el archivo."

        conn = get_connection()
        cursor = conn.cursor()
        ingresados = 0

        for _, row in df.iterrows():
            bodega = str(row.get('bodega_origen', 'Bodega AZ09 (Fármacos)')).strip()
            tipo_prod = str(row.get('tipo_producto', 'Fármaco')).strip()
            codigo = str(row.get('codigo_reyimen', '')).strip()
            descripcion = str(row.get('descripcion', '')).strip()
            tipo_compra = str(row.get('tipo_documento', 'CENABAST')).strip()
            unidad = str(row.get('unidad', 'UNIDAD')).strip()
            
            try:
                cantidad = float(row.get('cantidad', 0))
            except ValueError:
                cantidad = 0.0

            motivo = str(row.get('motivo_informe', 'Gestión pronto vencimiento')).strip()
            lote = str(row.get('lote', '')).strip()

            venc_raw = row.get('vencimiento', '')
            if isinstance(venc_raw, datetime) or isinstance(venc_raw, pd.Timestamp):
                vencimiento = venc_raw.strftime('%Y-%m-%d')
            else:
                vencimiento = str(venc_raw).strip()[:10]

            if codigo and descripcion and lote:
                cursor.execute("""
                INSERT INTO productos (
                    bodega_origen, tipo_producto, codigo_reyimen, descripcion, 
                    unidad, cantidad, vencimiento, lote, motivo_informe, 
                    tipo_documento, usuario_registro, paso_actual
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1)
                """, (bodega, tipo_prod, codigo, descripcion, unidad, cantidad, vencimiento, lote, motivo, tipo_compra, usuario_registro))
                ingresados += 1

        conn.commit()
        conn.close()

        if ingresados > 0:
            return True, f"¡Éxito! Se ingresaron {ingresados} productos correctamente al Paso 1."
        else:
            return False, "No se pudo ingresar ningún registro válido. Verifique los campos obligatorios."

    except Exception as e:
        return False, f"Error al procesar el archivo: {str(e)}"