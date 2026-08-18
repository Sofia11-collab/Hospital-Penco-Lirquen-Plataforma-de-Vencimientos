"""
Módulo para la carga masiva de productos mediante planillas Excel o CSV.
"""
import pandas as pd
import streamlit as st
from src.database import get_connection

def procesar_carga_masiva(file, usuario_actual: str):
    try:
        if file.name.endswith('.csv'):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file)
            
        df.columns = [str(c).strip().upper() for c in df.columns]
        
        conn = get_connection()
        cursor = conn.cursor()
        
        registros_exitosos = 0
        for _, row in df.iterrows():
            bodega = str(row.get('BODEGA', 'Bodega de excluidos'))
            codigo = str(row.get('CODIGO REYIMEN', row.get('CÓDIGO REYIMEN', '')))
            desc = str(row.get('DESCRIPCIÓN', row.get('DESCRIPCION', '')))
            unidad = str(row.get('UNIDAD', 'UNIDAD'))
            cantidad = float(row.get('CANTIDAD', 0))
            venc = str(row.get('VENCIMIENTO', ''))
            lote = str(row.get('LOTE', 'S/L'))
            motivo = str(row.get('MOTIVO DE INFORME', 'Carga masiva'))
            
            if desc and codigo:
                cursor.execute("""
                INSERT INTO productos (bodega_origen, codigo_reyimen, descripcion, unidad, cantidad, vencimiento, lote, motivo_informe, usuario_registro, paso_actual, estado_global)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 1, 'En trámite')
                """, (bodega, codigo, desc, unidad, cantidad, venc, lote, motivo, usuario_actual))
                registros_exitosos += 1
                
        conn.commit()
        conn.close()
        return True, f"Se cargaron exitosamente {registros_exitosos} productos."
    except Exception as e:
        return False, f"Error al procesar el archivo: {str(e)}"