"""
Catálogo base de productos Reyimen (Hospital Penco Lirquén).
Contiene un catálogo integrado de prueba (50 productos) con:
Código Reyimen, Descripción, Unidad de Medida y Costo Unitario.

Diseñado a prueba de fallos: si get_catalogo() lanza una excepción
en cualquier punto de la app, el llamador debe capturarla y pasar
a modo de ingreso manual (ver src/ui.py).
"""

from typing import List, Dict, Optional

OPCION_MANUAL = "🔧 Ingreso manual / Código no listado"

# Catálogo de prueba: 50 productos (fármacos, sueros e insumos clínicos)
_CATALOGO_BASE: List[Dict] = [
    {"codigo": "10001", "descripcion": "Paracetamol 500 mg comprimido", "unidad": "Comprimido", "costo": 15},
    {"codigo": "10002", "descripcion": "Ibuprofeno 400 mg comprimido", "unidad": "Comprimido", "costo": 20},
    {"codigo": "10003", "descripcion": "Amoxicilina 500 mg cápsula", "unidad": "Cápsula", "costo": 45},
    {"codigo": "10004", "descripcion": "Omeprazol 20 mg cápsula", "unidad": "Cápsula", "costo": 35},
    {"codigo": "10005", "descripcion": "Losartán 50 mg comprimido", "unidad": "Comprimido", "costo": 25},
    {"codigo": "10006", "descripcion": "Atorvastatina 20 mg comprimido", "unidad": "Comprimido", "costo": 40},
    {"codigo": "10007", "descripcion": "Metformina 850 mg comprimido", "unidad": "Comprimido", "costo": 18},
    {"codigo": "10008", "descripcion": "Enalapril 10 mg comprimido", "unidad": "Comprimido", "costo": 22},
    {"codigo": "10009", "descripcion": "Salbutamol inhalador 100 mcg", "unidad": "Inhalador", "costo": 3200},
    {"codigo": "10010", "descripcion": "Furosemida 40 mg comprimido", "unidad": "Comprimido", "costo": 19},
    {"codigo": "10011", "descripcion": "Insulina NPH 100 UI/ml frasco", "unidad": "Frasco", "costo": 4500},
    {"codigo": "10012", "descripcion": "Insulina Cristalina 100 UI/ml frasco", "unidad": "Frasco", "costo": 4300},
    {"codigo": "10013", "descripcion": "Warfarina 5 mg comprimido", "unidad": "Comprimido", "costo": 28},
    {"codigo": "10014", "descripcion": "Diazepam 10 mg comprimido", "unidad": "Comprimido", "costo": 30},
    {"codigo": "10015", "descripcion": "Tramadol 50 mg cápsula", "unidad": "Cápsula", "costo": 55},
    {"codigo": "10016", "descripcion": "Morfina 10 mg/ml ampolla", "unidad": "Ampolla", "costo": 980},
    {"codigo": "10017", "descripcion": "Ceftriaxona 1 g ampolla", "unidad": "Ampolla", "costo": 1200},
    {"codigo": "10018", "descripcion": "Vancomicina 500 mg ampolla", "unidad": "Ampolla", "costo": 3400},
    {"codigo": "10019", "descripcion": "Heparina sódica 5000 UI ampolla", "unidad": "Ampolla", "costo": 650},
    {"codigo": "10020", "descripcion": "Adrenalina 1 mg/ml ampolla", "unidad": "Ampolla", "costo": 720},
    {"codigo": "10021", "descripcion": "Dexametasona 4 mg/ml ampolla", "unidad": "Ampolla", "costo": 380},
    {"codigo": "10022", "descripcion": "Hidrocortisona 100 mg ampolla", "unidad": "Ampolla", "costo": 950},
    {"codigo": "10023", "descripcion": "Metoclopramida 10 mg/2ml ampolla", "unidad": "Ampolla", "costo": 210},
    {"codigo": "10024", "descripcion": "Ondansetrón 4 mg/2ml ampolla", "unidad": "Ampolla", "costo": 890},
    {"codigo": "10025", "descripcion": "Ketorolaco 30 mg/ml ampolla", "unidad": "Ampolla", "costo": 410},
    {"codigo": "10026", "descripcion": "Cloruro de Sodio 0.9% 1000 ml", "unidad": "Bolsa", "costo": 1100},
    {"codigo": "10027", "descripcion": "Ringer Lactato 1000 ml", "unidad": "Bolsa", "costo": 1150},
    {"codigo": "10028", "descripcion": "Glucosa 5% 500 ml", "unidad": "Bolsa", "costo": 980},
    {"codigo": "10029", "descripcion": "Manitol 20% 250 ml", "unidad": "Bolsa", "costo": 2300},
    {"codigo": "10030", "descripcion": "Bicarbonato de Sodio 8.4% ampolla", "unidad": "Ampolla", "costo": 560},
    {"codigo": "10031", "descripcion": "Jeringa desechable 5 ml", "unidad": "Unidad", "costo": 55},
    {"codigo": "10032", "descripcion": "Jeringa desechable 10 ml", "unidad": "Unidad", "costo": 70},
    {"codigo": "10033", "descripcion": "Jeringa desechable 20 ml", "unidad": "Unidad", "costo": 110},
    {"codigo": "10034", "descripcion": "Aguja hipodérmica 21G", "unidad": "Unidad", "costo": 25},
    {"codigo": "10035", "descripcion": "Guantes de examinación látex talla M (caja)", "unidad": "Caja", "costo": 6500},
    {"codigo": "10036", "descripcion": "Guantes estériles talla 7.5 (par)", "unidad": "Par", "costo": 420},
    {"codigo": "10037", "descripcion": "Mascarilla quirúrgica 3 pliegues", "unidad": "Unidad", "costo": 60},
    {"codigo": "10038", "descripcion": "Mascarilla N95", "unidad": "Unidad", "costo": 850},
    {"codigo": "10039", "descripcion": "Gasa estéril 10x10 cm", "unidad": "Unidad", "costo": 90},
    {"codigo": "10040", "descripcion": "Apósito transparente adhesivo", "unidad": "Unidad", "costo": 350},
    {"codigo": "10041", "descripcion": "Tela adhesiva micropore", "unidad": "Rollo", "costo": 900},
    {"codigo": "10042", "descripcion": "Alcohol gel 70% 1 L", "unidad": "Litro", "costo": 2100},
    {"codigo": "10043", "descripcion": "Alcohol etílico 70% 1 L", "unidad": "Litro", "costo": 1800},
    {"codigo": "10044", "descripcion": "Clorhexidina 2% 1 L", "unidad": "Litro", "costo": 3200},
    {"codigo": "10045", "descripcion": "Suero fisiológico ampolla 10 ml", "unidad": "Ampolla", "costo": 180},
    {"codigo": "10046", "descripcion": "Catéter venoso periférico 18G", "unidad": "Unidad", "costo": 320},
    {"codigo": "10047", "descripcion": "Catéter venoso periférico 20G", "unidad": "Unidad", "costo": 310},
    {"codigo": "10048", "descripcion": "Sonda Foley 14 Fr", "unidad": "Unidad", "costo": 1450},
    {"codigo": "10049", "descripcion": "Sonda nasogástrica 14 Fr", "unidad": "Unidad", "costo": 980},
    {"codigo": "10050", "descripcion": "Llave de tres pasos", "unidad": "Unidad", "costo": 260},
]


def get_catalogo() -> List[Dict]:
    """
    Devuelve el catálogo de productos.
    Envuelto en try/except: si algo falla (ej. catálogo corrupto o
    fuente externa caída en una versión futura), retorna lista vacía
    en lugar de lanzar la excepción hacia arriba. El llamador (ui.py)
    debe interpretar lista vacía como "activar modo manual".
    """
    try:
        # En una versión futura esto podría leer desde un Excel/CSV/DB externa.
        # Se valida la integridad mínima de cada registro antes de entregarlo.
        catalogo_validado = []
        for item in _CATALOGO_BASE:
            if all(k in item for k in ("codigo", "descripcion", "unidad", "costo")):
                catalogo_validado.append(item)
        return catalogo_validado
    except Exception:
        return []


def get_opciones_selectbox(catalogo: List[Dict]) -> List[str]:
    """Genera las etiquetas 'codigo - descripcion' para el selectbox + opción manual."""
    opciones = [f"{item['codigo']} - {item['descripcion']}" for item in catalogo]
    opciones.append(OPCION_MANUAL)
    return opciones


def buscar_por_codigo(catalogo: List[Dict], codigo: str) -> Optional[Dict]:
    """Busca un producto por su código Reyimen exacto."""
    for item in catalogo:
        if item["codigo"] == codigo:
            return item
    return None


def buscar_por_etiqueta(catalogo: List[Dict], etiqueta: str) -> Optional[Dict]:
    """Dada la etiqueta 'codigo - descripcion' del selectbox, retorna el producto."""
    if not etiqueta or etiqueta == OPCION_MANUAL:
        return None
    codigo = etiqueta.split(" - ", 1)[0].strip()
    return buscar_por_codigo(catalogo, codigo)