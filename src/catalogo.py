"""
Catálogo Reyimen real + listas oficiales (bodegas, proveedores, tipos de
documento) del Hospital Penco Lirquén.

El catálogo se lee desde un archivo Excel del repositorio. Se probaron
en orden las siguientes rutas candidatas (la primera que exista se usa):

    data/catalogo.xlsx
    data/catalogo.csv
    CONTROL_DE_INVENTARIO_POR_ESTABLECIMIENTO_Hospital Penco Lirquén.xlsx
    FARMACIA_E_INSUMOS_CLINICOS.xltx

El archivo original del hospital no trae el encabezado en la fila 1
(tiene filas de título y filas en blanco antes), así que la lectura
detecta automáticamente la fila que contiene "CODIGO REYIMEN" y lee
desde ahí. Si ningún archivo existe o la lectura falla, se retorna una
lista vacía: la interfaz (src/ui.py) debe interpretar esto como
"activar modo de ingreso manual" y NUNCA debe dejar de mostrar el
selector de Bodega por este motivo.
"""

from pathlib import Path
from typing import List, Dict, Optional
import streamlit as st

try:
    import openpyxl
except ImportError:  # pragma: no cover
    openpyxl = None

OPCION_MANUAL = "🔧 Ingreso manual / Código no listado"

RUTAS_CANDIDATAS = [
    "data/catalogo.xlsx",
    "data/catalogo.csv",
    "CONTROL_DE_INVENTARIO_POR_ESTABLECIMIENTO_Hospital Penco Lirquén.xlsx",
    "FARMACIA_E_INSUMOS_CLINICOS.xltx",
]

ENCABEZADOS_ESPERADOS = {"CODIGO REYIMEN", "CÓDIGO REYIMEN"}

# ---------------------------------------------------------------------------
# Listas oficiales
# ---------------------------------------------------------------------------
BODEGAS_OFICIALES = [
    "Bodega de excluidos",
    "Bodega BZ02 (Insumos clínicos)",
    "Bodega CS08 (Insumos clínicos)",
    "Bodega CZ69 (Insumos clínicos)",
    "Bodega BZ28 (Insumos clínicos)",
    "Bodega AZ10 (Fármacos)",
    "Bodega AZ09 (Fármacos)",
    "Bodega BZ03 (Sueros)",
]

TIPOS_DOCUMENTO = ["Factura", "Guía de despacho"]

PROVEEDORES_OFICIALES = [
    "SYNTHON CHILE LIMITADA", "ESPRITE DE VIE S.A", "ASCEND LABORATORIES SAP",
    "LABORATORIO BIOVAL SPA", "HOSPITAL PENCO-LIRQUEN", "PHARMA TRADE S.A.",
    "FARMACIA RECCIUS LTDA.", "FARMACEUTICA INSUVAL LTDA.", "DIFEM LABORATORIOS S.A.",
    "IMPORTADORA Y COMERCIALIZADORA CRISTIAN MAURICIO RETAMAL PEREZ E.I.R.L",
    "DFM PHARMA", "LABORATORIOS ANDROMACO S.A.", "BECRUX LABS",
    "DROGUERIA ACONCAGUA SPA", "WINPHARM SPA.", "BESTPHARMA SPA.", "GADOR LTDA.",
    "FRESENIUS KABI CHILE LTDA.", "DROGUERIA FARMOQUIMICA DEL PACIFICO LIMITADA",
    "GRIFOLS CHILE S.A.", "LABORATORIO ACONFAR CHILE LIMITADA",
    "AWAD ARTICULOS MEDICOS LTDA", "DISTRIBUIDORA BRISSA LIMITADA",
    "VERDEJO Y VERDEJO LTDA.", "COMERCIALIZADORA DE INSUMOS MEDICOS LTDA",
    "LABORATORIO CHILE S.A.", "SCM PHARMA SPA", "VITALIS",
    "LABORATORIO BIOSANO S.A.", "ACRUX LABS", "ETHON PHARMACEUTICALS LTDA.",
    "PHARMACOR SPA", "FARMACEUTICA CARIBEAN LTDA.", "ARAMA NATURAL",
    "LUXYPHARM SPA", "ALEMBIC PHARMACEUTICALS SPA", "MAYORDENT CHILE LIMITADA",
    "INVERSIONES C & F SPA", "LABORATORIOS SAVAL S.A.", "PHARMA NETWORK SPA",
    "EMCURE PHARMA CHILE SPA", "BIOQUIMICA.CL S.A.", "MEDIKS S.A.",
    "GRUPO LOGISTICO DE ABASTECIMIENTO VERSATIL SPA", "ABASTECIMIENTO ÁGIL SPA",
    "MDC HEALTH SPA.", "CEGAPHARMA SPA", "MUNNICH PHARMA MEDICAL LTDA.",
    "ADN FARMACEUTICA SPA", "LABORATORIOS RECALCINE S.A.",
    "SERVICIOS FARMACEUTICOS Y DE SALUD QAFARMA SPA", "LABORATORIO SANDERSON S.A.",
    "GRÜNENTHAL", "SANDOZ CHILE SPA", "INSTITUTO SANITAS S.A.", "ZERICUM SPA",
    "HOSPITAL DE TOME", "COMERCIALIZADORA E INVERSIONES GHALENO LTDA.",
    "PISA FARMACEUTICA", "ALCON LABORATORIOS CHILE LTDA.", "NOVOFARMA SERVICE S.A.",
    "SOCIEDAD FARMACEUTICA BULFOR LIMITADA", "SOC.COMERCIALMEDIKAR LTDA",
    "MEDINOVA LIMITADA", "LABORATORIOS SILESIA S.A.", "INVERSIONES PHARMAVISAN S.A.",
    "PFIZER CHILE S.A.", "LABORATORIO HOSPIFARMA CHILE LTDA.", "CHEMOPHARMA S.A.",
    "DISTRIBUIDORA ISLA DEL REY S.A.", "NEOETHICALS CHILE SPA", "COMERCIAL ETHOS S.A.",
    "RECETARIO MAGISTRAL ENDOVENOSO SOCIEDAD ANONIMA",
    "INDUSTRIAL Y COMERCIAL BAXTER DE CHILE LIMITADA", "SEVEN PHARMA CHILE SPA",
    "LABORATORIO PASTEUR S.A.", "OPKO CHILE S.A.", "LABORATORIO LAFI LTDA",
    "CRISTALIA CHILE SPA", "PHARMATECH CHILE S.A.", "DR. REDDYS LABORATORIES CHILE SPA",
    "ALPHA PHARMA CHILE SPA", "UNIFARMA SPA", "HOSPITAL SAN JOSÉ DE CORONEL",
    "DISTRIBUIDORA QUALIMED LIMITADA", "SICMAFARMA", "ITF-LABOMED FARMACEUTICA LTDA.",
    "BLAU FARMACEUTICA CHILE SPA", "SOCIEDAD COMERCIAL DISTRIMED LIMITADA",
    "LIBRA CHILE S.A.", "HOSPITAL LAS HIGUERAS TALCAHUANO", "INSUAMERICA SPA",
    "LABORATORIOS BOCELI SPA", "HOSPITAL CL. GUILLERMO GRANT BENAVENTE",
    "DROGUERIA GLOBAL PHARMA SPA", "FARMACEUTICA SANTIAGO DOS SPA",
    "GLAXOSMITHKLINE CHILE FARMACEUTICA LTDA.", "VESALIUS PHARMA S.A.",
    "SERVICIOS Y COMERCIALIZADORA DE INSUMOS MEDICOS LI", "CESFAM LIRQUEN", "3CC SPA",
    "NOVO NORDISK FARMACEUTICA LIMITADA", "GE HEALTHCARE INTERNATIONAL LLC AGENCIA CHILE",
    "FAES FARMA CHILE", "CENTRAL DE ABASTECIMIENTO", "FLEXPHARMA", "FLEXING CHILE SPA",
    "SALLES ZAPATA Y COMPAÑIA LIMITADA", "THEA PHARMA SPA", "MEDIPHARM",
    "CLINICAL MARKET S.A", "EXELTIS CHILE SPA", "FARMEDICAL SPA", "DENTAL LAVAL LTDA.",
    "DISTRIPHAR SPA", "VITAFARMA S.A.", "REDLAB S.A.", "PINNACLE CHILE SPA",
    "RECBEN XENERICS FARMACEUTICA LTDA.", "FARMACIA ENGELNAT",
    "GALENICUM HEALTH CHILE SPA", "ASPEN CHILE", "LABORATORIO C&D PHARMA",
    "ALPHALAB SPA", "TECNOFARMA S.A.", "JOHNSON Y JOHNSON DE CHILE S.A.",
    "MEGALABS CHILE S.A", "INDOPHARMA S.A.", "PISA LIFE SPA", "SALCOBRAND S.A.",
    "FARMANA SPA", "HSE PHARMA SPA", "MERCK S.A.", "PHARMA GO",
    "INTERCONTINENTAL GROUP SPA",
]
# Elimina duplicados exactos manteniendo el orden original de la planilla.
PROVEEDORES_OFICIALES = list(dict.fromkeys(PROVEEDORES_OFICIALES))


# ---------------------------------------------------------------------------
# Lectura del catálogo
# ---------------------------------------------------------------------------
def _encontrar_fila_encabezado(filas: List[tuple]) -> Optional[int]:
    """Busca en las primeras filas cuál corresponde al encabezado real."""
    for idx, fila in enumerate(filas[:30]):
        if not fila:
            continue
        primera_celda = str(fila[0]).strip().upper() if fila[0] is not None else ""
        if primera_celda in ENCABEZADOS_ESPERADOS:
            return idx
    return None


def _leer_excel(path: Path) -> List[Dict]:
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    filas = list(ws.iter_rows(values_only=True))
    idx_header = _encontrar_fila_encabezado(filas)
    if idx_header is None:
        return []

    encabezado = [str(c).strip().upper() if c else "" for c in filas[idx_header]]
    try:
        i_codigo = encabezado.index("CODIGO REYIMEN") if "CODIGO REYIMEN" in encabezado \
            else encabezado.index("CÓDIGO REYIMEN")
    except ValueError:
        return []
    i_desc = next((i for i, c in enumerate(encabezado) if "DESCRIPCION" in c or "DESCRIPCIÓN" in c), None)
    i_unidad = next((i for i, c in enumerate(encabezado) if "UNIDAD" in c), None)
    i_bodega = next((i for i, c in enumerate(encabezado) if "BODEGA" in c), None)

    if i_desc is None or i_unidad is None:
        return []

    catalogo = []
    for fila in filas[idx_header + 1:]:
        if not fila or fila[i_codigo] is None:
            continue
        try:
            codigo = str(int(fila[i_codigo])).strip()
        except (ValueError, TypeError):
            codigo = str(fila[i_codigo]).strip()
        if not codigo or codigo.upper() == "TOTAL:":
            continue
        descripcion = str(fila[i_desc]).strip() if fila[i_desc] else ""
        unidad = str(fila[i_unidad]).strip() if fila[i_unidad] else ""
        area_origen = str(fila[i_bodega]).strip() if (i_bodega is not None and fila[i_bodega]) else ""
        if not descripcion:
            continue
        catalogo.append({
            "codigo": codigo, "descripcion": descripcion,
            "unidad": unidad, "area_origen": area_origen,
        })
    return catalogo


def _leer_csv(path: Path) -> List[Dict]:
    import csv
    with open(path, encoding="utf-8") as f:
        lector = csv.reader(f)
        filas = list(lector)
    idx_header = _encontrar_fila_encabezado(filas)
    if idx_header is None:
        return []
    encabezado = [str(c).strip().upper() for c in filas[idx_header]]
    try:
        i_codigo = encabezado.index("CODIGO REYIMEN")
    except ValueError:
        return []
    i_desc = next((i for i, c in enumerate(encabezado) if "DESCRIPCION" in c), None)
    i_unidad = next((i for i, c in enumerate(encabezado) if "UNIDAD" in c), None)
    i_bodega = next((i for i, c in enumerate(encabezado) if "BODEGA" in c), None)
    if i_desc is None or i_unidad is None:
        return []

    catalogo = []
    for fila in filas[idx_header + 1:]:
        if not fila or len(fila) <= i_codigo or not fila[i_codigo]:
            continue
        codigo = str(fila[i_codigo]).strip()
        if not codigo or codigo.upper() == "TOTAL:":
            continue
        descripcion = fila[i_desc].strip() if len(fila) > i_desc else ""
        unidad = fila[i_unidad].strip() if len(fila) > i_unidad else ""
        area_origen = fila[i_bodega].strip() if (i_bodega is not None and len(fila) > i_bodega) else ""
        if not descripcion:
            continue
        catalogo.append({
            "codigo": codigo, "descripcion": descripcion,
            "unidad": unidad, "area_origen": area_origen,
        })
    return catalogo


@st.cache_data(show_spinner="Cargando catálogo Reyimen...")
def get_catalogo() -> List[Dict]:
    """
    Intenta cargar el catálogo real desde las rutas candidatas del
    repositorio, en orden. Si ninguna existe o la lectura falla por
    cualquier motivo, retorna [] (modo manual) en lugar de propagar
    la excepción.
    """
    if openpyxl is None:
        return []
    for ruta in RUTAS_CANDIDATAS:
        p = Path(ruta)
        if not p.exists():
            continue
        try:
            if p.suffix.lower() == ".csv":
                catalogo = _leer_csv(p)
            else:
                catalogo = _leer_excel(p)
            if catalogo:
                return catalogo
        except Exception:
            continue
    return []


def get_opciones_selectbox(catalogo: List[Dict]) -> List[str]:
    opciones = [f"{item['codigo']} - {item['descripcion']}" for item in catalogo]
    opciones.append(OPCION_MANUAL)
    return opciones


def buscar_por_codigo(catalogo: List[Dict], codigo: str) -> Optional[Dict]:
    for item in catalogo:
        if item["codigo"] == codigo:
            return item
    return None


def buscar_por_etiqueta(catalogo: List[Dict], etiqueta: str) -> Optional[Dict]:
    if not etiqueta or etiqueta == OPCION_MANUAL:
        return None
    codigo = etiqueta.split(" - ", 1)[0].strip()
    return buscar_por_codigo(catalogo, codigo)