"""
Lectura del histórico del Tablero manual (TABLERO_ING) para el reporte de
tendencia de Stock y Asignaciones.

El archivo que se carga es el mismo libro que se arma a mano en Excel: una
hoja por semana, nombrada "Tablero DDMMYYYY", y cada hoja trae DOS bloques
de grilla (Semana Anterior + la semana de la hoja). Se toma siempre el
SEGUNDO bloque (la semana propia de la hoja) para no contar cada semana dos
veces. La fecha de cada punto se saca del nombre de la hoja, no del título
en español dentro de la hoja (que varía de formato).
"""
import re
import unicodedata
from datetime import date

import pandas as pd

COL_DIVISION = 3
COL_AJUSTADOR = 4
COL_ASIGNADOS = 5
COL_Q = 6
COL_HON_UF = 9

RE_SHEET_FECHA = re.compile(r"(\d{2})(\d{2})(\d{4})")

# Nombres tal como aparecen en el Tablero manual (validado contra las 41
# semanas reales: el equipo no cambió en todo el período). Se identifica al
# ajustador por prefijo, sin acentos, porque el nombre en pantalla trae
# pegado el tramo (ej. "Francisco Silva <= 1000 UF").
AJUSTADORES_INGENIERIA = ["Francisco Silva", "Dagoberto Pérez", "Patrick Swain", "Camilo Abarzúa", "Nelson Canio"]
AJUSTADORES_MOVIL = ["Mauricio Muñoz", "Leopoldo Soto"]


def _sin_acentos(texto):
    return "".join(c for c in unicodedata.normalize("NFD", texto) if unicodedata.category(c) != "Mn")


def _identificar_ajustador(celda, nombres_conocidos):
    normalizado = _sin_acentos(str(celda)).strip().lower()
    for nombre in nombres_conocidos:
        if normalizado.startswith(_sin_acentos(nombre).lower()):
            return nombre
    return None


def _fecha_de_hoja(nombre_hoja):
    m = RE_SHEET_FECHA.search(nombre_hoja)
    if not m:
        return None
    dia, mes, anio = (int(x) for x in m.groups())
    try:
        return date(anio, mes, dia)
    except ValueError:
        return None


def _extraer_fila(df, prefijo_etiqueta):
    """Última fila (antes de 'AVANCE...') cuyo texto en la columna División
    empieza con `prefijo_etiqueta` — la última es siempre la de la semana
    propia de la hoja (la primera es la de 'Semana Anterior')."""
    col = df[COL_DIVISION].astype(str)
    limite_idx = col[col.str.startswith("AVANCE")].index
    limite = limite_idx[0] if len(limite_idx) else len(df)

    candidatos = col[col.str.startswith(prefijo_etiqueta) & (col.index < limite)].index
    if not len(candidatos):
        return None
    fila = candidatos[-1]

    valores = []
    for c in (COL_ASIGNADOS, COL_Q, COL_HON_UF):
        v = pd.to_numeric(df.iloc[fila, c], errors="coerce")
        valores.append(0.0 if pd.isna(v) else float(v))
    return valores


def parsear_historico(archivo):
    """`archivo`: ruta o file-like (.xlsx) con una hoja 'Tablero DDMMYYYY'
    por semana. Devuelve un DataFrame con columnas Fecha/Division/Asignados/
    Stock_Q/Stock_UF — una fila por (semana, división), más una fila
    'Total Gerencia' por semana. Ordenado cronológicamente."""
    xl = pd.ExcelFile(archivo)
    filas = []
    for nombre_hoja in xl.sheet_names:
        fecha = _fecha_de_hoja(nombre_hoja)
        if fecha is None:
            continue
        df = pd.read_excel(xl, sheet_name=nombre_hoja, header=None)
        if df.shape[1] <= COL_HON_UF:
            continue

        for etiqueta, prefijo in [
            ("Ingeniería y Energía", "Subtotal Ingenier"),
            ("Equipo Móvil", "Subtotal Equipo M"),
            ("Total Gerencia", "Total"),
        ]:
            valores = _extraer_fila(df, prefijo)
            if valores is None:
                continue
            filas.append({
                "Fecha": fecha, "Division": etiqueta,
                "Asignados": valores[0], "Stock_Q": valores[1], "Stock_UF": round(valores[2], 2),
            })

    columnas = ["Fecha", "Division", "Asignados", "Stock_Q", "Stock_UF"]
    if not filas:
        return pd.DataFrame(columns=columnas)
    return pd.DataFrame(filas)[columnas].sort_values("Fecha").reset_index(drop=True)


def _rango_semana_actual(df):
    """(inicio, fin) de las filas de detalle por ajustador de Ingeniería y de
    Equipo Móvil, para el bloque de la SEMANA PROPIA de la hoja (el segundo).
    None si a la hoja le falta algún marcador esperado."""
    col = df[COL_DIVISION].astype(str)
    limite_idx = col[col.str.startswith("AVANCE")].index
    limite = limite_idx[0] if len(limite_idx) else len(df)

    idx_total = col[col.str.startswith("Total") & (col.index < limite)].index
    idx_ing = col[col.str.startswith("Subtotal Ingenier") & (col.index < limite)].index
    idx_mov = col[col.str.startswith("Subtotal Equipo M") & (col.index < limite)].index
    if len(idx_total) < 2 or len(idx_ing) < 2 or len(idx_mov) < 2:
        return None
    return (idx_total[0], idx_ing[-1]), (idx_ing[-1], idx_mov[-1])


def _stock_por_ajustador(df, inicio, fin, nombres_conocidos):
    """{ajustador: Stock_Q total} sumando todas sus filas de tramo, dentro
    del rango [inicio, fin) de la semana actual."""
    totales = {}
    for i in range(inicio + 1, fin):
        nombre = _identificar_ajustador(df.iloc[i, COL_AJUSTADOR], nombres_conocidos)
        if nombre is None:
            continue
        q = pd.to_numeric(df.iloc[i, COL_Q], errors="coerce")
        totales[nombre] = totales.get(nombre, 0.0) + (0.0 if pd.isna(q) else float(q))
    return totales


def parsear_promedio_casos_por_ajustador(archivo, excluir=("Dagoberto Pérez",)):
    """Evolución del Stock (Q) PROMEDIO por ajustador, semana a semana, para
    Ingeniería, Equipo Móvil y el total de la gerencia (todos los ajustadores
    de ambas divisiones juntos) — excluyendo a quienes estén en `excluir`
    (por defecto, Dagoberto Pérez, cuya carga no es representativa).
    El promedio de cada semana se calcula solo entre los ajustadores que
    tuvieron algún caso esa semana (no se cuenta a alguien con 0 como 0)."""
    xl = pd.ExcelFile(archivo)
    filas = []
    for nombre_hoja in xl.sheet_names:
        fecha = _fecha_de_hoja(nombre_hoja)
        if fecha is None:
            continue
        df = pd.read_excel(xl, sheet_name=nombre_hoja, header=None)
        if df.shape[1] <= COL_HON_UF:
            continue

        rango = _rango_semana_actual(df)
        if rango is None:
            continue
        (ini_ing, fin_ing), (ini_mov, fin_mov) = rango

        stock_ing = _stock_por_ajustador(df, ini_ing, fin_ing, AJUSTADORES_INGENIERIA)
        stock_mov = _stock_por_ajustador(df, ini_mov, fin_mov, AJUSTADORES_MOVIL)

        for etiqueta, stock in [
            ("Ingeniería y Energía", stock_ing),
            ("Equipo Móvil", stock_mov),
            ("Total Gerencia", {**stock_ing, **stock_mov}),
        ]:
            valores = [v for k, v in stock.items() if k not in excluir]
            if not valores:
                continue
            filas.append({
                "Fecha": fecha, "Division": etiqueta,
                "Promedio_Q": round(sum(valores) / len(valores), 2),
                "N_Ajustadores": len(valores),
            })

    columnas = ["Fecha", "Division", "Promedio_Q", "N_Ajustadores"]
    if not filas:
        return pd.DataFrame(columns=columnas)
    return pd.DataFrame(filas)[columnas].sort_values("Fecha").reset_index(drop=True)
