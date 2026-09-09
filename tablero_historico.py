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
from datetime import date

import pandas as pd

COL_DIVISION = 3
COL_ASIGNADOS = 5
COL_Q = 6
COL_HON_UF = 9

RE_SHEET_FECHA = re.compile(r"(\d{2})(\d{2})(\d{4})")


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
