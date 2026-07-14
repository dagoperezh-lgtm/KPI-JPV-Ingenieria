"""
Motor de cálculo del Tablero Gerencial.

Traduce los datos crudos de OpsControl (Base Maestra + planes semanales)
a la misma grilla que hoy se arma a mano en TABLERO_ING: por División >
Ajustador > Tramo UF, con Stock, bloque Ajuste (trabajo en proceso) y
bloque IFL (informes finales / cierres facturables).

Regla de clasificación Ajuste vs IFL: se reutiliza EXACTAMENTE la misma
condición que OpsControl usa para separar "Ingreso Efectivo Facturable"
(IFL) de "Valor Potencial Traccionado / WIP" (Ajuste):
    accion contiene 'Informe Final de Liquidación' o 'Carta de Cobertura (Rechazo)'
"""
from functools import reduce

import pandas as pd

from tablero_datos import calcular_tramo, col, dict_divisiones, honorarios_del_caso

REGEX_IFL = r"Informe Final de Liquidación|Carta de Cobertura \(Rechazo\)"

COLUMNAS_GRILLA = [
    "Division", "Ajustador", "Tramo",
    "Asignados", "Stock_Q", "Dias_prom", "Stock_Hon_UF",
    "Ajuste_Qp", "Ajuste_Meta", "Ajuste_HonpUF", "Ajuste_Qr", "Ajuste_HonrUF",
    "IFL_Qp", "IFL_Meta", "IFL_HonpUF", "IFL_Qr", "IFL_HonrUF",
]

COLUMNAS_NUMERICAS = [c for c in COLUMNAS_GRILLA if c not in ("Division", "Ajustador", "Tramo", "Dias_prom")]


def _preparar_stock(df_maestro, dias_semana):
    c_aj = col(df_maestro, "ajustador")
    c_div = col(df_maestro, "division")
    c_estado = col(df_maestro, "estado")
    c_creado = col(df_maestro, "creado_en")

    df = df_maestro.copy()
    df["_ajustador"] = df[c_aj].astype(str).str.strip() if c_aj else ""
    df["_division"] = df[c_div].astype(str).str.strip() if c_div else "Sin División"
    df["_estado"] = df[c_estado].astype(str).str.strip().str.lower() if c_estado else ""

    vigentes = df[df["_estado"] != "cerrado"].copy()
    if vigentes.empty:
        return pd.DataFrame(columns=["Division", "Ajustador", "Tramo", "Asignados", "Stock_Q", "Dias_prom", "Stock_Hon_UF"])

    tramos, _ = zip(*vigentes.apply(calcular_tramo, axis=1))
    vigentes["_tramo"] = tramos
    vigentes["_honorarios"] = vigentes.apply(lambda r: honorarios_del_caso(r, df.columns), axis=1)

    if c_creado and c_creado in vigentes.columns:
        vigentes["_fecha_creado"] = pd.to_datetime(vigentes[c_creado], errors="coerce", dayfirst=True)
    else:
        vigentes["_fecha_creado"] = pd.NaT

    hoy = pd.Timestamp(dias_semana[-1])
    vigentes["_dias_abierto"] = (hoy - vigentes["_fecha_creado"]).dt.days
    ini, fin = dias_semana[0], dias_semana[-1]
    fecha_solo_dia = vigentes["_fecha_creado"].dt.date
    vigentes["_asignado_semana"] = fecha_solo_dia.between(ini, fin)

    grp = vigentes.groupby(["_division", "_ajustador", "_tramo"]).agg(
        Stock_Q=("_tramo", "size"),
        Dias_prom=("_dias_abierto", "mean"),
        Stock_Hon_UF=("_honorarios", "sum"),
        Asignados=("_asignado_semana", "sum"),
    ).reset_index()
    return grp.rename(columns={"_division": "Division", "_ajustador": "Ajustador", "_tramo": "Tramo"})


def _resumen_bloque(df_tareas, es_ifl, col_q, col_hon):
    """Cuenta casos únicos y suma honorarios (por caso único, tomando el máximo
    honorario reportado en la semana para ese caso) agrupado por Ajustador+Tramo."""
    sub = df_tareas[df_tareas["_es_ifl"] == es_ifl]
    if sub.empty:
        return pd.DataFrame(columns=["Ajustador", "Tramo", col_q, col_hon])

    q = sub.groupby(["Ajustador", "_tramo"])["numero_caso"].nunique().reset_index(name=col_q)
    hon_por_caso = sub.groupby(["Ajustador", "_tramo", "numero_caso"])["honorarios_estimados"].max().reset_index()
    hon = hon_por_caso.groupby(["Ajustador", "_tramo"])["honorarios_estimados"].sum().reset_index(name=col_hon)
    return q.merge(hon, on=["Ajustador", "_tramo"], how="left").rename(columns={"_tramo": "Tramo"})


def _preparar_ejecucion(df_plan_semana):
    columnas_vacias = ["Ajustador", "Tramo", "Ajuste_Qp", "Ajuste_HonpUF", "Ajuste_Qr", "Ajuste_HonrUF",
                        "IFL_Qp", "IFL_HonpUF", "IFL_Qr", "IFL_HonrUF"]
    if df_plan_semana is None or df_plan_semana.empty:
        return pd.DataFrame(columns=columnas_vacias)

    df = df_plan_semana[df_plan_semana.get("categoria", "") == "Operativa"].copy()
    if df.empty:
        return pd.DataFrame(columns=columnas_vacias)

    df["_es_ifl"] = df["accion"].astype(str).str.contains(REGEX_IFL, case=False, na=False)
    df["_tramo"] = df["tramo_uf"].replace("", "N/D").fillna("N/D")

    programadas = df[df["tipo_actividad"] == "Programada"]
    realizadas = df[df["estado_cumplimiento"] == "Realizado"]

    piezas = [
        _resumen_bloque(programadas, False, "Ajuste_Qp", "Ajuste_HonpUF"),
        _resumen_bloque(realizadas, False, "Ajuste_Qr", "Ajuste_HonrUF"),
        _resumen_bloque(programadas, True, "IFL_Qp", "IFL_HonpUF"),
        _resumen_bloque(realizadas, True, "IFL_Qr", "IFL_HonrUF"),
    ]
    resultado = reduce(lambda l, r: pd.merge(l, r, on=["Ajustador", "Tramo"], how="outer"), piezas)
    return resultado


def construir_grilla(df_maestro, df_plan_semana, dias_semana, metas_config=None):
    """Devuelve la grilla Division/Ajustador/Tramo con Stock + Ajuste + IFL,
    lista para mostrar tal como el Excel TABLERO_ING."""
    stock = _preparar_stock(df_maestro, dias_semana)
    ejecucion = _preparar_ejecucion(df_plan_semana)

    dict_div = dict(zip(stock["Ajustador"], stock["Division"])) if not stock.empty else {}
    if not ejecucion.empty:
        ejecucion = ejecucion.copy()
        ejecucion["Division"] = ejecucion["Ajustador"].map(dict_div).fillna("Sin División")

    grilla = pd.merge(stock, ejecucion, on=["Division", "Ajustador", "Tramo"], how="outer")
    if grilla.empty:
        return pd.DataFrame(columns=COLUMNAS_GRILLA)

    for c in COLUMNAS_NUMERICAS:
        if c not in grilla.columns:
            grilla[c] = 0
        grilla[c] = pd.to_numeric(grilla[c], errors="coerce").fillna(0)

    if "Dias_prom" in grilla.columns:
        grilla["Dias_prom"] = grilla["Dias_prom"].round(0)

    metas_config = metas_config or {}
    grilla["Ajuste_Meta"] = grilla.apply(lambda r: metas_config.get((r["Ajustador"], r["Tramo"], "Ajuste"), None), axis=1)
    grilla["IFL_Meta"] = grilla.apply(lambda r: metas_config.get((r["Ajustador"], r["Tramo"], "IFL"), None), axis=1)

    for c in COLUMNAS_GRILLA:
        if c not in grilla.columns:
            grilla[c] = None

    return grilla[COLUMNAS_GRILLA].sort_values(["Division", "Ajustador", "Tramo"]).reset_index(drop=True)


def bitacora_gestion_comercial(df_plan_semana):
    if df_plan_semana is None or df_plan_semana.empty:
        return pd.DataFrame(columns=["Ajustador", "Detalle", "Fecha"])
    df = df_plan_semana[df_plan_semana.get("categoria", "") == "Gestión Comercial"].copy()
    df = df[df["accion"].astype(str).str.strip() != ""]
    if df.empty:
        return pd.DataFrame(columns=["Ajustador", "Detalle", "Fecha"])
    return (
        df[["Ajustador", "accion", "fecha_compromiso"]]
        .rename(columns={"accion": "Detalle", "fecha_compromiso": "Fecha"})
        .drop_duplicates()
        .sort_values("Fecha")
        .reset_index(drop=True)
    )


def calcular_top5(df_maestro, filtro_division_contains=None):
    """Top 5 Compañías de seguros y Top 5 Corredoras por división, sobre casos vigentes."""
    c_div = col(df_maestro, "division")
    c_estado = col(df_maestro, "estado")
    c_comp = col(df_maestro, "compania")
    c_corr = col(df_maestro, "corredora")

    df = df_maestro.copy()
    if filtro_division_contains and c_div:
        df = df[df[c_div].astype(str).str.contains(filtro_division_contains, case=False, na=False)]
    if c_estado:
        df = df[df[c_estado].astype(str).str.strip().str.lower() != "cerrado"]

    top_comp = df[c_comp].dropna().astype(str).replace("", pd.NA).dropna().value_counts().head(5) if c_comp else pd.Series(dtype=int)
    top_corr = df[c_corr].dropna().astype(str).replace("", pd.NA).dropna().value_counts().head(5) if c_corr else pd.Series(dtype=int)
    return top_comp, top_corr


def avance_produccion(df_maestro, df_tareas_realizadas, hoy):
    """Producción acumulada anual y mensual en UF (solo IFL = trabajo facturado), por división."""
    columnas = ["Division", "Produccion_Acumulada_UF", "Produccion_Mes_UF"]
    if df_tareas_realizadas is None or df_tareas_realizadas.empty:
        return pd.DataFrame(columns=columnas)

    df = df_tareas_realizadas.copy()
    df["_es_ifl"] = df["accion"].astype(str).str.contains(REGEX_IFL, case=False, na=False)
    df = df[df["_es_ifl"]]
    if df.empty:
        return pd.DataFrame(columns=columnas)

    df["_fecha_ej"] = pd.to_datetime(df.get("fecha_ejecucion", pd.Series(dtype=str)), errors="coerce")
    df_anio = df[df["_fecha_ej"].dt.year == hoy.year].copy()

    dict_div = dict_divisiones(df_maestro)
    df_anio["Division"] = df_anio["Ajustador"].map(dict_div).fillna("Sin División")
    df_mes = df_anio[df_anio["_fecha_ej"].dt.month == hoy.month]

    acumulada = df_anio.groupby("Division")["honorarios_estimados"].sum()
    mes = df_mes.groupby("Division")["honorarios_estimados"].sum()
    areas = sorted(set(acumulada.index) | set(mes.index))

    filas = [
        {"Division": area, "Produccion_Acumulada_UF": round(acumulada.get(area, 0.0), 2), "Produccion_Mes_UF": round(mes.get(area, 0.0), 2)}
        for area in areas
    ]
    return pd.DataFrame(filas)
