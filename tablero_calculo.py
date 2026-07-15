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

from tablero_datos import calcular_tramo, col, dict_divisiones, honorarios_del_caso, parsear_fecha_flexible

REGEX_IFL = r"Informe Final de Liquidación|Carta de Cobertura \(Rechazo\)"

COLUMNAS_GRILLA = [
    "Division", "Ajustador", "Tramo",
    "Asignados", "Stock_Q", "Dias_prom", "Stock_Hon_UF",
    "Ajuste_Qp", "Ajuste_Meta", "Ajuste_HonpUF", "Ajuste_Qr", "Ajuste_HonrUF",
    "IFL_Qp", "IFL_Meta", "IFL_HonpUF", "IFL_Qr", "IFL_HonrUF",
]

COLUMNAS_NUMERICAS = [c for c in COLUMNAS_GRILLA if c not in ("Division", "Ajustador", "Tramo", "Dias_prom")]


def _es_equipo_movil(texto_division):
    return bool(pd.notna(texto_division)) and ("MÓVIL" in str(texto_division).upper() or "MOVIL" in str(texto_division).upper())


def _normalizar_division(texto_division):
    """La división real puede traer texto descriptivo adicional (p.ej.
    'Equipo Móvil, Construcción e Industría'); para buscar la meta estándar
    se usa una categoría fija, no el texto exacto."""
    if _es_equipo_movil(texto_division):
        return "Equipo Móvil"
    if texto_division and "INGENIER" in str(texto_division).upper():
        return "Ingeniería y Energía"
    return texto_division


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
    # Equipo Móvil no se segmenta por tramo de pérdida: una sola fila por ajustador.
    vigentes.loc[vigentes["_division"].apply(_es_equipo_movil), "_tramo"] = "General"
    vigentes["_honorarios"] = vigentes.apply(lambda r: honorarios_del_caso(r, df.columns), axis=1)

    if c_creado and c_creado in vigentes.columns:
        vigentes["_fecha_creado"] = pd.to_datetime(vigentes[c_creado].apply(parsear_fecha_flexible))
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


def _mapa_tramo_por_caso(df_maestro):
    """{numero_caso: tramo} recalculado con calcular_tramo() (UF unificado),
    para no depender del tramo que OpsControl guardó en su momento dentro de
    cada tarea del plan (ese valor es de OTRA app, con su propia lógica de
    tramos que todavía separa USD)."""
    if df_maestro is None or df_maestro.empty:
        return {}
    c_num = col(df_maestro, "numero_caso")
    if not c_num:
        return {}
    tramos, _ = zip(*df_maestro.apply(calcular_tramo, axis=1)) if len(df_maestro) else ([], [])
    return dict(zip(df_maestro[c_num].astype(str).str.strip(), tramos))


# Para casos que ya no están en la Base Maestra (se cerraron y OpsControl los
# excluye del export): no hay monto para recalcular su tramo exacto, así que
# se traduce el tramo heredado (con USD) a su equivalente más cercano en UF,
# en vez de mostrar la etiqueta USD tal cual.
REMAPEO_TRAMO_LEGADO = {
    "<= 200.000 USD": "> 1000 Y <= 5000 UF",
    "> 200.000 USD (MCL)": "> 5000 UF (MCL)",
}


def _preparar_ejecucion(df_plan_semana, mapa_tramo_caso=None, dict_div=None):
    columnas_vacias = ["Ajustador", "Tramo", "Ajuste_Qp", "Ajuste_HonpUF", "Ajuste_Qr", "Ajuste_HonrUF",
                        "IFL_Qp", "IFL_HonpUF", "IFL_Qr", "IFL_HonrUF"]
    if df_plan_semana is None or df_plan_semana.empty:
        return pd.DataFrame(columns=columnas_vacias)

    df = df_plan_semana[df_plan_semana.get("categoria", "") == "Operativa"].copy()
    if df.empty:
        return pd.DataFrame(columns=columnas_vacias)

    mapa_tramo_caso = mapa_tramo_caso or {}
    dict_div = dict_div or {}
    df["_es_ifl"] = df["accion"].astype(str).str.contains(REGEX_IFL, case=False, na=False)
    # Tramo recalculado a partir del caso en la Base Maestra (fuente de verdad,
    # siempre en UF). Si el caso ya no está en la Base Maestra (p.ej. se cerró
    # y salió del export), se traduce el tramo heredado de la tarea a UF.
    numero_caso = df["numero_caso"].astype(str).str.strip()
    tramo_heredado = df["tramo_uf"].replace(REMAPEO_TRAMO_LEGADO)
    df["_tramo"] = numero_caso.map(mapa_tramo_caso)
    df["_tramo"] = df["_tramo"].fillna(tramo_heredado).replace("", "N/D").fillna("N/D")
    # Equipo Móvil no se segmenta por tramo: una sola fila por ajustador.
    es_movil = df["Ajustador"].map(dict_div).apply(_es_equipo_movil)
    df.loc[es_movil, "_tramo"] = "General"

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


def construir_grilla(df_maestro, df_plan_semana, dias_semana, metas_ajustador=None, metas_tramo=None):
    """Devuelve la grilla Division/Ajustador/Tramo con Stock + Ajuste + IFL,
    lista para mostrar tal como el Excel TABLERO_ING. Equipo Móvil no se
    segmenta por tramo (Tramo='General').

    metas_ajustador: {(ajustador,tramo,tipo): valor} — excepción puntual.
    metas_tramo: {(division,tramo,tipo): valor} — estándar por tramo/división,
    se usa cuando no hay excepción para ese ajustador."""
    dict_div = dict_divisiones(df_maestro)
    stock = _preparar_stock(df_maestro, dias_semana)
    ejecucion = _preparar_ejecucion(df_plan_semana, _mapa_tramo_por_caso(df_maestro), dict_div)

    dict_div_stock = dict(zip(stock["Ajustador"], stock["Division"])) if not stock.empty else {}
    ejecucion = ejecucion.copy()
    if not ejecucion.empty:
        ejecucion["Division"] = ejecucion["Ajustador"].map(dict_div_stock)
        ejecucion["Division"] = ejecucion["Division"].fillna(ejecucion["Ajustador"].map(dict_div)).fillna("Sin División")
    else:
        ejecucion["Division"] = pd.Series(dtype=str)

    if stock.empty and ejecucion.empty:
        return pd.DataFrame(columns=COLUMNAS_GRILLA)

    grilla = pd.merge(stock, ejecucion, on=["Division", "Ajustador", "Tramo"], how="outer")
    if grilla.empty:
        return pd.DataFrame(columns=COLUMNAS_GRILLA)

    for c in COLUMNAS_NUMERICAS:
        if c not in grilla.columns:
            grilla[c] = 0
        grilla[c] = pd.to_numeric(grilla[c], errors="coerce").fillna(0)

    if "Dias_prom" in grilla.columns:
        grilla["Dias_prom"] = grilla["Dias_prom"].round(0)

    metas_ajustador = metas_ajustador or {}
    metas_tramo = metas_tramo or {}

    def _meta(fila, tipo):
        clave_aj = (fila["Ajustador"], fila["Tramo"], tipo)
        if clave_aj in metas_ajustador:
            return metas_ajustador[clave_aj]
        return metas_tramo.get((_normalizar_division(fila["Division"]), fila["Tramo"], tipo), None)

    grilla["Ajuste_Meta"] = grilla.apply(lambda r: _meta(r, "Ajuste"), axis=1)
    grilla["IFL_Meta"] = grilla.apply(lambda r: _meta(r, "IFL"), axis=1)

    for c in COLUMNAS_GRILLA:
        if c not in grilla.columns:
            grilla[c] = None

    return grilla[COLUMNAS_GRILLA].sort_values(["Division", "Ajustador", "Tramo"]).reset_index(drop=True)


def notas_por_ajustador(df_plan_semana):
    """{ajustador: {'comercial': texto, 'extra': texto}} — igual que las columnas
    'Gestión Comercial' y 'Extra' del Excel: notas libres de la semana. 'Comercial'
    sale de categoria=='Gestión Comercial'; 'Extra' de categoria=='Gestión
    Administrativa' (no hay un campo separado 'Extra' en el planificador, así
    que se usa la gestión administrativa como equivalente más cercano)."""
    vacio = {}
    if df_plan_semana is None or df_plan_semana.empty:
        return vacio

    def _juntar(categoria):
        df = df_plan_semana[df_plan_semana.get("categoria", "") == categoria].copy()
        df = df[df["accion"].astype(str).str.strip() != ""]
        if df.empty:
            return {}
        return df.groupby("Ajustador")["accion"].apply(lambda s: " | ".join(sorted(set(s.astype(str))))).to_dict()

    comercial = _juntar("Gestión Comercial")
    extra = _juntar("Gestión Administrativa")
    ajustadores = set(comercial) | set(extra)
    return {aj: {"comercial": comercial.get(aj, ""), "extra": extra.get(aj, "")} for aj in ajustadores}


def construir_grilla_con_totales(df_maestro, df_plan_semana, dias_semana, metas_ajustador=None, metas_tramo=None):
    """Igual que construir_grilla(), pero en el formato del Excel: una fila
    por Ajustador+Tramo (el nombre va PEGADO a su tramo, no en una fila
    aparte por encima de la segmentación). Las notas de Gestión Comercial /
    Extra de la semana se muestran en la PRIMERA fila de cada ajustador.
    Se descartan las filas sin ningún dato (tramo 'N/D' completamente vacío)."""
    grilla = construir_grilla(df_maestro, df_plan_semana, dias_semana, metas_ajustador, metas_tramo)
    if grilla.empty:
        return grilla

    grilla = grilla.copy()
    # Días prom nunca puede ser negativo (un caso no puede "crearse en el futuro");
    # si sale negativo es una fecha mal registrada en el sistema, se descarta el valor.
    grilla.loc[grilla["Dias_prom"] < 0, "Dias_prom"] = None

    # Descarta filas fantasma sin ningún dato (pueden aparecer para cualquier
    # tramo por el cruce Stock/Ejecución, no solo en 'N/D').
    cols_actividad = ["Stock_Q", "Ajuste_Qp", "Ajuste_Qr", "IFL_Qp", "IFL_Qr"]
    fila_vacia = grilla[cols_actividad].fillna(0).sum(axis=1) == 0
    grilla = grilla[~fila_vacia].reset_index(drop=True)
    if grilla.empty:
        return grilla

    notas = notas_por_ajustador(df_plan_semana)
    grilla["Gestion_Comercial"] = ""
    grilla["Extra"] = ""
    orden_tramo = {"General": 0, "<= 1000 UF": 0, "> 1000 Y <= 5000 UF": 1, "> 5000 UF (MCL)": 2, "N/D": 9}
    grilla["_orden"] = grilla["Tramo"].map(orden_tramo).fillna(5)
    grilla = grilla.sort_values(["Division", "Ajustador", "_orden"]).drop(columns="_orden").reset_index(drop=True)

    primera_fila_idx = grilla.groupby(["Division", "Ajustador"], sort=False).head(1).index
    for idx in primera_fila_idx:
        ajustador = grilla.loc[idx, "Ajustador"]
        grilla.loc[idx, "Gestion_Comercial"] = notas.get(ajustador, {}).get("comercial", "")
        grilla.loc[idx, "Extra"] = notas.get(ajustador, {}).get("extra", "")

    return grilla


def subtotales_por_division(grilla_con_totales):
    """Subtotal_Ingeniería / Subtotal_Equipo_Móvil / Total General."""
    if grilla_con_totales.empty:
        return pd.DataFrame(columns=["Division"] + COLUMNAS_NUMERICAS)
    suma_cols = [c for c in COLUMNAS_NUMERICAS if c not in ("Ajuste_Meta", "IFL_Meta")]
    subtotal = grilla_con_totales.groupby("Division")[suma_cols].sum().reset_index()
    total = pd.DataFrame([{"Division": "TOTAL GENERAL", **{c: grilla_con_totales[c].sum() for c in suma_cols}}])
    return pd.concat([subtotal, total], ignore_index=True)


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
    df_anio["Division"] = df_anio["Ajustador"].map(dict_div).fillna("Sin División").apply(_normalizar_division)
    df_mes = df_anio[df_anio["_fecha_ej"].dt.month == hoy.month]

    acumulada = df_anio.groupby("Division")["honorarios_estimados"].sum()
    mes = df_mes.groupby("Division")["honorarios_estimados"].sum()
    areas = sorted(set(acumulada.index) | set(mes.index))

    filas = [
        {"Division": area, "Produccion_Acumulada_UF": round(acumulada.get(area, 0.0), 2), "Produccion_Mes_UF": round(mes.get(area, 0.0), 2)}
        for area in areas
    ]
    return pd.DataFrame(filas)
