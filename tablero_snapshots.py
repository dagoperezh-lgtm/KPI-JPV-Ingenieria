"""
Congelamiento semanal del Tablero.

A partir de que una semana deja de ser la "Semana Actual", su grilla se
guarda tal cual (Stock, Asignados, Q programado/Meta/Hon programado,
notas) en una hoja del Google Sheet ("Tablero_Snapshots"). De ahí en
adelante esa foto ya NO se recalcula — solo se refresca el Q
realizado/Hon realizado (lo que se va marcando como completado),
igual como se reportó la semana en su momento.
"""
import json

import pandas as pd
import streamlit as st

from tablero_datos import get_google_client, _abrir_documento

NOMBRE_HOJA = "Tablero_Snapshots"

# Columnas que quedan congeladas para siempre una vez guardada la semana.
COLUMNAS_CONGELADAS = [
    "Division", "Ajustador", "Tramo",
    "Asignados", "Stock_Q", "Dias_prom", "Stock_Hon_UF",
    "Ajuste_Qp", "Ajuste_Meta", "Ajuste_HonpUF",
    "IFL_Qp", "IFL_Meta", "IFL_HonpUF",
    "Gestion_Comercial", "Extra",
]
# Columnas que siempre se recalculan en vivo (lo único que puede seguir cambiando).
COLUMNAS_VIVAS = ["Ajuste_Qr", "Ajuste_HonrUF", "IFL_Qr", "IFL_HonrUF"]


def _get_worksheet(client):
    doc = _abrir_documento(client)
    try:
        return doc.worksheet(NOMBRE_HOJA)
    except Exception:
        ws = doc.add_worksheet(title=NOMBRE_HOJA, rows=200, cols=3)
        ws.append_row(["Semana", "JSON_Data", "Fecha_Guardado"])
        return ws


@st.cache_data(ttl=300, show_spinner=False)
def load_snapshot(week_id):
    """Devuelve el DataFrame congelado de esa semana, o None si nunca se guardó."""
    client = get_google_client()
    if client is None:
        return None
    try:
        ws = _get_worksheet(client)
        registros = ws.get_all_records()
        for fila in registros:
            if str(fila.get("Semana", "")) == week_id:
                datos = json.loads(fila.get("JSON_Data", "[]"))
                return pd.DataFrame(datos) if datos else None
    except Exception:
        return None
    return None


def existe_snapshot(week_id):
    return load_snapshot(week_id) is not None


def save_snapshot(week_id, grilla_con_totales):
    """Guarda (o sobrescribe) la foto congelada de esa semana. Solo se
    persisten las columnas que deben quedar fijas; Qr/HonrUF NO se guardan
    porque siempre se recalculan en vivo."""
    client = get_google_client()
    if client is None or grilla_con_totales is None or grilla_con_totales.empty:
        return False
    try:
        columnas_presentes = [c for c in COLUMNAS_CONGELADAS if c in grilla_con_totales.columns]
        datos = grilla_con_totales[columnas_presentes].to_dict(orient="records")
        ws = _get_worksheet(client)
        from datetime import datetime
        fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        celdas = ws.col_values(1)
        json_str = json.dumps(datos, ensure_ascii=False, default=str)
        if week_id in celdas:
            fila_idx = celdas.index(week_id) + 1
            ws.update(f"A{fila_idx}:C{fila_idx}", [[week_id, json_str, fecha]])
        else:
            ws.append_row([week_id, json_str, fecha])
        load_snapshot.clear()
        return True
    except Exception as e:
        st.warning(f"⚠️ No se pudo guardar la foto de la semana en la nube: {e}")
        return False


def combinar_congelado_con_vivo(grilla_congelada, ejecucion_viva):
    """grilla_congelada: DataFrame ya cargado con load_snapshot() (columnas
    congeladas). ejecucion_viva: salida de tablero_calculo._preparar_ejecucion()
    para la misma semana (Ajustador, Tramo, *_Qr, *_HonrUF actualizados a hoy).
    Devuelve la grilla final: estructura y planificado congelados, realizado
    siempre al día. Si aparece una fila nueva en vivo que no estaba en la foto
    (p.ej. una actividad adicional en un tramo que no existía), se agrega."""
    congelada = grilla_congelada.copy()
    for c in COLUMNAS_VIVAS:
        if c not in congelada.columns:
            congelada[c] = 0.0

    if ejecucion_viva is None or ejecucion_viva.empty:
        for c in COLUMNAS_VIVAS:
            congelada[c] = 0.0
        return congelada

    vivo = ejecucion_viva.set_index(["Ajustador", "Tramo"])
    cols_qr = [c for c in COLUMNAS_VIVAS if c in ejecucion_viva.columns]

    congelada = congelada.set_index(["Ajustador", "Tramo"])
    for c in cols_qr:
        congelada[c] = vivo[c].reindex(congelada.index).fillna(0.0)
    congelada = congelada.reset_index()

    # Filas de ejecución en vivo que no calzaron con ninguna fila congelada
    # (tramo distinto al reportado en su momento): se agregan aparte para no perder el dato.
    claves_congeladas = set(zip(grilla_congelada["Ajustador"], grilla_congelada["Tramo"]))
    extra = ejecucion_viva[~ejecucion_viva.apply(lambda r: (r["Ajustador"], r["Tramo"]) in claves_congeladas, axis=1)].copy()
    if not extra.empty:
        dict_div = dict(zip(grilla_congelada["Ajustador"], grilla_congelada["Division"]))
        extra["Division"] = extra["Ajustador"].map(dict_div).fillna("Sin División")
        for c in COLUMNAS_CONGELADAS:
            if c not in extra.columns:
                extra[c] = 0 if c not in ("Gestion_Comercial", "Extra") else ""
        congelada = pd.concat([congelada, extra], ignore_index=True, sort=False)

    return congelada
