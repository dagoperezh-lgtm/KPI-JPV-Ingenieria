"""
Configuración manual de Metas del Tablero Gerencial.

Estos valores NO existen en los datos de OpsControl (son objetivos que
define la gerencia, no datos de casos), así que se guardan aparte:
  - Meta semanal de Ajuste/IFL por Ajustador + Tramo UF (cuota de entregables).
  - Meta Anual y Meta Mensual de producción (UF) por División.

Se persisten localmente y, si hay conexión a Google Sheets configurada,
también se respaldan en una hoja "Metas_Tablero" del mismo Sheet que usa
OpsControl (una pestaña propia, no se toca nada de OpsControl).
"""
import json
import os

import pandas as pd
import streamlit as st

from tablero_datos import get_google_client, _abrir_documento

PERSISTENCE_DIR = "persistence_tablero"
METAS_PATH = os.path.join(PERSISTENCE_DIR, "metas_tablero.json")

# Valores copiados del Tablero manual (TABLERO_ING) — según el usuario, estas
# metas no cambian en el resto del año. La meta semanal es por Tramo (misma
# para todos los ajustadores de esa división/tramo); "metas_semanales" queda
# disponible para excepciones puntuales por ajustador que sí difieran del
# estándar de su tramo (tiene prioridad sobre "metas_por_tramo").
METAS_DEFAULT = {
    "metas_semanales": [],
    "metas_por_tramo": [
        {"division": "Ingeniería y Energía", "tramo": "<= 1000 UF", "tipo": "Ajuste", "valor": 4},
        {"division": "Ingeniería y Energía", "tramo": "<= 1000 UF", "tipo": "IFL", "valor": 4},
        {"division": "Ingeniería y Energía", "tramo": "> 1000 Y <= 5000 UF", "tipo": "Ajuste", "valor": 3},
        {"division": "Ingeniería y Energía", "tramo": "> 1000 Y <= 5000 UF", "tipo": "IFL", "valor": 3},
        {"division": "Ingeniería y Energía", "tramo": "> 5000 UF (MCL)", "tipo": "Ajuste", "valor": 1},
        {"division": "Ingeniería y Energía", "tramo": "> 5000 UF (MCL)", "tipo": "IFL", "valor": 1},
        {"division": "Equipo Móvil", "tramo": "General", "tipo": "Ajuste", "valor": 4},
        {"division": "Equipo Móvil", "tramo": "General", "tipo": "IFL", "valor": 4},
    ],
    "metas_anuales": [
        {"division": "Ingeniería y Energía", "meta_anual_uf": 32000, "meta_mensual_uf": round(32000 / 12, 2)},
        {"division": "Equipo Móvil", "meta_anual_uf": 7400, "meta_mensual_uf": round(7400 / 12, 2)},
    ],
}


def _init():
    os.makedirs(PERSISTENCE_DIR, exist_ok=True)


def _cargar_de_nube():
    client = get_google_client()
    if client is None:
        return None
    try:
        doc = _abrir_documento(client)
        try:
            ws = doc.worksheet("Metas_Tablero")
        except Exception:
            return None
        contenido = ws.acell("A1").value
        if not contenido:
            return None
        return json.loads(contenido)
    except Exception:
        return None


def _guardar_en_nube(metas):
    client = get_google_client()
    if client is None:
        return False
    try:
        doc = _abrir_documento(client)
        try:
            ws = doc.worksheet("Metas_Tablero")
        except Exception:
            ws = doc.add_worksheet(title="Metas_Tablero", rows=10, cols=2)
        ws.update("A1", [[json.dumps(metas, ensure_ascii=False)]])
        return True
    except Exception as e:
        st.warning(f"⚠️ No se pudo respaldar las metas en la nube: {e}")
        return False


def _completar_con_default(metas):
    """Agrega las claves que falten (p.ej. metas_por_tramo en un archivo
    guardado antes de que existiera) sin pisar lo que el usuario ya guardó."""
    completo = dict(METAS_DEFAULT)
    completo.update(metas or {})
    for clave in METAS_DEFAULT:
        if clave not in metas or metas[clave] is None:
            completo[clave] = METAS_DEFAULT[clave]
    return completo


def load_metas():
    _init()
    if os.path.exists(METAS_PATH):
        try:
            with open(METAS_PATH, "r", encoding="utf-8") as f:
                return _completar_con_default(json.load(f))
        except Exception:
            pass
    metas_nube = _cargar_de_nube()
    if metas_nube is not None:
        metas_nube = _completar_con_default(metas_nube)
        save_metas(metas_nube, respaldar_nube=False)
        return metas_nube
    return dict(METAS_DEFAULT)


def save_metas(metas, respaldar_nube=True):
    _init()
    with open(METAS_PATH, "w", encoding="utf-8") as f:
        json.dump(metas, f, ensure_ascii=False, indent=2)
    if respaldar_nube:
        _guardar_en_nube(metas)


def metas_semanales_como_dict(metas):
    """[{ajustador,tramo,tipo,valor}] -> {(ajustador,tramo,tipo): valor} para construir_grilla()."""
    return {
        (m["ajustador"], m["tramo"], m["tipo"]): m["valor"]
        for m in metas.get("metas_semanales", [])
        if m.get("valor") not in (None, "")
    }


def metas_semanales_a_df(metas):
    filas = metas.get("metas_semanales", [])
    if not filas:
        return pd.DataFrame(columns=["ajustador", "tramo", "tipo", "valor"])
    return pd.DataFrame(filas)


def metas_semanales_desde_df(df):
    df = df.copy()
    df["valor"] = pd.to_numeric(df["valor"], errors="coerce")
    return df.dropna(subset=["ajustador", "tramo", "tipo"]).to_dict(orient="records")


def metas_por_tramo_como_dict(metas):
    """[{division,tramo,tipo,valor}] -> {(division,tramo,tipo): valor}."""
    return {
        (m["division"], m["tramo"], m["tipo"]): m["valor"]
        for m in metas.get("metas_por_tramo", [])
        if m.get("valor") not in (None, "")
    }


def metas_por_tramo_a_df(metas):
    filas = metas.get("metas_por_tramo", [])
    if not filas:
        return pd.DataFrame(columns=["division", "tramo", "tipo", "valor"])
    return pd.DataFrame(filas)


def metas_por_tramo_desde_df(df):
    df = df.copy()
    df["valor"] = pd.to_numeric(df["valor"], errors="coerce")
    return df.dropna(subset=["division", "tramo", "tipo"]).to_dict(orient="records")


def metas_anuales_a_df(metas):
    filas = metas.get("metas_anuales", [])
    if not filas:
        return pd.DataFrame(columns=["division", "meta_anual_uf", "meta_mensual_uf"])
    return pd.DataFrame(filas)


def metas_anuales_desde_df(df):
    df = df.copy()
    df["meta_anual_uf"] = pd.to_numeric(df["meta_anual_uf"], errors="coerce").fillna(0)
    df["meta_mensual_uf"] = pd.to_numeric(df["meta_mensual_uf"], errors="coerce").fillna(0)
    return df.dropna(subset=["division"]).to_dict(orient="records")
