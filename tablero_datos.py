"""
Capa de datos del Tablero Gerencial.

Este módulo es de solo lectura: se conecta al MISMO Google Sheet que ya
usa OpsControl (el planificador) para leer la Base Maestra (export del
sistema de casos) y los planes semanales guardados por cada ajustador
(plan_<ajustador>_<semana>.json). No escribe nada en OpsControl.

Requiere en st.secrets las mismas claves que OpsControl:
    st.secrets["gcp_service_account"]
    st.secrets["google_sheet_url"]
"""
import json
from datetime import datetime, timedelta

import pandas as pd
import streamlit as st

try:
    import gspread
    from google.oauth2.service_account import Credentials
except ImportError:
    gspread = None
    Credentials = None


# ---------------------------------------------------------
# CONEXIÓN A GOOGLE SHEETS (solo lectura)
# ---------------------------------------------------------
def get_google_client():
    if gspread is None:
        return None
    try:
        if "gcp_service_account" in st.secrets and "google_sheet_url" in st.secrets:
            scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
            creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
            return gspread.authorize(creds)
    except Exception as e:
        st.error(f"⚠️ Error de conexión a Google Cloud: {e}")
    return None


def _abrir_documento(client):
    return client.open_by_url(st.secrets["google_sheet_url"])


@st.cache_data(ttl=300, show_spinner="Leyendo Base Maestra desde OpsControl...")
def load_base_maestra():
    """Lee la hoja 'Base_Maestra' tal como la guarda OpsControl (fila 1 = metadata,
    fila 2 = encabezados, resto = datos)."""
    client = get_google_client()
    if client is None:
        return None, None
    try:
        doc = _abrir_documento(client)
        ws = doc.worksheet("Base_Maestra")
        metadata = ws.row_values(1)
        fecha_actualizacion = None
        if len(metadata) >= 2 and metadata[0] == "FECHA_ACTUALIZACION":
            fecha_actualizacion = metadata[1]
        registros = ws.get_all_records(head=2)
        df = pd.DataFrame(registros)
        return (df if not df.empty else None), fecha_actualizacion
    except Exception as e:
        st.error(f"⚠️ No se pudo leer la Base Maestra: {e}")
        return None, None


@st.cache_data(ttl=300, show_spinner="Leyendo planes semanales desde OpsControl...")
def _leer_todos_los_planes_raw():
    """Descarga la hoja principal (Archivo / JSON_Data) completa una sola vez."""
    client = get_google_client()
    if client is None:
        return []
    try:
        doc = _abrir_documento(client)
        return doc.sheet1.get_all_records()
    except Exception as e:
        st.error(f"⚠️ No se pudo leer los planes: {e}")
        return []


def parse_ajustador_de_archivo(nombre_archivo):
    partes = nombre_archivo.replace("plan_mensual_mcl_", "").replace("plan_", "").replace(".json", "")
    return partes.split("_20")[0].replace("_", " ")


def load_planes_semana(week_id):
    """Consolida en un DataFrame todas las tareas (de todos los ajustadores)
    del plan SEMANAL correspondiente a week_id (formato get_week_identifier)."""
    registros = _leer_todos_los_planes_raw()
    tareas = []
    for fila in registros:
        nombre = str(fila.get("Archivo", ""))
        if not nombre.endswith(".json"):
            continue
        if nombre == "BASE_MAESTRA.json" or nombre.startswith("plan_mensual_mcl_"):
            continue
        if not nombre.startswith("plan_") or not nombre.endswith(f"_{week_id}.json"):
            continue
        ajustador = parse_ajustador_de_archivo(nombre)
        try:
            plan = json.loads(fila.get("JSON_Data", "[]"))
        except Exception:
            continue
        for tarea in plan:
            t = dict(tarea)
            t["Ajustador"] = ajustador
            tareas.append(t)

    df = pd.DataFrame(tareas)
    if not df.empty:
        if "honorarios_estimados" not in df.columns:
            df["honorarios_estimados"] = 0.0
        df["honorarios_estimados"] = pd.to_numeric(df["honorarios_estimados"], errors="coerce").fillna(0.0)
        for col in ["tramo_uf", "tipo_actividad", "estado_cumplimiento", "categoria", "accion", "numero_caso", "asegurado"]:
            if col not in df.columns:
                df[col] = ""
    return df


# ---------------------------------------------------------
# UTILIDADES DE SEMANA (idénticas a OpsControl para calzar filenames)
# ---------------------------------------------------------
def get_week_identifier(offset_weeks=0, referencia=None):
    target_date = (referencia or datetime.now()) + timedelta(weeks=offset_weeks)
    return target_date.strftime("%Y_W%W")


def dias_habiles_semana(offset_weeks=0, referencia=None):
    base = (referencia or datetime.now()) + timedelta(weeks=offset_weeks)
    lunes = base - timedelta(days=base.weekday())
    return [(lunes + timedelta(days=i)).date() for i in range(5)]


def rango_semana_legible(offset_weeks=0, referencia=None):
    dias = dias_habiles_semana(offset_weeks, referencia)
    return f"{dias[0].strftime('%d de %B')} al {dias[-1].strftime('%d de %B de %Y')}"


# ---------------------------------------------------------
# CLASIFICACIÓN DE TRAMO UF / MCL (idéntica a OpsControl)
# ---------------------------------------------------------
def limpiar_monto(valor):
    if pd.isna(valor) or str(valor).strip() in ["", "nan", "NaN", "NaT"]:
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    v_str = str(valor).strip().replace("$", "").replace(" ", "")
    if "." in v_str and "," in v_str:
        if v_str.rfind(",") > v_str.rfind("."):
            v_str = v_str.replace(".", "").replace(",", ".")
        else:
            v_str = v_str.replace(",", "")
    elif "," in v_str:
        v_str = v_str.replace(",", ".")
    try:
        return float(v_str)
    except Exception:
        return 0.0


def parsear_fecha_flexible(valor):
    """Igual que parsear_fecha_mov() de OpsControl: acepta texto dd/mm/aaaa
    (formato chileno) y también números de serie de fecha de Excel (cuando
    la celda llega como número plano en vez de fecha formateada)."""
    if valor is None or str(valor).strip() in ("", "nan", "NaT", "None"):
        return pd.NaT
    s = str(valor).strip()
    try:
        if s.isdigit() or (s.replace(".", "", 1).isdigit() and "." in s):
            serial = int(float(s))
            if 20000 < serial < 60000:
                return pd.Timestamp("1899-12-30") + pd.Timedelta(days=serial)
    except Exception:
        pass
    return pd.to_datetime(s, dayfirst=True, errors="coerce")


def calcular_tramo(fila, col_perdida="Perdida bruta (en moneda del caso)", col_divisa="Divisa"):
    divisa = str(fila.get(col_divisa, "")).upper()
    valor = limpiar_monto(fila.get(col_perdida, 0)) if col_perdida in fila else 0.0

    is_mcl = False
    if "USD" in divisa or "US$" in divisa or "DÓLAR" in divisa or "DOLAR" in divisa:
        if valor > 200000:
            is_mcl = True
            tramo_str = "> 200.000 USD (MCL)"
        else:
            tramo_str = "<= 200.000 USD"
    else:
        if valor <= 1000:
            tramo_str = "<= 1000 UF"
        elif valor <= 5000:
            tramo_str = "> 1000 Y <= 5000 UF"
        else:
            is_mcl = True
            tramo_str = "> 5000 UF (MCL)"
    return tramo_str, is_mcl


# ---------------------------------------------------------
# LOCALIZACIÓN FLEXIBLE DE COLUMNAS EN LA BASE MAESTRA
# (el export del sistema puede variar de nombre entre versiones)
# ---------------------------------------------------------
COLUMNAS_BASE_MAESTRA = {
    "numero_caso": (["Número de caso", "Numero de caso"], 0),
    "nickname": (["Nickname"], 2),
    "division": (["División", "Division"], 3),
    "compania": (["Compañía de seguros"], 6),
    "corredora": (["Corredora"], 7),
    "ajustador": (["Ajustador senior"], 9),
    "asegurado": (["Asegurado"], 11),
    "estado": (["Estado"], 18),
    "sub_estado": (["Sub estado"], 19),
    "creado_en": (["Creado en", "Fecha de denuncio"], None),
    "ultimo_movimiento": (["Último movimiento"], 70),
    "contenido_ultimo_movimiento": (["Contenido último movimiento"], 71),
    "honorarios": ([], 66),
}


def col(df, clave):
    nombres, idx_fallback = COLUMNAS_BASE_MAESTRA[clave]
    for n in nombres:
        if n in df.columns:
            return n
    if idx_fallback is not None and len(df.columns) > idx_fallback:
        return df.columns[idx_fallback]
    return None


def honorarios_del_caso(fila, df_columns):
    """Rescata el honorario estimado del caso igual que OpsControl (columna 67, iloc[66])."""
    try:
        if len(df_columns) >= 67:
            return limpiar_monto(fila.iloc[66])
    except Exception:
        pass
    return 0.0


def dict_divisiones(df_maestro):
    """{ajustador: división} a partir de la Base Maestra, igual que el mapeo usado
    en la Carta Gantt de OpsControl."""
    if df_maestro is None or df_maestro.empty:
        return {}
    c_div = col(df_maestro, "division")
    c_aj = col(df_maestro, "ajustador")
    if not c_div or not c_aj:
        return {}
    resultado = {}
    for _, row in df_maestro.iterrows():
        aj = str(row[c_aj]).strip()
        div = str(row[c_div]).strip()
        if aj and div:
            resultado[aj] = div
    return resultado


@st.cache_data(ttl=300, show_spinner="Leyendo histórico anual de planes...")
def load_todas_las_tareas_realizadas():
    """Consolida TODAS las tareas Realizadas de todos los planes semanales
    (de todos los ajustadores, de todas las semanas disponibles en la nube).
    Se usa para calcular producción acumulada anual/mensual."""
    registros = _leer_todos_los_planes_raw()
    tareas = []
    for fila in registros:
        nombre = str(fila.get("Archivo", ""))
        if not nombre.endswith(".json") or nombre == "BASE_MAESTRA.json":
            continue
        if not nombre.startswith("plan_") or nombre.startswith("plan_mensual_mcl_"):
            continue
        ajustador = parse_ajustador_de_archivo(nombre)
        try:
            plan = json.loads(fila.get("JSON_Data", "[]"))
        except Exception:
            continue
        for tarea in plan:
            if tarea.get("estado_cumplimiento") != "Realizado":
                continue
            t = dict(tarea)
            t["Ajustador"] = ajustador
            tareas.append(t)

    df = pd.DataFrame(tareas)
    if not df.empty:
        if "honorarios_estimados" not in df.columns:
            df["honorarios_estimados"] = 0.0
        df["honorarios_estimados"] = pd.to_numeric(df["honorarios_estimados"], errors="coerce").fillna(0.0)
    return df
