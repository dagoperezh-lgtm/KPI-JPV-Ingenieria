"""
Generador del Reporte Ejecutivo "Estado de Cartera".

Cada slide se construye 100% por código con branding.py (header navy + logo
JPV Asociados · Sedgwick Alliance, footer teal, cuerpo claro, 16:9), en vez
de depender de un archivo .pptx plantilla con shape_id fijos: así se puede
agregar o quitar slides (Detalle Completo, Top 5, MCL/No MCL) sin tocar un
archivo binario aparte.
Sirve para cualquier corte de cartera: por Corredora, Compañía de seguros
(Aseguradora), Asegurado, o una combinación de estos filtros.
"""
import io
import re

import pandas as pd
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Emu, Pt

import branding as br

try:
    import gspread
    from google.oauth2.service_account import Credentials
except ImportError:
    gspread = None
    Credentials = None

# Pipeline actualizado por el equipo en Google Sheets, compartido en modo
# Lector con la cuenta de servicio opscontrol-bot@... (misma cuenta que ya
# usa OpsControl para la Base Maestra en tablero_datos.py). Los datos viven
# en la pestaña Pipeline_Backup, con la misma estructura que Base_Maestra:
# fila 1 = metadata (FECHA_ACTUALIZACION), fila 2 = encabezados reales.
PIPELINE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1aZYWn1JI_w4S9mQFpaCJRAQfr6E0iKqXMcMJ_xxlaNI/edit?usp=sharing"
PIPELINE_WORKSHEET_NAME = "Pipeline_Backup"
PIPELINE_HEADER_ROW = 2

MCL_UF = 5000
MCL_USD = 200000

# La columna "Indicación Probabilidad" del Pipeline usa estas 4 categorías.
INDICACION_A_PROB = {
    "cierta": 100,
    "altamente probable": 75,
    "podría ser": 50,
    "podria ser": 50,
    "nula": 0,
}

_TIMESTAMP_PREFIJO_RE = re.compile(r"^\s*\[\d{1,2}/\d{1,2}/\d{2,4}[^\]]*\]\s*")

MESES_ES = {
    1: "enero", 2: "febrero", 3: "marzo", 4: "abril", 5: "mayo", 6: "junio",
    7: "julio", 8: "agosto", 9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre",
}


def fmt_uf(valor):
    return f"{valor:,.0f}".replace(",", ".") + " UF"


def fmt_usd_m(valor):
    m = valor / 1_000_000
    s = f"{m:.1f}".replace(".", ",")
    return f"USD {s} M"


def fmt_fecha_larga(fecha):
    return f"{fecha.day} de {MESES_ES[fecha.month]} de {fecha.year}"


def _fila_siniestro_nickname(siniestro, nickname, max_chars=70):
    """Combina el N° de siniestro con el nickname para las columnas 'Siniestro
    / Nickname' de las tablas (Tiempos de Residencia, Atención Especial y
    Gestiones Iniciales). El N° de siniestro nunca se trunca (es un
    identificador); el nickname se recorta con el espacio restante."""
    siniestro = str(siniestro).strip()
    nickname = str(nickname).strip()
    if not siniestro or siniestro.lower() in ("nan", "none"):
        return _truncar_texto(nickname, max_chars)
    prefijo = f"Sin. {siniestro} — "
    espacio_nickname = max(15, max_chars - len(prefijo))
    return prefijo + _truncar_texto(nickname, espacio_nickname)


def _truncar_texto(texto, max_chars):
    """Recorta observaciones largas para que la fila no crezca más de lo que
    permite la tabla (si no, las últimas filas quedan fuera del área visible
    de la slide en modo presentación)."""
    texto = str(texto).strip()
    if len(texto) <= max_chars:
        return texto
    return texto[: max_chars - 1].rstrip() + "…"


def cargar_pipeline(archivo, sheet_name=None):
    """Lee el Excel de Pipeline (una fila por caso activo, con Probabilidad de
    cierre y Observaciones ya cargadas por el equipo)."""
    xl = pd.ExcelFile(archivo)
    hoja = sheet_name or xl.sheet_names[0]
    return pd.read_excel(archivo, sheet_name=hoja)


def _cliente_google_sheets(credenciales_sa):
    if gspread is None:
        raise RuntimeError("Faltan las librerías 'gspread' y 'google-auth' (revisa requirements.txt).")
    scope = [
        "https://www.googleapis.com/auth/spreadsheets.readonly",
        "https://www.googleapis.com/auth/drive.readonly",
    ]
    creds = Credentials.from_service_account_info(credenciales_sa, scopes=scope)
    return gspread.authorize(creds)


def cargar_hoja_desde_google_sheets(credenciales_sa, worksheet_name, header_row=1, sheet_url=None):
    """Descarga cualquier pestaña del Google Sheet de OpsControl (Pipeline_Backup,
    Base_Maestra, etc.) usando una cuenta de servicio (st.secrets["gcp_service_account"]).
    No requiere que la hoja esté compartida públicamente: basta con compartirla
    como Lector con el email de la cuenta de servicio.

    credenciales_sa: dict con el JSON de la cuenta de servicio.
    worksheet_name: nombre de la pestaña a leer.
    header_row: fila donde están los encabezados reales (1 si es la primera
    fila; 2 si, como Base_Maestra o Pipeline_Backup, la fila 1 es metadata).
    sheet_url: enlace del Google Sheet (por defecto PIPELINE_SHEET_URL).
    """
    client = _cliente_google_sheets(credenciales_sa)
    doc = client.open_by_url(sheet_url or PIPELINE_SHEET_URL)
    hoja = doc.worksheet(worksheet_name)
    return pd.DataFrame(hoja.get_all_records(head=header_row))


def cargar_pipeline_desde_google_sheets(credenciales_sa, sheet_url=None, worksheet_name=None):
    """Descarga el Pipeline directamente desde Google Sheets (ver
    cargar_hoja_desde_google_sheets)."""
    return cargar_hoja_desde_google_sheets(
        credenciales_sa,
        worksheet_name or PIPELINE_WORKSHEET_NAME,
        header_row=PIPELINE_HEADER_ROW,
        sheet_url=sheet_url,
    )


def filtrar_pipeline(df_pipeline, corredoras=None, aseguradoras=None, asegurados=None, ajustadores=None):
    """Filtra el Pipeline por Corredora / Compañía de seguros / Asegurado / Ajustador senior.

    Cada parámetro es una lista opcional de valores exactos (tal como aparecen
    en el archivo). Si una lista viene vacía o None, ese filtro no se aplica.
    Cuando se combinan varios filtros, se aplican con AND (intersección).
    """
    df = df_pipeline.copy()
    if corredoras and "Corredora" in df.columns:
        df = df[df["Corredora"].isin(corredoras)]
    if aseguradoras and "Compañía de seguros" in df.columns:
        df = df[df["Compañía de seguros"].isin(aseguradoras)]
    if asegurados and "Asegurado" in df.columns:
        df = df[df["Asegurado"].isin(asegurados)]
    if ajustadores and "Ajustador senior" in df.columns:
        df = df[df["Ajustador senior"].isin(ajustadores)]
    return df


def sugerir_titulo_cartera(corredoras=None, aseguradoras=None, asegurados=None):
    """Sugiere un título de portada a partir de los filtros activos (editable por el usuario)."""
    if asegurados:
        return " · ".join(asegurados)
    partes = []
    if corredoras:
        partes.append(" · ".join(corredoras))
    if aseguradoras:
        partes.append(" · ".join(aseguradoras))
    return " · ".join(partes) if partes else "Cartera General"


def sugerir_titulo_ajustador(ajustadores=None):
    """Sugiere un título de portada a partir de los Ajustadores senior seleccionados."""
    return " · ".join(ajustadores) if ajustadores else "Cartera General"


def limpiar_observacion(texto):
    return _TIMESTAMP_PREFIJO_RE.sub("", str(texto)).strip()


def preparar_tabla_casos(df_pipeline_filtrado, fecha_corte):
    """Construye la tabla base (una fila por caso) a partir del Pipeline.
    Probabilidad de cierre y Observación ya vienen cargadas por el equipo
    (columnas 'Indicación Probabilidad' / 'Probabilidad cierre 2026' y
    'Observaciones'); el usuario solo las revisa/ajusta antes de generar el pptx.
    """
    df = df_pipeline_filtrado.copy()

    for col, default in [
        ("Divisa", "UF"),
        ("Perdida bruta (en moneda del caso)", 0),
        ("Asegurado", "Sin asegurado"),
        ("Nickname", ""),
        ("Número de siniestro", ""),
        ("Observaciones", ""),
        ("Contenido último movimiento", ""),
        ("División", "Sin División"),
        ("Corredora", "Sin corredora"),
        ("Compañía de seguros", "Sin aseguradora"),
        ("Honorarios (UF)", 0),
    ]:
        if col not in df.columns:
            df[col] = default
        df[col] = df[col].fillna(default)

    divisa = df["Divisa"].astype(str).str.strip().str.upper()
    divisa = divisa.replace({"US$": "USD", "US $": "USD", "USD$": "USD", "U$": "USD"})

    perdida = pd.to_numeric(df["Perdida bruta (en moneda del caso)"], errors="coerce").fillna(0)
    fecha_ingreso = pd.to_datetime(df["Creado en"], errors="coerce")
    dias = (pd.Timestamp(fecha_corte) - fecha_ingreso).dt.days.clip(lower=0).fillna(0).astype(int)
    mcl = ((divisa == "UF") & (perdida > MCL_UF)) | ((divisa == "USD") & (perdida > MCL_USD))

    nickname = df["Nickname"].astype(str).str.strip()
    asegurado = df["Asegurado"].astype(str).str.strip()
    nickname_default = nickname.where(nickname.ne("") & nickname.ne("None") & nickname.ne("nan"), asegurado)

    # Puede llegar como número (float, con ".0" al final) si Excel/Sheets lo
    # detectó como columna numérica; se normaliza siempre a texto limpio.
    siniestro = df["Número de siniestro"].astype(str).str.strip()
    siniestro = siniestro.str.replace(r"\.0$", "", regex=True)
    siniestro = siniestro.where(~siniestro.str.lower().isin(["nan", "none"]), "")

    indicacion = df.get("Indicación Probabilidad", pd.Series("", index=df.index)).astype(str).str.strip().str.lower()
    prob_indicacion = indicacion.map(INDICACION_A_PROB)
    prob_numerica = pd.to_numeric(df.get("Probabilidad cierre 2026", pd.Series(1.0, index=df.index)), errors="coerce") * 100
    prob = prob_indicacion.fillna(prob_numerica).fillna(100).round().astype(int)

    observacion = df["Observaciones"].astype(str).apply(limpiar_observacion)
    fallback = df["Contenido último movimiento"].astype(str).apply(limpiar_observacion)
    observacion_final = observacion.where(observacion.ne("") & observacion.ne("nan"), fallback)

    tabla = pd.DataFrame({
        "Caso": df["Número de caso"].astype(str),
        "Numero_Siniestro": siniestro,
        "Asegurado": asegurado,
        "Nickname": nickname_default,
        "Divisa": divisa,
        "Perdida_bruta": perdida,
        "Dias": dias,
        "Division": df["División"].astype(str),
        "MCL": mcl,
        "Observacion_sugerida": observacion_final,
        "Prob": prob,
        "Observacion": observacion_final,
        "Corredora": df["Corredora"].astype(str).str.strip(),
        "Aseguradora": df["Compañía de seguros"].astype(str).str.strip(),
        "Honorarios": pd.to_numeric(df["Honorarios (UF)"], errors="coerce").fillna(0),
    })
    return tabla.sort_values("Dias", ascending=False).reset_index(drop=True)


def calcular_kpis(tabla):
    """Calcula todos los KPIs objetivos del reporte a partir de la tabla
    (ya editada por el usuario, con Prob/Observacion definitivos)."""
    total = len(tabla)
    uf_mask = tabla["Divisa"] == "UF"
    usd_mask = tabla["Divisa"] == "USD"

    uf_count = int(uf_mask.sum())
    usd_count = int(usd_mask.sum())
    uf_total = float(tabla.loc[uf_mask, "Perdida_bruta"].sum())
    usd_total = float(tabla.loc[usd_mask, "Perdida_bruta"].sum())

    div_ing = int(tabla["Division"].str.contains("Ingenier", case=False, na=False).sum())
    div_otros = total - div_ing

    mcl = tabla[tabla["MCL"]]
    otros = tabla[~tabla["MCL"]]
    mcl_count = len(mcl)
    otros_count = len(otros)
    mcl_pct = round(100 * mcl_count / total) if total else 0
    otros_pct = 100 - mcl_pct if total else 0

    mcl_uf = float(mcl.loc[mcl["Divisa"] == "UF", "Perdida_bruta"].sum())
    mcl_usd = float(mcl.loc[mcl["Divisa"] == "USD", "Perdida_bruta"].sum())
    otros_uf = float(otros.loc[otros["Divisa"] == "UF", "Perdida_bruta"].sum())
    otros_usd = float(otros.loc[otros["Divisa"] == "USD", "Perdida_bruta"].sum())

    top_aseg = tabla.groupby("Asegurado").size().sort_values(ascending=False)
    top4 = list(top_aseg.head(4).items())
    while len(top4) < 4:
        top4.append(("—", 0))
    aseg_otros = int(top_aseg.iloc[4:].sum())

    # Top 5 por Asegurado / Corredor / Aseguradora (solo se usan en el reporte
    # de Cartera General, en una slide independiente con 3 gráficos nativos).
    top5_asegurado = list(top_aseg.head(5).items())
    top5_corredor = list(tabla.groupby("Corredora").size().sort_values(ascending=False).head(5).items())
    top5_aseguradora = list(tabla.groupby("Aseguradora").size().sort_values(ascending=False).head(5).items())

    dias_max = int(tabla["Dias"].max()) if total else 0
    dias_prom = round(tabla["Dias"].mean()) if total else 0
    dias_600 = int((tabla["Dias"] > 600).sum())

    prob = pd.to_numeric(tabla["Prob"], errors="coerce").fillna(100)
    tier_counts = {
        100: int((prob >= 90).sum()),
        75: int(((prob >= 60) & (prob < 90)).sum()),
        50: int(((prob >= 25) & (prob < 60)).sum()),
        0: int((prob < 25).sum()),
    }

    # Mismos tramos de probabilidad 2026, pero acotados a MCL / No MCL por
    # separado (para las slides "Casos MCL/No MCL por Probabilidad de Cierre
    # 2026", solo Cartera General).
    def _tiers_de(subset):
        prob_subset = pd.to_numeric(subset["Prob"], errors="coerce").fillna(100)
        masks = {
            100: prob_subset >= 90,
            75: (prob_subset >= 60) & (prob_subset < 90),
            50: (prob_subset >= 25) & (prob_subset < 60),
            0: prob_subset < 25,
        }
        counts = {tier: int(mask.sum()) for tier, mask in masks.items()}
        honorarios = {tier: float(subset.loc[mask, "Honorarios"].sum()) for tier, mask in masks.items()}
        return counts, honorarios

    mcl_tier_counts, mcl_tier_honorarios = _tiers_de(mcl)
    otros_tier_counts, otros_tier_honorarios = _tiers_de(otros)

    return dict(
        total=total, uf_count=uf_count, usd_count=usd_count,
        uf_total=uf_total, usd_total=usd_total,
        div_ing=div_ing, div_otros=div_otros,
        mcl_count=mcl_count, otros_count=otros_count,
        mcl_pct=mcl_pct, otros_pct=otros_pct,
        mcl_uf=mcl_uf, mcl_usd=mcl_usd, otros_uf=otros_uf, otros_usd=otros_usd,
        top4=top4, aseg_otros=aseg_otros,
        top5_asegurado=top5_asegurado, top5_corredor=top5_corredor, top5_aseguradora=top5_aseguradora,
        dias_max=dias_max, dias_prom=dias_prom, dias_600=dias_600,
        tier_counts=tier_counts,
        mcl_tier_counts=mcl_tier_counts, mcl_tier_honorarios=mcl_tier_honorarios,
        otros_tier_counts=otros_tier_counts, otros_tier_honorarios=otros_tier_honorarios,
    )


# ---------------------------------------------------------------------------
# Motor de slides (branding.py): header + cuerpo + footer en cada una.
# ---------------------------------------------------------------------------

CONTENT_LEFT = br.MARGEN_LATERAL
CONTENT_TOP = br.CONTENIDO_TOP + Emu(228600)
CONTENT_WIDTH = br.CONTENIDO_ANCHO
CONTENT_BOTTOM = br.CONTENIDO_BOTTOM - Emu(91440)
CONTENT_HEIGHT = CONTENT_BOTTOM - CONTENT_TOP

_GAP = Emu(182880)
_KPI_ALTO = Emu(1097280)


def _set_text(text_frame, texto):
    p = text_frame.paragraphs[0]
    if p.runs:
        p.runs[0].text = texto
        for r in p.runs[1:]:
            r.text = ""
    else:
        p.text = texto


def _slide_base(prs, titulo, etiqueta=None, numero=None, fondo_claro=True):
    slide = br.nueva_slide(prs)
    br.agregar_header(slide, titulo, etiqueta=etiqueta, fondo_claro=fondo_claro)
    br.agregar_footer(slide, numero=numero)
    return slide


def _tarjeta_kpi(slide, left, top, width, height, valor, etiqueta, oscuro=False):
    fondo_color = br.NAVY_OSCURO if oscuro else br.BLANCO
    color_valor = br.BLANCO if oscuro else br.NAVY
    color_etiqueta = br.TEAL_PALIDO if oscuro else br.GRIS_SLATE
    br.rect(slide, left, top, width, height, fondo_color)
    br.rect(slide, left, top, Emu(54864), height, br.TEAL)
    pad = Emu(137160)
    # Los valores largos (p.ej. "3.408.000 UF + USD 2,6 M") se reducen de
    # tamaño para no invadir la etiqueta de abajo.
    tam_valor = 22 if len(valor) <= 14 else (18 if len(valor) <= 22 else 14)
    br.caja_texto(slide, left + pad, top + Emu(91440), width - 2 * pad, Emu(640080),
                  valor, color_valor, tam_valor, bold=True, fuente=br.FUENTE_TITULOS)
    br.caja_texto(slide, left + pad, top + height - Emu(320040), width - 2 * pad, Emu(274320),
                  etiqueta, color_etiqueta, 10.5, fuente=br.FUENTE_TEXTO)


def _fila_tarjetas(slide, top, height, tarjetas, oscuro=False):
    """tarjetas: lista de (valor, etiqueta). Reparte el ancho del contenido
    en partes iguales, con un espacio fijo entre tarjetas."""
    n = len(tarjetas)
    ancho = (CONTENT_WIDTH - (n - 1) * _GAP) // n
    for i, (valor, etiqueta) in enumerate(tarjetas):
        left = CONTENT_LEFT + i * (ancho + _GAP)
        _tarjeta_kpi(slide, left, top, ancho, height, valor, etiqueta, oscuro=oscuro)


def _agregar_grafico_barras(slide, left, top, width, height, titulo, datos, horizontal=True):
    """Agrega un gráfico de barras nativo (no una imagen ni barras dibujadas a
    mano): así el ancho/alto de cada barra siempre es fiel al valor real, sin
    importar cuántas categorías traiga `datos` (lista de tuplas
    (categoria, valor), ya ordenada de mayor a menor)."""
    categorias = [str(c) if c not in (None, "") else "Sin dato" for c, _ in datos]
    valores = [v for _, v in datos]
    chart_data = CategoryChartData()
    if horizontal:
        categorias = [_truncar_texto(c, 26) for c in categorias]
        # Un gráfico de barras horizontales dibuja la primera categoría abajo:
        # se invierte el orden para que el puesto #1 quede arriba.
        chart_data.categories = list(reversed(categorias))
        chart_data.add_series("Serie", list(reversed(valores)))
        tipo = XL_CHART_TYPE.BAR_CLUSTERED
    else:
        chart_data.categories = categorias
        chart_data.add_series("Serie", valores)
        tipo = XL_CHART_TYPE.COLUMN_CLUSTERED

    chart = slide.shapes.add_chart(tipo, left, top, width, height, chart_data).chart
    chart.has_legend = False
    chart.has_title = True
    chart.chart_title.text_frame.text = titulo
    run_titulo = chart.chart_title.text_frame.paragraphs[0].runs[0]
    run_titulo.font.size, run_titulo.font.bold = Pt(12), True
    run_titulo.font.color.rgb = br.NAVY
    run_titulo.font.name = br.FUENTE_TEXTO

    plot = chart.plots[0]
    plot.has_data_labels = True
    plot.data_labels.font.size = Pt(9)
    plot.data_labels.number_format = "#,##0"
    plot.data_labels.number_format_is_linked = False
    chart.series[0].format.fill.solid()
    chart.series[0].format.fill.fore_color.rgb = br.TEAL

    chart.category_axis.tick_labels.font.size = Pt(9)
    chart.value_axis.visible = False
    chart.value_axis.has_major_gridlines = False
    return chart


def _tabla(slide, top, height, columnas, filas, filas_max, fuente_header=9.5, fuente_dato=9):
    """columnas: lista de (encabezado, ancho_emu_o_None, max_chars_o_None).
    Un ancho None reparte el espacio restante en partes iguales entre las
    columnas flexibles. filas: lista de listas de valores (ya en el orden de
    columnas); se truncan solo las columnas con max_chars."""
    n_cols = len(columnas)
    n_filas = filas_max + 1
    anchos_fijos = [a for _, a, _ in columnas if a is not None]
    n_flex = n_cols - len(anchos_fijos)
    ancho_flex = (CONTENT_WIDTH - sum(anchos_fijos)) // n_flex if n_flex else 0
    anchos = [a if a is not None else ancho_flex for _, a, _ in columnas]

    graphic_frame = slide.shapes.add_table(n_filas, n_cols, CONTENT_LEFT, top, CONTENT_WIDTH, height)
    tabla_pptx = graphic_frame.table
    tabla_pptx.horz_banding = False
    for col, ancho in zip(tabla_pptx.columns, anchos):
        col.width = ancho

    alto_header = Emu(320040)
    alto_dato = Emu(max(1, int((height - alto_header) / filas_max)))
    tabla_pptx.rows[0].height = alto_header
    for fila in list(tabla_pptx.rows)[1:]:
        fila.height = alto_dato

    for c, (encabezado, _, _) in enumerate(columnas):
        cell = tabla_pptx.cell(0, c)
        _set_text(cell.text_frame, encabezado)
        cell.fill.solid()
        cell.fill.fore_color.rgb = br.NAVY
        cell.margin_left, cell.margin_right = Pt(4), Pt(4)
        cell.margin_top, cell.margin_bottom = Pt(2), Pt(2)
        run = cell.text_frame.paragraphs[0].runs[0]
        run.font.size, run.font.bold = Pt(fuente_header), True
        run.font.color.rgb = br.BLANCO
        run.font.name = br.FUENTE_TEXTO

    for i in range(filas_max):
        for c, (_, _, max_chars) in enumerate(columnas):
            if i < len(filas):
                valor = filas[i][c]
                if max_chars:
                    valor = _truncar_texto(valor, max_chars)
            else:
                valor = ""
            cell = tabla_pptx.cell(i + 1, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = br.GRIS_CLARO if i % 2 == 0 else br.BLANCO
            cell.margin_left, cell.margin_right = Pt(4), Pt(4)
            cell.margin_top, cell.margin_bottom = Pt(1), Pt(1)
            cell.text_frame.word_wrap = False
            _set_text(cell.text_frame, str(valor))
            if cell.text_frame.paragraphs[0].runs:
                run = cell.text_frame.paragraphs[0].runs[0]
                run.font.size = Pt(fuente_dato)
                run.font.color.rgb = br.NAVY
                run.font.name = br.FUENTE_TEXTO
    return tabla_pptx


# --- Columnas estándar reutilizadas por varias slides ---
_COLS_CASOS = [
    ("#", Emu(457200), 3),
    ("Caso", Emu(822960), 12),
    ("Siniestro / Nickname", None, 42),
    ("Divisa", Emu(548640), 6),
    ("Días", Emu(548640), 6),
    ("Prob.", Emu(548640), 6),
    ("Observación", Emu(5516880), 82),
]
_COLS_ATENCION = [
    ("Caso", Emu(822960), 12),
    ("Siniestro / Nickname", None, 46),
    ("Prob.", Emu(548640), 6),
    ("Divisa", Emu(548640), 6),
    ("MCL", Emu(457200), 3),
    ("Observación", Emu(6065520), 90),
]
_COLS_DETALLE = [
    ("#", Emu(457200), 3),
    ("Caso", Emu(822960), 12),
    ("Siniestro", Emu(1097280), 18),
    ("Nickname / Asegurado", None, 110),
    ("Divisa", Emu(548640), 6),
    ("Días", Emu(548640), 6),
    ("Prob.", Emu(548640), 6),
]


def _slide_portada(prs, titulo_cartera, fecha_corte, kpis):
    slide = br.nueva_slide(prs)
    br.fondo(slide, br.NAVY)
    br.agregar_portada_header(slide)
    br.agregar_footer(slide)

    y = Emu(329184) + br.HEADER_ALTO + Emu(365760)
    br.caja_texto(slide, br.MARGEN_LATERAL, y, br.CONTENIDO_ANCHO, Emu(274320),
                  "ESTADO DE CARTERA", br.TEAL_CLARO, 13, bold=True)
    y += Emu(365760)
    br.caja_texto(slide, br.MARGEN_LATERAL, y, br.CONTENIDO_ANCHO, Emu(822960),
                  titulo_cartera, br.BLANCO, 34, bold=True, fuente=br.FUENTE_TITULOS)
    y += Emu(914400)
    br.caja_texto(slide, br.MARGEN_LATERAL, y, br.CONTENIDO_ANCHO, Emu(274320),
                  f"Análisis al {fmt_fecha_larga(fecha_corte)}", br.TEAL_PALIDO, 13)

    y_tarjetas = Emu(4389120)
    _fila_tarjetas(slide, y_tarjetas, Emu(1189992), [
        (str(kpis["total"]), "CASOS ACTIVOS"),
        (fmt_uf(kpis["uf_total"]), "PÉRDIDA BRUTA UF"),
        (fmt_usd_m(kpis["usd_total"]), "PÉRDIDA BRUTA USD"),
    ], oscuro=True)


def _slide_resumen(prs, titulo_cartera, kpis, numero):
    slide = _slide_base(prs, "RESUMEN GENERAL DE CARTERA", etiqueta=titulo_cartera, numero=numero)
    _fila_tarjetas(slide, CONTENT_TOP, _KPI_ALTO, [
        (str(kpis["total"]), "CARTERA ACTIVA (CASOS)"),
        (f"{kpis['div_ing']} / {kpis['div_otros']}", "INGENIERÍA / OTRAS DIVISIONES"),
        (f"{kpis['uf_count']} UF · {kpis['usd_count']} USD", "CASOS POR DIVISA"),
    ])
    _fila_tarjetas(slide, CONTENT_TOP + _KPI_ALTO + _GAP, _KPI_ALTO, [
        (fmt_uf(kpis["uf_total"]), "EXPOSICIÓN TOTAL UF"),
        (fmt_usd_m(kpis["usd_total"]), "EXPOSICIÓN TOTAL USD"),
        (f"{kpis['mcl_pct']}% · {kpis['mcl_count']} casos", "MAYOR CUANTÍA (MCL)"),
    ])


def _slide_distribucion(prs, kpis, numero):
    slide = _slide_base(prs, "DISTRIBUCIÓN DE CARTERA", etiqueta="MCL vs. Otros · Top 4 Asegurados", numero=numero)
    _fila_tarjetas(slide, CONTENT_TOP, _KPI_ALTO, [
        (f"{kpis['mcl_count']} casos · {kpis['mcl_pct']}%",
         f"MAYOR CUANTÍA  —  {fmt_uf(kpis['mcl_uf'])} + {fmt_usd_m(kpis['mcl_usd'])}"),
        (f"{kpis['otros_count']} casos · {kpis['otros_pct']}%",
         f"OTROS CASOS  —  {fmt_uf(kpis['otros_uf'])} + {fmt_usd_m(kpis['otros_usd'])}"),
    ])
    datos = list(kpis["top4"]) + [("Otros asegurados", kpis["aseg_otros"])]
    top_chart = CONTENT_TOP + _KPI_ALTO + _GAP
    _agregar_grafico_barras(slide, CONTENT_LEFT, top_chart, CONTENT_WIDTH, CONTENT_BOTTOM - top_chart,
                             "Concentración de cartera por Asegurado", datos, horizontal=True)


def _slide_probabilidad(prs, kpis, numero):
    slide = _slide_base(prs, "PROBABILIDAD DE CIERRE 2026", numero=numero)
    tc = kpis["tier_counts"]
    categorias = [
        f"100% – Cierta ({tc[100]})", f"75% – Alt. probable ({tc[75]})",
        f"50% – Podría ser ({tc[50]})", f"0% – Nula ({tc[0]})",
    ]
    datos = list(zip(categorias, [tc[100], tc[75], tc[50], tc[0]]))
    _agregar_grafico_barras(slide, CONTENT_LEFT, CONTENT_TOP, CONTENT_WIDTH, CONTENT_HEIGHT,
                             "Casos por probabilidad de cierre 2026", datos, horizontal=True)


def _slide_top5(prs, kpis, numero):
    slide = _slide_base(prs, "TOP 5 — CONCENTRACIÓN DE CARTERA", etiqueta="Cartera General", numero=numero)
    ancho_col = (CONTENT_WIDTH - 2 * _GAP) // 3
    lefts = [CONTENT_LEFT, CONTENT_LEFT + ancho_col + _GAP, CONTENT_LEFT + 2 * (ancho_col + _GAP)]
    graficos = [
        ("Top 5 Asegurados", kpis["top5_asegurado"]),
        ("Top 5 Corredores", kpis["top5_corredor"]),
        ("Top 5 Aseguradoras", kpis["top5_aseguradora"]),
    ]
    for left, (titulo, datos) in zip(lefts, graficos):
        _agregar_grafico_barras(slide, left, CONTENT_TOP, ancho_col, CONTENT_HEIGHT, titulo, datos, horizontal=True)


def _slide_probabilidad_por_tier(prs, titulo, etiqueta, tier_counts, tier_honorarios, numero):
    slide = _slide_base(prs, titulo, etiqueta="Cartera General", numero=numero)
    categorias = ["100% – Cierta", "75% – Alt. probable", "50% – Podría ser", "0% – Nula"]
    tc, th = tier_counts, tier_honorarios
    datos_cantidad = list(zip(categorias, [tc[100], tc[75], tc[50], tc[0]]))
    datos_honorarios = list(zip(categorias, [round(th[100]), round(th[75]), round(th[50]), round(th[0])]))

    ancho_col = (CONTENT_WIDTH - _GAP) // 2
    _agregar_grafico_barras(slide, CONTENT_LEFT, CONTENT_TOP, ancho_col, CONTENT_HEIGHT,
                             f"Cantidad de casos {etiqueta}", datos_cantidad, horizontal=False)
    _agregar_grafico_barras(slide, CONTENT_LEFT + ancho_col + _GAP, CONTENT_TOP, ancho_col, CONTENT_HEIGHT,
                             f"Honorarios totales {etiqueta} (UF)", datos_honorarios, horizontal=False)


def _slide_tiempos_residencia(prs, titulo_cartera, kpis, tabla, numero):
    etiqueta_promedio = "Promedio cartera" if titulo_cartera == "Cartera General" else f"Promedio cartera {titulo_cartera}"
    slide = _slide_base(prs, "TIEMPOS DE RESIDENCIA", etiqueta="Top 10 más antiguos", numero=numero)
    _fila_tarjetas(slide, CONTENT_TOP, _KPI_ALTO, [
        (f"{kpis['dias_max']} días", "CASO MÁS ANTIGUO"),
        (f"~{kpis['dias_prom']} días", etiqueta_promedio.upper()),
        (f"{kpis['dias_600']} caso" + ("s" if kpis["dias_600"] != 1 else ""), "CON MÁS DE 600 DÍAS"),
    ])

    top10_antiguos = tabla.sort_values("Dias", ascending=False).head(10)
    filas = []
    for i, (_, row) in enumerate(top10_antiguos.iterrows(), start=1):
        observacion = row["Observacion"] or row["Observacion_sugerida"]
        filas.append([
            i, row["Caso"], _fila_siniestro_nickname(row["Numero_Siniestro"], row["Nickname"]), row["Divisa"],
            row["Dias"], f"{int(row['Prob'])}%", observacion,
        ])
    top_tabla = CONTENT_TOP + _KPI_ALTO + _GAP
    _tabla(slide, top_tabla, CONTENT_BOTTOM - top_tabla, _COLS_CASOS, filas, filas_max=10)


def _slide_atencion(prs, kpis, tabla, alerta_prioritaria, numero):
    slide = _slide_base(prs, "⚠  CASOS QUE REQUIEREN ATENCIÓN ESPECIAL", numero=numero)
    atencion = tabla[pd.to_numeric(tabla["Prob"], errors="coerce").fillna(100) < 100]
    atencion = atencion.sort_values("Perdida_bruta", ascending=False).head(10)

    intro = (
        f"{len(atencion)} casos con probabilidad de cierre inferior al 100% acumulan exposición "
        f"significativa. Ordenados por magnitud de pérdida bruta."
    )
    br.caja_texto(slide, CONTENT_LEFT, CONTENT_TOP, CONTENT_WIDTH, Emu(411480), intro, br.GRIS_SLATE, 11)

    top_alerta = CONTENT_TOP + Emu(457200)
    alto_alerta = Emu(685800)
    br.rect(slide, CONTENT_LEFT, top_alerta, CONTENT_WIDTH, alto_alerta, br.GRIS_CLARO)
    br.rect(slide, CONTENT_LEFT, top_alerta, Emu(54864), alto_alerta, br.TEAL_OSCURO)
    br.caja_texto(slide, CONTENT_LEFT + Emu(182880), top_alerta + Emu(45720), CONTENT_WIDTH - Emu(365760), Emu(594360),
                  alerta_prioritaria or "Sin alertas prioritarias registradas.", br.NAVY, 11, italic=True)

    filas = []
    for _, row in atencion.iterrows():
        observacion = row["Observacion"] or row["Observacion_sugerida"]
        filas.append([
            row["Caso"], _fila_siniestro_nickname(row["Numero_Siniestro"], row["Nickname"]), f"{int(row['Prob'])}%",
            row["Divisa"], "✓" if row["MCL"] else "", observacion,
        ])
    top_tabla = top_alerta + alto_alerta + _GAP
    _tabla(slide, top_tabla, CONTENT_BOTTOM - top_tabla, _COLS_ATENCION, filas, filas_max=10)


def _slide_gestiones_iniciales(prs, tabla, numero):
    slide = _slide_base(prs, "GESTIONES INICIALES", etiqueta="< 2 semanas de antigüedad", numero=numero)

    recientes = tabla[tabla["Dias"] < 14].sort_values("Dias", ascending=False)
    sin_gestion = recientes[
        (recientes["Observacion"].astype(str).str.strip() == "")
        & (recientes["Observacion_sugerida"].astype(str).str.strip() == "")
    ]
    dias_prom_recientes = round(recientes["Dias"].mean()) if len(recientes) else 0

    _fila_tarjetas(slide, CONTENT_TOP, _KPI_ALTO, [
        (f"{len(recientes)} caso" + ("s" if len(recientes) != 1 else ""), "INGRESADOS < 2 SEMANAS"),
        (f"{len(sin_gestion)} caso" + ("s" if len(sin_gestion) != 1 else ""), "SIN GESTIÓN INICIAL REGISTRADA"),
        (f"~{dias_prom_recientes} días", "ANTIGÜEDAD PROMEDIO DEL GRUPO"),
    ])

    filas = []
    for i, (_, row) in enumerate(recientes.head(10).iterrows(), start=1):
        observacion = row["Observacion"] or row["Observacion_sugerida"] or "Sin gestión registrada"
        filas.append([
            i, row["Caso"], _fila_siniestro_nickname(row["Numero_Siniestro"], row["Nickname"]), row["Divisa"],
            row["Dias"], f"{int(row['Prob'])}%", observacion,
        ])
    top_tabla = CONTENT_TOP + _KPI_ALTO + _GAP
    _tabla(slide, top_tabla, CONTENT_BOTTOM - top_tabla, _COLS_CASOS, filas, filas_max=10)


def _slide_detalle_completo(prs, pagina, num_pagina, total_paginas, total_casos, kpis, numero):
    slide = _slide_base(
        prs, "DETALLE COMPLETO DE CARTERA",
        etiqueta=f"Página {num_pagina} de {total_paginas}  ·  {total_casos} casos", numero=numero,
    )
    _fila_tarjetas(slide, CONTENT_TOP, _KPI_ALTO, [
        (f"{total_casos} caso" + ("s" if total_casos != 1 else ""), "TOTAL CARTERA FILTRADA"),
        (f"{num_pagina} / {total_paginas}", "PÁGINA"),
        (f"{fmt_uf(kpis['uf_total'])}  +  {fmt_usd_m(kpis['usd_total'])}", "EXPOSICIÓN TOTAL (UF + USD)"),
    ])

    filas = []
    inicio = (num_pagina - 1) * 20 + 1
    for i, (_, row) in enumerate(pagina.iterrows(), start=inicio):
        filas.append([
            i, row["Caso"], row["Numero_Siniestro"], row["Nickname"], row["Divisa"], row["Dias"], f"{int(row['Prob'])}%",
        ])
    top_tabla = CONTENT_TOP + _KPI_ALTO + _GAP
    _tabla(slide, top_tabla, CONTENT_BOTTOM - top_tabla, _COLS_DETALLE, filas, filas_max=20,
           fuente_header=8.5, fuente_dato=8)


def _slide_proximos_pasos(prs, pasos, fecha_corte, titulo_cartera, numero):
    slide = _slide_base(prs, "PRÓXIMOS PASOS Y FOCOS DE GESTIÓN", numero=numero)
    pasos = (pasos + [{}] * 5)[:5]
    alto_fila = (CONTENT_HEIGHT - 2 * _GAP) // 3
    ancho_col = (CONTENT_WIDTH - _GAP) // 2
    posiciones = [
        (CONTENT_LEFT, CONTENT_TOP),
        (CONTENT_LEFT + ancho_col + _GAP, CONTENT_TOP),
        (CONTENT_LEFT, CONTENT_TOP + alto_fila + _GAP),
        (CONTENT_LEFT + ancho_col + _GAP, CONTENT_TOP + alto_fila + _GAP),
        (CONTENT_LEFT, CONTENT_TOP + 2 * (alto_fila + _GAP)),
    ]
    for (left, top), paso in zip(posiciones, pasos):
        titulo = paso.get("titulo", "") if paso else ""
        desc = paso.get("desc", "") if paso else ""
        if not (titulo or desc):
            continue
        br.rect(slide, left, top, ancho_col, alto_fila, br.BLANCO)
        br.rect(slide, left, top, ancho_col, Emu(54864), br.TEAL)
        pad = Emu(182880)
        br.caja_texto(slide, left + pad, top + Emu(109728), ancho_col - 2 * pad, Emu(365760),
                      titulo, br.NAVY, 13, bold=True, fuente=br.FUENTE_TITULOS)
        br.caja_texto(slide, left + pad, top + Emu(475488), ancho_col - 2 * pad, alto_fila - Emu(594360),
                      desc, br.GRIS_SLATE, 10.5)

    pie = (
        f"JPV Asociados · Sedgwick Alliance  ·  Análisis al {fmt_fecha_larga(fecha_corte)}  ·  "
        f"{titulo_cartera if titulo_cartera.lower().startswith('cartera') else f'Cartera {titulo_cartera}'}"
    )
    br.caja_texto(slide, CONTENT_LEFT, CONTENT_TOP + 3 * alto_fila + 2 * _GAP - Emu(228600),
                  CONTENT_WIDTH, Emu(228600), pie, br.GRIS_SLATE, 9, italic=True)


def generar_pptx(fecha_corte, titulo_cartera, tabla, pasos, alerta_prioritaria):
    """
    fecha_corte: datetime.date/datetime
    titulo_cartera: texto para la portada/resumen/pie (ej. 'MARSH S.A.', 'Essbio S.A.',
                    'Aon Risk Services · Essbio S.A.', 'Cartera General', etc.)
    tabla: DataFrame de preparar_tabla_casos() ya editado por el usuario
           (columnas Caso, Nickname, Divisa, Perdida_bruta, Dias, Division, MCL, Prob, Observacion)
    pasos: lista de hasta 5 dicts {'titulo':..., 'desc':...}
    alerta_prioritaria: texto libre para el callout de "Atención Especial"
    """
    kpis = calcular_kpis(tabla)
    prs = br.nueva_presentacion()
    n = 1

    _slide_portada(prs, titulo_cartera, fecha_corte, kpis)
    n += 1
    _slide_resumen(prs, titulo_cartera, kpis, n); n += 1
    _slide_distribucion(prs, kpis, n); n += 1
    _slide_probabilidad(prs, kpis, n); n += 1

    es_cartera_general = titulo_cartera == "Cartera General"
    if es_cartera_general:
        _slide_top5(prs, kpis, n); n += 1
        _slide_probabilidad_por_tier(
            prs, "CASOS MCL POR PROBABILIDAD DE CIERRE 2026", "MCL",
            kpis["mcl_tier_counts"], kpis["mcl_tier_honorarios"], n,
        )
        n += 1
        _slide_probabilidad_por_tier(
            prs, "CASOS NO MCL POR PROBABILIDAD DE CIERRE 2026", "No MCL",
            kpis["otros_tier_counts"], kpis["otros_tier_honorarios"], n,
        )
        n += 1

    _slide_tiempos_residencia(prs, titulo_cartera, kpis, tabla, n); n += 1
    _slide_atencion(prs, kpis, tabla, alerta_prioritaria, n); n += 1
    _slide_gestiones_iniciales(prs, tabla, n); n += 1

    detalle = tabla.sort_values("Dias", ascending=False).reset_index(drop=True)
    filas_max_detalle = 20
    paginas_detalle = [
        detalle.iloc[i:i + filas_max_detalle] for i in range(0, len(detalle), filas_max_detalle)
    ]
    for num_pagina, pagina in enumerate(paginas_detalle, start=1):
        _slide_detalle_completo(prs, pagina, num_pagina, len(paginas_detalle), len(detalle), kpis, n)
        n += 1

    _slide_proximos_pasos(prs, pasos, fecha_corte, titulo_cartera, n)

    output = io.BytesIO()
    prs.save(output)
    return output.getvalue()
