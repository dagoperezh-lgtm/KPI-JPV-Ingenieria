"""
PPTX de tendencia: evolución de Stock (Q y UF) y Asignaciones semanales por
área y total de la gerencia, a partir del histórico armado en
tablero_historico.py.

Reutiliza el logo y la paleta de colores (navy + teal) de assets/logo_jpv.png
y ficha_caso.py, para que quede consistente con el resto de los reportes
del ecosistema JPV (Reporte de Cartera, Ficha de Caso).
"""
import io
import os

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.util import Emu, Inches, Pt

LOGO_PATH = os.path.join(os.path.dirname(__file__), "assets", "logo_jpv.png")

NAVY_OSCURO = RGBColor(0x0D, 0x1F, 0x38)
NAVY = RGBColor(0x1B, 0x2A, 0x4A)
TEAL = RGBColor(0x14, 0xA8, 0xA0)
TEAL_OSCURO = RGBColor(0x0D, 0x73, 0x77)
GRIS_TEXTO = RGBColor(0x33, 0x33, 0x33)
GRIS_CLARO = RGBColor(0xF4, 0xF6, 0xF8)
BLANCO = RGBColor(0xFF, 0xFF, 0xFF)

ANCHO_SLIDE = Inches(13.333)
ALTO_SLIDE = Inches(7.5)
HEADER_ALTO = Inches(1.1)


def _agregar_header(slide, titulo, subtitulo):
    barra = slide.shapes.add_shape(1, 0, 0, ANCHO_SLIDE, HEADER_ALTO)
    barra.fill.solid()
    barra.fill.fore_color.rgb = NAVY
    barra.line.fill.background()
    barra.shadow.inherit = False

    divisor = slide.shapes.add_shape(1, 0, HEADER_ALTO, ANCHO_SLIDE, Emu(45720))
    divisor.fill.solid()
    divisor.fill.fore_color.rgb = TEAL_OSCURO
    divisor.line.fill.background()
    divisor.shadow.inherit = False

    if os.path.exists(LOGO_PATH):
        slide.shapes.add_picture(LOGO_PATH, Inches(0.35), Inches(0.2), height=Inches(0.7))

    caja_titulo = slide.shapes.add_textbox(Inches(2.2), Inches(0.18), ANCHO_SLIDE - Inches(2.6), Inches(0.85))
    tf = caja_titulo.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = titulo
    p.font.size, p.font.bold, p.font.color.rgb = Pt(26), True, BLANCO
    if subtitulo:
        p2 = tf.add_paragraph()
        p2.text = subtitulo
        p2.font.size, p2.font.italic, p2.font.color.rgb = Pt(13), True, TEAL


def _slide_base(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    fondo = slide.background.fill
    fondo.solid()
    fondo.fore_color.rgb = BLANCO
    return slide


def _agregar_titulo_chico(slide, texto, left, top, width):
    caja = slide.shapes.add_textbox(left, top, width, Inches(0.35))
    p = caja.text_frame.paragraphs[0]
    p.text = texto
    p.font.size, p.font.bold, p.font.color.rgb = Pt(14), True, NAVY


def _agregar_grafico(slide, tipo, categorias, series, num_format=None,
                      left=Inches(0.5), top=None, width=None, height=None,
                      leyenda=True, color_offset=0):
    if top is None:
        top = HEADER_ALTO + Inches(0.35)
    if width is None:
        width = ANCHO_SLIDE - Inches(1)
    if height is None:
        height = ALTO_SLIDE - HEADER_ALTO - Inches(0.85)

    cd = CategoryChartData()
    cd.categories = categorias
    for nombre, valores in series:
        cd.add_series(nombre, valores)

    grafico_frame = slide.shapes.add_chart(tipo, left, top, width, height, cd)
    chart = grafico_frame.chart
    chart.has_legend = leyenda
    if leyenda:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(12)
        chart.legend.font.color.rgb = GRIS_TEXTO

    try:
        eje_cat = chart.category_axis
        eje_cat.tick_labels.font.size = Pt(8)
        eje_cat.tick_labels.font.color.rgb = GRIS_TEXTO
    except Exception:
        pass
    try:
        eje_valor = chart.value_axis
        eje_valor.tick_labels.font.size = Pt(10)
        eje_valor.tick_labels.font.color.rgb = GRIS_TEXTO
        if num_format:
            eje_valor.tick_labels.number_format = num_format
            eje_valor.tick_labels.number_format_is_linked = False
    except Exception:
        pass

    colores = [TEAL, NAVY, TEAL_OSCURO]
    for i, serie in enumerate(chart.series):
        color = colores[(i + color_offset) % len(colores)]
        try:
            if tipo in (XL_CHART_TYPE.LINE, XL_CHART_TYPE.LINE_MARKERS):
                serie.format.line.color.rgb = color
                serie.format.line.width = Pt(2.5)
                serie.marker.format.fill.solid()
                serie.marker.format.fill.fore_color.rgb = color
            else:
                serie.format.fill.solid()
                serie.format.fill.fore_color.rgb = color
        except Exception:
            pass
    return chart


def _agregar_par_dividido(slide, tipo, etiquetas, serie_ing, serie_movil, num_format=None):
    """Dos gráficos lado a lado (Ingeniería / Equipo Móvil), cada uno con su
    propio eje: cuando la escala de una división es varias veces la de la
    otra, un solo eje combinado aplana la tendencia de la más chica."""
    mitad = Emu((ANCHO_SLIDE - Inches(1.2)) // 2)
    top_grafico = HEADER_ALTO + Inches(0.75)
    alto_grafico = ALTO_SLIDE - top_grafico - Inches(0.4)

    _agregar_titulo_chico(slide, "Ingeniería y Energía", Inches(0.5), HEADER_ALTO + Inches(0.3), mitad)
    _agregar_grafico(slide, tipo, etiquetas, [("Ingeniería y Energía", serie_ing)],
                      num_format=num_format, left=Inches(0.5), top=top_grafico, width=mitad, height=alto_grafico,
                      leyenda=False, color_offset=0)

    _agregar_titulo_chico(slide, "Equipo Móvil", Inches(0.7) + mitad, HEADER_ALTO + Inches(0.3), mitad)
    _agregar_grafico(slide, tipo, etiquetas, [("Equipo Móvil", serie_movil)],
                      num_format=num_format, left=Inches(0.7) + mitad, top=top_grafico, width=mitad, height=alto_grafico,
                      leyenda=False, color_offset=1)


def _agregar_par_dividido_multiserie(slide, tipo, etiquetas, series_ing, series_mov, num_format=None):
    """Dos gráficos lado a lado (Ingeniería / Equipo Móvil), cada uno con
    VARIAS series (ej. Asignaciones vs IFL dentro de la misma área) — a
    diferencia de _agregar_par_dividido, aquí sí hace falta leyenda porque
    cada gráfico tiene más de una serie que distinguir.
    series_ing / series_mov: [(nombre, valores), ...]."""
    mitad = Emu((ANCHO_SLIDE - Inches(1.2)) // 2)
    top_grafico = HEADER_ALTO + Inches(0.75)
    alto_grafico = ALTO_SLIDE - top_grafico - Inches(0.4)

    _agregar_titulo_chico(slide, "Ingeniería y Energía", Inches(0.5), HEADER_ALTO + Inches(0.3), mitad)
    _agregar_grafico(slide, tipo, etiquetas, series_ing, num_format=num_format,
                      left=Inches(0.5), top=top_grafico, width=mitad, height=alto_grafico)

    _agregar_titulo_chico(slide, "Equipo Móvil", Inches(0.7) + mitad, HEADER_ALTO + Inches(0.3), mitad)
    _agregar_grafico(slide, tipo, etiquetas, series_mov, num_format=num_format,
                      left=Inches(0.7) + mitad, top=top_grafico, width=mitad, height=alto_grafico)


MESES_ES = {1: "ene", 2: "feb", 3: "mar", 4: "abr", 5: "may", 6: "jun",
            7: "jul", 8: "ago", 9: "sep", 10: "oct", 11: "nov", 12: "dic"}


def _promedio_semanal_y_percapita(df_historico, df_promedio, columna, fechas):
    """(promedio semanal total, promedio semanal por ajustador vigente) de
    `columna` para 'Total Gerencia' en el rango `fechas`. El per-cápita se
    calcula semana a semana (valor de la semana / ajustadores vigentes esa
    semana, excluyendo a Dagoberto Pérez) y luego se promedia — no divide el
    promedio total por el promedio de ajustadores."""
    sub_valor = df_historico[df_historico["Division"] == "Total Gerencia"].set_index("Fecha")[columna]
    valores = [sub_valor.get(f) for f in fechas]
    valores_validos = [v for v in valores if v is not None]
    promedio_total = sum(valores_validos) / len(valores_validos) if valores_validos else None

    promedio_percapita = None
    if df_promedio is not None and not df_promedio.empty:
        sub_n = df_promedio[df_promedio["Division"] == "Total Gerencia"].set_index("Fecha")["N_Ajustadores"]
        razones = [v / sub_n[f] for f, v in zip(fechas, valores) if v is not None and sub_n.get(f)]
        promedio_percapita = sum(razones) / len(razones) if razones else None
    return promedio_total, promedio_percapita


def generar_pptx_historico(df_historico, df_promedio=None):
    """df_historico: salida de tablero_historico.parsear_historico().
    df_promedio: salida opcional de
    tablero_historico.parsear_promedio_casos_por_ajustador() — si se pasa,
    agrega una slide con la evolución del Stock (Q) promedio por ajustador."""
    if df_historico is None or df_historico.empty:
        raise ValueError("No hay datos históricos para generar el PPTX.")

    prs = Presentation()
    prs.slide_width = ANCHO_SLIDE
    prs.slide_height = ALTO_SLIDE

    fechas = sorted(df_historico["Fecha"].unique())
    etiquetas = [f"{f.day:02d}-{MESES_ES[f.month]}" for f in fechas]

    def serie_de(division, columna):
        sub = df_historico[df_historico["Division"] == division].set_index("Fecha")[columna]
        return [sub.get(f, None) for f in fechas]

    rango = f"{etiquetas[0]} al {etiquetas[-1]} de {fechas[-1].year} · {len(fechas)} semanas"

    slide = _slide_base(prs)
    _agregar_header(slide, "Tablero Gerencial — Evolución Histórica", f"Ingeniería y Equipo Móvil · {rango}")
    caja = slide.shapes.add_textbox(Inches(0.5), Inches(3), ANCHO_SLIDE - Inches(1), Inches(2))
    tf = caja.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "JPV Asociados — Ajustadores Especializados"
    p.font.size, p.font.bold, p.font.color.rgb = Pt(18), True, NAVY

    slide = _slide_base(prs)
    _agregar_header(slide, "Evolución de Stock — Honorarios (UF)", "Total Gerencia")
    _agregar_grafico(slide, XL_CHART_TYPE.LINE_MARKERS, etiquetas, [
        ("Total Gerencia", serie_de("Total Gerencia", "Stock_UF")),
    ], num_format="#,##0", color_offset=2)

    slide = _slide_base(prs)
    _agregar_header(slide, "Evolución de Stock — Cantidad de Casos (Q)", "Total Gerencia")
    _agregar_grafico(slide, XL_CHART_TYPE.LINE_MARKERS, etiquetas, [
        ("Total Gerencia", serie_de("Total Gerencia", "Stock_Q")),
    ], color_offset=2)

    slide = _slide_base(prs)
    _agregar_header(slide, "Evolución de Stock — Cantidad de Casos (Q)", None)
    _agregar_grafico(slide, XL_CHART_TYPE.LINE_MARKERS, etiquetas, [
        ("Ingeniería y Energía", serie_de("Ingeniería y Energía", "Stock_Q")),
        ("Equipo Móvil", serie_de("Equipo Móvil", "Stock_Q")),
    ])

    slide = _slide_base(prs)
    _agregar_header(slide, "Evolución de Stock — Honorarios (UF)", None)
    # Un gráfico por división: la escala de Ingeniería en UF es muchas veces
    # la de Equipo Móvil, y un solo eje combinado aplanaba su tendencia.
    _agregar_par_dividido(slide, XL_CHART_TYPE.LINE_MARKERS, etiquetas,
                           serie_de("Ingeniería y Energía", "Stock_UF"), serie_de("Equipo Móvil", "Stock_UF"),
                           num_format="#,##0")

    slide = _slide_base(prs)
    _agregar_header(slide, "Asignaciones Semanales por Área", None)
    _agregar_par_dividido(slide, XL_CHART_TYPE.COLUMN_CLUSTERED, etiquetas,
                           serie_de("Ingeniería y Energía", "Asignados"), serie_de("Equipo Móvil", "Asignados"))

    slide = _slide_base(prs)
    _agregar_header(slide, "Asignaciones Semanales — Total Gerencia", None)
    _agregar_grafico(slide, XL_CHART_TYPE.COLUMN_CLUSTERED, etiquetas, [
        ("Total Gerencia", serie_de("Total Gerencia", "Asignados")),
    ], leyenda=False)

    slide = _slide_base(prs)
    _agregar_header(slide, "IFL Emitidos — Cantidad de Casos", None)
    _agregar_par_dividido(slide, XL_CHART_TYPE.COLUMN_CLUSTERED, etiquetas,
                           serie_de("Ingeniería y Energía", "IFL_Q"), serie_de("Equipo Móvil", "IFL_Q"))

    slide = _slide_base(prs)
    _agregar_header(slide, "IFL Emitidos — Total Gerencia", "Cantidad de Casos")
    _agregar_grafico(slide, XL_CHART_TYPE.LINE_MARKERS, etiquetas, [
        ("Total Gerencia", serie_de("Total Gerencia", "IFL_Q")),
    ], leyenda=False, color_offset=2)

    slide = _slide_base(prs)
    promedio_total, promedio_percapita = _promedio_semanal_y_percapita(df_historico, df_promedio, "IFL_UF", fechas)
    partes_subtitulo = []
    if promedio_total is not None:
        partes_subtitulo.append(f"Promedio semanal: {promedio_total:,.0f} UF")
    if promedio_percapita is not None:
        partes_subtitulo.append(f"{promedio_percapita:,.0f} UF por ajustador vigente")
    _agregar_header(slide, "IFL Emitidos — Honorarios (UF)", " · ".join(partes_subtitulo) or None)
    # Un gráfico por división: la escala de Ingeniería en UF es muchas veces
    # la de Equipo Móvil, igual que en Stock UF.
    _agregar_par_dividido(slide, XL_CHART_TYPE.COLUMN_CLUSTERED, etiquetas,
                           serie_de("Ingeniería y Energía", "IFL_UF"), serie_de("Equipo Móvil", "IFL_UF"),
                           num_format="#,##0")

    slide = _slide_base(prs)
    _agregar_header(slide, "IFL Emitidos — Total Gerencia", "Honorarios (UF)")
    _agregar_grafico(slide, XL_CHART_TYPE.LINE_MARKERS, etiquetas, [
        ("Total Gerencia", serie_de("Total Gerencia", "IFL_UF")),
    ], num_format="#,##0", leyenda=False, color_offset=2)

    slide = _slide_base(prs)
    _agregar_header(slide, "Ingresos vs Egresos — Total Gerencia", "Asignaciones (ingresos) vs IFL Emitidos (egresos)")
    _agregar_grafico(slide, XL_CHART_TYPE.LINE_MARKERS, etiquetas, [
        ("Asignaciones (ingresos)", serie_de("Total Gerencia", "Asignados")),
        ("IFL Emitidos (egresos)", serie_de("Total Gerencia", "IFL_Q")),
    ])

    slide = _slide_base(prs)
    _agregar_header(slide, "Ingresos vs Egresos por Área", "Asignaciones (ingresos) vs IFL Emitidos (egresos)")
    _agregar_par_dividido_multiserie(slide, XL_CHART_TYPE.LINE_MARKERS, etiquetas, [
        ("Asignaciones (ingresos)", serie_de("Ingeniería y Energía", "Asignados")),
        ("IFL Emitidos (egresos)", serie_de("Ingeniería y Energía", "IFL_Q")),
    ], [
        ("Asignaciones (ingresos)", serie_de("Equipo Móvil", "Asignados")),
        ("IFL Emitidos (egresos)", serie_de("Equipo Móvil", "IFL_Q")),
    ])

    if df_promedio is not None and not df_promedio.empty:
        fechas_prom = sorted(df_promedio["Fecha"].unique())
        etiquetas_prom = [f"{f.day:02d}-{MESES_ES[f.month]}" for f in fechas_prom]

        def serie_promedio(division):
            sub = df_promedio[df_promedio["Division"] == division].set_index("Fecha")["Promedio_Q"]
            return [sub.get(f, None) for f in fechas_prom]

        slide = _slide_base(prs)
        _agregar_header(slide, "Stock Promedio por Ajustador", "Excluye a Dagoberto Pérez (carga no representativa)")
        _agregar_grafico(slide, XL_CHART_TYPE.LINE_MARKERS, etiquetas_prom, [
            ("Ingeniería y Energía", serie_promedio("Ingeniería y Energía")),
            ("Equipo Móvil", serie_promedio("Equipo Móvil")),
        ])

    buffer = io.BytesIO()
    prs.save(buffer)
    return buffer.getvalue()
