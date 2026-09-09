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


def generar_pptx_historico(df_historico):
    """df_historico: salida de tablero_historico.parsear_historico()."""
    if df_historico is None or df_historico.empty:
        raise ValueError("No hay datos históricos para generar el PPTX.")

    prs = Presentation()
    prs.slide_width = ANCHO_SLIDE
    prs.slide_height = ALTO_SLIDE

    fechas = sorted(df_historico["Fecha"].unique())
    meses_es = {1: "ene", 2: "feb", 3: "mar", 4: "abr", 5: "may", 6: "jun",
                7: "jul", 8: "ago", 9: "sep", 10: "oct", 11: "nov", 12: "dic"}
    etiquetas = [f"{f.day:02d}-{meses_es[f.month]}" for f in fechas]

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
    ])

    buffer = io.BytesIO()
    prs.save(buffer)
    return buffer.getvalue()
