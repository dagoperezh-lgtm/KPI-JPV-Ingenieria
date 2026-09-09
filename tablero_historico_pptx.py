"""
PPTX de tendencia: evolución de Stock (Q y UF) y Asignaciones semanales por
área y total de la gerencia, a partir del histórico armado en
tablero_historico.py.
"""
import io

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.util import Inches, Pt

AZUL = RGBColor(0x1F, 0x38, 0x64)
AZUL_CLARO = RGBColor(0x4A, 0x7A, 0xB5)
GRIS = RGBColor(0x59, 0x59, 0x59)

ANCHO_SLIDE = Inches(13.333)
ALTO_SLIDE = Inches(7.5)


def _slide_titulo(prs, titulo, subtitulo=None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    caja = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), ANCHO_SLIDE - Inches(1), Inches(1))
    tf = caja.text_frame
    tf.word_wrap = True
    tf.text = titulo
    tf.paragraphs[0].font.size = Pt(28)
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.color.rgb = AZUL
    if subtitulo:
        p = tf.add_paragraph()
        p.text = subtitulo
        p.font.size = Pt(14)
        p.font.color.rgb = GRIS
    return slide


def _agregar_grafico(slide, tipo, categorias, series, num_format=None):
    cd = CategoryChartData()
    cd.categories = categorias
    for nombre, valores in series:
        cd.add_series(nombre, valores)

    grafico_frame = slide.shapes.add_chart(
        tipo, Inches(0.5), Inches(1.35), ANCHO_SLIDE - Inches(1), ALTO_SLIDE - Inches(1.85), cd
    )
    chart = grafico_frame.chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    chart.legend.font.size = Pt(12)

    try:
        chart.category_axis.tick_labels.font.size = Pt(8)
    except Exception:
        pass
    try:
        eje_valor = chart.value_axis
        eje_valor.tick_labels.font.size = Pt(10)
        if num_format:
            eje_valor.tick_labels.number_format = num_format
            eje_valor.tick_labels.number_format_is_linked = False
    except Exception:
        pass

    colores = [AZUL, AZUL_CLARO, GRIS]
    for i, serie in enumerate(chart.series):
        try:
            if tipo in (XL_CHART_TYPE.LINE, XL_CHART_TYPE.LINE_MARKERS):
                serie.format.line.color.rgb = colores[i % len(colores)]
                serie.format.line.width = Pt(2.5)
            else:
                serie.format.fill.solid()
                serie.format.fill.fore_color.rgb = colores[i % len(colores)]
        except Exception:
            pass
    return chart


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
    _slide_titulo(prs, "Tablero Gerencial — Evolución Histórica", f"Ingeniería y Equipo Móvil · {rango}")

    slide = _slide_titulo(prs, "Evolución de Stock — Cantidad de Casos (Q)")
    _agregar_grafico(slide, XL_CHART_TYPE.LINE_MARKERS, etiquetas, [
        ("Ingeniería y Energía", serie_de("Ingeniería y Energía", "Stock_Q")),
        ("Equipo Móvil", serie_de("Equipo Móvil", "Stock_Q")),
    ])

    slide = _slide_titulo(prs, "Evolución de Stock — Honorarios (UF)")
    _agregar_grafico(slide, XL_CHART_TYPE.LINE_MARKERS, etiquetas, [
        ("Ingeniería y Energía", serie_de("Ingeniería y Energía", "Stock_UF")),
        ("Equipo Móvil", serie_de("Equipo Móvil", "Stock_UF")),
    ], num_format="#,##0")

    slide = _slide_titulo(prs, "Asignaciones Semanales por Área")
    _agregar_grafico(slide, XL_CHART_TYPE.COLUMN_STACKED, etiquetas, [
        ("Ingeniería y Energía", serie_de("Ingeniería y Energía", "Asignados")),
        ("Equipo Móvil", serie_de("Equipo Móvil", "Asignados")),
    ])

    slide = _slide_titulo(prs, "Asignaciones Semanales — Total Gerencia")
    _agregar_grafico(slide, XL_CHART_TYPE.COLUMN_CLUSTERED, etiquetas, [
        ("Total Gerencia", serie_de("Total Gerencia", "Asignados")),
    ])

    buffer = io.BytesIO()
    prs.save(buffer)
    return buffer.getvalue()
