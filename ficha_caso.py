"""
Generador de la "Ficha de Caso" individual (3 slides, pptx independiente
del Reporte de Cartera):

1. Resumen del caso.
2. Espacio en blanco para Registro Fotográfico (6 fotos, a llenar a mano
   en PowerPoint una vez descargado el pptx).
3. Espacio en blanco para el Detalle de Gestiones Realizadas (6 líneas de
   Fecha + Detalle, también a llenar a mano).

No usa una plantilla .pptx: arma las 3 slides desde cero con python-pptx,
reutilizando el mismo logo y paleta de colores (navy + teal) que
assets/plantilla_estado_cartera.pptx.
"""
import io
import os

import pandas as pd
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Emu, Inches, Pt

LOGO_PATH = os.path.join(os.path.dirname(__file__), "assets", "logo_jpv.png")

SLIDE_WIDTH = Emu(9144000)
SLIDE_HEIGHT = Emu(5143500)

NAVY_OSCURO = RGBColor(0x0D, 0x1F, 0x38)
NAVY = RGBColor(0x1B, 0x2A, 0x4A)
TEAL = RGBColor(0x14, 0xA8, 0xA0)
TEAL_OSCURO = RGBColor(0x0D, 0x73, 0x77)
GRIS_TEXTO = RGBColor(0x33, 0x33, 0x33)
GRIS_CLARO = RGBColor(0xF4, 0xF6, 0xF8)
BLANCO = RGBColor(0xFF, 0xFF, 0xFF)

HEADER_ALTO = Inches(0.82)


def _sin_relleno_ni_borde(shape):
    shape.fill.background()
    shape.line.fill.background()


def _agregar_header(slide, titulo, subtitulo):
    barra = slide.shapes.add_shape(1, 0, 0, SLIDE_WIDTH, HEADER_ALTO)
    barra.fill.solid()
    barra.fill.fore_color.rgb = NAVY
    barra.line.fill.background()
    barra.shadow.inherit = False

    divisor = slide.shapes.add_shape(1, 0, HEADER_ALTO, SLIDE_WIDTH, Emu(36576))
    divisor.fill.solid()
    divisor.fill.fore_color.rgb = TEAL_OSCURO
    divisor.line.fill.background()
    divisor.shadow.inherit = False

    if os.path.exists(LOGO_PATH):
        slide.shapes.add_picture(LOGO_PATH, Emu(48142), Emu(76341), height=Emu(528226))

    caja_titulo = slide.shapes.add_textbox(Inches(1.25), Inches(0.12), Inches(7.5), Inches(0.6))
    tf = caja_titulo.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = titulo
    p.font.size, p.font.bold, p.font.color.rgb = Pt(20), True, BLANCO

    p2 = tf.add_paragraph()
    p2.text = subtitulo
    p2.font.size, p2.font.italic, p2.font.color.rgb = Pt(12), True, TEAL


def _campo(slide, left, top, width, etiqueta, valor):
    caja = slide.shapes.add_textbox(left, top, width, Inches(0.62))
    tf = caja.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = etiqueta.upper()
    p.font.size, p.font.bold, p.font.color.rgb = Pt(9), True, TEAL_OSCURO

    p2 = tf.add_paragraph()
    p2.text = str(valor) if valor not in (None, "", "nan", "NaT") else "—"
    p2.font.size, p2.font.color.rgb = Pt(14), GRIS_TEXTO


def _slide_resumen(prs, datos):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _agregar_header(slide, "FICHA DE CASO", datos.get("nickname") or datos.get("asegurado") or "")

    col1_x, col2_x = Inches(0.5), Inches(5.15)
    ancho_col = Inches(4.3)
    top = Inches(1.15)
    paso = Inches(0.82)

    columna_1 = [
        ("Asegurado", datos.get("asegurado")),
        ("N° de Siniestro", datos.get("numero_siniestro")),
        ("Caso JPV", datos.get("caso_jpv")),
        ("Estado", datos.get("estado")),
        ("Monto asegurado", datos.get("monto_asegurado_fmt")),
    ]
    columna_2 = [
        ("Fecha de ocurrencia", datos.get("fecha_ocurrencia")),
        ("Fecha de denuncio", datos.get("fecha_denuncio")),
        ("Fecha de asignación", datos.get("fecha_asignacion")),
        ("Días desde asignación", datos.get("dias_asignacion")),
        ("Pérdida bruta", datos.get("perdida_bruta_fmt")),
    ]
    for i, (etiqueta, valor) in enumerate(columna_1):
        _campo(slide, col1_x, top + i * paso, ancho_col, etiqueta, valor)
    for i, (etiqueta, valor) in enumerate(columna_2):
        _campo(slide, col2_x, top + i * paso, ancho_col, etiqueta, valor)

    pie = slide.shapes.add_textbox(Inches(0.5), Inches(5.25), Inches(9), Inches(0.3))
    p = pie.text_frame.paragraphs[0]
    p.text = "JPV Asociados · Ajustadores Especializados"
    p.font.size, p.font.color.rgb = Pt(8), RGBColor(0x99, 0x99, 0x99)
    return slide


def _slide_fotos(prs, datos):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _agregar_header(slide, "REGISTRO FOTOGRÁFICO", datos.get("nickname") or datos.get("asegurado") or "")

    cols, filas = 3, 2
    margen = Inches(0.4)
    espacio = Inches(0.2)
    ancho_disponible = SLIDE_WIDTH - 2 * margen - (cols - 1) * espacio
    alto_disponible = SLIDE_HEIGHT - HEADER_ALTO - Emu(36576) - Inches(0.6) - (filas - 1) * espacio
    ancho_foto = Emu(int(ancho_disponible / cols))
    alto_foto = Emu(int(alto_disponible / filas))
    top_inicial = HEADER_ALTO + Emu(36576) + Inches(0.3)

    n = 1
    for f in range(filas):
        for c in range(cols):
            left = margen + c * (ancho_foto + espacio)
            top = top_inicial + f * (alto_foto + espacio)
            marco = slide.shapes.add_shape(1, left, top, ancho_foto, alto_foto)
            marco.fill.solid()
            marco.fill.fore_color.rgb = GRIS_CLARO
            marco.line.color.rgb = RGBColor(0xB0, 0xB8, 0xC2)
            marco.line.width = Pt(1)
            marco.line.dash_style = MSO_LINE_DASH_STYLE.DASH
            marco.shadow.inherit = False
            tf = marco.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            p.text = f"Foto {n}"
            p.font.size, p.font.color.rgb = Pt(13), RGBColor(0x8A, 0x93, 0x9E)
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            n += 1
    return slide


def _slide_gestiones(prs, datos):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _agregar_header(slide, "DETALLE DE GESTIONES REALIZADAS", datos.get("nickname") or datos.get("asegurado") or "")

    n_filas_dato = 6
    left, top = Inches(0.4), HEADER_ALTO + Emu(36576) + Inches(0.3)
    width = SLIDE_WIDTH - Inches(0.8)
    height = SLIDE_HEIGHT - top - Inches(0.3)

    graphic_frame = slide.shapes.add_table(n_filas_dato + 1, 2, left, top, width, height)
    tabla = graphic_frame.table
    tabla.columns[0].width = Inches(1.6)
    tabla.columns[1].width = width - Inches(1.6)

    for c, texto in enumerate(["Fecha", "Detalle de la gestión"]):
        celda = tabla.cell(0, c)
        celda.text_frame.paragraphs[0].text = texto
        celda.fill.solid()
        celda.fill.fore_color.rgb = NAVY_OSCURO
        run = celda.text_frame.paragraphs[0].runs[0]
        run.font.size, run.font.bold, run.font.color.rgb = Pt(12), True, BLANCO

    for r in range(1, n_filas_dato + 1):
        for c in range(2):
            celda = tabla.cell(r, c)
            celda.fill.solid()
            celda.fill.fore_color.rgb = BLANCO
            celda.text_frame.paragraphs[0].text = ""
    return slide


def _fmt_monto(valor, divisa):
    try:
        numero = float(valor)
    except (TypeError, ValueError):
        return None
    texto = f"{numero:,.0f}".replace(",", ".")
    return f"{texto} {divisa}".strip()


def generar_ficha_pptx(fila):
    """fila: dict o pandas.Series con los datos del caso, usando los mismos
    nombres de columna que trae la Base Maestra (ver app.py)."""
    divisa = str(fila.get("Divisa") or "").strip()
    datos = dict(
        asegurado=fila.get("Asegurado"),
        nickname=fila.get("Nickname"),
        numero_siniestro=fila.get("Número de siniestro"),
        caso_jpv=fila.get("ID_Caso"),
        estado=fila.get("Estado_Actual") or fila.get("Estado"),
        fecha_ocurrencia=_fmt_fecha(fila.get("Fecha de ocurrencia")),
        fecha_denuncio=_fmt_fecha(fila.get("Fecha de denuncio")),
        fecha_asignacion=_fmt_fecha(fila.get("Fecha de asignación")),
        dias_asignacion=fila.get("Días desde asignación"),
        monto_asegurado_fmt=_fmt_monto(fila.get("Monto asegurado (en moneda del caso)"), divisa),
        perdida_bruta_fmt=_fmt_monto(fila.get("Perdida bruta (en moneda del caso)"), divisa),
    )

    prs = Presentation()
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT

    _slide_resumen(prs, datos)
    _slide_fotos(prs, datos)
    _slide_gestiones(prs, datos)

    output = io.BytesIO()
    prs.save(output)
    return output.getvalue()


def _fmt_fecha(valor):
    if valor in (None, "", "nan", "NaT"):
        return None
    fecha = pd.to_datetime(valor, errors="coerce")
    if pd.isna(fecha):
        return str(valor)
    return fecha.strftime("%d-%m-%Y")
