"""
Identidad visual única para todos los pptx que genera la app (Estado de
Cartera / Reporte por Ajustador, Ficha de Caso y Tablero Gerencial), tomada
de la plantilla corporativa JPV_Plantilla_LogoFondoBlanco.potx (logo
JPV Asociados · Sedgwick Alliance, tema de color "JPV Asociados").

Cada pptx se construye 100% por código (sin depender de un archivo .pptx
plantilla con shape_id fijos): así el mismo patrón de header/footer se
puede aplicar de forma consistente a cualquier slide, sin importar cuántas
haga falta agregar.
"""
import os

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Pt

# --- Colores (tema "JPV Asociados" del .potx) ---
NAVY = RGBColor(0x1B, 0x2A, 0x4B)          # dk1 — header, texto principal
NAVY_OSCURO = RGBColor(0x11, 0x1D, 0x37)   # dk2 — variante más oscura
BLANCO = RGBColor(0xFF, 0xFF, 0xFF)        # lt1
GRIS_CLARO = RGBColor(0xF8, 0xF9, 0xFB)    # lt2 — fondo del cuerpo
TEAL_OSCURO = RGBColor(0x0D, 0x73, 0x77)   # accent1 — footer
TEAL = RGBColor(0x10, 0xAA, 0x9E)          # accent2 — divisores/acentos
TEAL_CLARO = RGBColor(0x6F, 0xC7, 0xC1)    # accent3 — etiquetas sobre navy
GRIS_SLATE = RGBColor(0x5B, 0x6B, 0x7F)    # accent4 — texto secundario
TEAL_PALIDO = RGBColor(0x9F, 0xD8, 0xD3)   # accent5
GRIS_PALIDO = RGBColor(0xC7, 0xD0, 0xDA)   # accent6

# --- Tipografía (tema "JPV Asociados" del .potx) ---
FUENTE_TITULOS = "Cambria"
FUENTE_TEXTO = "Calibri"

LOGO_PATH = os.path.join(os.path.dirname(__file__), "assets", "logo_jpv.png")

# --- Geometría base (16:9 widescreen, igual al .potx) ---
SLIDE_WIDTH = Emu(12192000)
SLIDE_HEIGHT = Emu(6858000)

HEADER_ALTO = Emu(914400)          # 1 in
LOGO_CAJA_ANCHO = Emu(2926080)     # ancho de la caja blanca del logo
DIVISOR_GROSOR = Emu(36576)        # línea teal vertical/horizontal
FOOTER_ALTO = Emu(365760)          # 0.4 in
MARGEN_LATERAL = Emu(457200)       # 0.5 in — margen estándar de contenido

CONTENIDO_TOP = HEADER_ALTO + DIVISOR_GROSOR
CONTENIDO_ANCHO = SLIDE_WIDTH - 2 * MARGEN_LATERAL
CONTENIDO_BOTTOM = SLIDE_HEIGHT - FOOTER_ALTO


def nueva_presentacion():
    prs = Presentation()
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT
    return prs


def nueva_slide(prs):
    """Agrega una slide completamente en blanco (layout en blanco del
    master por defecto de python-pptx, slide_layouts[6])."""
    return prs.slides.add_slide(prs.slide_layouts[6])


def fondo(slide, color):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color


def _rect(slide, left, top, width, height, color, line=False):
    from pptx.enum.shapes import MSO_SHAPE
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shp.fill.solid()
    shp.fill.fore_color.rgb = color
    if not line:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = color
    shp.shadow.inherit = False
    return shp


def _texto(shape, texto, color, size, bold=False, italic=False, align=PP_ALIGN.LEFT,
           fuente=FUENTE_TEXTO, anchor=None):
    tf = shape.text_frame
    tf.word_wrap = True
    if anchor is not None:
        tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = texto
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = fuente
    return run


def rect(slide, left, top, width, height, color, line=False):
    """Versión pública de _rect: cualquier módulo puede dibujar bloques de
    color sueltos (tarjetas KPI, acentos) con la misma paleta."""
    return _rect(slide, left, top, width, height, color, line=line)


def caja_texto(slide, left, top, width, height, texto_, color, size, bold=False, italic=False,
                align=PP_ALIGN.LEFT, fuente=FUENTE_TEXTO, anchor=None, word_wrap=True):
    """Crea un textbox y le pone el texto de una vez (atajo usado por todos
    los módulos para tarjetas KPI, celdas de resumen, etc.)."""
    box = slide.shapes.add_textbox(left, top, width, height)
    box.text_frame.word_wrap = word_wrap
    _texto(box, texto_, color, size, bold=bold, italic=italic, align=align, fuente=fuente, anchor=anchor)
    return box


def agregar_portada_header(slide):
    """Encabezado minimal para portadas/separadores de sección: franja teal
    fina arriba + caja blanca con el logo, sobre fondo navy de ancho
    completo (el fondo lo define el llamador con `fondo(slide, NAVY)`)."""
    _rect(slide, 0, 0, SLIDE_WIDTH, Emu(27432), TEAL)
    _rect(slide, MARGEN_LATERAL, Emu(329184), LOGO_CAJA_ANCHO, HEADER_ALTO, BLANCO)

    if os.path.exists(LOGO_PATH):
        alto_logo = Emu(int(HEADER_ALTO * 0.55))
        top_logo = Emu(int(329184 + (HEADER_ALTO - alto_logo) / 2))
        pic = slide.shapes.add_picture(LOGO_PATH, 0, top_logo, height=alto_logo)
        pic.left = Emu(int(MARGEN_LATERAL + (LOGO_CAJA_ANCHO - pic.width) / 2))


def agregar_header(slide, titulo, etiqueta=None, fondo_claro=True):
    """Header estándar: banda navy de ancho completo, caja blanca con el
    logo JPV · Sedgwick a la izquierda, divisor teal vertical y uno
    horizontal justo debajo, título en blanco y una etiqueta opcional a la
    derecha (tono teal claro). `fondo_claro` define el color de fondo del
    cuerpo de la slide (True = gris claro corporativo, False = navy, para
    portadas/separadores)."""
    fondo(slide, GRIS_CLARO if fondo_claro else NAVY)

    _rect(slide, 0, 0, SLIDE_WIDTH, HEADER_ALTO, NAVY)
    _rect(slide, 0, 0, LOGO_CAJA_ANCHO, HEADER_ALTO, BLANCO)
    _rect(slide, LOGO_CAJA_ANCHO, 0, DIVISOR_GROSOR, HEADER_ALTO, TEAL)
    _rect(slide, 0, HEADER_ALTO, SLIDE_WIDTH, DIVISOR_GROSOR, TEAL)

    if os.path.exists(LOGO_PATH):
        alto_logo = Emu(int(HEADER_ALTO * 0.55))
        top_logo = Emu(int((HEADER_ALTO - alto_logo) / 2))
        # add_picture sin "width" preserva el aspect ratio nativo de la imagen.
        pic = slide.shapes.add_picture(LOGO_PATH, 0, top_logo, height=alto_logo)
        pic.left = Emu(int((LOGO_CAJA_ANCHO - pic.width) / 2))

    etiqueta_ancho = Emu(2286000) if etiqueta else 0
    titulo_left = LOGO_CAJA_ANCHO + DIVISOR_GROSOR + MARGEN_LATERAL
    titulo_ancho = SLIDE_WIDTH - titulo_left - etiqueta_ancho - MARGEN_LATERAL
    titulo_box = slide.shapes.add_textbox(titulo_left, 0, titulo_ancho, HEADER_ALTO)
    _texto(titulo_box, titulo, BLANCO, 20, bold=True, fuente=FUENTE_TITULOS,
           anchor=MSO_ANCHOR.MIDDLE)

    if etiqueta:
        etq_box = slide.shapes.add_textbox(SLIDE_WIDTH - etiqueta_ancho - MARGEN_LATERAL, 0, etiqueta_ancho, HEADER_ALTO)
        _texto(etq_box, etiqueta, TEAL_CLARO, 11, bold=True, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)


def agregar_footer(slide, texto_izquierda="JPV Asociados · Sedgwick Alliance", numero=None):
    """Footer estándar: banda teal oscuro de ancho completo al pie, con
    texto blanco a la izquierda y, opcionalmente, un número de página a
    la derecha."""
    top = SLIDE_HEIGHT - FOOTER_ALTO
    _rect(slide, 0, top, SLIDE_WIDTH, FOOTER_ALTO, TEAL_OSCURO)

    texto_box = slide.shapes.add_textbox(MARGEN_LATERAL, top, Emu(9601200), FOOTER_ALTO)
    _texto(texto_box, texto_izquierda, BLANCO, 10, anchor=MSO_ANCHOR.MIDDLE)

    if numero is not None:
        num_box = slide.shapes.add_textbox(SLIDE_WIDTH - Emu(548640) - MARGEN_LATERAL, top, Emu(548640), FOOTER_ALTO)
        _texto(num_box, str(numero), BLANCO, 10, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
