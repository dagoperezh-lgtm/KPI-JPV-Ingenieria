"""
Generador de la "Ficha de Caso" individual (7 slides, pptx independiente
del Reporte de Cartera):

1. Resumen del caso (Ficha del Caso).
2. Descripción del Siniestro / Materia Asegurada (texto libre, en blanco).
3. Registro Fotográfico: si se le pasan fotos ya seleccionadas (p.ej.
   extraídas de un Acta de Inspección con acta_inspeccion.py), se generan
   tantas páginas de 6 fotos (con su pie de foto) como hagan falta; si no,
   una sola slide con 6 casilleros en blanco para completar a mano.
4. Reserva del caso: valor y desglose extraídos de la planilla base
   (Pérdida bruta, Deducible, Monto asegurado, Pérdida neta/Reserva,
   Gastos, Honorarios), con una columna en blanco para que el ajustador
   justifique cada concepto.
5. Estado Actual del Siniestro: precargado con la observación completa
   del Pipeline para ese caso (si existe), editable; si no hay dato,
   queda en blanco como texto libre.
6. Detalle de Gestiones Realizadas (tabla de 6 filas vacías, Fecha +
   Detalle, a llenar a mano).
7. Próximas Acciones (texto libre, en blanco).

Las slides de texto libre en blanco (2 y 7, y la 5 cuando no hay
observación del Pipeline) llevan la nota "(Texto libre)".

No usa una plantilla .pptx: arma las slides desde cero con python-pptx,
reutilizando branding.py (header navy + logo JPV Asociados · Sedgwick
Alliance, footer teal, cuerpo claro, 16:9) para que tenga la misma
identidad visual que el Reporte de Cartera.
"""
import io

import pandas as pd
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.util import Emu, Inches, Pt

import branding as br

CONTENT_LEFT = br.MARGEN_LATERAL
CONTENT_TOP = br.CONTENIDO_TOP + Emu(228600)
CONTENT_WIDTH = br.CONTENIDO_ANCHO
CONTENT_BOTTOM = br.CONTENIDO_BOTTOM - Emu(91440)
CONTENT_HEIGHT = CONTENT_BOTTOM - CONTENT_TOP

GRIS_TEXTO = RGBColor(0x33, 0x33, 0x33)
GRIS_BORDE = RGBColor(0xB0, 0xB8, 0xC2)
GRIS_NOTA = RGBColor(0x8A, 0x93, 0x9E)


def _truncar(texto, max_chars):
    texto = str(texto).strip()
    if len(texto) <= max_chars:
        return texto
    return texto[: max_chars - 1].rstrip() + "…"


def _subtitulo(datos):
    return datos.get("nickname") or datos.get("asegurado") or ""


def _campo(slide, left, top, width, etiqueta, valor):
    br.caja_texto(slide, left, top, width, Emu(173736), etiqueta.upper(), br.TEAL_OSCURO, 9, bold=True)
    valor_texto = str(valor) if valor not in (None, "", "nan", "NaT") else "—"
    br.caja_texto(slide, left, top + Emu(228600), width, Emu(347472), valor_texto, GRIS_TEXTO, 14)


def _slide_resumen(prs, datos, numero):
    slide = br.nueva_slide(prs)
    br.agregar_header(slide, "FICHA DE CASO", etiqueta=_subtitulo(datos))
    br.agregar_footer(slide, numero=numero)

    ancho_col = (CONTENT_WIDTH - br.MARGEN_LATERAL) // 2
    col1_x, col2_x = CONTENT_LEFT, CONTENT_LEFT + ancho_col + br.MARGEN_LATERAL
    top = CONTENT_TOP
    paso = Emu(749808)

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
    return slide


def _caja_manual(slide, left, top, width, height, titulo, texto_precargado=None, fuente=None):
    """Etiqueta + caja con borde punteado. Sin `texto_precargado`, queda
    vacía con la nota "(Texto libre)" para completar a mano. Con
    `texto_precargado` (p.ej. "Hechos y Circunstancias" del Acta de
    Inspección), se precarga ese texto —editable— y se indica la `fuente`
    en una nota al pie."""
    br.caja_texto(slide, left, top, width, Emu(274320), titulo.upper(), br.TEAL_OSCURO, 11, bold=True)

    alto_nota = Emu(201168) if (texto_precargado and fuente) else Emu(0)
    top_caja = top + Emu(292608)
    alto_caja = height - Emu(292608) - alto_nota
    caja = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top_caja, width, alto_caja)
    caja.fill.solid()
    caja.fill.fore_color.rgb = br.GRIS_CLARO
    caja.line.color.rgb = GRIS_BORDE
    caja.line.width = Pt(1)
    caja.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    caja.shadow.inherit = False
    tf = caja.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    tf.margin_left, tf.margin_right = Pt(8), Pt(8)
    tf.margin_top, tf.margin_bottom = Pt(6), Pt(6)
    p = tf.paragraphs[0]
    if texto_precargado:
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        run = p.add_run()
        run.text = texto_precargado
        run.font.size, run.font.color.rgb = Pt(12), GRIS_TEXTO
    else:
        run = p.add_run()
        run.text = "(Texto libre)"
        run.font.size, run.font.italic, run.font.color.rgb = Pt(11), True, GRIS_NOTA

    if texto_precargado and fuente:
        br.caja_texto(slide, left, top_caja + alto_caja + Emu(18288), width, alto_nota,
                      f"Fuente: {fuente} — edítalo si corresponde.", GRIS_NOTA, 8, italic=True)


def _slide_descripcion(prs, datos, numero, descripcion_siniestro=None, descripcion_fuente=None):
    slide = br.nueva_slide(prs)
    br.agregar_header(slide, "DESCRIPCIÓN DEL SINIESTRO Y MATERIA ASEGURADA", etiqueta=_subtitulo(datos))
    br.agregar_footer(slide, numero=numero)

    alto_caja, espacio = (CONTENT_HEIGHT - Emu(228600)) // 2, Emu(228600)

    _caja_manual(
        slide, CONTENT_LEFT, CONTENT_TOP, CONTENT_WIDTH, alto_caja, "Descripción del Siniestro",
        texto_precargado=descripcion_siniestro, fuente=descripcion_fuente,
    )
    _caja_manual(slide, CONTENT_LEFT, CONTENT_TOP + alto_caja + espacio, CONTENT_WIDTH, alto_caja, "Materia Asegurada")
    return slide


_FOTOS_POR_PAGINA = 6
_FOTOS_COLS, _FOTOS_FILAS = 3, 2


def _posiciones_grilla_fotos():
    espacio = Emu(182880)
    ancho_disponible = CONTENT_WIDTH - (_FOTOS_COLS - 1) * espacio
    alto_disponible = CONTENT_HEIGHT - (_FOTOS_FILAS - 1) * espacio
    ancho, alto = Emu(int(ancho_disponible / _FOTOS_COLS)), Emu(int(alto_disponible / _FOTOS_FILAS))
    posiciones = []
    for f in range(_FOTOS_FILAS):
        for c in range(_FOTOS_COLS):
            left = CONTENT_LEFT + c * (ancho + espacio)
            top = CONTENT_TOP + f * (alto + espacio)
            posiciones.append((left, top, ancho, alto))
    return posiciones


def _casillero_vacio(slide, left, top, ancho, alto, numero):
    marco = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, ancho, alto)
    marco.fill.solid()
    marco.fill.fore_color.rgb = br.GRIS_CLARO
    marco.line.color.rgb = GRIS_BORDE
    marco.line.width = Pt(1)
    marco.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    marco.shadow.inherit = False
    tf = marco.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = f"Foto {numero}"
    run.font.size, run.font.color.rgb = Pt(13), GRIS_NOTA


def _casillero_con_foto(slide, left, top, ancho, alto, imagen_bytes, pie):
    """Sin marco/recuadro: la foto va directo sobre el fondo de la slide,
    con su pie de foto debajo (más limpio en modo presentación)."""
    alto_pie = Emu(365760)
    alto_imagen = alto - alto_pie

    try:
        from PIL import Image
        iw, ih = Image.open(io.BytesIO(imagen_bytes)).size
        escala = min(ancho / iw, alto_imagen / ih)
        w, h = int(iw * escala), int(ih * escala)
        x = left + (ancho - w) // 2
        y = top + (alto_imagen - h) // 2
        slide.shapes.add_picture(io.BytesIO(imagen_bytes), x, y, width=w, height=h)
    except Exception:
        slide.shapes.add_picture(io.BytesIO(imagen_bytes), left, top, width=ancho, height=alto_imagen)

    br.caja_texto(slide, left, top + alto_imagen, ancho, alto_pie, _truncar(pie, 90) if pie else "",
                  GRIS_TEXTO, 8, align=PP_ALIGN.CENTER)


def _slide_fotos(prs, datos, numero, fotos=None, pagina=None, total_paginas=None):
    """Si `fotos` viene vacío/None: una slide con 6 casilleros en blanco
    (comportamiento original, para completar a mano). Si trae fotos (ya
    seleccionadas por el ajustador): esta función se llama una vez por
    página de hasta 6 fotos — ver generar_ficha_pptx."""
    slide = br.nueva_slide(prs)
    subtitulo = _subtitulo(datos)
    if total_paginas and total_paginas > 1:
        subtitulo = f"{subtitulo} — Página {pagina} de {total_paginas}"
    br.agregar_header(slide, "REGISTRO FOTOGRÁFICO", etiqueta=subtitulo)
    br.agregar_footer(slide, numero=numero)

    posiciones = _posiciones_grilla_fotos()
    for i, (left, top, ancho, alto) in enumerate(posiciones, start=1):
        if fotos and i <= len(fotos):
            _casillero_con_foto(slide, left, top, ancho, alto, fotos[i - 1]["imagen"], fotos[i - 1].get("pie", ""))
        elif not fotos:
            _casillero_vacio(slide, left, top, ancho, alto, i)
    return slide


def _slide_gestiones(prs, datos, numero):
    slide = br.nueva_slide(prs)
    br.agregar_header(slide, "DETALLE DE GESTIONES REALIZADAS", etiqueta=_subtitulo(datos))
    br.agregar_footer(slide, numero=numero)

    n_filas_dato = 6
    graphic_frame = slide.shapes.add_table(n_filas_dato + 1, 2, CONTENT_LEFT, CONTENT_TOP, CONTENT_WIDTH, CONTENT_HEIGHT)
    tabla = graphic_frame.table
    ancho_fecha = Emu(1645920)
    tabla.columns[0].width = ancho_fecha
    tabla.columns[1].width = CONTENT_WIDTH - ancho_fecha

    for c, texto in enumerate(["Fecha", "Detalle de la gestión"]):
        celda = tabla.cell(0, c)
        celda.text_frame.paragraphs[0].text = texto
        celda.fill.solid()
        celda.fill.fore_color.rgb = br.NAVY_OSCURO
        run = celda.text_frame.paragraphs[0].runs[0]
        run.font.size, run.font.bold, run.font.color.rgb = Pt(12), True, br.BLANCO

    for r in range(1, n_filas_dato + 1):
        for c in range(2):
            celda = tabla.cell(r, c)
            celda.fill.solid()
            celda.fill.fore_color.rgb = br.BLANCO
            celda.text_frame.paragraphs[0].text = ""
    return slide


def _slide_reserva(prs, datos, numero):
    slide = br.nueva_slide(prs)
    br.agregar_header(slide, "RESERVA DEL CASO", etiqueta=_subtitulo(datos))
    br.agregar_footer(slide, numero=numero)

    # --- Valor destacado (extraído de la planilla base) ---
    alto_callout = Emu(822960)
    callout = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, CONTENT_LEFT, CONTENT_TOP, CONTENT_WIDTH, alto_callout)
    callout.fill.solid()
    callout.fill.fore_color.rgb = br.NAVY
    callout.line.fill.background()
    callout.shadow.inherit = False
    tf = callout.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = "RESERVA (Pérdida neta)"
    run.font.size, run.font.bold, run.font.color.rgb = Pt(12), True, br.TEAL_CLARO
    p2 = tf.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    run2 = p2.add_run()
    run2.text = datos.get("perdida_neta_fmt") or "—"
    run2.font.size, run2.font.bold, run2.font.color.rgb = Pt(30), True, br.BLANCO

    # --- Desglose extraído de la planilla + columna de justificación manual ---
    espacio = Emu(228600)
    top_tabla = CONTENT_TOP + alto_callout + espacio
    height_tabla = CONTENT_BOTTOM - top_tabla
    filas_valores = [
        ("Pérdida bruta", datos.get("perdida_bruta_fmt")),
        ("Deducible", datos.get("deducible_fmt")),
        ("Monto asegurado", datos.get("monto_asegurado_fmt")),
        ("Pérdida neta (Reserva)", datos.get("perdida_neta_fmt")),
        ("Gastos", datos.get("gastos_fmt")),
        ("Honorarios", datos.get("honorarios_fmt")),
    ]
    graphic_frame = slide.shapes.add_table(len(filas_valores) + 1, 3, CONTENT_LEFT, top_tabla, CONTENT_WIDTH, height_tabla)
    tabla = graphic_frame.table
    ancho_concepto, ancho_monto = Emu(2971800), Emu(2194560)
    tabla.columns[0].width = ancho_concepto
    tabla.columns[1].width = ancho_monto
    tabla.columns[2].width = CONTENT_WIDTH - ancho_concepto - ancho_monto

    for c, texto in enumerate(["Concepto", "Monto (planilla)", "Justificación del ajustador"]):
        celda = tabla.cell(0, c)
        celda.text_frame.paragraphs[0].text = texto
        celda.fill.solid()
        celda.fill.fore_color.rgb = br.NAVY_OSCURO
        run = celda.text_frame.paragraphs[0].runs[0]
        run.font.size, run.font.bold, run.font.color.rgb = Pt(11), True, br.BLANCO

    for r, (concepto, monto) in enumerate(filas_valores, start=1):
        celda_concepto = tabla.cell(r, 0)
        celda_concepto.text_frame.paragraphs[0].text = concepto
        celda_concepto.fill.solid()
        celda_concepto.fill.fore_color.rgb = br.GRIS_CLARO
        celda_concepto.text_frame.paragraphs[0].runs[0].font.size = Pt(11)

        celda_monto = tabla.cell(r, 1)
        celda_monto.text_frame.paragraphs[0].text = monto or "—"
        celda_monto.fill.solid()
        celda_monto.fill.fore_color.rgb = br.GRIS_CLARO
        celda_monto.text_frame.paragraphs[0].runs[0].font.size = Pt(11)

        celda_just = tabla.cell(r, 2)
        celda_just.text_frame.paragraphs[0].text = ""
        celda_just.fill.solid()
        celda_just.fill.fore_color.rgb = br.BLANCO
    return slide


def _slide_texto_libre(prs, datos, titulo, numero, texto_precargado=None, fuente=None):
    """Slide con un único cuadro grande de texto (usado por Estado Actual del
    Siniestro y Próximas Acciones). Si `texto_precargado` viene vacío, el
    cuadro queda en blanco con la nota "(Texto libre)"; si trae contenido
    (p.ej. la observación del Pipeline), se precarga ese texto —editable— y
    se indica su `fuente` en una nota al pie."""
    slide = br.nueva_slide(prs)
    br.agregar_header(slide, titulo, etiqueta=_subtitulo(datos))
    br.agregar_footer(slide, numero=numero)

    alto_nota = Emu(274320) if (texto_precargado and fuente) else Emu(0)
    height = CONTENT_HEIGHT - alto_nota

    caja = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, CONTENT_LEFT, CONTENT_TOP, CONTENT_WIDTH, height)
    caja.fill.solid()
    caja.fill.fore_color.rgb = br.GRIS_CLARO
    caja.line.color.rgb = GRIS_BORDE
    caja.line.width = Pt(1)
    caja.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    caja.shadow.inherit = False
    tf = caja.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    tf.margin_left, tf.margin_right = Pt(10), Pt(10)
    tf.margin_top, tf.margin_bottom = Pt(8), Pt(8)
    p = tf.paragraphs[0]
    if texto_precargado:
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        run = p.add_run()
        run.text = texto_precargado
        run.font.size, run.font.color.rgb = Pt(14), GRIS_TEXTO
    else:
        run = p.add_run()
        run.text = "(Texto libre)"
        run.font.size, run.font.italic, run.font.color.rgb = Pt(12), True, GRIS_NOTA

    if texto_precargado and fuente:
        br.caja_texto(slide, CONTENT_LEFT, CONTENT_TOP + height + Emu(27432), CONTENT_WIDTH, alto_nota,
                      f"Fuente: {fuente} — edítalo si corresponde.", GRIS_NOTA, 9, italic=True)
    return slide


def _fmt_monto(valor, divisa):
    try:
        numero = float(valor)
    except (TypeError, ValueError):
        return None
    texto = f"{numero:,.0f}".replace(",", ".")
    return f"{texto} {divisa}".strip()


def generar_ficha_pptx(
    fila, observacion_pipeline=None, observacion_fuente=None, fotos_seleccionadas=None,
    descripcion_siniestro=None, descripcion_fuente=None,
):
    """fila: dict o pandas.Series con los datos del caso, usando los mismos
    nombres de columna que trae la Base Maestra (ver app.py).

    observacion_pipeline: texto completo (opcional) para precargar la slide
    "Estado Actual del Siniestro" — normalmente la observación del Pipeline
    para ese Caso JPV. observacion_fuente describe de dónde salió (se
    muestra como nota al pie, p.ej. "Observaciones del Pipeline").

    fotos_seleccionadas: lista opcional de dicts {"imagen": bytes, "pie": str}
    (p.ej. de acta_inspeccion.extraer_fotos_acta, ya filtrada por el
    ajustador) para precargar el Registro Fotográfico. Se reparten en
    páginas de 6 fotos cada una; sin este argumento, queda una sola slide
    con 6 casilleros en blanco para completar a mano.

    descripcion_siniestro: texto completo (opcional) para precargar
    "Descripción del Siniestro" — normalmente el campo "Hechos y
    Circunstancias" del Acta de Inspección (ver
    acta_inspeccion.extraer_descripcion_siniestro). descripcion_fuente
    describe de dónde salió.
    """
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
        deducible_fmt=_fmt_monto(fila.get("Deducible (en moneda del caso)"), divisa),
        perdida_neta_fmt=_fmt_monto(fila.get("Perdida neta (en moneda del caso)"), divisa),
        gastos_fmt=_fmt_monto(fila.get("Gastos (UF)"), "UF"),
        honorarios_fmt=_fmt_monto(fila.get("Honorarios (UF)"), "UF"),
    )

    prs = br.nueva_presentacion()
    n = 1

    _slide_resumen(prs, datos, n); n += 1
    _slide_descripcion(prs, datos, n, descripcion_siniestro=descripcion_siniestro, descripcion_fuente=descripcion_fuente)
    n += 1
    if fotos_seleccionadas:
        paginas = [
            fotos_seleccionadas[i:i + _FOTOS_POR_PAGINA]
            for i in range(0, len(fotos_seleccionadas), _FOTOS_POR_PAGINA)
        ]
        for num_pagina, pagina in enumerate(paginas, start=1):
            _slide_fotos(prs, datos, n, fotos=pagina, pagina=num_pagina, total_paginas=len(paginas))
            n += 1
    else:
        _slide_fotos(prs, datos, n)
        n += 1
    _slide_reserva(prs, datos, n); n += 1
    _slide_texto_libre(
        prs, datos, "ESTADO ACTUAL DEL SINIESTRO", n,
        texto_precargado=observacion_pipeline, fuente=observacion_fuente,
    )
    n += 1
    _slide_gestiones(prs, datos, n); n += 1
    _slide_texto_libre(prs, datos, "PRÓXIMAS ACCIONES", n)

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
