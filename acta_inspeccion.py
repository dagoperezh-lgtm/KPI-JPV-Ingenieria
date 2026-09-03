"""
Extrae fotos + pie de foto desde un Acta de Inspección en Word (.docx), para
precargar la slide de "Registro Fotográfico" de la Ficha de Caso.

Se apoya en un patrón consistente que usan estas actas: cada bloque de fotos
es una tabla donde unas celdas traen solo imagen(es) (sin texto) y otras
traen solo el pie de foto (sin imagen). Cualquier tabla cuya primera fila SÍ
tenga texto (p.ej. "Antecedentes Generales", "Firmas") se descarta completa:
son tablas de secciones/datos o de firmas, no de fotos — así se evita, por
ejemplo, colar las imágenes de las firmas como si fueran fotos del siniestro.

Dentro de una tabla de fotos, cada imagen se empareja con el pie de foto
cuya columna esté más cerca (no simplemente "el primer texto de la tabla"):
varios bloques traen 2+ fotos con pies de foto DISTINTOS uno al lado del
otro (columnas separadas), y no una sola descripción compartida — si se le
asigna a todas las fotos el mismo texto (o el texto de la tabla entera
concatenado), el pie que se ve junto a cada foto deja de corresponderle.
"""
from docx import Document
from docx.oxml.ns import qn


def _imagenes_en_celda(cell):
    rids = []
    for blip in cell._element.findall(".//" + qn("a:blip")):
        rid = blip.get(qn("r:embed"))
        if rid:
            rids.append(rid)
    return rids


def _celdas_distintas_con_columnas(table):
    """Recorre la tabla y devuelve, para cada celda DISTINTA (una celda
    combinada horizontalmente aparece repetida en varias posiciones de
    row.cells; acá se reporta una sola vez), en qué columnas de esa fila
    aparece — necesario para saber a qué imagen le corresponde qué pie de
    foto cuando conviven varias en la misma tabla."""
    resultado = []  # [(cols_ocupadas: set[int], cell)]
    for row in table.rows:
        tcs_fila = [c._tc for c in row.cells]  # referencias vivas: comparar por identidad, no por id()
        vistas = []
        for c, cell in enumerate(row.cells):
            if any(cell._tc is tc for tc in vistas):
                continue
            vistas.append(cell._tc)
            cols_ocupadas = {j for j, tc in enumerate(tcs_fila) if tc is cell._tc}
            resultado.append((cols_ocupadas, cell))
    return resultado


def _centro(cols_ocupadas):
    return sum(cols_ocupadas) / len(cols_ocupadas)


def extraer_fotos_acta(archivo):
    """archivo: ruta o file-like (.docx). Devuelve una lista de dicts
    {"imagen": bytes, "pie": str} en el orden en que aparecen en el
    documento."""
    doc = Document(archivo)
    fotos = []

    for table in doc.tables:
        if len(table.rows) == 0:
            continue
        fila0 = table.rows[0]
        if any(cell.text.strip() for cell in fila0.cells):
            continue  # tabla de sección/firmas, no es un bloque de fotos

        celdas_imagen = []  # [(cols_ocupadas, [rids])]
        celdas_texto = []  # [(cols_ocupadas, texto)]
        for cols_ocupadas, cell in _celdas_distintas_con_columnas(table):
            rids = _imagenes_en_celda(cell)
            if rids:
                celdas_imagen.append((cols_ocupadas, rids))
                continue
            texto = cell.text.strip()
            if texto:
                celdas_texto.append((cols_ocupadas, texto))

        for cols_img, rids in celdas_imagen:
            pie = ""
            if celdas_texto:
                centro_img = _centro(cols_img)
                _, pie = min(celdas_texto, key=lambda ct: abs(_centro(ct[0]) - centro_img))
            for rid in rids:
                parte = doc.part.related_parts[rid]
                fotos.append({"imagen": parte.blob, "pie": pie})

    return fotos
