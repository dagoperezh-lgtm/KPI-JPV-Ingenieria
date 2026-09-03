"""
Extrae fotos + pie de foto desde un Acta de Inspección en Word (.docx), para
precargar la slide de "Registro Fotográfico" de la Ficha de Caso.

Se apoya en un patrón consistente que usan estas actas: cada bloque de fotos
es una tabla de 2 filas — la fila 0 trae solo la(s) imagen(es) (sin texto) y
la fila 1 trae solo el pie de foto (sin imagen). Cualquier tabla cuya fila 0
SÍ tenga texto (p.ej. "Antecedentes Generales", "Firmas") se descarta
completa: son tablas de secciones/datos o de firmas, no de fotos — así se
evita, por ejemplo, colar las imágenes de las firmas como si fueran fotos
del siniestro.
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

        filas_con_imagen = []
        texto_pie = ""
        for row in table.rows:
            rids_fila, textos_fila, vistos = [], [], set()
            for cell in row.cells:
                rids = _imagenes_en_celda(cell)
                if rids:
                    for r in rids:
                        if r not in vistos:
                            vistos.add(r)
                            rids_fila.append(r)
                else:
                    texto = cell.text.strip()
                    if texto and texto not in textos_fila:
                        textos_fila.append(texto)
            if rids_fila:
                filas_con_imagen.append(rids_fila)
            elif textos_fila and not texto_pie:
                texto_pie = " ".join(textos_fila)

        for rids in filas_con_imagen:
            for rid in rids:
                parte = doc.part.related_parts[rid]
                fotos.append({"imagen": parte.blob, "pie": texto_pie})

    return fotos
