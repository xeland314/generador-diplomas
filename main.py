import copy
import io
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


def duplicar_slide(prs, slide):
    """Duplica una diapositiva copiando formas y reconstruyendo

    las relaciones de las imágenes para que no se pierdan.
    """
    layout = slide.slide_layout
    nueva = prs.slides.add_slide(layout)

    # Elimina los placeholders creados automáticamente por el layout
    for shape in list(nueva.shapes):
        el = shape._element
        el.getparent().remove(el)

    # Copia todos los elementos de la diapositiva original
    for shape in slide.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            image_stream = io.BytesIO(shape.image.blob)
            nueva.shapes.add_picture(
                image_stream, shape.left, shape.top, shape.width, shape.height
            )
        else:
            nuevo_el = copy.deepcopy(shape._element)
            nueva.shapes._spTree.insert_element_before(nuevo_el, "p:extLst")

    return nueva


def reemplazar_etiquetas(slide, datos):
    """Busca y reemplaza las etiquetas de texto de forma segura."""
    for shape in slide.shapes:
        if shape.has_text_frame:
            text_frame = getattr(shape, "text_frame", None)
            if text_frame:
                for paragraph in text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.text = (
                            run.text.replace("{{nombre}}", datos.get("nombre", ""))
                            .replace("{{curso}}", datos.get("curso", ""))
                            .replace("{{fecha}}", datos.get("fecha", ""))
                        )


def generar_certificados_pptx(
    lista_participantes,
    ruta_plantilla="templates/plantilla_v3.pptx",
    archivo_salida="Certificados.pptx",
):
    prs = Presentation(ruta_plantilla)

    # La primera diapositiva es nuestra base/plantilla
    plantilla = prs.slides[0]

    # Generar una nueva diapositiva por cada participante
    for persona in lista_participantes:
        slide = duplicar_slide(prs, plantilla)
        reemplazar_etiquetas(slide, persona)

    # Eliminar la primera diapositiva (la plantilla original con los {{tags}})
    id_slide = prs.slides._sldIdLst[0]
    prs.slides._sldIdLst.remove(id_slide)

    prs.save(archivo_salida)
    print(
        f"¡Proceso completado! Se generó '{archivo_salida}' con {len(lista_participantes)} certificados."
    )


if __name__ == "__main__":
    import os

    datos_ejemplo = [
        {
            "nombre": "Carlos Mendoza",
            "curso": "Python Básico",
            "fecha": "12 de Julio de 2026",
        },
        {
            "nombre": "Ana María Silva",
            "curso": "Python Intermedio",
            "fecha": "15 de Julio de 2026",
        },
        {
            "nombre": "Luis Alfredo Pérez",
            "curso": "Python Básico",
            "fecha": "12 de Julio de 2026",
        },
    ]
    if os.path.exists("templates/plantilla_v3.pptx"):
        generar_certificados_pptx(datos_ejemplo)
    else:
        print("Por favor, asegúrate de tener el archivo 'templates/plantilla_v3.pptx'.")
