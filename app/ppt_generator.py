from pptx import Presentation
from pptx.util import Inches
from .utils import download_image
import os

def generar_pptx(data, filename):
    pr = Presentation()

    # --- INICIO DE LAS VALIDACIONES DE ENTRADA ---
    if not isinstance(data, list) or not data:
        raise ValueError("Input data must be a non-empty list.")

    # Intentar obtener los datos del resumen global
    try:
        datos = data[0]
        if not isinstance(datos, dict):
            raise ValueError("First item in data must be a dictionary.")
    except IndexError:
        raise ValueError("Data list is empty for global summary.")

    # Intentar obtener las URLs de las imágenes, manejando errores
    imagenes = []
    for d in data[1:]:
        if isinstance(d, dict) and "url" in d:
            imagenes.append(d["url"])
        else:
            # Aquí puedes decidir qué hacer:
            # 1. Ignorar el elemento malformado (como está ahora con el print de advertencia).
            # 2. Levantar un error para detener la ejecución si las imágenes son obligatorias.
            #    raise ValueError(f"Image item has incorrect format or missing 'url': {d}")
            print(f"Advertencia: Elemento de imagen con formato incorrecto o sin 'url': {d}")
    # --- FIN DE LAS VALIDACIONES DE ENTRADA ---

    # Slide resumen global
    slide = pr.slides.add_slide(pr.slide_layouts[5])
    title = slide.shapes.title
    title.text = "Resumen Global"
    cuerpo = (
        f"🗓️ Período: {datos['fechaInicial']} al {datos['fechaFinal']}\n"
        f"📰 Noticias totales: {datos['totalGlobalNoticias']}\n"
        f"👥 Audiencia total: {datos['totalGlobalAudiencia']}\n"
        f"💸 VPE Total: {datos['totalGlobalVPE']} | VC Total: {datos['totalGlobalVC']}"
    )
    txBox = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(8), Inches(2))
    tf = txBox.text_frame
    tf.text = cuerpo

    # Slide por medio
    for medio in ["TV", "Radio", "Prensa", "Medios Digitales"]:
        medio_data = datos.get(medio + "_raw")
        if not medio_data:
            continue
        slide = pr.slides.add_slide(pr.slide_layouts[5])
        slide.shapes.title.text = medio
        cuerpo = (
            f"📰 Noticias: {medio_data['cantidad_noticias']}\n"
            f"👥 Audiencia: {medio_data['total_audiencia']}\n"
            f"💸 VPE: {medio_data['total_vpe']} | VC: {medio_data['total_vc']}"
        )
        txBox = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(8), Inches(2))
        tf = txBox.text_frame
        tf.text = cuerpo
        for noticia in medio_data.get("noticias", []):
            tf.add_paragraph().text = f"- {noticia['fecha']}: {noticia['titulo']}"

    # Slides de gráficos
    for url in imagenes:
        slide = pr.slides.add_slide(pr.slide_layouts[5])
        img_path = download_image(url)
        if img_path:
            slide.shapes.add_picture(img_path, Inches(1), Inches(1.5), height=Inches(4.5))
            # Considera eliminar la imagen temporal después de usarla
            try:
                os.remove(img_path)
            except OSError as e:
                print(f"Error al eliminar la imagen temporal {img_path}: {e}")

    output_path = f"/tmp/{filename}"
    pr.save(output_path)
    return output_path
