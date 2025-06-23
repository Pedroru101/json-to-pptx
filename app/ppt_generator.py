from pptx import Presentation
from pptx.util import Inches
from .utils import download_image
import os

def generar_pptx(data, filename):
    pr = Presentation()
    datos = data[0]
    imagenes = [d["url"] for d in data[1:]]

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

    output_path = f"/tmp/{filename}"
    pr.save(output_path)
    return output_path
