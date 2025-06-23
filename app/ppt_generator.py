from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from PIL import Image
import os
import logging
from app.utils import download_image

# Configurar logging para ver los errores en la consola (y en los logs de Render)
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- CONFIGURACIÓN DE ESTILOS Y TEMAS (SIMPLIFICADO) ---
# Colores (ejemplos, puedes cambiarlos a los de tu marca)
# CORRECCIÓN AQUÍ: Pasar R, G, B como argumentos separados
COLOR_PRINCIPAL = RGBColor(0x00, 0x56, 0x91) # Un azul corporativo (0x005691)
COLOR_SECUNDARIO = RGBColor(0xEE, 0xEE, 0xEE) # Un gris claro para fondos o texto sutil (0xEEEEEE)
COLOR_TEXTO_OSCURO = RGBColor(0x33, 0x33, 0x33) # Gris oscuro para texto principal (0x333333)

# Fuentes (asegúrate de que estén disponibles en el entorno de Render o usa fuentes comunes)
FUENTE_TITULO = 'Arial' # O 'Calibri Light', 'Roboto', etc.
FUENTE_CUERPO = 'Arial' # O 'Calibri', 'Open Sans', etc.

# --- FUNCIÓN PRINCIPAL ---
def generar_pptx(data, filename):
    pr = Presentation()

    # --- INICIO DE LAS VALIDACIONES DE ENTRADA ---
    datos = None
    imagenes = []

    # Permitir tanto lista como objeto como entrada
    if isinstance(data, list):
        if not data:
            logging.error("Input data list is empty.")
            raise ValueError("Input data list is empty.")
        datos = data[0] if isinstance(data[0], dict) else None
        # Extraer imágenes del resto de la lista
        for d in data[1:]:
            if isinstance(d, dict) and "url" in d:
                imagenes.append(d["url"])
            elif isinstance(d, str) and d.startswith("http"):
                imagenes.append(d)
            else:
                logging.warning(f"Elemento de imagen con formato incorrecto: {d}")
    elif isinstance(data, dict):
        datos = data
        # Buscar claves que puedan contener urls de imagen
        posibles_claves = ["imagenes", "images", "urls", "graficos", "charts"]
        for clave in posibles_claves:
            if clave in data and isinstance(data[clave], list):
                for img in data[clave]:
                    if isinstance(img, str) and img.startswith("http"):
                        imagenes.append(img)
                    elif isinstance(img, dict) and "url" in img and img["url"].startswith("http"):
                        imagenes.append(img["url"])
        # Buscar claves individuales que sean urls
        for v in data.values():
            if isinstance(v, str) and v.startswith("http"):
                imagenes.append(v)
    else:
        logging.error("Input data must be a list or a dict.")
        raise ValueError("Input data must be a list or a dict.")

    if not datos or not isinstance(datos, dict):
        logging.error("No se pudo extraer el objeto de datos principal.")
        raise ValueError("No se pudo extraer el objeto de datos principal.")
    # --- FIN DE LAS VALIDACIONES DE ENTRADA ---

    # --- DISEÑO PROFESIONAL DEL PPT ---
    # Fondo degradado azul-gris para todas las slides
    SLIDE_BG_COLOR_1 = RGBColor(0x00, 0x56, 0x91)  # Azul corporativo
    SLIDE_BG_COLOR_2 = RGBColor(0xEE, 0xEE, 0xEE)  # Gris claro
    SLIDE_BG_COLOR_3 = RGBColor(0xF4, 0xF8, 0xFB)  # Azul muy claro

    def set_slide_background(slide, color=SLIDE_BG_COLOR_3):
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = color

    # --- SLIDE RESUMEN GLOBAL ---
    slide_layout_title_only = pr.slide_layouts[5]
    slide = pr.slides.add_slide(slide_layout_title_only)
    set_slide_background(slide, SLIDE_BG_COLOR_3)

    title = slide.shapes.title
    title.text = "Informe de Medios"
    title.text_frame.paragraphs[0].font.name = FUENTE_TITULO
    title.text_frame.paragraphs[0].font.size = Pt(36)
    title.text_frame.paragraphs[0].font.color.rgb = COLOR_PRINCIPAL
    title.text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT

    # Caja con borde y fondo para datos resumen
    left = Inches(0.8)
    top = Inches(1.8)
    width = Inches(8.2)
    height = Inches(2.2)
    txBox = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    txBox.fill.solid()
    txBox.fill.fore_color.rgb = SLIDE_BG_COLOR_2
    txBox.line.color.rgb = COLOR_PRINCIPAL
    txBox.shadow.inherit = False
    tf = txBox.text_frame
    tf.clear()
    tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
    tf.word_wrap = True

    p = tf.add_paragraph()
    p.text = f"🗓️ Período: {datos.get('fechaInicial', 'N/A')} al {datos.get('fechaFinal', 'N/A')}"
    p.font.name = FUENTE_CUERPO
    p.font.size = Pt(18)
    p.font.color.rgb = COLOR_TEXTO_OSCURO

    p = tf.add_paragraph()
    p.text = f"📰 Noticias totales: {datos.get('totalGlobalNoticias', 'N/A')}"
    p.font.name = FUENTE_CUERPO
    p.font.size = Pt(18)
    p.font.color.rgb = COLOR_TEXTO_OSCURO

    p = tf.add_paragraph()
    p.text = f"👥 Audiencia total: {datos.get('totalGlobalAudiencia', 'N/A')}"
    p.font.name = FUENTE_CUERPO
    p.font.size = Pt(18)
    p.font.color.rgb = COLOR_TEXTO_OSCURO

    p = tf.add_paragraph()
    p.text = f"💸 VPE Total: {datos.get('totalGlobalVPE', 'N/A')} | VC Total: {datos.get('totalGlobalVC', 'N/A')}"
    p.font.name = FUENTE_CUERPO
    p.font.size = Pt(18)
    p.font.color.rgb = COLOR_TEXTO_OSCURO

    add_footer(slide, "Informe Generado Automáticamente")

    # --- SLIDES POR MEDIO ---
    for medio_nombre in ["TV", "Radio", "Prensa", "Medios Digitales"]:
        medio_data = datos.get(medio_nombre + "_raw")
        if not medio_data:
            logging.info(f"No hay datos para el medio: {medio_nombre}. Saltando diapositiva.")
            continue
        slide = pr.slides.add_slide(slide_layout_title_only)
        set_slide_background(slide, SLIDE_BG_COLOR_1 if medio_nombre=="TV" else SLIDE_BG_COLOR_2)
        slide.shapes.title.text = f"📊 Análisis: {medio_nombre}"
        slide.shapes.title.text_frame.paragraphs[0].font.name = FUENTE_TITULO
        slide.shapes.title.text_frame.paragraphs[0].font.size = Pt(32)
        slide.shapes.title.text_frame.paragraphs[0].font.color.rgb = COLOR_PRINCIPAL

        # Caja de datos
        left = Inches(0.8)
        top = Inches(1.6)
        width = Inches(8.2)
        height = Inches(2.0)
        txBox = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
        txBox.fill.solid()
        txBox.fill.fore_color.rgb = SLIDE_BG_COLOR_3
        txBox.line.color.rgb = COLOR_PRINCIPAL
        txBox.shadow.inherit = False
        tf = txBox.text_frame
        tf.clear()
        tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
        tf.word_wrap = True

        p = tf.add_paragraph()
        p.text = (
            f"📰 Noticias: {medio_data.get('cantidad_noticias', 'N/A')}\n"
            f"👥 Audiencia: {medio_data.get('total_audiencia', 'N/A')}\n"
            f"💸 VPE: {medio_data.get('total_vpe', 'N/A')} | VC: {medio_data.get('total_vc', 'N/A')}"
        )
        p.font.name = FUENTE_CUERPO
        p.font.size = Pt(16)
        p.font.color.rgb = COLOR_TEXTO_OSCURO

        noticias_list = medio_data.get("noticias", [])
        if noticias_list:
            p = tf.add_paragraph()
            p.text = "\nNoticias Destacadas:"
            p.font.name = FUENTE_CUERPO
            p.font.size = Pt(14)
            p.font.bold = True
            p.font.color.rgb = COLOR_PRINCIPAL
            for noticia in noticias_list:
                p = tf.add_paragraph()
                p.text = f"- {noticia.get('fecha', 'N/A')}: {noticia.get('titulo', 'N/A')}"
                p.font.name = FUENTE_CUERPO
                p.font.size = Pt(12)
                p.font.color.rgb = COLOR_TEXTO_OSCURO
        add_footer(slide, f"{medio_nombre} - Informe de Medios")

    # --- SLIDES DE GRÁFICOS ---
    slide_layout_blank = pr.slide_layouts[6]
    for idx, url in enumerate(imagenes):
        slide = pr.slides.add_slide(slide_layout_blank)
        set_slide_background(slide, SLIDE_BG_COLOR_1)
        logging.info(f"[IMG {idx+1}/{len(imagenes)}] Intentando incrustar imagen desde: {url}")
        img_path = download_image(url)
        if img_path:
            try:
                img = Image.open(img_path)
                width_px, height_px = img.size
                aspect_ratio = width_px / height_px
                slide_width = pr.slide_width
                slide_height = pr.slide_height
                margin_left_right_inches = Inches(0.5)
                margin_top_bottom_inches = Inches(0.5)
                max_width_emu = slide_width - (margin_left_right_inches * 2).emu
                max_height_emu = slide_height - (margin_top_bottom_inches * 2).emu
                if (max_width_emu / aspect_ratio) <= max_height_emu:
                    scaled_width_emu = max_width_emu
                    scaled_height_emu = scaled_width_emu / aspect_ratio
                else:
                    scaled_height_emu = max_height_emu
                    scaled_width_emu = scaled_height_emu * aspect_ratio
                left = (slide_width - scaled_width_emu) / 2
                top = (slide_height - scaled_height_emu) / 2
                slide.shapes.add_picture(img_path, left, top, width=scaled_width_emu, height=scaled_height_emu)
                logging.info(f"[IMG {idx+1}] Imagen '{url}' incrustada exitosamente.")
            except Exception as e:
                logging.error(f"[IMG {idx+1}] Error al procesar o añadir imagen {url}: {e}")
                # Slide de error visual con URL
                error_box = slide.shapes.add_textbox(Inches(1), Inches(3), Inches(8), Inches(1.2))
                tf = error_box.text_frame
                tf.text = f"Error al procesar imagen:\n{url}"
                tf.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
                tf.paragraphs[0].font.size = Pt(18)
            finally:
                if os.path.exists(img_path):
                    try:
                        os.remove(img_path)
                        logging.info(f"[IMG {idx+1}] Imagen temporal eliminada: {img_path}")
                    except OSError as e:
                        logging.error(f"[IMG {idx+1}] Error al eliminar la imagen temporal {img_path}: {e}")
        else:
            logging.warning(f"[IMG {idx+1}] No se pudo descargar la imagen de la URL: {url}")
            # Slide de error visual con URL
            error_box = slide.shapes.add_textbox(Inches(1), Inches(3), Inches(8), Inches(1.2))
            tf = error_box.text_frame
            tf.text = f"No se pudo descargar la imagen:\n{url}"
            tf.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
            tf.paragraphs[0].font.size = Pt(18)
        add_footer(slide, "Gráfico de Medios")

    output_path = f"/tmp/{filename}"
    try:
        pr.save(output_path)
        logging.info(f"Presentación PPTX guardada en: {output_path}")
    except Exception as e:
        logging.error(f"Error al guardar la presentación PPTX: {e}")
        raise
    return output_path

    # --- SLIDE RESUMEN GLOBAL ---
    slide_layout_title_only = pr.slide_layouts[5]
    slide = pr.slides.add_slide(slide_layout_title_only)

    title = slide.shapes.title
    title.text = "Informe de Medios"
    title.text_frame.paragraphs[0].font.name = FUENTE_TITULO
    title.text_frame.paragraphs[0].font.size = Pt(36)
    title.text_frame.paragraphs[0].font.color.rgb = COLOR_PRINCIPAL
    title.text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT

    left = Inches(1.0)
    top = Inches(2.0)
    width = Inches(8.0)
    height = Inches(4.0)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
    tf.word_wrap = True

    p = tf.add_paragraph()
    p.text = f"🗓️ Período: {datos.get('fechaInicial', 'N/A')} al {datos.get('fechaFinal', 'N/A')}"
    p.font.name = FUENTE_CUERPO
    p.font.size = Pt(18)
    p.font.color.rgb = COLOR_TEXTO_OSCURO

    p = tf.add_paragraph()
    p.text = f"📰 Noticias totales: {datos.get('totalGlobalNoticias', 'N/A')}"
    p.font.name = FUENTE_CUERPO
    p.font.size = Pt(18)
    p.font.color.rgb = COLOR_TEXTO_OSCURO

    p = tf.add_paragraph()
    p.text = f"👥 Audiencia total: {datos.get('totalGlobalAudiencia', 'N/A')}"
    p.font.name = FUENTE_CUERPO
    p.font.size = Pt(18)
    p.font.color.rgb = COLOR_TEXTO_OSCURO

    p = tf.add_paragraph()
    p.text = f"💸 VPE Total: {datos.get('totalGlobalVPE', 'N/A')} | VC Total: {datos.get('totalGlobalVC', 'N/A')}"
    p.font.name = FUENTE_CUERPO
    p.font.size = Pt(18)
    p.font.color.rgb = COLOR_TEXTO_OSCURO

    add_footer(slide, "Informe Generado Automáticamente")


    # --- SLIDES POR MEDIO ---
    for medio_nombre in ["TV", "Radio", "Prensa", "Medios Digitales"]:
        medio_data = datos.get(medio_nombre + "_raw")
        if not medio_data:
            logging.info(f"No hay datos para el medio: {medio_nombre}. Saltando diapositiva.")
            continue

        slide = pr.slides.add_slide(slide_layout_title_only)
        slide.shapes.title.text = f"📊 Análisis: {medio_nombre}"
        slide.shapes.title.text_frame.paragraphs[0].font.name = FUENTE_TITULO
        slide.shapes.title.text_frame.paragraphs[0].font.size = Pt(32)
        slide.shapes.title.text_frame.paragraphs[0].font.color.rgb = COLOR_PRINCIPAL

        left = Inches(1.0)
        top = Inches(1.8)
        width = Inches(8.0)
        height = Inches(4.5)
        txBox = slide.shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame
        tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
        tf.word_wrap = True

        p = tf.add_paragraph()
        p.text = (
            f"📰 Noticias: {medio_data.get('cantidad_noticias', 'N/A')}\n"
            f"👥 Audiencia: {medio_data.get('total_audiencia', 'N/A')}\n"
            f"💸 VPE: {medio_data.get('total_vpe', 'N/A')} | VC: {medio_data.get('total_vc', 'N/A')}"
        )
        p.font.name = FUENTE_CUERPO
        p.font.size = Pt(16)
        p.font.color.rgb = COLOR_TEXTO_OSCURO

        noticias_list = medio_data.get("noticias", [])
        if noticias_list:
            p = tf.add_paragraph()
            p.text = "\nNoticias Destacadas:"
            p.font.name = FUENTE_CUERPO
            p.font.size = Pt(14)
            p.font.bold = True
            p.font.color.rgb = COLOR_PRINCIPAL

            for noticia in noticias_list:
                p = tf.add_paragraph()
                p.text = f"- {noticia.get('fecha', 'N/A')}: {noticia.get('titulo', 'N/A')}"
                p.font.name = FUENTE_CUERPO
                p.font.size = Pt(12)
                p.font.color.rgb = COLOR_TEXTO_OSCURO

        add_footer(slide, f"{medio_nombre} - Informe de Medios")


    # --- SLIDES DE GRÁFICOS ---
    slide_layout_blank = pr.slide_layouts[6]

    for url in imagenes:
        slide = pr.slides.add_slide(slide_layout_blank)
        logging.info(f"Intentando incrustar imagen desde: {url}")
        img_path = download_image(url)

        if img_path:
            try:
                img = Image.open(img_path)
                width_px, height_px = img.size
                aspect_ratio = width_px / height_px

                slide_width = pr.slide_width
                slide_height = pr.slide_height

                margin_left_right_inches = Inches(0.5)
                margin_top_bottom_inches = Inches(0.5)

                max_width_emu = slide_width - (margin_left_right_inches * 2).emu
                max_height_emu = slide_height - (margin_top_bottom_inches * 2).emu

                if (max_width_emu / aspect_ratio) <= max_height_emu:
                    scaled_width_emu = max_width_emu
                    scaled_height_emu = scaled_width_emu / aspect_ratio
                else:
                    scaled_height_emu = max_height_emu
                    scaled_width_emu = scaled_height_emu * aspect_ratio

                left = (slide_width - scaled_width_emu) / 2
                top = (slide_height - scaled_height_emu) / 2

                slide.shapes.add_picture(img_path, left, top, width=scaled_width_emu, height=scaled_height_emu)
                logging.info(f"Imagen '{url}' incrustada exitosamente.")

            except Exception as e:
                logging.error(f"Error al procesar o añadir imagen {url}: {e}")
            finally:
                if os.path.exists(img_path):
                    try:
                        os.remove(img_path)
                        logging.info(f"Imagen temporal eliminada: {img_path}")
                    except OSError as e:
                        logging.error(f"Error al eliminar la imagen temporal {img_path}: {e}")
        else:
            logging.warning(f"Advertencia: No se pudo incrustar la imagen de la URL: {url}")
            txBox = slide.shapes.add_textbox(Inches(1), Inches(3), Inches(8), Inches(1))
            tf = txBox.text_frame
            tf.text = "Error: Imagen no disponible."
            tf.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
            tf.paragraphs[0].font.size = Pt(24)

        add_footer(slide, "Gráfico de Medios")


    output_path = f"/tmp/{filename}"
    try:
        pr.save(output_path)
        logging.info(f"Presentación PPTX guardada en: {output_path}")
    except Exception as e:
        logging.error(f"Error al guardar la presentación PPTX: {e}")
        raise
    return output_path

def add_footer(slide, text_content):
    """Añade un pie de página simple a la diapositiva."""
    left = Inches(0.5)
    top = Inches(7.0)
    width = Inches(9.0)
    height = Inches(0.5)

    footer_shape = slide.shapes.add_textbox(left, top, width, height)
    tf = footer_shape.text_frame
    tf.text = text_content
    tf.paragraphs[0].font.name = FUENTE_CUERPO
    tf.paragraphs[0].font.size = Pt(10)
    tf.paragraphs[0].font.color.rgb = COLOR_TEXTO_OSCURO
    tf.paragraphs[0].alignment = 3
