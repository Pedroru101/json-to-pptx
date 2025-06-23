from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from PIL import Image # Importar Pillow
import os
import logging # Para un mejor manejo de logs

# Configurar logging para ver los errores en la consola (y en los logs de Render)
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- CONFIGURACIÓN DE ESTILOS Y TEMAS (SIMPLIFICADO) ---
# Colores (ejemplos, puedes cambiarlos a los de tu marca)
COLOR_PRINCIPAL = RGBColor(0x005691) # Un azul corporativo
COLOR_SECUNDARIO = RGBColor(0xEEEEEE) # Un gris claro para fondos o texto sutil
COLOR_TEXTO_OSCURO = RGBColor(0x333333) # Gris oscuro para texto principal

# Fuentes (asegúrate de que estén disponibles en el entorno de Render o usa fuentes comunes)
FUENTE_TITULO = 'Arial' # O 'Calibri Light', 'Roboto', etc.
FUENTE_CUERPO = 'Arial' # O 'Calibri', 'Open Sans', etc.

# --- FUNCIÓN PRINCIPAL ---
def generar_pptx(data, filename):
    pr = Presentation()

    # --- INICIO DE LAS VALIDACIONES DE ENTRADA ---
    if not isinstance(data, list) or not data:
        logging.error("Input data must be a non-empty list.")
        raise ValueError("Input data must be a non-empty list.")

    # Intentar obtener los datos del resumen global
    try:
        datos = data[0]
        if not isinstance(datos, dict):
            logging.error(f"First item in data is not a dictionary: {type(datos)}")
            raise ValueError("First item in data must be a dictionary.")
    except IndexError:
        logging.error("Data list is empty for global summary.")
        raise ValueError("Data list is empty for global summary.")

    # Intentar obtener las URLs de las imágenes, manejando errores
    imagenes = []
    for d in data[1:]:
        if isinstance(d, dict) and "url" in d:
            imagenes.append(d["url"])
        else:
            logging.warning(f"Elemento de imagen con formato incorrecto o sin 'url': {d}")
    # --- FIN DE LAS VALIDACIONES DE ENTRADA ---

    # --- SLIDE RESUMEN GLOBAL ---
    # Usar un layout con solo título o incluso en blanco para más control
    # slide = pr.slides.add_slide(pr.slide_layouts[5]) # Layout original: Title and Content
    slide_layout_title_only = pr.slide_layouts[5] # Podría ser un "Título y Contenido"
    slide = pr.slides.add_slide(slide_layout_title_only)

    title = slide.shapes.title
    title.text = "Informe de Medios" # Título más general
    title.text_frame.paragraphs[0].font.name = FUENTE_TITULO
    title.text_frame.paragraphs[0].font.size = Pt(36)
    title.text_frame.paragraphs[0].font.color.rgb = COLOR_PRINCIPAL
    title.text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT # Ajustar tamaño del cuadro de texto

    # Contenido del resumen global
    left = Inches(1.0)
    top = Inches(2.0)
    width = Inches(8.0)
    height = Inches(4.0)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT # Ajustar tamaño del cuadro de texto
    tf.word_wrap = True # Habilitar ajuste de línea

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

    # Pie de página simple para la diapositiva de resumen
    add_footer(slide, "Informe Generado Automáticamente")


    # --- SLIDES POR MEDIO ---
    for medio_nombre in ["TV", "Radio", "Prensa", "Medios Digitales"]:
        medio_data = datos.get(medio_nombre + "_raw")
        if not medio_data:
            logging.info(f"No hay datos para el medio: {medio_nombre}. Saltando diapositiva.")
            continue

        slide = pr.slides.add_slide(slide_layout_title_only) # Usar el mismo layout
        slide.shapes.title.text = f"📊 Análisis: {medio_nombre}" # Título más descriptivo
        slide.shapes.title.text_frame.paragraphs[0].font.name = FUENTE_TITULO
        slide.shapes.title.text_frame.paragraphs[0].font.size = Pt(32)
        slide.shapes.title.text_frame.paragraphs[0].font.color.rgb = COLOR_PRINCIPAL

        # Contenido del medio
        left = Inches(1.0)
        top = Inches(1.8) # Un poco más arriba para dejar espacio al título
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

        # Añadir noticias si existen
        noticias_list = medio_data.get("noticias", [])
        if noticias_list:
            p = tf.add_paragraph() # Un párrafo para el encabezado de noticias
            p.text = "\nNoticias Destacadas:"
            p.font.name = FUENTE_CUERPO
            p.font.size = Pt(14)
            p.font.bold = True
            p.font.color.rgb = COLOR_PRINCIPAL # Color diferente para encabezado

            for noticia in noticias_list:
                p = tf.add_paragraph()
                p.text = f"- {noticia.get('fecha', 'N/A')}: {noticia.get('titulo', 'N/A')}"
                p.font.name = FUENTE_CUERPO
                p.font.size = Pt(12)
                p.font.color.rgb = COLOR_TEXTO_OSCURO

        add_footer(slide, f"{medio_nombre} - Informe de Medios")


    # --- SLIDES DE GRÁFICOS ---
    # Usamos un layout completamente en blanco para las imágenes
    slide_layout_blank = pr.slide_layouts[6]

    for url in imagenes:
        slide = pr.slides.add_slide(slide_layout_blank) # Layout en blanco para gráficos
        logging.info(f"Intentando incrustar imagen desde: {url}")
        img_path = download_image(url)

        if img_path:
            try:
                # Obtener las dimensiones de la imagen para calcular la relación de aspecto
                img = Image.open(img_path)
                width_px, height_px = img.size
                aspect_ratio = width_px / height_px

                # Dimensiones de la diapositiva para calcular el espacio disponible
                slide_width = pr.slide_width
                slide_height = pr.slide_height

                # Definir márgenes para la imagen (ej: 0.5 pulgadas por lado)
                margin_left_right_inches = Inches(0.5)
                margin_top_bottom_inches = Inches(0.5) # Dejar espacio para título o pie de página si se añade

                # Calcular el espacio máximo disponible para la imagen
                max_width_emu = slide_width - (margin_left_right_inches * 2).emu
                max_height_emu = slide_height - (margin_top_bottom_inches * 2).emu

                # Calcular el tamaño de la imagen manteniendo la relación de aspecto
                # Si la imagen es más ancha que alta para el espacio disponible
                if (max_width_emu / aspect_ratio) <= max_height_emu:
                    # El ancho es el factor limitante
                    scaled_width_emu = max_width_emu
                    scaled_height_emu = scaled_width_emu / aspect_ratio
                else:
                    # La altura es el factor limitante
                    scaled_height_emu = max_height_emu
                    scaled_width_emu = scaled_height_emu * aspect_ratio

                # Convertir EMUs a pulgadas para la posición
                scaled_width_inches = Inches(scaled_width_emu / Inches(1).emu)
                scaled_height_inches = Inches(scaled_height_emu / Inches(1).emu)

                # Centrar la imagen en la diapositiva
                left = (slide_width - scaled_width_emu) / 2
                top = (slide_height - scaled_height_emu) / 2

                slide.shapes.add_picture(img_path, left, top, width=scaled_width_emu, height=scaled_height_emu)
                logging.info(f"Imagen '{url}' incrustada exitosamente.")

            except Exception as e:
                logging.error(f"Error al procesar o añadir imagen {url}: {e}")
            finally:
                # Asegurarse de eliminar la imagen temporal incluso si hubo un error al añadirla
                if os.path.exists(img_path):
                    try:
                        os.remove(img_path)
                        logging.info(f"Imagen temporal eliminada: {img_path}")
                    except OSError as e:
                        logging.error(f"Error al eliminar la imagen temporal {img_path}: {e}")
        else:
            logging.warning(f"Advertencia: No se pudo incrustar la imagen de la URL: {url}")
            # Puedes añadir un placeholder de texto o forma si no se carga la imagen
            txBox = slide.shapes.add_textbox(Inches(1), Inches(3), Inches(8), Inches(1))
            tf = txBox.text_frame
            tf.text = "Error: Imagen no disponible."
            tf.paragraphs[0].font.color.rgb = RGBColor(0xFF0000) # Rojo para error
            tf.paragraphs[0].font.size = Pt(24)

        add_footer(slide, "Gráfico de Medios")


    output_path = f"/tmp/{filename}"
    try:
        pr.save(output_path)
        logging.info(f"Presentación PPTX guardada en: {output_path}")
    except Exception as e:
        logging.error(f"Error al guardar la presentación PPTX: {e}")
        raise # Volver a lanzar el error para que FastAPI lo capture
    return output_path

# --- FUNCIONES DE UTILIDAD (DENTRO DEL MISMO ARCHIVO O EN utils.py) ---
# Si esta función ya está en utils.py, asegúrate de que sea la versión mejorada.
# Si no, puedes añadirla aquí para probar, pero es mejor mantenerla en utils.py.
# (Mantendré el import .utils para usar la versión externa si existe)
# def download_image(url):
#     # ... (código de download_image con logging y timeout) ...


def add_footer(slide, text_content):
    """Añade un pie de página simple a la diapositiva."""
    left = Inches(0.5)
    top = Inches(7.0) # Posición cerca del final de la diapositiva
    width = Inches(9.0)
    height = Inches(0.5)

    footer_shape = slide.shapes.add_textbox(left, top, width, height)
    tf = footer_shape.text_frame
    tf.text = text_content
    tf.paragraphs[0].font.name = FUENTE_CUERPO
    tf.paragraphs[0].font.size = Pt(10)
    tf.paragraphs[0].font.color.rgb = COLOR_TEXTO_OSCURO
    tf.paragraphs[0].alignment = 3 # CENTER (puede ser LEFT o RIGHT)

    # Opcional: añadir un número de página (requiere más lógica para cada slide)
    # page_num_shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(9), Inches(7.0), Inches(0.5), Inches(0.5))
    # page_num_tf = page_num_shape.text_frame
    # page_num_tf.text = str(slide.slide_id) # Esto no es un contador de página real
