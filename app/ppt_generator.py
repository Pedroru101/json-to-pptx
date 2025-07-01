from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from PIL import Image
import os
import logging
from app.utils import download_image

# Configuración de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Paleta de colores corporativa
COLORES = {
    'principal': RGBColor(0x00, 0x3D, 0x7D),    # Azul corporativo oscuro
    'secundario': RGBColor(0x00, 0x84, 0xD1),   # Azul corporativo medio
    'acento': RGBColor(0xFF, 0x8C, 0x00),       # Naranja acento
    'fondo_claro': RGBColor(0xF5, 0xF9, 0xFF),  # Azul muy claro para fondos
    'texto_oscuro': RGBColor(0x2C, 0x3E, 0x50), # Gris azulado para texto
    'blanco': RGBColor(0xFF, 0xFF, 0xFF),       # Blanco puro
    'gris_claro': RGBColor(0xE5, 0xE5, 0xE5)    # Gris claro para bordes
}

# Configuración de fuentes
FUENTES = {
    'titulo': 'Segoe UI',
    'subtitulo': 'Segoe UI Light',
    'cuerpo': 'Segoe UI'
}

def aplicar_estilo_slide(slide):
    """Aplica el estilo base a una diapositiva."""
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = COLORES['fondo_claro']

def add_footer(slide, text_content):
    """Añade un pie de página mejorado."""
    footer = slide.shapes.add_textbox(Inches(0.5), Inches(6.9), Inches(9), Inches(0.3))
    tf = footer.text_frame
    tf.text = text_content
    p = tf.paragraphs[0]
    p.font.size = Pt(9)
    p.font.name = FUENTES['cuerpo']
    p.font.color.rgb = COLORES['texto_oscuro']
    p.alignment = PP_ALIGN.CENTER

def agregar_logo(slide):
    """Agrega el logo corporativo en la esquina superior derecha de cada diapositiva."""
    logo_url = "https://mmi-e.com/wp-content/uploads/2021/01/logo-mmi.png"
    logo_path = download_image(logo_url)
    
    if logo_path and os.path.exists(logo_path):
        try:
            # Posición en esquina superior derecha
            left = Inches(8.5)  # 10 inches (ancho total) - 1.5 inches (ancho del logo)
            top = Inches(0.2)
            width = Inches(1.2)
            
            # Añadir logo con tamaño fijo
            logo = slide.shapes.add_picture(logo_path, left, top, width=width)
            
        except Exception as e:
            logging.error(f"Error al añadir logo: {e}")
        finally:
            try:
                os.remove(logo_path)
            except OSError as e:
                logging.error(f"Error al eliminar logo temporal: {e}")

def crear_portada(pr, datos):
    slide = pr.slides.add_slide(pr.slide_layouts[5])  # Usar layout en blanco
    aplicar_estilo_slide(slide)
    
    # Barra superior decorativa
    top_rect = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        0,
        0,
        Inches(10),
        Inches(1.5)
    )
    top_rect.fill.solid()
    top_rect.fill.fore_color.rgb = COLORES['principal']
    top_rect.line.fill.background()
    
    # Título principal centrado horizontalmente
    title = slide.shapes.add_textbox(
        0,
        Inches(2.5),
        Inches(10),
        Inches(1)
    )
    tf = title.text_frame
    tf.text = "Informe de Medios"
    p = tf.paragraphs[0]
    p.font.name = FUENTES['titulo']
    p.font.size = Pt(44)
    p.font.color.rgb = COLORES['principal']
    p.alignment = PP_ALIGN.CENTER
    
    # Subtítulo con período
    subtitle = slide.shapes.add_textbox(
        0,
        Inches(4),
        Inches(10),
        Inches(0.75)
    )
    tf = subtitle.text_frame
    tf.text = f"Período: {datos.get('fechaInicial', 'N/A')} - {datos.get('fechaFinal', 'N/A')}"
    p = tf.paragraphs[0]
    p.font.name = FUENTES['subtitulo']
    p.font.size = Pt(24)
    p.font.color.rgb = COLORES['secundario']
    p.alignment = PP_ALIGN.CENTER
    
    # Añadir logo
    agregar_logo(slide)
    add_footer(slide, "Informe de Medios")

def crear_metodologia(pr):
    slide = pr.slides.add_slide(pr.slide_layouts[2])
    aplicar_estilo_slide(slide)
    
    # Título con barra de color
    title_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0.5), Inches(10), Inches(0.8))
    title_box.fill.solid()
    title_box.fill.fore_color.rgb = COLORES['principal']
    title_box.line.fill.background()
    
    title = slide.shapes.title
    title.top = Inches(0.6)
    title.text = "Metodología"
    title.text_frame.paragraphs[0].font.name = FUENTES['titulo']
    title.text_frame.paragraphs[0].font.size = Pt(32)
    title.text_frame.paragraphs[0].font.color.rgb = COLORES['blanco']
    title.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    # Contenido en caja con estilo
    content_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(2), Inches(8), Inches(4))
    content_box.fill.solid()
    content_box.fill.fore_color.rgb = COLORES['blanco']
    content_box.line.color.rgb = COLORES['gris_claro']
    
    tf = content_box.text_frame
    tf.word_wrap = True
    
    items = [
        "• Monitoreo continuo de medios",
        "• Análisis cualitativo y cuantitativo",
        "• Métricas de evaluación:",
        "   - Valor Publicitario Equivalente (VPE)",
        "   - Alcance y audiencia",
        "   - Sentimiento y tono",
        "   - Presencia en diferentes soportes"
    ]
    
    for item in items:
        p = tf.add_paragraph()
        p.text = item
        p.font.name = FUENTES['cuerpo']
        p.font.size = Pt(18)
        p.font.color.rgb = COLORES['texto_oscuro']
        if item.startswith("   -"):
            p.level = 1
    
    add_footer(slide, "Metodología - Informe de Medios")

def crear_datos_cobertura(pr, datos, tipo_medio):
    slide = pr.slides.add_slide(pr.slide_layouts[5])
    aplicar_estilo_slide(slide)
    
    medio_data = datos.get(f"{tipo_medio}_raw", {})
    if not medio_data:
        return
    
    # Barra de título horizontal
    title_box = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        0,
        0,
        Inches(10),
        Inches(1)
    )
    title_box.fill.solid()
    title_box.fill.fore_color.rgb = COLORES['principal']
    title_box.line.fill.background()
    
    # Título centrado horizontalmente
    title = slide.shapes.add_textbox(
        0,
        Inches(0.2),
        Inches(10),
        Inches(0.6)
    )
    tf = title.text_frame
    tf.text = f"Datos de Cobertura - {tipo_medio}"
    p = tf.paragraphs[0]
    p.font.name = FUENTES['titulo']
    p.font.size = Pt(28)
    p.font.color.rgb = COLORES['blanco']
    p.alignment = PP_ALIGN.CENTER
    
    # Caja de resumen con datos principales
    summary_box = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(1),
        Inches(1.5),
        Inches(8),
        Inches(1.5)
    )
    summary_box.fill.solid()
    summary_box.fill.fore_color.rgb = COLORES['blanco']
    summary_box.line.color.rgb = COLORES['secundario']
    
    tf = summary_box.text_frame
    tf.word_wrap = True
    
    # Datos de resumen con iconos
    p = tf.add_paragraph()
    p.text = f"📊 Total de Noticias: {medio_data.get('cantidad_noticias', 'N/A')}"
    p.font.name = FUENTES['cuerpo']
    p.font.size = Pt(14)
    p.font.color.rgb = COLORES['texto_oscuro']
    
    p = tf.add_paragraph()
    p.text = f"👥 Audiencia Total: {medio_data.get('total_audiencia', 'N/A')}"
    p.font.name = FUENTES['cuerpo']
    p.font.size = Pt(14)
    p.font.color.rgb = COLORES['texto_oscuro']
    
    p = tf.add_paragraph()
    p.text = f"💰 VPE: {medio_data.get('total_vpe', 'N/A')} | VC: {medio_data.get('total_vc', 'N/A')}"
    p.font.name = FUENTES['cuerpo']
    p.font.size = Pt(14)
    p.font.color.rgb = COLORES['texto_oscuro']
    
    # Lista de noticias con scroll
    noticias_list = medio_data.get("noticias", [])
    if noticias_list:
        news_box = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(1),
            Inches(3.2),
            Inches(8),
            Inches(3)
        )
        news_box.fill.solid()
        news_box.fill.fore_color.rgb = COLORES['blanco']
        news_box.line.color.rgb = COLORES['gris_claro']
        
        tf = news_box.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
        
        p = tf.add_paragraph()
        p.text = "📰 Noticias Destacadas"
        p.font.name = FUENTES['subtitulo']
        p.font.size = Pt(16)
        p.font.color.rgb = COLORES['secundario']
        p.space_after = Pt(12)
        
        items_per_slide = 4
        current_items = 0
        
        for noticia in noticias_list:
            if current_items >= items_per_slide:
                # Nueva diapositiva para más noticias
                slide = pr.slides.add_slide(pr.slide_layouts[5])
                aplicar_estilo_slide(slide)
                
                # Barra de título
                title_box = slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE,
                    0,
                    0,
                    Inches(10),
                    Inches(1)
                )
                title_box.fill.solid()
                title_box.fill.fore_color.rgb = COLORES['principal']
                title_box.line.fill.background()
                
                # Título
                title = slide.shapes.add_textbox(
                    0,
                    Inches(0.2),
                    Inches(10),
                    Inches(0.6)
                )
                tf = title.text_frame
                tf.text = f"Datos de Cobertura - {tipo_medio} (Continuación)"
                p = tf.paragraphs[0]
                p.font.name = FUENTES['titulo']
                p.font.size = Pt(28)
                p.font.color.rgb = COLORES['blanco']
                p.alignment = PP_ALIGN.CENTER
                
                # Nueva caja de noticias
                news_box = slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE,
                    Inches(1),
                    Inches(1.5),
                    Inches(8),
                    Inches(4.5)
                )
                news_box.fill.solid()
                news_box.fill.fore_color.rgb = COLORES['blanco']
                news_box.line.color.rgb = COLORES['gris_claro']
                
                tf = news_box.text_frame
                tf.word_wrap = True
                tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
                current_items = 0
                
                # Añadir logo a la nueva diapositiva
                agregar_logo(slide)
            
            p = tf.add_paragraph()
            p.text = f"📅 {noticia.get('fecha', 'N/A')} - {noticia.get('titulo', 'N/A')}"
            p.font.name = FUENTES['cuerpo']
            p.font.size = Pt(12)
            p.font.color.rgb = COLORES['texto_oscuro']
            p.space_before = Pt(6)
            p.space_after = Pt(2)
            
            p = tf.add_paragraph()
            p.text = f"     {noticia.get('titular', 'N/A')}"
            p.font.name = FUENTES['cuerpo']
            p.font.size = Pt(11)
            p.font.color.rgb = COLORES['secundario']
            p.font.italic = True
            p.space_after = Pt(12)
            
            current_items += 1
    
    # Añadir logo
    agregar_logo(slide)
    add_footer(slide, f"Cobertura {tipo_medio} - Informe de Medios")

def crear_graficos(pr, datos):
    urls = datos.get("urls", [])
    if not urls:
        return
    
    for url in urls:
        slide = pr.slides.add_slide(pr.slide_layouts[5])
        aplicar_estilo_slide(slide)
        
        # Determinar tipo de gráfico y título
        tipo_grafico = ""
        if "vpe_barra" in url:
            tipo_grafico = "VPE por Medio"
        elif "vpe_torta" in url:
            tipo_grafico = "Distribución de VPE"
        elif "impactos_barra" in url:
            tipo_grafico = "Impactos por Medio"
        elif "impactos_torta" in url:
            tipo_grafico = "Distribución de Impactos"
        elif "top10" in url:
            medio = url.split("top10_vpe_")[1].split(".")[0].replace("_", " ").title()
            tipo_grafico = f"Top 10 VPE - {medio}"
        
        # Barra de título
        title_box = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            0,
            0,
            Inches(10),
            Inches(1)
        )
        title_box.fill.solid()
        title_box.fill.fore_color.rgb = COLORES['principal']
        title_box.line.fill.background()
        
        # Título
        title = slide.shapes.add_textbox(
            0,
            Inches(0.2),
            Inches(10),
            Inches(0.6)
        )
        tf = title.text_frame
        tf.text = tipo_grafico
        p = tf.paragraphs[0]
        p.font.name = FUENTES['titulo']
        p.font.size = Pt(28)
        p.font.color.rgb = COLORES['blanco']
        p.alignment = PP_ALIGN.CENTER
        
        # Contenedor para el gráfico con fondo y sombra
        chart_container = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(0.5),
            Inches(1.5),
            Inches(9),
            Inches(5)
        )
        chart_container.fill.solid()
        chart_container.fill.fore_color.rgb = COLORES['blanco']
        chart_container.line.color.rgb = COLORES['gris_claro']
        chart_container.shadow.inherit = False
        
        # Descargar e insertar gráfico
        img_path = download_image(url)
        if img_path and os.path.exists(img_path):
            try:
                # Calcular dimensiones manteniendo proporción
                img = Image.open(img_path)
                img_width, img_height = img.size
                aspect_ratio = img_width / img_height
                
                # Ajustar dimensiones
                target_width = Inches(8.5)
                target_height = target_width / aspect_ratio
                
                # Centrar el gráfico
                left = Inches(0.75)
                top = Inches(1.7)
                
                # Insertar gráfico con z-order alto para estar en primer plano
                pic = slide.shapes.add_picture(
                    img_path,
                    left,
                    top,
                    width=target_width,
                    height=min(target_height, Inches(4.6))
                )
                
            except Exception as e:
                logging.error(f"Error al procesar gráfico {url}: {e}")
                error_box = slide.shapes.add_textbox(
                    Inches(2),
                    Inches(3),
                    Inches(6),
                    Inches(1)
                )
                tf = error_box.text_frame
                tf.text = "Error al cargar el gráfico"
                tf.paragraphs[0].font.color.rgb = COLORES['acento']
                tf.paragraphs[0].font.size = Pt(14)
                tf.paragraphs[0].alignment = PP_ALIGN.CENTER
            finally:
                try:
                    os.remove(img_path)
                except OSError as e:
                    logging.error(f"Error al eliminar imagen temporal {img_path}: {e}")
        
        # Añadir logo
        agregar_logo(slide)
        add_footer(slide, f"{tipo_grafico} - Informe de Medios")

def crear_vpe_totales(pr, datos):
    slide = pr.slides.add_slide(pr.slide_layouts[2])
    aplicar_estilo_slide(slide)
    
    # Barra de título
    title_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0.5), Inches(10), Inches(0.8))
    title_box.fill.solid()
    title_box.fill.fore_color.rgb = COLORES['principal']
    title_box.line.fill.background()
    
    title = slide.shapes.title
    title.top = Inches(0.6)
    title.text = "Valor Publicitario Equivalente (VPE) Total"
    title.text_frame.paragraphs[0].font.name = FUENTES['titulo']
    title.text_frame.paragraphs[0].font.size = Pt(32)
    title.text_frame.paragraphs[0].font.color.rgb = COLORES['blanco']
    title.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    # Caja de datos VPE
    vpe_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(2), Inches(8), Inches(4))
    vpe_box.fill.solid()
    vpe_box.fill.fore_color.rgb = COLORES['blanco']
    vpe_box.line.color.rgb = COLORES['secundario']
    
    tf = vpe_box.text_frame
    
    p = tf.add_paragraph()
    p.text = "Resumen de Valor Publicitario"
    p.font.name = FUENTES['subtitulo']
    p.font.size = Pt(20)
    p.font.color.rgb = COLORES['secundario']
    p.alignment = PP_ALIGN.CENTER
    
    p = tf.add_paragraph()
    p.text = f"VPE Total: {datos.get('totalGlobalVPE', 'N/A')}"
    p.font.name = FUENTES['cuerpo']
    p.font.size = Pt(28)
    p.font.color.rgb = COLORES['principal']
    p.alignment = PP_ALIGN.CENTER
    
    add_footer(slide, "VPE Total - Informe de Medios")

def generar_pptx(data, filename):
    pr = Presentation()
    
    # Validación de datos de entrada
    if not isinstance(data, (list, dict)):
        logging.error("Input data must be a list or dict.")
        raise ValueError("Input data must be a list or dict.")
    
    datos = data[0] if isinstance(data, list) and data else data
    if not isinstance(datos, dict):
        logging.error("No se pudo extraer el objeto de datos principal.")
        raise ValueError("No se pudo extraer el objeto de datos principal.")
    
    # Generar estructura de presentación
    crear_portada(pr, datos)
    crear_metodologia(pr)
    
    # Datos de cobertura por tipo de medio
    for medio in ["TV", "Radio", "Prensa", "Medios Digitales"]:
        crear_datos_cobertura(pr, datos, medio)
    
    crear_vpe_totales(pr, datos)
    crear_graficos(pr, datos)
    
    # Guardar presentación
    output_path = f"/tmp/{filename}"
    try:
        pr.save(output_path)
        logging.info(f"Presentación PPTX guardada en: {output_path}")
    except Exception as e:
        logging.error(f"Error al guardar la presentación PPTX: {e}")
        raise
    
    return output_path
