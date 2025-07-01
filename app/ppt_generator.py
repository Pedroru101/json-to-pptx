from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from PIL import Image
import os
import logging
from app.utils import download_image
import tempfile
import requests

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
    # URL directa al logo de MMI
    logo_url = "https://mmi-e.com/wp-content/uploads/2021/01/logo-mmi.png"
    
    try:
        # Ruta temporal para almacenar el logo descargado
        temp_dir = tempfile.gettempdir()
        logo_path = os.path.join(temp_dir, "logo_mmi.png")
        
        # Intentar descargar el logo si no existe
        if not os.path.exists(logo_path) or os.path.getsize(logo_path) == 0:
            logging.info(f"Descargando logo desde {logo_url}")
            response = requests.get(logo_url, stream=True, timeout=15)
            if response.status_code == 200:
                with open(logo_path, 'wb') as f:
                    for chunk in response.iter_content(chunk_size=8192):
                        f.write(chunk)
                logging.info(f"Logo descargado correctamente a {logo_path}")
            else:
                logging.error(f"Error al descargar logo: {response.status_code}")
                # Intentar con URL alternativa
                alt_logo_url = "https://mmi-e.com/wp-content/uploads/2020/09/logo-mmi.png"
                logging.info(f"Intentando con URL alternativa: {alt_logo_url}")
                response = requests.get(alt_logo_url, stream=True, timeout=15)
                if response.status_code == 200:
                    with open(logo_path, 'wb') as f:
                        for chunk in response.iter_content(chunk_size=8192):
                            f.write(chunk)
                    logging.info(f"Logo alternativo descargado correctamente")
                else:
                    logging.error(f"Error al descargar logo alternativo: {response.status_code}")
                    return
        
        # Verificar que el archivo existe y tiene contenido
        if os.path.exists(logo_path) and os.path.getsize(logo_path) > 0:
            # Posición en esquina superior derecha
            left = Inches(8.5)
            top = Inches(0.2)
            width = Inches(1.2)
            
            # Añadir logo con tamaño fijo
            logo = slide.shapes.add_picture(logo_path, left, top, width=width)
            logging.info(f"Logo añadido correctamente a la diapositiva")
        else:
            logging.error(f"El archivo del logo no existe o está vacío: {logo_path}")
            
    except Exception as e:
        logging.error(f"Error al añadir logo: {e}")

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
    summary_box.shadow.inherit = False
    
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
    
    # Lista de noticias
    noticias_list = medio_data.get("noticias", [])
    if noticias_list:
        # Calcular cuántas noticias podemos mostrar por diapositiva (máx 4)
        max_noticias_por_slide = min(4, len(noticias_list))
        
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
        news_box.shadow.inherit = False
        
        tf = news_box.text_frame
        tf.word_wrap = True
        
        p = tf.add_paragraph()
        p.text = "📰 Noticias Destacadas"
        p.font.name = FUENTES['subtitulo']
        p.font.size = Pt(16)
        p.font.color.rgb = COLORES['secundario']
        p.space_after = Pt(10)
        
        # Añadir las primeras noticias
        for i in range(max_noticias_por_slide):
            if i >= len(noticias_list):
                break
                
            noticia = noticias_list[i]
            
            p = tf.add_paragraph()
            p.text = f"📅 {noticia.get('fecha', 'N/A')} - {noticia.get('titulo', 'N/A')}"
            p.font.name = FUENTES['cuerpo']
            p.font.size = Pt(12)
            p.font.color.rgb = COLORES['texto_oscuro']
            p.font.bold = True
            p.space_before = Pt(5)
            p.space_after = Pt(2)
            
            # Párrafo con hipervínculo
            p = tf.add_paragraph()
            run = p.add_run()
            run.text = f"     {noticia.get('titular', 'N/A')}"
            run.font.name = FUENTES['cuerpo']
            run.font.size = Pt(11)
            run.font.color.rgb = COLORES['secundario']
            run.font.italic = True
            
            # Añadir hipervínculo al titular
            if 'url' in noticia and noticia['url']:
                hlink = run.hyperlink
                hlink.address = noticia['url']
            elif 'link' in noticia and noticia['link']:
                hlink = run.hyperlink
                hlink.address = noticia['link']
            
            p.space_after = Pt(8)
        
        # Si hay más noticias, crear una nueva diapositiva
        if len(noticias_list) > max_noticias_por_slide:
            slide_continuacion = pr.slides.add_slide(pr.slide_layouts[5])
            aplicar_estilo_slide(slide_continuacion)
            
            # Barra de título
            title_box = slide_continuacion.shapes.add_shape(
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
            title = slide_continuacion.shapes.add_textbox(
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
            
            # Caja de noticias
            news_box = slide_continuacion.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                Inches(1),
                Inches(1.5),
                Inches(8),
                Inches(5)
            )
            news_box.fill.solid()
            news_box.fill.fore_color.rgb = COLORES['blanco']
            news_box.line.color.rgb = COLORES['gris_claro']
            news_box.shadow.inherit = False
            
            tf = news_box.text_frame
            tf.word_wrap = True
            
            p = tf.add_paragraph()
            p.text = "📰 Noticias Destacadas (Continuación)"
            p.font.name = FUENTES['subtitulo']
            p.font.size = Pt(16)
            p.font.color.rgb = COLORES['secundario']
            p.space_after = Pt(10)
            
            # Añadir las noticias restantes
            for i in range(max_noticias_por_slide, len(noticias_list)):
                noticia = noticias_list[i]
                
                p = tf.add_paragraph()
                p.text = f"📅 {noticia.get('fecha', 'N/A')} - {noticia.get('titulo', 'N/A')}"
                p.font.name = FUENTES['cuerpo']
                p.font.size = Pt(12)
                p.font.color.rgb = COLORES['texto_oscuro']
                p.font.bold = True
                p.space_before = Pt(5)
                p.space_after = Pt(2)
                
                # Párrafo con hipervínculo
                p = tf.add_paragraph()
                run = p.add_run()
                run.text = f"     {noticia.get('titular', 'N/A')}"
                run.font.name = FUENTES['cuerpo']
                run.font.size = Pt(11)
                run.font.color.rgb = COLORES['secundario']
                run.font.italic = True
                
                # Añadir hipervínculo al titular
                if 'url' in noticia and noticia['url']:
                    hlink = run.hyperlink
                    hlink.address = noticia['url']
                elif 'link' in noticia and noticia['link']:
                    hlink = run.hyperlink
                    hlink.address = noticia['link']
                
                p.space_after = Pt(8)
            
            # Añadir logo a la diapositiva de continuación
            agregar_logo(slide_continuacion)
            add_footer(slide_continuacion, f"Cobertura {tipo_medio} - Informe de Medios")
    
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
        
        # Primero descargamos la imagen del gráfico
        img_path = None
        try:
            # Usar requests para descargar la imagen
            temp_dir = tempfile.gettempdir()
            img_path = os.path.join(temp_dir, f"grafico_{hash(url)}.png")
            
            response = requests.get(url, stream=True, timeout=15)
            if response.status_code == 200:
                with open(img_path, 'wb') as f:
                    for chunk in response.iter_content(chunk_size=8192):
                        f.write(chunk)
                logging.info(f"Gráfico descargado correctamente: {url}")
            else:
                logging.error(f"Error al descargar gráfico: {response.status_code}")
                img_path = None
        except Exception as e:
            logging.error(f"Error descargando gráfico {url}: {e}")
            img_path = None
        
        # Área de contenido principal (centrado en la diapositiva)
        content_area_top = Inches(1.5)
        content_area_height = Inches(5)
        content_area_left = Inches(0.5)
        content_area_width = Inches(9)
        
        # Añadir marco decorativo para el gráfico
        chart_frame = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            content_area_left,
            content_area_top,
            content_area_width,
            content_area_height
        )
        chart_frame.fill.solid()
        chart_frame.fill.fore_color.rgb = COLORES['blanco']
        chart_frame.line.color.rgb = COLORES['secundario']
        chart_frame.line.width = Pt(2)
        chart_frame.shadow.inherit = False
        
        # Intentar insertar el gráfico si se descargó correctamente
        if img_path and os.path.exists(img_path) and os.path.getsize(img_path) > 0:
            try:
                # Obtener dimensiones de la imagen
                img = Image.open(img_path)
                img_width, img_height = img.size
                aspect_ratio = img_width / img_height
                
                # Calcular tamaño manteniendo proporción
                # Dejamos margen de 0.5 pulgadas a cada lado
                max_width = content_area_width - Inches(1)
                max_height = content_area_height - Inches(0.5)
                
                # Calcular dimensiones para mantener proporción y centrar
                if aspect_ratio > 1:  # Imagen más ancha que alta
                    target_width = max_width
                    target_height = target_width / aspect_ratio
                    if target_height > max_height:
                        target_height = max_height
                        target_width = target_height * aspect_ratio
                else:  # Imagen más alta que ancha
                    target_height = max_height
                    target_width = target_height * aspect_ratio
                    if target_width > max_width:
                        target_width = max_width
                        target_height = target_width / aspect_ratio
                
                # Calcular posición para centrar perfectamente
                left = content_area_left + (content_area_width - target_width) / 2
                top = content_area_top + (content_area_height - target_height) / 2
                
                # Insertar imagen del gráfico
                pic = slide.shapes.add_picture(
                    img_path,
                    left,
                    top,
                    width=target_width,
                    height=target_height
                )
                
                # Asegurarse de que el gráfico esté en primer plano
                pic.z_order = -1  # Poner en primer plano
                
            except Exception as e:
                logging.error(f"Error al insertar gráfico {url}: {e}")
                
                # Mostrar mensaje de error
                error_box = slide.shapes.add_textbox(
                    content_area_left + Inches(1.5),
                    content_area_top + Inches(2),
                    Inches(6),
                    Inches(1)
                )
                tf = error_box.text_frame
                tf.text = "Error al cargar el gráfico"
                p = tf.paragraphs[0]
                p.font.color.rgb = COLORES['acento']
                p.font.size = Pt(16)
                p.font.bold = True
                p.alignment = PP_ALIGN.CENTER
            finally:
                # Limpiar archivo temporal
                try:
                    if os.path.exists(img_path):
                        os.remove(img_path)
                except:
                    pass
        else:
            # Mostrar mensaje de error si no se pudo descargar
            error_box = slide.shapes.add_textbox(
                content_area_left + Inches(1.5),
                content_area_top + Inches(2),
                Inches(6),
                Inches(1)
            )
            tf = error_box.text_frame
            tf.text = "Error al descargar el gráfico"
            p = tf.paragraphs[0]
            p.font.color.rgb = COLORES['acento']
            p.font.size = Pt(16)
            p.font.bold = True
            p.alignment = PP_ALIGN.CENTER
        
        # Añadir logo
        agregar_logo(slide)
        add_footer(slide, f"{tipo_grafico} - Informe de Medios")

def crear_vpe_totales(pr, datos):
    slide = pr.slides.add_slide(pr.slide_layouts[5])
    aplicar_estilo_slide(slide)
    
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
    tf.text = "Valor Publicitario Equivalente (VPE) Total"
    p = tf.paragraphs[0]
    p.font.name = FUENTES['titulo']
    p.font.size = Pt(28)
    p.font.color.rgb = COLORES['blanco']
    p.alignment = PP_ALIGN.CENTER
    
    # Caja de datos VPE
    vpe_box = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(1),
        Inches(2),
        Inches(8),
        Inches(3.5)
    )
    vpe_box.fill.solid()
    vpe_box.fill.fore_color.rgb = COLORES['blanco']
    vpe_box.line.color.rgb = COLORES['secundario']
    vpe_box.shadow.inherit = False
    
    tf = vpe_box.text_frame
    tf.word_wrap = True
    
    p = tf.add_paragraph()
    p.text = "Resumen de Valor Publicitario"
    p.font.name = FUENTES['subtitulo']
    p.font.size = Pt(20)
    p.font.color.rgb = COLORES['secundario']
    p.alignment = PP_ALIGN.CENTER
    p.space_after = Pt(20)
    
    p = tf.add_paragraph()
    p.text = f"VPE Total: {datos.get('totalGlobalVPE', 'N/A')}"
    p.font.name = FUENTES['cuerpo']
    p.font.size = Pt(28)
    p.font.color.rgb = COLORES['principal']
    p.font.bold = True
    p.alignment = PP_ALIGN.CENTER
    
    # Añadir logo
    agregar_logo(slide)
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
