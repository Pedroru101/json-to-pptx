from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from PIL import Image
import os
import logging
from app.utils import download_image

# Configuración de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Configuración de estilos
COLOR_PRINCIPAL = RGBColor(0x00, 0x56, 0x91)
COLOR_SECUNDARIO = RGBColor(0xEE, 0xEE, 0xEE)
COLOR_TEXTO_OSCURO = RGBColor(0x33, 0x33, 0x33)
COLOR_ACENTO = RGBColor(0x00, 0x8A, 0xD7)

FUENTE_TITULO = 'Arial'
FUENTE_CUERPO = 'Arial'

def add_footer(slide, text_content):
    footer = slide.shapes.add_textbox(Inches(0.5), Inches(6.9), Inches(9), Inches(0.3))
    tf = footer.text_frame
    tf.text = text_content
    p = tf.paragraphs[0]
    p.font.size = Pt(10)
    p.font.name = FUENTE_CUERPO
    p.font.color.rgb = COLOR_TEXTO_OSCURO

def crear_portada(pr, datos):
    slide = pr.slides.add_slide(pr.slide_layouts[0])
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    
    title.text = "Informe de Medios"
    subtitle.text = f"Período: {datos.get('fechaInicial', 'N/A')} - {datos.get('fechaFinal', 'N/A')}"
    
    for shape in [title, subtitle]:
        shape.text_frame.paragraphs[0].font.name = FUENTE_TITULO
        shape.text_frame.paragraphs[0].font.color.rgb = COLOR_PRINCIPAL
    
    title.text_frame.paragraphs[0].font.size = Pt(44)
    subtitle.text_frame.paragraphs[0].font.size = Pt(24)
    
    add_footer(slide, "Portada - Informe de Medios")

def crear_metodologia(pr):
    slide = pr.slides.add_slide(pr.slide_layouts[1])
    title = slide.shapes.title
    content = slide.placeholders[1]
    
    title.text = "Metodología"
    title.text_frame.paragraphs[0].font.name = FUENTE_TITULO
    title.text_frame.paragraphs[0].font.color.rgb = COLOR_PRINCIPAL
    
    metodologia_text = [
        "• Monitoreo continuo de medios",
        "• Análisis cualitativo y cuantitativo",
        "• Métricas de evaluación:",
        "   - Valor Publicitario Equivalente (VPE)",
        "   - Alcance y audiencia",
        "   - Sentimiento y tono",
        "   - Presencia en diferentes soportes"
    ]
    
    content.text = "\n".join(metodologia_text)
    for p in content.text_frame.paragraphs:
        p.font.name = FUENTE_CUERPO
        p.font.size = Pt(18)
    
    add_footer(slide, "Metodología - Informe de Medios")

def crear_datos_cobertura(pr, datos, tipo_medio):
    slide = pr.slides.add_slide(pr.slide_layouts[2])
    title = slide.shapes.title
    
    title.text = f"Datos de Cobertura - {tipo_medio}"
    title.text_frame.paragraphs[0].font.name = FUENTE_TITULO
    title.text_frame.paragraphs[0].font.color.rgb = COLOR_PRINCIPAL
    
    medio_data = datos.get(f"{tipo_medio}_raw", {})
    if not medio_data:
        return
    
    # Crear caja de datos principales
    left = Inches(0.8)
    top = Inches(1.5)
    width = Inches(8.2)
    height = Inches(1.5)
    
    txBox = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    txBox.fill.solid()
    txBox.fill.fore_color.rgb = COLOR_SECUNDARIO
    txBox.line.color.rgb = COLOR_PRINCIPAL
    
    tf = txBox.text_frame
    tf.word_wrap = True
    
    # Añadir datos específicos del medio con formato mejorado
    p = tf.add_paragraph()
    p.text = f"📊 Resumen de Impacto en {tipo_medio}"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.name = FUENTE_CUERPO
    
    p = tf.add_paragraph()
    p.text = f"Total de Noticias: {medio_data.get('cantidad_noticias', 'N/A')} | Alcance: {medio_data.get('total_audiencia', 'N/A')}"
    p.font.size = Pt(14)
    p.font.name = FUENTE_CUERPO
    
    p = tf.add_paragraph()
    p.text = f"VPE: {medio_data.get('total_vpe', 'N/A')} | Valor Cualitativo: {medio_data.get('total_vc', 'N/A')}"
    p.font.size = Pt(14)
    p.font.name = FUENTE_CUERPO
    
    # Lista de noticias
    noticias_list = medio_data.get("noticias", [])
    if noticias_list:
        # Crear caja para noticias
        left = Inches(0.8)
        top = Inches(3.2)
        width = Inches(8.2)
        height = Inches(3.5)
        
        newsBox = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
        newsBox.fill.solid()
        newsBox.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        newsBox.line.color.rgb = COLOR_PRINCIPAL
        
        tf = newsBox.text_frame
        tf.word_wrap = True
        
        p = tf.add_paragraph()
        p.text = "📰 Noticias Destacadas"
        p.font.size = Pt(16)
        p.font.bold = True
        p.font.name = FUENTE_CUERPO
        
        items_per_slide = 6
        current_items = 0
        
        for noticia in noticias_list:
            if current_items >= items_per_slide:
                # Crear nueva diapositiva para más noticias
                slide = pr.slides.add_slide(pr.slide_layouts[2])
                title = slide.shapes.title
                title.text = f"Datos de Cobertura - {tipo_medio} (Continuación)"
                title.text_frame.paragraphs[0].font.name = FUENTE_TITULO
                title.text_frame.paragraphs[0].font.color.rgb = COLOR_PRINCIPAL
                
                newsBox = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, Inches(1.5), width, height)
                newsBox.fill.solid()
                newsBox.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                newsBox.line.color.rgb = COLOR_PRINCIPAL
                
                tf = newsBox.text_frame
                tf.word_wrap = True
                current_items = 0
            
            p = tf.add_paragraph()
            p.text = f"📅 {noticia.get('fecha', 'N/A')} - {noticia.get('titulo', 'N/A')}"
            p.font.size = Pt(12)
            p.font.name = FUENTE_CUERPO
            
            p = tf.add_paragraph()
            p.text = f"    {noticia.get('titular', 'N/A')}"
            p.font.size = Pt(11)
            p.font.name = FUENTE_CUERPO
            p.font.italic = True
            
            current_items += 1
    
    add_footer(slide, f"Cobertura {tipo_medio} - Informe de Medios")

def crear_vpe_totales(pr, datos):
    slide = pr.slides.add_slide(pr.slide_layouts[2])
    title = slide.shapes.title
    
    title.text = "Valor Publicitario Equivalente (VPE) Total"
    title.text_frame.paragraphs[0].font.name = FUENTE_TITULO
    title.text_frame.paragraphs[0].font.color.rgb = COLOR_PRINCIPAL
    
    left = Inches(0.8)
    top = Inches(1.8)
    width = Inches(8.2)
    height = Inches(4)
    
    txBox = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    txBox.fill.solid()
    txBox.fill.fore_color.rgb = COLOR_SECUNDARIO
    
    tf = txBox.text_frame
    p = tf.add_paragraph()
    p.text = f"VPE Total: {datos.get('totalGlobalVPE', 'N/A')}"
    p.font.size = Pt(20)
    p.font.name = FUENTE_CUERPO
    
    add_footer(slide, "VPE Total - Informe de Medios")

def crear_analisis_rrss(pr, datos):
    slide = pr.slides.add_slide(pr.slide_layouts[2])
    title = slide.shapes.title
    
    title.text = "Análisis de Redes Sociales"
    title.text_frame.paragraphs[0].font.name = FUENTE_TITULO
    title.text_frame.paragraphs[0].font.color.rgb = COLOR_PRINCIPAL
    
    # Aquí se añadirían los datos específicos de RRSS cuando estén disponibles
    rrss_data = datos.get('RRSS_raw', {})
    if not rrss_data:
        return
    
    add_footer(slide, "Análisis RRSS - Informe de Medios")

def crear_analisis_offline(pr, datos):
    slide = pr.slides.add_slide(pr.slide_layouts[2])
    title = slide.shapes.title
    
    title.text = "Análisis de Elementos Offline"
    title.text_frame.paragraphs[0].font.name = FUENTE_TITULO
    title.text_frame.paragraphs[0].font.color.rgb = COLOR_PRINCIPAL
    
    # Aquí se añadirían los datos específicos de elementos offline
    offline_data = datos.get('Offline_raw', {})
    if not offline_data:
        return
    
    add_footer(slide, "Análisis Offline - Informe de Medios")

def crear_otros_elementos(pr, datos):
    slide = pr.slides.add_slide(pr.slide_layouts[2])
    title = slide.shapes.title
    
    title.text = "Análisis de Otros Elementos"
    title.text_frame.paragraphs[0].font.name = FUENTE_TITULO
    title.text_frame.paragraphs[0].font.color.rgb = COLOR_PRINCIPAL
    
    # Aquí se añadirían los datos de web, anuncios, etc.
    otros_data = datos.get('Otros_raw', {})
    if not otros_data:
        return
    
    add_footer(slide, "Otros Elementos - Informe de Medios")

def crear_tabla_sumatoria(pr, datos):
    slide = pr.slides.add_slide(pr.slide_layouts[2])
    title = slide.shapes.title
    
    title.text = "Resumen Final - VPE y Audiencias"
    title.text_frame.paragraphs[0].font.name = FUENTE_TITULO
    title.text_frame.paragraphs[0].font.color.rgb = COLOR_PRINCIPAL
    
    left = Inches(0.8)
    top = Inches(1.8)
    width = Inches(8.2)
    height = Inches(4)
    
    txBox = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    txBox.fill.solid()
    txBox.fill.fore_color.rgb = COLOR_SECUNDARIO
    
    tf = txBox.text_frame
    tf.clear()
    
    p = tf.add_paragraph()
    p.text = "Resumen Global"
    p.font.size = Pt(20)
    p.font.bold = True
    
    p = tf.add_paragraph()
    p.text = f"Total Noticias: {datos.get('totalGlobalNoticias', 'N/A')}"
    p.font.size = Pt(16)
    
    p = tf.add_paragraph()
    p.text = f"Audiencia Total: {datos.get('totalGlobalAudiencia', 'N/A')}"
    p.font.size = Pt(16)
    
    p = tf.add_paragraph()
    p.text = f"VPE Total: {datos.get('totalGlobalVPE', 'N/A')}"
    p.font.size = Pt(16)
    
    add_footer(slide, "Resumen Final - Informe de Medios")

def crear_roi(pr, datos):
    slide = pr.slides.add_slide(pr.slide_layouts[2])
    title = slide.shapes.title
    
    title.text = "Análisis ROI"
    title.text_frame.paragraphs[0].font.name = FUENTE_TITULO
    title.text_frame.paragraphs[0].font.color.rgb = COLOR_PRINCIPAL
    
    # Aquí se implementaría la fórmula de ROI específica
    roi_data = datos.get('ROI_raw', {})
    if not roi_data:
        return
    
    add_footer(slide, "Análisis ROI - Informe de Medios")

def crear_graficos(pr, datos):
    urls = datos.get("urls", [])
    if not urls:
        return
    
    for url in urls:
        slide = pr.slides.add_slide(pr.slide_layouts[6])  # Blank layout
        
        # Determinar el tipo de gráfico basado en la URL
        tipo_grafico = ""
        if "vpe_barra" in url:
            tipo_grafico = "VPE por Medio (Gráfico de Barras)"
        elif "vpe_torta" in url:
            tipo_grafico = "Distribución de VPE (Gráfico Circular)"
        elif "impactos_barra" in url:
            tipo_grafico = "Impactos por Medio (Gráfico de Barras)"
        elif "impactos_torta" in url:
            tipo_grafico = "Distribución de Impactos (Gráfico Circular)"
        elif "top10" in url:
            medio = url.split("top10_vpe_")[1].split(".")[0].replace("_", " ").title()
            tipo_grafico = f"Top 10 VPE - {medio}"
        
        # Añadir título al gráfico
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.75))
        tf = title_box.text_frame
        tf.text = tipo_grafico
        tf.paragraphs[0].font.name = FUENTE_TITULO
        tf.paragraphs[0].font.size = Pt(24)
        tf.paragraphs[0].font.color.rgb = COLOR_PRINCIPAL
        
        # Descargar e insertar imagen
        img_path = download_image(url)
        if img_path:
            try:
                left = Inches(1)
                top = Inches(1.5)
                width = Inches(8)
                height = Inches(5)
                slide.shapes.add_picture(img_path, left, top, width=width, height=height)
            except Exception as e:
                logging.error(f"Error al procesar gráfico {url}: {e}")
            finally:
                if os.path.exists(img_path):
                    try:
                        os.remove(img_path)
                    except OSError as e:
                        logging.error(f"Error al eliminar imagen temporal {img_path}: {e}")
        
        add_footer(slide, f"{tipo_grafico} - Informe de Medios")

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
    crear_graficos(pr, datos)  # Añadir sección de gráficos
    crear_analisis_rrss(pr, datos)
    crear_analisis_offline(pr, datos)
    crear_otros_elementos(pr, datos)
    crear_tabla_sumatoria(pr, datos)
    crear_roi(pr, datos)
    
    # Guardar presentación
    output_path = f"/tmp/{filename}"
    try:
        pr.save(output_path)
        logging.info(f"Presentación PPTX guardada en: {output_path}")
    except Exception as e:
        logging.error(f"Error al guardar la presentación PPTX: {e}")
        raise
    
    return output_path
