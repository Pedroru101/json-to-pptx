import requests
from uuid import uuid4
import logging # Añadir esta importación para logging

# Configurar logging (opcional, pero ayuda a ver los mensajes en los logs de Render)
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def download_image(url):
    try:
        # Añadir un timeout para evitar que las solicitudes se queden colgadas indefinidamente
        response = requests.get(url, timeout=15) # Puedes ajustar el valor del timeout
        if response.status_code == 200:
            path = f"/tmp/{uuid4()}.png"
            with open(path, "wb") as f:
                f.write(response.content)
            logging.info(f"Imagen descargada exitosamente: {url} a {path}")
            return path
        else:
            # Registrar el código de estado si no es 200
            logging.error(f"Error al descargar imagen {url}: Status {response.status_code}. Contenido: {response.text[:200]}")
    except requests.exceptions.Timeout:
        logging.error(f"Timeout al descargar imagen desde: {url}")
    except requests.exceptions.ConnectionError as e:
        logging.error(f"Error de conexión al descargar imagen desde {url}: {e}")
    except requests.exceptions.RequestException as e:
        # Captura otros errores de solicitud de requests (ej. invalid URL)
        logging.error(f"Error de solicitud al descargar imagen {url}: {e}")
    except Exception as e:
        # Captura cualquier otro error inesperado
        logging.error(f"Error inesperado en download_image para {url}: {e}")
    return None
