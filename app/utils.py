import requests
from uuid import uuid4

def download_image(url):
    try:
        response = requests.get(url)
        if response.status_code == 200:
            path = f"/tmp/{uuid4()}.png"
            with open(path, "wb") as f:
                f.write(response.content)
            return path
    except:
        pass
    return None
