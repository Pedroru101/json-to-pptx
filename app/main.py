from fastapi import FastAPI, Request
from fastapi.responses import FileResponse
import uuid
from app.ppt_generator import generar_pptx

app = FastAPI()

@app.post("/generar-pptx")
async def generar_pptx_endpoint(request: Request):
    data = await request.json()
    nombre_archivo = f"reporte_{uuid.uuid4()}.pptx"
    ruta_pptx = generar_pptx(data, nombre_archivo)
    return FileResponse(ruta_pptx, media_type='application/vnd.openxmlformats-officedocument.presentationml.presentation', filename=nombre_archivo)
