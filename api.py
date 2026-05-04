from fastapi import FastAPI
from BagheeraBrain import generar_contrato_automatico

app = FastAPI()

@app.get("/")
def root():
    return {"status": "Bagheera API activa 🐱"}

@app.post("/contrato")
def generar(data: dict):

    archivo = generar_contrato_automatico(
        data["nombre"],
        data["fecha_inicio"],
        data["fecha_termino"]
    )

    return {
        "status": "ok",
        "archivo": archivo
    }
