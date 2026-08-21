import os
from pathlib import Path
from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
import uvicorn

from database import engine, Base
import models  # Registrar modelos en Base
from routers import tutelas, casos, memoriales, athento, jurisprudencia

BASE_DIR = Path(__file__).resolve().parent

# Crear tablas en la BD
Base.metadata.create_all(bind=engine)

# Crear carpetas necesarias
(BASE_DIR / "plantillas").mkdir(exist_ok=True)
(BASE_DIR / "contestaciones").mkdir(exist_ok=True)
(BASE_DIR / "memoriales").mkdir(exist_ok=True)
(BASE_DIR / "static").mkdir(exist_ok=True)

app = FastAPI(
    title="Automatizador de Tutelas — EPS SURA",
    description="API para generación automática de contestaciones de acciones de tutela.",
    version="1.0.1",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(tutelas.router)
app.include_router(casos.router)
app.include_router(memoriales.router)
app.include_router(athento.router)
app.include_router(jurisprudencia.router)

app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")


@app.get("/", include_in_schema=False)
async def root():
    return FileResponse(str(BASE_DIR / "static" / "index.html"))


@app.get("/health")
async def health():
    return {"status": "ok", "app": "Automatizador de Tutelas EPS SURA"}


if __name__ == "__main__":
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
