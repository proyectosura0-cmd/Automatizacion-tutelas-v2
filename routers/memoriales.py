from pathlib import Path
from fastapi import APIRouter, HTTPException
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import Optional
from services.memorial_generator import generar_memorial, MEMORIALES_DIR, TIPOS_MEMORIAL

router = APIRouter(prefix="/api/memoriales", tags=["memoriales"])


class DatosMemorial(BaseModel):
    # Tipo
    tipo_memorial: Optional[str] = ""
    # Cabecera común
    fecha_memorial: Optional[str] = ""
    numero_rs: Optional[str] = ""
    juzgado: Optional[str] = ""
    ciudad_departamento: Optional[str] = ""
    radicado: Optional[str] = ""
    accionante: Optional[str] = ""
    tipo_documento: Optional[str] = ""
    numero_documento: Optional[str] = ""
    # Representante
    representante_legal: Optional[str] = ""
    cedula_representante: Optional[str] = ""
    cargo_representante: Optional[str] = ""
    regional: Optional[str] = ""
    # Cumplimiento (plantilla real)
    tarjeta_profesional: Optional[str] = ""
    fecha_orden_judicial: Optional[str] = ""
    descripcion_orden: Optional[str] = ""
    cumplimiento_1: Optional[str] = ""
    cumplimiento_2: Optional[str] = ""
    cumplimiento_3: Optional[str] = ""
    # Cumplimiento (python-docx legacy)
    descripcion_cumplimiento: Optional[str] = ""
    fecha_cumplimiento_efectivo: Optional[str] = ""
    soportes_adjuntos: Optional[str] = ""
    # Impugnación
    fecha_fallo_impugnado: Optional[str] = ""
    tipo_fallo: Optional[str] = ""
    argumentos_impugnacion: Optional[str] = ""
    pretension_impugnacion: Optional[str] = ""
    # Incidente de desacato
    fecha_apertura_incidente: Optional[str] = ""
    descripcion_incidente: Optional[str] = ""
    acciones_realizadas: Optional[str] = ""
    justificacion_desacato: Optional[str] = ""
    # Común al final
    observaciones_memorial: Optional[str] = ""


@router.get("/tipos")
async def listar_tipos():
    return {"tipos": TIPOS_MEMORIAL}


@router.post("/generar")
async def generar(datos: DatosMemorial):
    try:
        nombre_archivo, ruta = generar_memorial(datos.model_dump())
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error al generar el memorial: {str(e)}")

    return {
        "archivo": nombre_archivo,
        "url_descarga": f"/api/memoriales/descargar/{nombre_archivo}",
    }


@router.get("/descargar/{nombre_archivo}")
async def descargar(nombre_archivo: str):
    if ".." in nombre_archivo or "/" in nombre_archivo or "\\" in nombre_archivo:
        raise HTTPException(status_code=400, detail="Nombre de archivo inválido.")
    ruta = MEMORIALES_DIR / nombre_archivo
    if not ruta.exists():
        raise HTTPException(status_code=404, detail="Archivo no encontrado.")
    return FileResponse(
        path=str(ruta),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=nombre_archivo,
    )
