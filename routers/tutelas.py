import os
from pathlib import Path
from fastapi import APIRouter, HTTPException, UploadFile, File
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import Optional
from services.word_generator import (
    generar_documento, escanear_marcadores,
    PLANTILLAS_DIR, CONTESTACIONES_DIR, NOMBRE_PLANTILLA, CONCEPTOS_DIR,
)

router = APIRouter(prefix="/api/tutelas", tags=["tutelas"])


class DatosTutela(BaseModel):
    # Campos del caso
    fecha_contestacion: Optional[str] = ""
    numero_rs: Optional[str] = ""
    concepto_nombre: Optional[str] = ""
    concepto_texto: Optional[str] = ""
    # Juzgado
    juzgado: Optional[str] = ""
    ciudad_departamento: Optional[str] = ""
    # Accionante
    accionante: Optional[str] = ""
    tipo_documento: Optional[str] = ""
    numero_documento: Optional[str] = ""
    radicado: Optional[str] = ""
    regimen_afiliacion: Optional[str] = ""
    calidad: Optional[str] = ""
    edad: Optional[str] = ""
    sexo: Optional[str] = ""
    # Representante
    representante_legal: Optional[str] = ""
    cedula_representante: Optional[str] = ""
    regional: Optional[str] = ""
    # Servicio
    diagnostico: Optional[str] = ""
    servicio_solicitado: Optional[str] = ""
    estado_servicio: Optional[str] = ""
    numero_autorizacion: Optional[str] = ""
    fecha_autorizacion: Optional[str] = ""
    prestador: Optional[str] = ""
    nit_prestador: Optional[str] = ""
    fecha_cita_entrega: Optional[str] = ""
    # Jurídico
    pretension_principal: Optional[str] = ""
    resumen_hechos: Optional[str] = ""
    solicitud_al_juez: Optional[str] = ""
    observaciones: Optional[str] = ""
    # Transporte
    origen: Optional[str] = ""
    destino: Optional[str] = ""
    frecuencia_transporte: Optional[str] = ""
    tipo_transporte: Optional[str] = ""
    requiere_acompanante: Optional[str] = ""
    solicita_alojamiento: Optional[str] = ""
    entidad_territorial: Optional[str] = ""
    # Medicamento importado
    nombre_medicamento: Optional[str] = ""
    estado_importacion: Optional[str] = ""
    proveedor_importacion: Optional[str] = ""
    fecha_estimada_entrega: Optional[str] = ""
    # No INVIMA
    indicacion_aprobada: Optional[str] = ""
    alternativa_medica: Optional[str] = ""
    razon_no_procedencia: Optional[str] = ""
    # Silla de ruedas
    tipo_silla: Optional[str] = ""
    requiere_cojin: Optional[str] = ""
    estado_cotizacion: Optional[str] = ""
    tiempo_fabricacion: Optional[str] = ""
    # Carencia actual de objeto
    hecho_sobreviniente: Optional[str] = ""
    fecha_hecho_sobreviniente: Optional[str] = ""
    soporte_disponible: Optional[str] = ""
    # No PBS
    presupuesto_maximo: Optional[str] = ""
    tecnologia_solicitada: Optional[str] = ""
    justificacion_exclusion: Optional[str] = ""


@router.post("/generar")
async def generar(datos: DatosTutela):
    try:
        nombre_archivo, ruta, total_reemplazos = generar_documento(datos.model_dump())
    except FileNotFoundError as e:
        raise HTTPException(status_code=404, detail=str(e))
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error al generar el documento: {str(e)}")

    return {
        "archivo": nombre_archivo,
        "url_descarga": f"/api/tutelas/descargar/{nombre_archivo}",
        "total_reemplazos": total_reemplazos,
    }


@router.get("/descargar/{nombre_archivo}")
async def descargar(nombre_archivo: str):
    # Validar que no haya path traversal
    if ".." in nombre_archivo or "/" in nombre_archivo or "\\" in nombre_archivo:
        raise HTTPException(status_code=400, detail="Nombre de archivo inválido.")
    ruta = CONTESTACIONES_DIR / nombre_archivo
    if not ruta.exists():
        raise HTTPException(status_code=404, detail="Archivo no encontrado.")

    # Detectar media_type basado en extensión
    if nombre_archivo.endswith(".odt"):
        media_type = "application/vnd.oasis.opendocument.text"
    else:
        media_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

    return FileResponse(
        path=str(ruta),
        media_type=media_type,
        filename=nombre_archivo,
    )


@router.get("/plantillas")
async def listar_plantillas():
    PLANTILLAS_DIR.mkdir(parents=True, exist_ok=True)
    archivos = [f.name for f in PLANTILLAS_DIR.glob("*.docx")]
    return {
        "plantillas": archivos,
        "modelos_requeridos": list(NOMBRE_PLANTILLA.values()),
        "modelos_disponibles": [
            modelo for modelo, nombre in NOMBRE_PLANTILLA.items()
            if (PLANTILLAS_DIR / nombre).exists()
        ],
    }


@router.post("/subir-plantilla")
async def subir_plantilla(archivo: UploadFile = File(...)):
    if not archivo.filename.endswith(".docx"):
        raise HTTPException(status_code=400, detail="Solo se aceptan archivos .docx")

    # Buscar si el nombre del archivo coincide con alguna plantilla canónica
    # (comparación insensible a mayúsculas, espacios y guiones bajos equivalentes)
    def _normalizar(s: str) -> str:
        return s.lower().replace(" ", "_").replace("-", "_").replace(".", "_")

    nombre_destino = archivo.filename
    nombre_norm = _normalizar(archivo.filename.replace(".docx", ""))
    for nombre_canonico in NOMBRE_PLANTILLA.values():
        if _normalizar(nombre_canonico.replace(".docx", "")) == nombre_norm:
            nombre_destino = nombre_canonico
            break

    PLANTILLAS_DIR.mkdir(parents=True, exist_ok=True)
    ruta_destino = PLANTILLAS_DIR / nombre_destino
    contenido = await archivo.read()

    if len(contenido) == 0:
        raise HTTPException(status_code=400, detail="El archivo está vacío.")

    with open(ruta_destino, "wb") as f:
        f.write(contenido)

    try:
        marcadores = escanear_marcadores(nombre_destino)
    except Exception:
        marcadores = []

    return {
        "mensaje": f"Plantilla '{nombre_destino}' cargada exitosamente.",
        "tamaño_bytes": len(contenido),
        "marcadores_encontrados": marcadores,
        "nombre_guardado": nombre_destino,
    }


@router.get("/escanear/{nombre_archivo}")
async def escanear_plantilla(nombre_archivo: str):
    """Devuelve todos los marcadores [VARIABLE] encontrados en una plantilla."""
    if ".." in nombre_archivo or "/" in nombre_archivo or "\\" in nombre_archivo:
        raise HTTPException(status_code=400, detail="Nombre inválido.")
    try:
        marcadores = escanear_marcadores(nombre_archivo)
    except FileNotFoundError as e:
        raise HTTPException(status_code=404, detail=str(e))
    return {"archivo": nombre_archivo, "marcadores": marcadores, "total": len(marcadores)}


@router.get("/conceptos")
async def listar_conceptos():
    """Lista todos los conceptos jurídicos disponibles."""
    CONCEPTOS_DIR.mkdir(parents=True, exist_ok=True)
    conceptos = [f.stem for f in CONCEPTOS_DIR.glob("*.docx")]
    return {
        "conceptos": sorted(conceptos),
        "total": len(conceptos),
    }
