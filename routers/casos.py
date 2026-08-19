import csv
import io
from fastapi import APIRouter, Depends, HTTPException
from fastapi.responses import StreamingResponse
from sqlalchemy.orm import Session
from pydantic import BaseModel
from typing import Optional
from database import get_db
from models import Caso

router = APIRouter(prefix="/api/casos", tags=["casos"])


class CasoSchema(BaseModel):
    numero_rs: Optional[str] = ""
    tipo_modelo: Optional[str] = ""
    juzgado: Optional[str] = ""
    ciudad_departamento: Optional[str] = ""
    accionante: Optional[str] = ""
    tipo_documento: Optional[str] = ""
    numero_documento: Optional[str] = ""
    radicado: Optional[str] = ""
    regimen_afiliacion: Optional[str] = ""
    calidad: Optional[str] = ""
    edad: Optional[str] = ""
    sexo: Optional[str] = ""
    representante_legal: Optional[str] = ""
    cedula_representante: Optional[str] = ""
    regional: Optional[str] = ""
    diagnostico: Optional[str] = ""
    servicio_solicitado: Optional[str] = ""
    estado_servicio: Optional[str] = ""
    numero_autorizacion: Optional[str] = ""
    fecha_autorizacion: Optional[str] = ""
    prestador: Optional[str] = ""
    nit_prestador: Optional[str] = ""
    fecha_cita_entrega: Optional[str] = ""
    pretension_principal: Optional[str] = ""
    resumen_hechos: Optional[str] = ""
    solicitud_al_juez: Optional[str] = ""
    observaciones: Optional[str] = ""
    nombre_archivo_generado: Optional[str] = ""
    ruta_archivo: Optional[str] = ""
    estado_caso: Optional[str] = "Guardado"

    class Config:
        from_attributes = True


def _serializar(caso: Caso) -> dict:
    return {
        "id": caso.id,
        "fecha_registro": caso.fecha_registro.isoformat() if caso.fecha_registro else "",
        "numero_rs": caso.numero_rs or "",
        "tipo_modelo": caso.tipo_modelo or "",
        "juzgado": caso.juzgado or "",
        "ciudad_departamento": caso.ciudad_departamento or "",
        "accionante": caso.accionante or "",
        "tipo_documento": caso.tipo_documento or "",
        "numero_documento": caso.numero_documento or "",
        "radicado": caso.radicado or "",
        "regimen_afiliacion": caso.regimen_afiliacion or "",
        "calidad": caso.calidad or "",
        "edad": caso.edad or "",
        "sexo": caso.sexo or "",
        "representante_legal": caso.representante_legal or "",
        "cedula_representante": caso.cedula_representante or "",
        "regional": caso.regional or "",
        "diagnostico": caso.diagnostico or "",
        "servicio_solicitado": caso.servicio_solicitado or "",
        "estado_servicio": caso.estado_servicio or "",
        "numero_autorizacion": caso.numero_autorizacion or "",
        "fecha_autorizacion": caso.fecha_autorizacion or "",
        "prestador": caso.prestador or "",
        "nit_prestador": caso.nit_prestador or "",
        "fecha_cita_entrega": caso.fecha_cita_entrega or "",
        "pretension_principal": caso.pretension_principal or "",
        "resumen_hechos": caso.resumen_hechos or "",
        "solicitud_al_juez": caso.solicitud_al_juez or "",
        "observaciones": caso.observaciones or "",
        "nombre_archivo_generado": caso.nombre_archivo_generado or "",
        "estado_caso": caso.estado_caso or "Guardado",
    }


@router.get("")
async def listar_casos(db: Session = Depends(get_db)):
    casos = db.query(Caso).order_by(Caso.id.desc()).all()
    return [_serializar(c) for c in casos]


@router.post("")
async def crear_caso(data: CasoSchema, db: Session = Depends(get_db)):
    caso = Caso(**data.model_dump())
    db.add(caso)
    db.commit()
    db.refresh(caso)
    return _serializar(caso)


@router.get("/exportar/csv")
async def exportar_csv(db: Session = Depends(get_db)):
    casos = db.query(Caso).order_by(Caso.id).all()
    columnas = [
        "id", "fecha_registro", "numero_rs", "tipo_modelo", "juzgado",
        "accionante", "radicado", "representante_legal", "servicio_solicitado",
        "estado_servicio", "solicitud_al_juez", "nombre_archivo_generado", "estado_caso",
    ]
    output = io.StringIO()
    writer = csv.DictWriter(output, fieldnames=columnas, extrasaction="ignore")
    writer.writeheader()
    for c in casos:
        fila = _serializar(c)
        fila["fecha_registro"] = fila["fecha_registro"][:10] if fila["fecha_registro"] else ""
        writer.writerow({k: fila.get(k, "") for k in columnas})

    output.seek(0)
    return StreamingResponse(
        iter(["﻿" + output.read()]),  # BOM para Excel
        media_type="text/csv; charset=utf-8",
        headers={"Content-Disposition": "attachment; filename=BaseDatos_Tutelas.csv"},
    )


@router.get("/{caso_id}")
async def obtener_caso(caso_id: int, db: Session = Depends(get_db)):
    caso = db.query(Caso).filter(Caso.id == caso_id).first()
    if not caso:
        raise HTTPException(status_code=404, detail="Caso no encontrado.")
    return _serializar(caso)


@router.delete("/{caso_id}")
async def eliminar_caso(caso_id: int, db: Session = Depends(get_db)):
    caso = db.query(Caso).filter(Caso.id == caso_id).first()
    if not caso:
        raise HTTPException(status_code=404, detail="Caso no encontrado.")
    db.delete(caso)
    db.commit()
    return {"mensaje": f"Caso {caso_id} eliminado."}
