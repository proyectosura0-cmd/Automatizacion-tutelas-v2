import json
from typing import Optional
from fastapi import APIRouter, Depends, HTTPException, Query
from sqlalchemy.orm import Session
from pydantic import BaseModel

from database import get_db
import models

router = APIRouter(prefix="/api/jurisprudencia", tags=["jurisprudencia"])

TIPOS = ["ley", "decreto", "resolucion", "circular", "lineamiento", "sentencia", "concepto", "documento"]


class JurisprudenciaIn(BaseModel):
    tipo:         str
    numero:       str
    tribunal:     str
    fecha:        Optional[str] = ""
    titulo:       str
    resumen:      Optional[str] = ""
    cita_clave:   Optional[str] = ""
    modelos:      list[str] = []
    compania:     str = "EPS"
    agregado_por: Optional[str] = "Abogado"


def _to_dict(j: models.Jurisprudencia) -> dict:
    return {
        "id":           j.id,
        "tipo":         j.tipo,
        "numero":       j.numero,
        "tribunal":     j.tribunal,
        "fecha":        j.fecha or "",
        "titulo":       j.titulo,
        "resumen":      j.resumen or "",
        "cita_clave":   j.cita_clave or "",
        "modelos":      json.loads(j.modelos) if j.modelos else [],
        "compania":     j.compania or "EPS",
        "activa":       j.activa,
        "agregado_por": j.agregado_por or "",
        "fecha_creado": j.fecha_creado.isoformat() if j.fecha_creado else "",
    }


# ─── LISTAR / BUSCAR ──────────────────────────────────────────────────────────
@router.get("")
def listar(
    q:        Optional[str] = Query(None, description="Texto libre"),
    modelo:   Optional[str] = Query(None, description="Filtrar por modelo de tutela"),
    tipo:     Optional[str] = Query(None, description="ley | decreto | resolucion | ..."),
    compania: str           = Query("EPS", description="EPS | ARL"),
    db: Session = Depends(get_db),
):
    rows = db.query(models.Jurisprudencia).filter(models.Jurisprudencia.activa == True)

    if compania:
        rows = rows.filter(
            (models.Jurisprudencia.compania == compania) |
            (models.Jurisprudencia.compania == "ambas")
        )
    if tipo:
        rows = rows.filter(models.Jurisprudencia.tipo == tipo)
    if q:
        q_lower = f"%{q.lower()}%"
        rows = rows.filter(
            models.Jurisprudencia.titulo.ilike(q_lower) |
            models.Jurisprudencia.numero.ilike(q_lower) |
            models.Jurisprudencia.resumen.ilike(q_lower) |
            models.Jurisprudencia.cita_clave.ilike(q_lower)
        )

    todos = [_to_dict(r) for r in rows.order_by(models.Jurisprudencia.fecha.desc()).all()]

    # Filtro por modelo (JSON array en columna de texto)
    if modelo:
        todos = [j for j in todos if modelo in j["modelos"]]

    return todos


# ─── POR MODELO (para el widget del formulario) ───────────────────────────────
@router.get("/por-modelo/{modelo}")
def por_modelo(
    modelo:   str,
    compania: str = Query("EPS"),
    db: Session = Depends(get_db),
):
    """Devuelve jurisprudencia relevante para un modelo de tutela específico."""
    rows = db.query(models.Jurisprudencia).filter(
        models.Jurisprudencia.activa == True,
        (models.Jurisprudencia.compania == compania) |
        (models.Jurisprudencia.compania == "ambas"),
    ).order_by(models.Jurisprudencia.fecha.desc()).all()

    resultado = [_to_dict(r) for r in rows if modelo in (json.loads(r.modelos) if r.modelos else [])]
    return resultado


# ─── DETALLE ──────────────────────────────────────────────────────────────────
@router.get("/{id}")
def detalle(id: int, db: Session = Depends(get_db)):
    j = db.query(models.Jurisprudencia).filter(models.Jurisprudencia.id == id).first()
    if not j:
        raise HTTPException(status_code=404, detail="No encontrado")
    return _to_dict(j)


# ─── CREAR ────────────────────────────────────────────────────────────────────
@router.post("", status_code=201)
def crear(body: JurisprudenciaIn, db: Session = Depends(get_db)):
    j = models.Jurisprudencia(
        tipo         = body.tipo,
        numero       = body.numero,
        tribunal     = body.tribunal,
        fecha        = body.fecha,
        titulo       = body.titulo,
        resumen      = body.resumen,
        cita_clave   = body.cita_clave,
        modelos      = json.dumps(body.modelos, ensure_ascii=False),
        compania     = body.compania,
        activa       = True,
        agregado_por = body.agregado_por,
    )
    db.add(j)
    db.commit()
    db.refresh(j)
    return _to_dict(j)


# ─── EDITAR ───────────────────────────────────────────────────────────────────
@router.put("/{id}")
def editar(id: int, body: JurisprudenciaIn, db: Session = Depends(get_db)):
    j = db.query(models.Jurisprudencia).filter(models.Jurisprudencia.id == id).first()
    if not j:
        raise HTTPException(status_code=404, detail="No encontrado")
    j.tipo       = body.tipo
    j.numero     = body.numero
    j.tribunal   = body.tribunal
    j.fecha      = body.fecha
    j.titulo     = body.titulo
    j.resumen    = body.resumen
    j.cita_clave = body.cita_clave
    j.modelos    = json.dumps(body.modelos, ensure_ascii=False)
    j.compania   = body.compania
    db.commit()
    return _to_dict(j)


# ─── ARCHIVAR (soft delete) ───────────────────────────────────────────────────
@router.delete("/{id}")
def archivar(id: int, db: Session = Depends(get_db)):
    j = db.query(models.Jurisprudencia).filter(models.Jurisprudencia.id == id).first()
    if not j:
        raise HTTPException(status_code=404, detail="No encontrado")
    j.activa = False
    db.commit()
    return {"ok": True}
