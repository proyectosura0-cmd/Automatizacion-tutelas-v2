from sqlalchemy import Column, Integer, String, DateTime, Boolean, Text
from sqlalchemy.sql import func
from database import Base


class Caso(Base):
    __tablename__ = "casos"

    id = Column(Integer, primary_key=True, index=True)
    fecha_registro = Column(DateTime(timezone=True), server_default=func.now())
    numero_rs = Column(String, index=True)
    tipo_modelo = Column(String)
    juzgado = Column(Text)
    ciudad_departamento = Column(String)
    accionante = Column(String, index=True)
    tipo_documento = Column(String)
    numero_documento = Column(String)
    radicado = Column(String, index=True)
    regimen_afiliacion = Column(String)
    calidad = Column(String)
    edad = Column(String)
    sexo = Column(String)
    representante_legal = Column(String)
    cedula_representante = Column(String)
    regional = Column(String)
    diagnostico = Column(String)
    servicio_solicitado = Column(Text)
    estado_servicio = Column(String)
    numero_autorizacion = Column(String)
    fecha_autorizacion = Column(String)
    prestador = Column(Text)
    nit_prestador = Column(String)
    fecha_cita_entrega = Column(String)
    pretension_principal = Column(Text)
    resumen_hechos = Column(Text)
    solicitud_al_juez = Column(String)
    observaciones = Column(Text)
    nombre_archivo_generado = Column(String)
    ruta_archivo = Column(String)
    estado_caso = Column(String, default="Guardado")


class AthentoMotivoCausal(Base):
    """Causales asignados manualmente a tutelas de Athento por los abogados."""
    __tablename__ = "athento_causales"

    id            = Column(Integer, primary_key=True, index=True)
    uuid_athento  = Column(String, unique=True, index=True)  # UUID del doc en Athento
    expediente    = Column(String, index=True)
    modelo        = Column(String)        # nombre del modelo: "Transporte", "Medicamento Importado"...
    subcausal     = Column(String)        # texto libre opcional
    asignado_por  = Column(String)
    fecha_asignado = Column(DateTime(timezone=True), server_default=func.now())
    notas         = Column(Text, default="")


class Jurisprudencia(Base):
    __tablename__ = "jurisprudencia"

    id            = Column(Integer, primary_key=True, index=True)
    tipo          = Column(String, index=True)      # sentencia | ley | decreto | circular | concepto
    numero        = Column(String, index=True)      # T-760/2008 | Ley 1751/2015
    tribunal      = Column(String)                  # Corte Constitucional | Consejo de Estado
    fecha         = Column(String)                  # YYYY-MM-DD o YYYY
    titulo        = Column(Text)                    # Nombre descriptivo
    resumen       = Column(Text)                    # Resumen del alcance
    cita_clave    = Column(Text)                    # La cita textual más útil para contestaciones
    modelos       = Column(String)                  # JSON: ["Transporte","No PBS"] — modelos que aplica
    compania      = Column(String, default="EPS")   # EPS | ARL | ambas
    activa        = Column(Boolean, default=True)
    agregado_por  = Column(String, default="Sistema")
    fecha_creado  = Column(DateTime(timezone=True), server_default=func.now())
    fecha_updated = Column(DateTime(timezone=True), onupdate=func.now())


class JuzgadoRiesgo(Base):
    """Scoring de riesgo por juzgado — SURA tiende a ganar/perder según juzgado."""
    __tablename__ = "juzgado_riesgo"

    id             = Column(Integer, primary_key=True, index=True)
    juzgado        = Column(String, unique=True, index=True)  # ej: "JUZGADO 014 CIVIL MUNICIPAL DE CALI"
    score          = Column(Integer, default=2)  # 1=favorable, 2=neutral, 3=desfavorable
    notas          = Column(Text, default="")
    actualizado_por = Column(String, default="Sistema")
    fecha_actualizacion = Column(DateTime(timezone=True), server_default=func.now(), onupdate=func.now())


class ModeloRiesgo(Base):
    """Scoring de riesgo por modelo de tutela."""
    __tablename__ = "modelo_riesgo"

    id             = Column(Integer, primary_key=True, index=True)
    modelo         = Column(String, unique=True, index=True)  # ej: "Medicamento Importado"
    score          = Column(Integer, default=2)  # 1=bajo riesgo, 2=medio, 3=alto
    notas          = Column(Text, default="")
    actualizado_por = Column(String, default="Sistema")
    fecha_actualizacion = Column(DateTime(timezone=True), server_default=func.now(), onupdate=func.now())


class ConceptoPreestablecido(Base):
    """Conceptos técnicos pre-establecidos por modelo de tutela."""
    __tablename__ = "conceptos_preestablecidos"

    id             = Column(Integer, primary_key=True, index=True)
    modelo         = Column(String, index=True)  # ej: "Medicamento Importado"
    nombre         = Column(String)              # ej: "Medicamento no disponible"
    variante       = Column(String, default="")  # ej: "Por agotamiento", "Por falta de registro"
    texto_concepto = Column(Text)               # El texto completo del concepto
    orden          = Column(Integer, default=0) # Para ordenar múltiples variantes
    activo         = Column(Boolean, default=True)
    fecha_creado   = Column(DateTime(timezone=True), server_default=func.now())
    fecha_actualizado = Column(DateTime(timezone=True), onupdate=func.now())
