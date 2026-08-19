"""
Extrae texto de los PDFs de conocimiento y siembra la tabla jurisprudencia.
Ejecutar una sola vez: python seed_jurisprudencia.py
"""
import sys, os, json, re
sys.stdout.reconfigure(encoding="utf-8")
os.chdir(os.path.dirname(os.path.abspath(__file__)))

import pdfplumber
from database import SessionLocal, engine
import models

# Crear tabla si no existe
models.Base.metadata.create_all(bind=engine)

CONOCIMIENTO = r"C:\Users\alvar\OneDrive\Escritorio\INSTRUCIONES IA\CONOCIMIENTO"

# Mapeo: fragmento del nombre de archivo → metadata manual
META = {
    "Ley 1751": {
        "tipo": "ley", "numero": "Ley 1751/2015", "tribunal": "Congreso de la República",
        "fecha": "2015-02-16",
        "titulo": "Ley Estatutaria de Salud — Derecho fundamental a la salud",
        "modelos": ["Transporte", "No PBS / Presupuestos Maximos", "Tratamiento Integral",
                    "Acceso a Citas", "Medicamento Importado", "No INVIMA", "Silla de Ruedas",
                    "Cuidador Primario"],
        "compania": "EPS",
        "cita_clave": "Artículo 6: El derecho fundamental a la salud incluye el acceso a los servicios de salud de manera oportuna, eficaz y con calidad para la preservación, el mejoramiento y la promoción de la salud. Artículo 15: El sistema de salud garantizará el suministro de tecnologías en salud que no estén financiadas con recursos públicos asignados a la salud cuando el médico tratante las prescriba como necesarias.",
    },
    "Ley_100": {
        "tipo": "ley", "numero": "Ley 100/1993", "tribunal": "Congreso de la República",
        "fecha": "1993-12-23",
        "titulo": "Sistema de Seguridad Social Integral",
        "modelos": ["Transporte", "No PBS / Presupuestos Maximos", "Acceso a Citas"],
        "compania": "EPS",
        "cita_clave": "Artículo 153: El sistema garantizará a todos los habitantes del territorio nacional, la protección integral de las contingencias que menoscaben la salud y la capacidad económica de los habitantes del territorio nacional.",
    },
    "Ley_1955": {
        "tipo": "ley", "numero": "Ley 1955/2019", "tribunal": "Congreso de la República",
        "fecha": "2019-05-25",
        "titulo": "Plan Nacional de Desarrollo 2018-2022 — Presupuestos Máximos en Salud",
        "modelos": ["No PBS / Presupuestos Maximos"],
        "compania": "EPS",
        "cita_clave": "Artículo 240: Los servicios y tecnologías en salud no financiados con cargo a la UPC serán sometidos a la aplicación de un presupuesto máximo definido por el Ministerio de Salud, cuya fuente de financiación serán los recursos de la ADRES.",
    },
    "Decreto_780": {
        "tipo": "decreto", "numero": "Decreto 780/2016", "tribunal": "Ministerio de Salud",
        "fecha": "2016-05-06",
        "titulo": "Decreto Único Reglamentario del Sector Salud y Protección Social",
        "modelos": ["Transporte", "No PBS / Presupuestos Maximos", "Acceso a Citas", "Medicamento Importado"],
        "compania": "EPS",
        "cita_clave": "Regula el sistema general de seguridad social en salud, incluyendo afiliación, prestación de servicios, red de prestadores y mecanismos de financiación del sistema.",
    },
    "Decreto No. 1652": {
        "tipo": "decreto", "numero": "Decreto 1652/2022", "tribunal": "Ministerio de Salud",
        "fecha": "2022-08-09",
        "titulo": "Transferencia de tecnologías en salud no financiadas con recursos de la UPC",
        "modelos": ["No PBS / Presupuestos Maximos", "Medicamento Importado", "No INVIMA"],
        "compania": "EPS",
        "cita_clave": "Regula los mecanismos de transferencia, reconocimiento y pago de tecnologías en salud no incluidas en el plan de beneficios, incluyendo procedimientos de recobro y gestión de presupuestos máximos.",
    },
    "Resolucion No 695": {
        "tipo": "resolucion", "numero": "Resolución 695/2026", "tribunal": "Ministerio de Salud",
        "fecha": "2026-01-01",
        "titulo": "Actualización del Plan de Beneficios en Salud con cargo a la UPC 2026",
        "modelos": ["No PBS / Presupuestos Maximos", "Medicamento Importado", "Acceso a Citas"],
        "compania": "EPS",
        "cita_clave": "Define las tecnologías en salud cubiertas con cargo a la UPC para la vigencia 2026, actualizando inclusiones y exclusiones del plan de beneficios.",
    },
    "Resolucion No. 243": {
        "tipo": "resolucion", "numero": "Resolución 243/2019", "tribunal": "Ministerio de Salud",
        "fecha": "2019-02-14",
        "titulo": "Lineamientos para el reconocimiento y pago de tecnologías no PBS",
        "modelos": ["No PBS / Presupuestos Maximos"],
        "compania": "EPS",
        "cita_clave": "Establece el procedimiento para la solicitud, autorización y reconocimiento económico de tecnologías en salud no cubiertas por el plan de beneficios.",
    },
    "Resolución No 067 de 2025": {
        "tipo": "resolucion", "numero": "Resolución 067/2025", "tribunal": "Ministerio de Salud",
        "fecha": "2025-01-01",
        "titulo": "Criterios de exclusión del Plan de Beneficios en Salud 2025",
        "modelos": ["No PBS / Presupuestos Maximos", "Medicamento Importado", "No INVIMA"],
        "compania": "EPS",
        "cita_clave": "Actualiza los criterios y listado de tecnologías expresamente excluidas del plan de beneficios en salud para la vigencia 2025.",
    },
    "Resolución No 2641": {
        "tipo": "resolucion", "numero": "Resolución 2641/2024", "tribunal": "Ministerio de Salud",
        "fecha": "2024-06-01",
        "titulo": "Actualización de servicios complementarios y transporte de pacientes",
        "modelos": ["Transporte"],
        "compania": "EPS",
        "cita_clave": "Define las condiciones y requisitos para la autorización del servicio de transporte de pacientes como parte de los servicios complementarios del sistema de salud.",
    },
    "Resolucion No. 738": {
        "tipo": "resolucion", "numero": "Resolución 738/2019", "tribunal": "Ministerio de Salud",
        "fecha": "2019-04-01",
        "titulo": "Condiciones para el transporte de pacientes en el SGSSS",
        "modelos": ["Transporte"],
        "compania": "EPS",
        "cita_clave": "El servicio de transporte se reconoce cuando el usuario requiere desplazarse para acceder a servicios de salud que no están disponibles en su municipio de residencia y la EPS no tiene red disponible.",
    },
    "CIRCULAR_Y_RESOLUCION_NO_PBS": {
        "tipo": "circular", "numero": "Circular PBS Antioquia", "tribunal": "Secretaría de Salud de Antioquia",
        "fecha": "2024-01-01",
        "titulo": "Circular y Resolución No PBS — Departamento de Antioquia",
        "modelos": ["No PBS / Presupuestos Maximos"],
        "compania": "EPS",
        "cita_clave": "Regula el proceso departamental para la gestión de tecnologías no PBS en el departamento de Antioquia, incluyendo flujos de autorización y reconocimiento.",
    },
    "circular-048-2025": {
        "tipo": "circular", "numero": "Circular 048/2025", "tribunal": "Superintendencia Nacional de Salud",
        "fecha": "2025-01-01",
        "titulo": "Instrucciones sobre tiempos de respuesta en autorizaciones de servicios de salud",
        "modelos": ["Acceso a Citas", "No PBS / Presupuestos Maximos", "Tratamiento Integral"],
        "compania": "EPS",
        "cita_clave": "Las EPS deben garantizar la oportunidad en la asignación de citas y en el trámite de autorizaciones de servicios de salud, conforme a los plazos establecidos en la normativa vigente.",
    },
    "circular-externa-0003-de-2025": {
        "tipo": "circular", "numero": "Circular Externa 003/2025", "tribunal": "Ministerio de Salud",
        "fecha": "2025-01-01",
        "titulo": "Lineamientos para la atención de acciones de tutela en salud 2025",
        "modelos": ["Transporte", "No PBS / Presupuestos Maximos", "Acceso a Citas", "Tratamiento Integral",
                    "Medicamento Importado"],
        "compania": "EPS",
        "cita_clave": "Imparte instrucciones a las EPS sobre el manejo, contestación y cumplimiento de fallos de tutela en materia de salud, estableciendo responsabilidades institucionales.",
    },
    "circular-externa-0046-de-2025": {
        "tipo": "circular", "numero": "Circular Externa 046/2025", "tribunal": "Superintendencia Nacional de Salud",
        "fecha": "2025-01-01",
        "titulo": "Seguimiento al cumplimiento de fallos de tutela en salud",
        "modelos": ["Carencia Actual de Objeto", "Tratamiento Integral"],
        "compania": "EPS",
        "cita_clave": "Establece mecanismos de supervisión y seguimiento al cumplimiento efectivo de fallos de tutela por parte de las EPS, con énfasis en casos de medidas provisionales.",
    },
    "resolucion-0037-de-2025": {
        "tipo": "resolucion", "numero": "Resolución 037/2025", "tribunal": "Ministerio de Salud",
        "fecha": "2025-01-01",
        "titulo": "Actualización de tecnologías en salud PBS — vigencia 2025",
        "modelos": ["No PBS / Presupuestos Maximos", "Medicamento Importado"],
        "compania": "EPS",
        "cita_clave": "Actualiza el listado de tecnologías incluidas en el plan de beneficios, determinando qué servicios son responsabilidad directa de la EPS con cargo a la UPC.",
    },
    "resolucion-0067-de-2025": {
        "tipo": "resolucion", "numero": "Resolución 067/2025", "tribunal": "Ministerio de Salud",
        "fecha": "2025-01-01",
        "titulo": "Criterios técnicos de exclusión de tecnologías del PBS",
        "modelos": ["No PBS / Presupuestos Maximos", "No INVIMA"],
        "compania": "EPS",
        "cita_clave": "Define los criterios bajo los cuales una tecnología en salud puede ser excluida del plan de beneficios, incluyendo ausencia de evidencia científica y tecnologías en fase experimental.",
    },
    "resolucion-0757-de-2025": {
        "tipo": "resolucion", "numero": "Resolución 757/2025", "tribunal": "Ministerio de Salud",
        "fecha": "2025-03-01",
        "titulo": "Tarifas SOAT y transferencia de recursos para tecnologías no PBS 2025",
        "modelos": ["No PBS / Presupuestos Maximos"],
        "compania": "EPS",
        "cita_clave": "Establece las tarifas de referencia y mecanismos de transferencia de recursos para el reconocimiento de tecnologías en salud no financiadas con cargo a la UPC en 2025.",
    },
    "resolucion-1905-de-2025": {
        "tipo": "resolucion", "numero": "Resolución 1905/2025", "tribunal": "Ministerio de Salud",
        "fecha": "2025-05-01",
        "titulo": "Procedimientos para acceso a medicamentos de alto costo",
        "modelos": ["Medicamento Importado", "No PBS / Presupuestos Maximos", "No INVIMA"],
        "compania": "EPS",
        "cita_clave": "Regula los procedimientos para la solicitud, autorización y suministro de medicamentos de alto costo, incluyendo requisitos de prescripción y tiempos máximos de respuesta.",
    },
    "resolucion-2605-de-2025": {
        "tipo": "resolucion", "numero": "Resolución 2605/2025", "tribunal": "Ministerio de Salud",
        "fecha": "2025-07-01",
        "titulo": "Actualización lineamientos técnicos enfermedades huérfanas",
        "modelos": ["Medicamento Importado", "No INVIMA", "Tratamiento Integral"],
        "compania": "EPS",
        "cita_clave": "Establece los lineamientos para la atención integral de pacientes con enfermedades huérfanas o raras, incluyendo acceso a medicamentos sin registro INVIMA y tratamientos especializados.",
    },
    "resolucion-2764-2025": {
        "tipo": "resolucion", "numero": "Resolución 2764/2025", "tribunal": "Ministerio de Salud",
        "fecha": "2025-07-15",
        "titulo": "Régimen de habilitación de dispositivos médicos y tecnologías de apoyo",
        "modelos": ["Silla de Ruedas", "No PBS / Presupuestos Maximos"],
        "compania": "EPS",
        "cita_clave": "Define los requisitos para la habilitación y suministro de dispositivos médicos y tecnologías de apoyo, incluyendo ayudas técnicas como sillas de ruedas y equipos especializados.",
    },
    "resolucion-3514-de-2019": {
        "tipo": "resolucion", "numero": "Resolución 3514/2019", "tribunal": "Ministerio de Salud",
        "fecha": "2019-08-01",
        "titulo": "Procedimientos para medicamentos sin registro INVIMA (importación)",
        "modelos": ["Medicamento Importado", "No INVIMA"],
        "compania": "EPS",
        "cita_clave": "Establece el procedimiento para la autorización de importación y uso de medicamentos que no cuentan con registro sanitario INVIMA vigente, con fines terapéuticos individuales.",
    },
    "lineamientos-tecnicos-atencion-integral-pacientes-hemofilia": {
        "tipo": "lineamiento", "numero": "Lineamiento Hemofilia", "tribunal": "Ministerio de Salud",
        "fecha": "2024-01-01",
        "titulo": "Lineamientos técnicos para atención integral de pacientes con Hemofilia",
        "modelos": ["Tratamiento Integral", "Medicamento Importado", "No PBS / Presupuestos Maximos"],
        "compania": "EPS",
        "cita_clave": "La atención integral del paciente con hemofilia incluye el suministro continuo de factor de coagulación, seguimiento especializado y cobertura de complicaciones, sin que la EPS pueda interrumpir el tratamiento.",
    },
    "lineamientos-tecnicos-atencion-pacientes-enfermedad-gaucher": {
        "tipo": "lineamiento", "numero": "Lineamiento Gaucher", "tribunal": "Ministerio de Salud",
        "fecha": "2024-01-01",
        "titulo": "Lineamientos técnicos para atención de pacientes con Enfermedad de Gaucher",
        "modelos": ["Tratamiento Integral", "Medicamento Importado"],
        "compania": "EPS",
        "cita_clave": "La enfermedad de Gaucher requiere terapia de reemplazo enzimático continua. La EPS debe garantizar el acceso oportuno al medicamento prescrito por el médico tratante especialista.",
    },
    "lineamientos-tecnicos-atencion-integral-pacientes-enfermedad-fabry": {
        "tipo": "lineamiento", "numero": "Lineamiento Fabry", "tribunal": "Ministerio de Salud",
        "fecha": "2024-01-01",
        "titulo": "Lineamientos técnicos para atención integral de pacientes con Enfermedad de Fabry",
        "modelos": ["Tratamiento Integral", "Medicamento Importado", "No INVIMA"],
        "compania": "EPS",
        "cita_clave": "La enfermedad de Fabry requiere terapia de reemplazo enzimático. El médico tratante especialista es quien determina la molécula indicada y la EPS debe suministrarla oportunamente.",
    },
    "lineamientos-tecnicos-atencion-personas-fibrosis-quistica": {
        "tipo": "lineamiento", "numero": "Lineamiento Fibrosis Quística", "tribunal": "Ministerio de Salud",
        "fecha": "2024-01-01",
        "titulo": "Lineamientos técnicos para atención de personas con Fibrosis Quística",
        "modelos": ["Tratamiento Integral", "Medicamento Importado"],
        "compania": "EPS",
        "cita_clave": "La atención integral en fibrosis quística incluye manejo multidisciplinario, acceso a moduladores de CFTR cuando estén indicados, y soporte nutricional continuo, sin interrupciones.",
    },
    "lineamientos-tecnicos-atencion-pacientes-atrofia-muscular-espinal": {
        "tipo": "lineamiento", "numero": "Lineamiento AME 5q", "tribunal": "Ministerio de Salud",
        "fecha": "2024-01-01",
        "titulo": "Lineamientos técnicos para atención de pacientes con Atrofia Muscular Espinal (AME 5q)",
        "modelos": ["Tratamiento Integral", "Medicamento Importado", "No PBS / Presupuestos Maximos"],
        "compania": "EPS",
        "cita_clave": "El tratamiento con nusinersen, risdiplam o onasemnogene abeparvovec debe iniciarse y mantenerse de acuerdo con el criterio del médico tratante especialista en neurología pediátrica.",
    },
    "lineamientos-tecnicos-atencion-personas-hemoglobinuria": {
        "tipo": "lineamiento", "numero": "Lineamiento HPN", "tribunal": "Ministerio de Salud",
        "fecha": "2024-01-01",
        "titulo": "Lineamientos técnicos para personas con Hemoglobinuria Paroxística Nocturna (HPN)",
        "modelos": ["Tratamiento Integral", "Medicamento Importado"],
        "compania": "EPS",
        "cita_clave": "La HPN requiere tratamiento con inhibidores del complemento de manera continua. La EPS debe garantizar el acceso oportuno y sin interrupciones al medicamento prescrito.",
    },
    "posicionamiento-terapeutico": {
        "tipo": "lineamiento", "numero": "Posicionamiento Cáncer Próstata", "tribunal": "Ministerio de Salud",
        "fecha": "2024-01-01",
        "titulo": "Posicionamiento terapéutico — Tratamiento farmacológico cáncer de próstata",
        "modelos": ["No PBS / Presupuestos Maximos", "Medicamento Importado", "Tratamiento Integral"],
        "compania": "EPS",
        "cita_clave": "El posicionamiento terapéutico define el algoritmo de tratamiento y las líneas de medicamentos indicados para el manejo del cáncer de próstata, siendo guía para las EPS en la cobertura de dichos tratamientos.",
    },
    "via-clinica-atencion-pacientes-fibrosis-quistica": {
        "tipo": "lineamiento", "numero": "Vía Clínica Fibrosis Quística", "tribunal": "Ministerio de Salud",
        "fecha": "2024-01-01",
        "titulo": "Vía clínica para atención de pacientes con Fibrosis Quística",
        "modelos": ["Tratamiento Integral", "Acceso a Citas"],
        "compania": "EPS",
        "cita_clave": "Define la ruta de atención integral para pacientes con fibrosis quística, incluyendo frecuencia de controles, especialidades requeridas y criterios de hospitalización.",
    },
}


def _clasificar_archivo(nombre_archivo):
    """Busca la metadata correcta para un archivo dado."""
    nombre = nombre_archivo.replace(".pdf", "")
    for clave, meta in META.items():
        if clave.lower() in nombre.lower() or nombre.lower() in clave.lower():
            return meta
    return None


def _extraer_texto(ruta_pdf, max_paginas=4):
    """Extrae texto de las primeras N páginas del PDF."""
    try:
        with pdfplumber.open(ruta_pdf) as pdf:
            textos = []
            for i, page in enumerate(pdf.pages[:max_paginas]):
                t = page.extract_text()
                if t:
                    textos.append(t.strip())
            return "\n".join(textos)[:3000]
    except Exception as e:
        return f"[No se pudo extraer texto: {e}]"


def _resumen_desde_texto(texto, titulo):
    """Extrae un resumen del texto extraído."""
    if not texto or texto.startswith("["):
        return titulo
    # Tomar las primeras líneas con contenido
    lineas = [l.strip() for l in texto.split("\n") if len(l.strip()) > 40]
    return " ".join(lineas[:4])[:600]


def sembrar():
    db = SessionLocal()
    try:
        existentes = db.query(models.Jurisprudencia).count()
        if existentes > 0:
            print(f"Ya existen {existentes} registros. Limpiando para re-sembrar...")
            db.query(models.Jurisprudencia).delete()
            db.commit()

        archivos = [f for f in os.listdir(CONOCIMIENTO) if f.endswith(".pdf")]
        sembrados = 0
        sin_meta = []

        for archivo in sorted(archivos):
            ruta = os.path.join(CONOCIMIENTO, archivo)
            meta = _clasificar_archivo(archivo)

            if not meta:
                sin_meta.append(archivo)
                # Crear entrada genérica
                texto = _extraer_texto(ruta, max_paginas=2)
                nombre = archivo.replace(".pdf", "").replace("-", " ").replace("_", " ").title()
                meta = {
                    "tipo": "documento",
                    "numero": nombre[:60],
                    "tribunal": "Ministerio de Salud",
                    "fecha": "2024-01-01",
                    "titulo": nombre,
                    "modelos": [],
                    "compania": "EPS",
                    "cita_clave": _resumen_desde_texto(texto, nombre),
                }

            texto_pdf = _extraer_texto(ruta)
            resumen = _resumen_desde_texto(texto_pdf, meta["titulo"])

            j = models.Jurisprudencia(
                tipo         = meta["tipo"],
                numero       = meta["numero"],
                tribunal     = meta["tribunal"],
                fecha        = meta.get("fecha", ""),
                titulo       = meta["titulo"],
                resumen      = resumen,
                cita_clave   = meta.get("cita_clave", resumen[:400]),
                modelos      = json.dumps(meta.get("modelos", []), ensure_ascii=False),
                compania     = meta.get("compania", "EPS"),
                activa       = True,
                agregado_por = "Sistema — carga inicial",
            )
            db.add(j)
            sembrados += 1
            print(f"  ✓ {meta['numero'][:50]:<52} [{meta['tipo']}]")

        db.commit()
        print(f"\n✅ {sembrados} documentos sembrados en la base de datos.")
        if sin_meta:
            print(f"⚠  Sin metadata manual ({len(sin_meta)}): {', '.join(sin_meta)}")
    finally:
        db.close()


if __name__ == "__main__":
    sembrar()
