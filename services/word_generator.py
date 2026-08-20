"""
Generador de contestaciones de tutela.

Estrategia de reemplazo:
- Abre el .docx como ZIP y procesa cada archivo XML directamente.
- Para cada <w:p>…</w:p> concatena el texto de todos sus <w:t> hijos,
  detecta marcadores [VARIABLE] aunque estén fragmentados en varios runs,
  hace el reemplazo y pone el texto resultante en el primer <w:t>,
  vaciando los demás. Esto conserva el formato del run donde empieza
  el marcador y resuelve la fragmentación de Word.
"""

import re
import zipfile
from datetime import date
from io import BytesIO
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
PLANTILLAS_DIR = BASE_DIR / "plantillas"
CONTESTACIONES_DIR = BASE_DIR / "contestaciones"

NOMBRE_PLANTILLA = {
    "Modelo Principal":              "MODELO_PRINCIPAL.docx",
    "Transporte":                    "Transporte.docx",
    "Medicamento Importado":         "Medicamento_Importado.docx",
    "No INVIMA":                     "No_INVIMA_2025.docx",
    "No PBS / Presupuestos Maximos": "NO_PBS_Presup_Max.docx",
    "Silla de Ruedas":               "Silla_de_Ruedas.docx",
    "Carencia Actual de Objeto":     "Carencia_Actual_de_Objeto.docx",
}

# ─────────────────────────────────────────────
# Helpers XML
# ─────────────────────────────────────────────

def _escape_xml(s: str) -> str:
    return (s
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;"))


def _procesar_xml(xml_str: str, mapa: dict) -> tuple[str, int]:
    """
    Recorre cada <w:p> del XML, concatena el texto de sus <w:t>,
    reemplaza marcadores y reconstruye poniendo todo en el primer <w:t>.
    """
    total = 0

    def procesar_parrafo(m: re.Match) -> str:
        nonlocal total
        p_xml = m.group(0)

        # Extraer todo el texto de los <w:t>...</w:t> dentro del párrafo
        t_matches = list(re.finditer(r'<w:t(?:\s[^>]*)?>([^<]*)</w:t>', p_xml))
        if not t_matches:
            return p_xml

        texto_completo = "".join(tm.group(1) for tm in t_matches)

        # Verificar si hay algún marcador presente (puede estar fragmentado)
        if not any(k in texto_completo for k in mapa):
            return p_xml

        # Realizar los reemplazos
        texto_final = texto_completo
        n = 0
        for marcador, valor in mapa.items():
            if marcador in texto_final:
                cnt = texto_final.count(marcador)
                repl = _escape_xml(str(valor)) if valor else _escape_xml(marcador)
                texto_final = texto_final.replace(marcador, repl)
                n += cnt

        if n == 0:
            return p_xml

        total += n

        # Poner todo el texto en el primer <w:t>, vaciar el resto
        first = [True]

        def replace_wt(mt: re.Match) -> str:
            if first[0]:
                first[0] = False
                return f'<w:t xml:space="preserve">{texto_final}</w:t>'
            return "<w:t></w:t>"

        return re.sub(r'<w:t(?:\s[^>]*)?>([^<]*)</w:t>', replace_wt, p_xml)

    xml_str = re.sub(
        r'<w:p(?:\s[^>]*)?>.*?</w:p>',
        procesar_parrafo,
        xml_str,
        flags=re.DOTALL,
    )

    return xml_str, total


def _insertar_advertencia(xml_str: str) -> str:
    """Inserta párrafo de borrador en rojo y negrita al inicio del body."""
    txt = _escape_xml(
        "*** BORRADOR — PARA REVISIÓN DEL ABOGADO RESPONSABLE — NO RADICAR ***"
    )
    parrafo = (
        "<w:p>"
        "<w:pPr><w:jc w:val=\"center\"/></w:pPr>"
        "<w:r>"
        "<w:rPr>"
        "<w:b/>"
        "<w:color w:val=\"FF0000\"/>"
        "<w:sz w:val=\"20\"/><w:szCs w:val=\"20\"/>"
        "</w:rPr>"
        f"<w:t xml:space=\"preserve\">{txt}</w:t>"
        "</w:r>"
        "</w:p>"
    )
    return xml_str.replace("<w:body>", f"<w:body>{parrafo}", 1)


# ─────────────────────────────────────────────
# Mapa de marcadores
# ─────────────────────────────────────────────

def _construir_mapa(datos: dict) -> dict:
    d = datos.get
    return {
        "[FECHA_CONTESTACION]":      d("fecha_contestacion", ""),
        "[NUMERO_RS]":               d("numero_rs", ""),
        "[JUZGADO]":                 d("juzgado", ""),
        "[CIUDAD_DEPARTAMENTO]":     d("ciudad_departamento", ""),
        "[ACCIONANTE]":              (d("accionante", "") or "").upper(),
        "[TIPO_DOCUMENTO]":          d("tipo_documento", ""),
        "[NUMERO_DOCUMENTO]":        d("numero_documento", ""),
        "[RADICADO]":                d("radicado", ""),
        "[REGIMEN_AFILIACION]":      d("regimen_afiliacion", ""),
        "[REPRESENTANTE_LEGAL]":     d("representante_legal", ""),
        "[CEDULA_REPRESENTANTE]":    d("cedula_representante", ""),
        "[REGIONAL]":                d("regional", ""),
        "[RESUMEN_HECHOS]":          d("resumen_hechos", ""),
        "[PRETENSIONES]":            d("pretension_principal", ""),
        "[SERVICIO_SOLICITADO]":     d("servicio_solicitado", ""),
        "[ESTADO_SERVICIO]":         d("estado_servicio", ""),
        "[NUMERO_AUTORIZACION]":     d("numero_autorizacion", ""),
        "[FECHA_AUTORIZACION]":      d("fecha_autorizacion", ""),
        "[PRESTADOR]":               d("prestador", ""),
        "[NIT_PRESTADOR]":           d("nit_prestador", ""),
        "[FECHA_CITA_ENTREGA]":      d("fecha_cita_entrega", ""),
        "[SOLICITUD_FINAL]":         d("solicitud_al_juez", ""),
        "[DIAGNOSTICO]":             d("diagnostico", ""),
        "[OBSERVACIONES]":           d("observaciones", ""),
        # Transporte
        "[ORIGEN]":                  d("origen", ""),
        "[DESTINO]":                 d("destino", ""),
        "[FRECUENCIA_TRANSPORTE]":   d("frecuencia_transporte", ""),
        "[TIPO_TRANSPORTE]":         d("tipo_transporte", ""),
        "[REQUIERE_ACOMPANANTE]":    d("requiere_acompanante", ""),
        "[SOLICITA_ALOJAMIENTO]":    d("solicita_alojamiento", ""),
        "[ENTIDAD_TERRITORIAL]":     d("entidad_territorial", ""),
        # Medicamento importado
        "[NOMBRE_MEDICAMENTO]":      d("nombre_medicamento", ""),
        "[ESTADO_IMPORTACION]":      d("estado_importacion", ""),
        "[PROVEEDOR_IMPORTACION]":   d("proveedor_importacion", ""),
        "[FECHA_ESTIMADA_ENTREGA]":  d("fecha_estimada_entrega", ""),
        # No INVIMA
        "[INDICACION_APROBADA]":     d("indicacion_aprobada", ""),
        "[ALTERNATIVA_MEDICA]":      d("alternativa_medica", ""),
        "[RAZON_NO_PROCEDENCIA]":    d("razon_no_procedencia", ""),
        # Silla de ruedas
        "[TIPO_SILLA]":              d("tipo_silla", ""),
        "[REQUIERE_COJIN]":          d("requiere_cojin", ""),
        "[ESTADO_COTIZACION]":       d("estado_cotizacion", ""),
        "[TIEMPO_FABRICACION]":      d("tiempo_fabricacion", ""),
        # Carencia actual de objeto
        "[HECHO_SOBREVINIENTE]":         d("hecho_sobreviniente", ""),
        "[FECHA_HECHO_SOBREVINIENTE]":   d("fecha_hecho_sobreviniente", ""),
        "[SOPORTE_DISPONIBLE]":          d("soporte_disponible", ""),
        # No PBS
        "[PRESUPUESTO_MAXIMO]":      d("presupuesto_maximo", ""),
        "[TECNOLOGIA_SOLICITADA]":   d("tecnologia_solicitada", ""),
        "[JUSTIFICACION_EXCLUSION]": d("justificacion_exclusion", ""),
    }


# ─────────────────────────────────────────────
# Función principal
# ─────────────────────────────────────────────

def generar_documento(datos: dict) -> tuple[str, str, int]:
    """
    Genera el .docx de contestación.
    Retorna (nombre_archivo, ruta_completa, total_reemplazos).
    """
    tipo_modelo = datos.get("tipo_modelo", "")
    # TEMPORAL: Usar Medicamento_Importado.docx para prueba
    if tipo_modelo == "Modelo Principal":
        nombre_plantilla = "Medicamento_Importado.docx"
    else:
        nombre_plantilla = NOMBRE_PLANTILLA.get(tipo_modelo)
    if not nombre_plantilla:
        raise ValueError(f"Tipo de modelo no reconocido: '{tipo_modelo}'")

    ruta_plantilla = PLANTILLAS_DIR / nombre_plantilla
    if not ruta_plantilla.exists():
        raise FileNotFoundError(
            f"Plantilla no encontrada: {ruta_plantilla}. "
            "Súbala desde la pestaña Plantillas de la interfaz."
        )

    mapa = _construir_mapa(datos)
    zip_in = BytesIO(ruta_plantilla.read_bytes())
    zip_out_buf = BytesIO()
    total_reemplazos = 0

    with zipfile.ZipFile(zip_in, "r") as zin:
        with zipfile.ZipFile(zip_out_buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)

                if item.filename.endswith(".xml"):
                    try:
                        xml_str = data.decode("utf-8")
                        xml_str, n = _procesar_xml(xml_str, mapa)
                        total_reemplazos += n
                        if item.filename == "word/document.xml":
                            xml_str = _insertar_advertencia(xml_str)
                        data = xml_str.encode("utf-8")
                    except Exception:
                        pass  # si falla el XML, conservar bytes originales

                zout.writestr(item, data)

    # Nombre del archivo de salida
    rs = re.sub(r"[^A-Za-z0-9]", "", datos.get("numero_rs", "SINUS"))
    accionante = re.sub(r"\s+", "_", (datos.get("accionante", "SIN") or "SIN").upper())[:20]
    accionante = re.sub(r"[^A-Za-z0-9_]", "", accionante)
    fecha_hoy = date.today().strftime("%Y%m%d")
    nombre_archivo = f"Contestacion_{rs}_{accionante}_{fecha_hoy}.docx"

    CONTESTACIONES_DIR.mkdir(parents=True, exist_ok=True)
    ruta_salida = CONTESTACIONES_DIR / nombre_archivo
    ruta_salida.write_bytes(zip_out_buf.getvalue())

    return nombre_archivo, str(ruta_salida), total_reemplazos


def escanear_marcadores(nombre_plantilla_archivo: str) -> list[str]:
    """
    Escanea un .docx y devuelve todos los marcadores [VARIABLE] encontrados.
    Útil para diagnosticar si la plantilla tiene los marcadores correctos.
    """
    ruta = PLANTILLAS_DIR / nombre_plantilla_archivo
    if not ruta.exists():
        raise FileNotFoundError(f"No se encontró: {ruta}")

    marcadores: set[str] = set()
    with zipfile.ZipFile(BytesIO(ruta.read_bytes()), "r") as z:
        for name in z.namelist():
            if name.endswith(".xml"):
                try:
                    xml_str = z.read(name).decode("utf-8")
                    # Concatenar texto de cada párrafo y buscar [VAR]
                    for m_p in re.finditer(r'<w:p(?:\s[^>]*)?>.*?</w:p>', xml_str, re.DOTALL):
                        texto = "".join(
                            tm.group(1)
                            for tm in re.finditer(r'<w:t(?:\s[^>]*)?>([^<]*)</w:t>', m_p.group(0))
                        )
                        for var in re.findall(r'\[[A-Z][A-Z0-9_]+\]', texto):
                            marcadores.add(var)
                except Exception:
                    pass

    return sorted(marcadores)
