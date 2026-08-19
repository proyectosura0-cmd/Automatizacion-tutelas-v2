"""Script de una sola ejecución: inserta marcadores [VARIABLE] en el docx
de memorial de cumplimiento y lo copia a plantillas/."""
import zipfile, re
from io import BytesIO
from pathlib import Path

SRC = Path(r"C:\Users\alvar\Downloads\memorial de cumpliento\Memorial_Cumplimiento_EPS_SURA (1).docx")
DST = Path("plantillas/Memorial_Cumplimiento.docx")

REEMPLAZOS = [
    ("JUEZ (XX) DE (XX) MUNICIPAL DE (XX)",          "[JUZGADO]"),
    ("NOMBRE DEL ACCIONANTE",                         "[ACCIONANTE]"),
    ("20XX-XXX",                                      "[RADICADO]"),
    ("LUISA FERNANDA GIRALDO GÓMEZ",                 "[REPRESENTANTE_LEGAL]"),
    ("1.020.450.788",                                 "[NUMERO_CEDULA]"),
    ("315.478 del C. S. de la J.",                   "[TARJETA_PROFESIONAL]"),
    ("apoderada judicial de EPS SURA",               "[CARGO_REPRESENTANTE] de EPS SURA"),
    ("(XX) de (XX) de 20XX, se ordenó",              "[FECHA_ORDEN_JUDICIAL], se ordenó"),
    ("(i) _______________________________ (indicar orden); (ii)", "[DESCRIPCION_ORDEN]"),
    # Ítems de cumplimiento — se manejan con regex más abajo

    ("(XX) de (XX) de 20XX",                         "[FECHA_MEMORIAL]"),
    ("C.C. 1.020.450.788 de Bogotá D.C.",            "C.C. [CEDULA_REPRESENTANTE]"),
    ("T.P. 315.478 del C. S. de la J.",              "T.P. [TARJETA_PROFESIONAL]"),
    ("Apoderada Judicial",                            "[CARGO_REPRESENTANTE_FORMAL]"),
]


def _escape_xml(s: str) -> str:
    return (s.replace("&", "&amp;").replace("<", "&lt;")
             .replace(">", "&gt;").replace('"', "&quot;"))


def _procesar_xml(xml_str: str) -> str:
    def reemplazar_parrafo(m):
        p = m.group(0)
        t_matches = list(re.finditer(r"<w:t(?:\s[^>]*)?>([^<]*)</w:t>", p))
        if not t_matches:
            return p
        texto = "".join(tm.group(1) for tm in t_matches)
        nuevo = texto
        for orig, marcador in REEMPLAZOS:
            if orig in nuevo:
                nuevo = nuevo.replace(orig, marcador)
        # Reemplazar ítems de cumplimiento con regex (cantidad de guiones variable)
        nuevo = re.sub(r'^\(iii\)\s+_+\.$', '[CUMPLIMIENTO_3]', nuevo)
        nuevo = re.sub(r'^\(ii\)\s+_+\.$',  '[CUMPLIMIENTO_2]', nuevo)
        nuevo = re.sub(r'^\(i\)\s+_+\.$',   '[CUMPLIMIENTO_1]', nuevo)
        # Párrafo de continuación de la orden (línea suelta de guiones)
        nuevo = re.sub(r'^_+\.$', '[DESCRIPCION_ORDEN_CONT]', nuevo)

        if nuevo == texto:
            return p
        first = [True]
        def repl_wt(mt):
            if first[0]:
                first[0] = False
                return f'<w:t xml:space="preserve">{_escape_xml(nuevo)}</w:t>'
            return "<w:t></w:t>"
        return re.sub(r"<w:t(?:\s[^>]*)?>([^<]*)</w:t>", repl_wt, p)

    return re.sub(
        r"<w:p(?:\s[^>]*)?>.*?</w:p>",
        reemplazar_parrafo,
        xml_str,
        flags=re.DOTALL,
    )


zin_buf = BytesIO(SRC.read_bytes())
zout_buf = BytesIO()

with zipfile.ZipFile(zin_buf, "r") as zin:
    with zipfile.ZipFile(zout_buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.endswith(".xml"):
                try:
                    xml = data.decode("utf-8")
                    xml = _procesar_xml(xml)
                    data = xml.encode("utf-8")
                except Exception as e:
                    print(f"  [WARN] {item.filename}: {e}")
            zout.writestr(item, data)

DST.write_bytes(zout_buf.getvalue())
print(f"Plantilla guardada: {DST}")
