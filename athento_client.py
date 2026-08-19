"""
Cliente Athento para EPS SURA.
Autenticación: token API (ATHENTO_TOKEN) o cookie de sesión (ATHENTO_SESSIONID).
El token se genera en Athento → perfil de usuario → API Token.
"""
import os
from datetime import date
from dotenv import load_dotenv
import requests
from bs4 import BeautifulSoup

load_dotenv()

BASE       = os.getenv("ATHENTO_BASE",       "https://sura.athento.com")
QUEUE_ID   = os.getenv("ATHENTO_QUEUE_ID",   "")
SESSIONID  = os.getenv("ATHENTO_SESSIONID",  "")
CSRFTOKEN  = os.getenv("ATHENTO_CSRFTOKEN",  "")
API_TOKEN  = os.getenv("ATHENTO_TOKEN",       "")

# Columnas exactas de la tabla content_queue (por orden)
_COLS = [
    "estado",
    "nro_expediente",
    "abogado",
    "compania",
    "fecha_expediente",
    "fecha_notificacion",
    "horas_otorgadas",
    "fecha_vencimiento",
    "medida_provisional",
    "regional",
    "radicado",
    "juzgado",
    "accionante",
    "usuario_tecnico_eps",
    "area_causal",
    "usuario_tecnico",
    "usuario_tecnico_arl",
    "fecha_contestacion",
    "tipo_respuesta",
]

from routers.athento import MAPA_SUBCAUSAL


def _session() -> requests.Session:
    s = requests.Session()
    if API_TOKEN:
        s.headers.update({"Authorization": f"Token {API_TOKEN}"})
    else:
        s.cookies.set("sessionid", SESSIONID, domain="sura.athento.com")
        s.cookies.set("csrftoken", CSRFTOKEN,  domain="sura.athento.com")
        s.headers.update({"X-CSRFToken": CSRFTOKEN})
    s.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Referer":    f"{BASE}/",
    })
    return s


def _clasificar(estado: str, area_causal: str) -> dict:
    """
    Determina si la tutela puede responderse sin concepto médico.
    Busca en MAPA_SUBCAUSAL el area_causal que viene del HTML de Athento.
    Formato: "EPS_SALUD_ACCESO A CITAS" o "EPS_SALUD_MEDICAMENTOS"
    """
    from routers.athento import MAPA_SUBCAUSAL

    estado_l = (estado or "").lower()
    sin_concepto = "estructurar" in estado_l

    # Buscar modelo en MAPA usando área_causal normalizada
    area_norm = (area_causal or "").lower().strip()
    modelo = MAPA_SUBCAUSAL.get(area_norm, "")

    # Si no encontró, intentar variantes
    if not modelo and area_norm:
        # Remover/reemplazar underscores con espacios
        area_alt = area_norm.replace("_", " ")
        modelo = MAPA_SUBCAUSAL.get(area_alt, "")
        # Probar quitando el prefijo "eps_salud_"
        if not modelo and "eps_salud" in area_norm:
            area_alt2 = area_norm.replace("eps_salud_", "").replace("eps_salud-", "")
            modelo = MAPA_SUBCAUSAL.get(area_alt2, "")

    return {
        "sin_concepto":    sin_concepto,
        "modelo_sugerido": modelo or "",
        "auto_generable":  bool(modelo),
    }


def _parse_queue_html(html: str) -> list[dict]:
    """Parsea la tabla HTML devuelta por /queues/api/queue/content_queue/."""
    soup = BeautifulSoup(html, "html.parser")
    tbl  = soup.find("table", class_="queues-document-list")
    if not tbl:
        return []

    rows = []
    for tr in tbl.find("tbody").find_all("tr") if tbl.find("tbody") else []:
        uuid  = tr.get("id", "")
        cells = tr.find_all("td")
        if not cells:
            continue

        row = {"_uuid": uuid}
        for i, td in enumerate(cells):
            key = _COLS[i] if i < len(_COLS) else f"col_{i}"
            row[key] = td.get_text(strip=True)

        # Descomponer nro_expediente: "EXP-T-26-000021933 - 2026-0351 - 1013460705"
        partes = [p.strip() for p in row.get("nro_expediente", "").split(" - ")]
        row["expediente"] = partes[0] if len(partes) > 0 else ""
        # radicado ya viene en la columna radicado; si no, tomar de partes[1]
        if not row.get("radicado") and len(partes) > 1:
            row["radicado"] = partes[1]
        row["cedula"] = partes[2] if len(partes) > 2 else ""

        rows.append(row)
    return rows


def _normalizar(row: dict) -> dict:
    estado    = row.get("estado", "")
    area_caus = row.get("area_causal", "")
    clasif    = _clasificar(estado, area_caus)

    juzgado   = " ".join(row.get("juzgado", "").split())

    return {
        "uuid":               row.get("_uuid", ""),
        "expediente":         row.get("expediente", ""),
        "radicado":           row.get("radicado", ""),
        "cedula":             row.get("cedula", ""),
        "accionante":         row.get("accionante", ""),
        "juzgado":            juzgado,
        "abogado":            row.get("abogado", ""),
        "compania":           row.get("compania", ""),
        "regional":           row.get("regional", ""),
        "estado":             estado,
        "medida_provisional": row.get("medida_provisional", ""),
        "horas_otorgadas":    row.get("horas_otorgadas", ""),
        "fecha_expediente":   row.get("fecha_expediente", ""),
        "fecha_notificacion": row.get("fecha_notificacion", ""),
        "fecha_vencimiento":  row.get("fecha_vencimiento", ""),
        "fecha_contestacion": row.get("fecha_contestacion", ""),
        "tipo_respuesta":     row.get("tipo_respuesta", ""),
        "area_causal":        area_caus,
        # clasificación
        "sin_concepto":       clasif["sin_concepto"],
        "modelo_sugerido":    clasif["modelo_sugerido"],
        "auto_generable":     clasif["auto_generable"],
    }


def _extraer_causal_documento(uuid: str) -> dict | None:
    """
    Usa Selenium para extraer "Motivo causal IA" y "Subcausal IA" de un documento en Athento.
    Retorna {"motivo_causal": "...", "subcausal_ia": "..."} o None si falla.
    """
    try:
        from selenium import webdriver
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        from webdriver_manager.chrome import ChromeDriverManager
        from selenium.webdriver.chrome.service import Service
        import time

        # Configurar Chrome en modo headless
        options = webdriver.ChromeOptions()
        options.add_argument("--headless")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--window-size=1920,1080")
        options.add_argument(f"user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64)")

        driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=options
        )

        # Navegar al sitio base primero para establecer cookies
        driver.get(f"{BASE}/")
        import time as _time
        _time.sleep(1)

        # Agregar cookies sin especificar dominio (Selenium lo infiere)
        driver.add_cookie({"name": "sessionid", "value": SESSIONID})
        driver.add_cookie({"name": "csrftoken", "value": CSRFTOKEN})

        # Navegar al documento
        url = f"{BASE}/file/view/{uuid}"
        driver.get(url)

        # Esperar a que cargue el formulario (máx 15 segundos)
        wait = WebDriverWait(driver, 15)

        # Buscar divs que contengan "Resumen Jurídico IA"
        try:
            wait.until(EC.presence_of_all_elements_located((By.TAG_NAME, "label")))
        except:
            pass

        motivo_causal = None
        subcausal_ia = None

        try:
            # Buscar todos los selects y sus labels
            labels = driver.find_elements(By.TAG_NAME, "label")
            print(f"[{uuid}] Encontradas {len(labels)} labels")

            for label in labels:
                label_text = label.text.strip().lower()
                print(f"[{uuid}] Label: {label_text[:60]}")

                if "motivo causal" in label_text:
                    # Buscar el select más cercano
                    try:
                        # Intentar find_element en el parent
                        parent = label.find_element(By.XPATH, "./ancestor::div[1]")
                        select = parent.find_element(By.TAG_NAME, "select")
                        # Obtener el valor seleccionado (primer option con valor)
                        options = select.find_elements(By.TAG_NAME, "option")
                        for opt in options:
                            val = opt.get_attribute("value")
                            if val and val.strip():
                                motivo_causal = opt.text.strip()
                                break
                        print(f"[{uuid}] Motivo causal: {motivo_causal}")
                    except Exception as e:
                        print(f"[{uuid}] Error extrayendo motivo_causal: {e}")

                if "subcausal" in label_text:
                    try:
                        parent = label.find_element(By.XPATH, "./ancestor::div[1]")
                        select = parent.find_element(By.TAG_NAME, "select")
                        options = select.find_elements(By.TAG_NAME, "option")
                        for opt in options:
                            val = opt.get_attribute("value")
                            if val and val.strip():
                                subcausal_ia = opt.text.strip()
                                break
                        print(f"[{uuid}] Subcausal IA: {subcausal_ia}")
                    except Exception as e:
                        print(f"[{uuid}] Error extrayendo subcausal_ia: {e}")
        except Exception as e:
            print(f"[{uuid}] Error buscando labels: {e}")

        driver.quit()

        if motivo_causal or subcausal_ia:
            return {
                "motivo_causal": motivo_causal or "",
                "subcausal_ia": subcausal_ia or "",
            }
        return None

    except Exception as e:
        print(f"[{uuid}] Error extrayendo causal: {e}")
        return None


def fetch_bandeja(
    fecha: str | None = None,
    abogado: str | None = None,
    solo_pendientes: bool = True,
    compania: str = "EPS",          # "EPS" | "ARL" | "" (todas)
) -> tuple[list, str]:
    """
    Retorna (tutelas, error_msg).
    fecha:           "YYYY-MM-DD" para filtrar por fecha_notificacion o fecha_expediente.
    abogado:         fragmento del nombre (case-insensitive). None = todos.
    solo_pendientes: si True, excluye las que ya tienen fecha_contestacion.
    compania:        "EPS" | "ARL" | "" para filtrar por compañía.
    """
    if not QUEUE_ID:
        return [], "no_queue_id"

    s = _session()
    url = f"{BASE}/queues/api/queue/content_queue/"

    try:
        r = s.get(url, params={"uuid": QUEUE_ID}, timeout=30)
    except Exception as e:
        return [], str(e)

    if r.status_code != 200:
        if "seus.sura.com" in r.url or "login" in r.url.lower():
            return [], "session_expired"
        return [], f"HTTP {r.status_code}"

    raw_rows = _parse_queue_html(r.text)
    if not raw_rows:
        return [], "sin_datos"

    tutelas = []
    for row in raw_rows:
        t = _normalizar(row)

        # Filtro por compañía (EPS / ARL / vacío = todas)
        if compania and t["compania"] and t["compania"].upper() != compania.upper():
            continue

        # Filtro: abogado
        if abogado and t["abogado"]:
            if abogado.lower() not in t["abogado"].lower():
                continue

        # Filtro: fecha (busca en fecha_notificacion y fecha_expediente)
        if fecha:
            match = any(
                t[col].startswith(fecha)
                for col in ("fecha_notificacion", "fecha_expediente", "fecha_vencimiento")
                if t.get(col)
            )
            if not match:
                continue

        # Filtro: solo pendientes (sin contestación)
        if solo_pendientes and t["fecha_contestacion"]:
            continue

        tutelas.append(t)

    # ── Enriquecer con causales guardados en BD + auto-guardar nuevos ──
    try:
        from database import SessionLocal
        import models as _models
        db = SessionLocal()
        # Construir mapa expediente → causal
        rows = db.query(_models.AthentoMotivoCausal).all()
        mapa_exp = {r.expediente.strip(): r for r in rows}
        mapa_uuid = {r.uuid_athento: r for r in rows if not r.uuid_athento.startswith("exp:")}

        for t in tutelas:
            # Buscar por UUID primero, luego por expediente
            c = mapa_uuid.get(t["uuid"]) or mapa_exp.get(t["expediente"].strip())
            if c and c.modelo:
                t["modelo_sugerido"] = c.modelo
                t["subcausal"]       = c.subcausal or ""
                t["sin_concepto"]    = True
                t["auto_generable"]  = True
            # ── Auto-guardar causal detectado en area_causal ──
            elif t.get("area_causal"):
                clasif = _clasificar("", t.get("area_causal", ""))
                if clasif["modelo_sugerido"] and not c:
                    try:
                        db.add(_models.AthentoMotivoCausal(
                            uuid_athento  = t["uuid"],
                            expediente    = t["expediente"].strip(),
                            modelo        = clasif["modelo_sugerido"],
                            subcausal     = "",
                            asignado_por  = f"Auto · {t['abogado'][:25]}",
                            notas         = f"area_causal={t.get('area_causal')}",
                        ))
                        db.commit()
                        t["modelo_sugerido"] = clasif["modelo_sugerido"]
                        t["sin_concepto"]    = True
                        t["auto_generable"]  = True
                    except Exception:
                        pass
        db.close()
    except Exception:
        pass

    return tutelas, "ok"


# Alias para compatibilidad con código anterior
def fetch_bandeja_hoy(fecha: str | None = None, abogado: str | None = None) -> tuple[list, str]:
    if not fecha:
        fecha = date.today().isoformat()
    return fetch_bandeja(fecha=fecha, abogado=abogado)
