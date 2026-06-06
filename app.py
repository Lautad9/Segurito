"""
Servidor web del Generador de AR.
Endpoints:
  GET  /                 -> sirve la página (frontend)
  POST /api/analizar     -> recibe el texto del PE, la IA devuelve el borrador (fases + riesgos)
  POST /api/generar      -> recibe el borrador confirmado, rellena la plantilla y devuelve Excel + PDF
  GET  /download/<file>  -> descarga del archivo generado

Variables de entorno:
  ANTHROPIC_API_KEY   (obligatoria para el paso de IA)
"""
import os, json, re
import requests
from flask import Flask, request, jsonify, send_from_directory, abort
from engine_runner import generar_todo, OUTDIR, CATALOGO

BASE = os.path.dirname(os.path.abspath(__file__))
app = Flask(__name__, static_folder=None)

# Catálogo de riesgos válidos (los 28 tipos) -> para validar lo que devuelve la IA
with open(os.path.join(BASE, "engine", "riesgos_validos.json"), encoding="utf-8") as f:
    NOMBRES_VALIDOS = json.load(f)

API_URL = "https://api.anthropic.com/v1/messages"
MODELO = "claude-haiku-4-5-20251001"

def construir_prompt(pe_text):
    lista = "\n".join(f"  - {n}" for n in NOMBRES_VALIDOS)
    return f"""Sos un experto en Higiene y Seguridad en el Trabajo (petróleo y construcción, Argentina).
Te paso el texto de un PROCEDIMIENTO DE TRABAJO (PE). Armá el borrador de un Análisis de Riesgo.

DEVOLVÉ SOLO un objeto JSON válido (sin markdown ni texto extra) con esta forma:
{{"titulo":"...","instalacion":"...","fases":[{{"titulo":"...","detalle":"una oración","personas":<numero o null>,"herramientas":["..."],"riesgos":["..."]}}]}}

REGLAS:
- 2 a 6 fases lógicas y secuenciales.
- En "riesgos" usá EXCLUSIVAMENTE valores EXACTOS de esta lista (copialos tal cual, incluso con errores de tipeo):
{lista}
- 3 a 8 riesgos aplicables por fase. "detalle" breve. Si el texto menciona la planta, ponela en "instalacion".

PROCEDIMIENTO:
\"\"\"{pe_text[:6000]}\"\"\""""

def validar_borrador(p):
    """Filtra riesgos fuera de catálogo para no dejar pasar invenciones."""
    fases = []
    for f in (p.get("fases") or []):
        riesgos = [r for r in (f.get("riesgos") or []) if r in NOMBRES_VALIDOS]
        fases.append({
            "titulo": f.get("titulo", ""), "detalle": f.get("detalle", ""),
            "personas": f.get("personas"), "herramientas": f.get("herramientas") or [],
            "riesgos": riesgos,
        })
    return {"titulo": p.get("titulo", ""), "instalacion": p.get("instalacion", ""), "fases": fases}

def extraer_texto_docx(file_bytes):
    """Extrae texto plano de un .docx."""
    import io
    from docx import Document
    doc = Document(io.BytesIO(file_bytes))
    return "\n".join(p.text for p in doc.paragraphs if p.text.strip())

@app.post("/api/analizar-archivo")
def analizar_archivo():
    """Recibe un .docx, extrae el texto y lo analiza con IA en un solo paso."""
    if "file" not in request.files:
        return jsonify(error="No se recibió ningún archivo."), 400
    f = request.files["file"]
    if not f.filename.lower().endswith((".docx", ".doc")):
        return jsonify(error="Solo se aceptan archivos .docx"), 400
    try:
        pe_text = extraer_texto_docx(f.read())
    except Exception as e:
        return jsonify(error=f"No se pudo leer el archivo: {e}"), 400
    if not pe_text.strip():
        return jsonify(error="El archivo está vacío o no tiene texto legible."), 400
    # reutilizar la lógica de análisis con IA
    key = os.environ.get("ANTHROPIC_API_KEY")
    if not key:
        return jsonify(error="Falta configurar ANTHROPIC_API_KEY en el servidor."), 500
    try:
        r = requests.post(API_URL, timeout=60, headers={
            "x-api-key": key, "anthropic-version": "2023-06-01", "content-type": "application/json",
        }, json={"model": MODELO, "max_tokens": 2000,
                 "messages": [{"role": "user", "content": construir_prompt(pe_text)}]})
        print(f"[Anthropic/archivo] status={r.status_code}")
        if r.status_code != 200:
            return jsonify(error=f"Error de Anthropic ({r.status_code}): {r.text[:200]}"), 502
        data = r.json()
        txt = "".join(b.get("text", "") for b in data.get("content", []))
        a, b = txt.find("{"), txt.rfind("}")
        if a < 0 or b < 0:
            return jsonify(error="La IA no pudo procesar el archivo. Intentá de nuevo."), 502
        parsed = json.loads(txt[a:b + 1])
        return jsonify({**validar_borrador(parsed), "pe_text": pe_text[:500]})
    except Exception as e:
        return jsonify(error=f"Error inesperado: {e}"), 502

@app.post("/api/analizar")
def analizar():
    pe_text = (request.json or {}).get("pe_text", "").strip()
    if not pe_text:
        return jsonify(error="Falta el texto del procedimiento."), 400
    key = os.environ.get("ANTHROPIC_API_KEY")
    if not key:
        return jsonify(error="Falta configurar ANTHROPIC_API_KEY en el servidor."), 500
    try:
        r = requests.post(API_URL, timeout=60, headers={
            "x-api-key": key, "anthropic-version": "2023-06-01", "content-type": "application/json",
        }, json={"model": MODELO, "max_tokens": 2000,
                 "messages": [{"role": "user", "content": construir_prompt(pe_text)}]})
        print(f"[Anthropic] status={r.status_code} body={r.text[:300]}")
        if r.status_code == 401:
            return jsonify(error="Clave de IA inválida. Verificá que la clave en INICIAR.bat sea correcta y esté activa en console.anthropic.com"), 502
        if r.status_code == 403:
            return jsonify(error="Clave sin permisos. Verificá que tengas créditos en console.anthropic.com → Billing"), 502
        if r.status_code != 200:
            return jsonify(error=f"Error de Anthropic ({r.status_code}): {r.text[:200]}"), 502
        data = r.json()
        if "error" in data:
            return jsonify(error=f"Anthropic dice: {data['error'].get('message','error desconocido')}"), 502
        txt = "".join(b.get("text", "") for b in data.get("content", []))
        a, b = txt.find("{"), txt.rfind("}")
        if a < 0 or b < 0:
            return jsonify(error="La IA no devolvió un resultado válido. Intentá de nuevo."), 502
        parsed = json.loads(txt[a:b + 1])
        return jsonify(validar_borrador(parsed))
    except requests.exceptions.ConnectionError:
        return jsonify(error="No se pudo conectar con la IA. Verificá que tengas internet."), 502
    except Exception as e:
        return jsonify(error=f"Error inesperado: {e}"), 502

@app.post("/api/generar")
def generar_endpoint():
    spec = request.json or {}
    if not spec.get("fases"):
        return jsonify(error="El borrador no tiene fases."), 400
    try:
        res = generar_todo(spec)
        jid = res["job_id"]
        return jsonify(job_id=jid,
                       xlsx_url=f"/download/AR_{jid}.xlsm",
                       pdf_url=f"/download/AR_{jid}.pdf" if res["pdf"] else None)
    except Exception as e:
        return jsonify(error=f"No se pudo generar el AR: {e}"), 500

@app.get("/download/<path:fname>")
def download(fname):
    if not re.fullmatch(r"AR_[a-f0-9]+\.(xlsm|pdf)", fname):
        abort(404)
    return send_from_directory(OUTDIR, fname, as_attachment=True)

@app.get("/")
def home():
    return send_from_directory(os.path.join(BASE, "frontend"), "index.html")

@app.get("/<path:asset>")
def assets(asset):
    return send_from_directory(os.path.join(BASE, "frontend"), asset)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8000)), debug=True)
