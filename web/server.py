"""
WEP AI — Servidor Web (v16.6)
==============================
Envuelve el motor de generación (brain + controllers) en una API web con FastAPI,
accesible desde cualquier sistema operativo y desde el móvil, sin instalación.

NOVEDADES v16.6 (versión web profesional):
    - Autenticación real: registro / login / logout con sesión por cookie firmada
      (HMAC), reutilizando storage/db.py (bcrypt, planes free/pro/biz).
    - Historial de conversaciones por usuario (persistido en SQLite).
    - Límite de documentos por plan (consumo atómico) con aviso de "mejorar plan".
    - Vista previa del documento generado (miniatura de la 1ª página, best-effort
      vía LibreOffice; degrada con elegancia si no está disponible).

ENFOQUE COLOMBIA:
    Toda la lógica de país/idioma ya está fijada en el motor (Colombia + español).
    Este servidor solo orquesta: auth → chat → generación → historial → descarga.

ENDPOINTS:
    GET  /api/health                 → estado del servidor y del motor
    POST /api/register               → crea cuenta + inicia sesión
    POST /api/login                  → inicia sesión
    POST /api/logout                 → cierra sesión
    GET  /api/me                     → datos del usuario + plan + uso
    GET  /api/conversations          → lista de conversaciones del usuario
    GET  /api/conversations/{id}     → mensajes + documentos de una conversación
    POST /api/chat                   → conversa; puede generar y devolver un documento
    GET  /api/download/{token}       → descarga el documento generado
    GET  /api/preview/{token}        → miniatura PNG de la 1ª página del documento
"""

import os
import sys
import uuid
import time
import json
import base64
import hmac
import hashlib
import shutil
import tempfile
import threading
import subprocess
from pathlib import Path
from collections import defaultdict, deque

from fastapi import FastAPI, HTTPException, Request, Response
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

# ── Hacer importable el motor (carpeta padre del proyecto) ────────────────────
PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from storage import db  # noqa: E402

# ── Inicializar la base de datos al arrancar ──────────────────────────────────
try:
    db.init_db()
except Exception as _e:  # pragma: no cover
    print(f"[server] No se pudo inicializar la BD: {_e}", file=sys.stderr)

# ── Directorio de descargas/temporales ────────────────────────────────────────
DOWNLOAD_DIR = Path(tempfile.gettempdir()) / "wepai_web_downloads"
DOWNLOAD_DIR.mkdir(exist_ok=True)

DOWNLOAD_TTL_SECONDS = 60 * 60          # 1 h de validez del enlace
_DOWNLOAD_TOKENS: dict = {}
_PREVIEW_TOKENS: dict = {}
_TOKENS_LOCK = threading.Lock()

# Rate limiting por IP en /api/chat para proteger el costo de la API.
RATE_WINDOW_SECONDS = 60
RATE_MAX_REQUESTS = 20
_RATE_BUCKET: dict = defaultdict(deque)
_RATE_LOCK = threading.Lock()

# ── Sesiones: cookie firmada con HMAC (sin dependencias extra) ────────────────
# El secreto debe venir de la variable de entorno WEPAI_SECRET_KEY en producción.
SECRET_KEY = os.environ.get("WEPAI_SECRET_KEY", "wepai-dev-secret-cambia-esto-en-produccion").encode()
SESSION_COOKIE = "wepai_session"
SESSION_TTL_SECONDS = 60 * 60 * 24 * 30  # 30 días
# secure=True exige HTTPS. Se activa salvo que WEPAI_INSECURE_COOKIES=1 (desarrollo local).
COOKIE_SECURE = os.environ.get("WEPAI_INSECURE_COOKIES", "0") != "1"


def _sign_session(user_id: int) -> str:
    payload = {"uid": user_id, "exp": int(time.time()) + SESSION_TTL_SECONDS}
    raw = base64.urlsafe_b64encode(json.dumps(payload).encode()).decode().rstrip("=")
    sig = hmac.new(SECRET_KEY, raw.encode(), hashlib.sha256).hexdigest()[:40]
    return f"{raw}.{sig}"


def _unsign_session(token: str):
    try:
        raw, sig = token.rsplit(".", 1)
        expected = hmac.new(SECRET_KEY, raw.encode(), hashlib.sha256).hexdigest()[:40]
        if not hmac.compare_digest(sig, expected):
            return None
        pad = "=" * (-len(raw) % 4)
        payload = json.loads(base64.urlsafe_b64decode(raw + pad))
        if payload.get("exp", 0) < time.time():
            return None
        return payload
    except Exception:
        return None


def _current_user_id(request: Request):
    token = request.cookies.get(SESSION_COOKIE)
    if not token:
        return None
    payload = _unsign_session(token)
    return payload.get("uid") if payload else None


def _require_user(request: Request) -> int:
    uid = _current_user_id(request)
    if not uid:
        raise HTTPException(status_code=401, detail="No autenticado.")
    return uid


def _set_session_cookie(resp: Response, user_id: int):
    resp.set_cookie(
        SESSION_COOKIE, _sign_session(user_id),
        max_age=SESSION_TTL_SECONDS, httponly=True,
        samesite="lax", secure=COOKIE_SECURE, path="/",
    )


# ── Mantenimiento de tokens ───────────────────────────────────────────────────
def _sweep_expired_downloads():
    now = time.time()
    with _TOKENS_LOCK:
        for store in (_DOWNLOAD_TOKENS, _PREVIEW_TOKENS):
            vencidos = [t for t, m in store.items()
                        if now - m.get("created", 0) > DOWNLOAD_TTL_SECONDS]
            for t in vencidos:
                meta = store.pop(t, None)
                if meta:
                    try:
                        p = meta.get("path")
                        if p and os.path.exists(p):
                            os.remove(p)
                    except OSError:
                        pass


def _rate_limited(client_ip: str) -> bool:
    now = time.time()
    with _RATE_LOCK:
        bucket = _RATE_BUCKET[client_ip]
        while bucket and now - bucket[0] > RATE_WINDOW_SECONDS:
            bucket.popleft()
        if len(bucket) >= RATE_MAX_REQUESTS:
            return True
        bucket.append(now)
        return False


# ── Vista previa: 1ª página → PNG (best-effort, requiere LibreOffice) ─────────
_SOFFICE = shutil.which("soffice") or shutil.which("libreoffice")
_PDFTOPPM = shutil.which("pdftoppm")


def _make_preview(src_path: str) -> str | None:
    """Genera una miniatura PNG de la primera página. Devuelve la URL o None.
    No bloquea el flujo: si falta LibreOffice/pdftoppm o algo falla, retorna None."""
    if not _SOFFICE or not _PDFTOPPM:
        return None
    tmp = tempfile.mkdtemp(prefix="wepai_prev_")
    try:
        profile = Path(tmp) / "lo_profile"
        subprocess.run(
            [_SOFFICE, "--headless", f"-env:UserInstallation=file://{profile}",
             "--convert-to", "pdf", "--outdir", tmp, src_path],
            timeout=45, capture_output=True,
        )
        pdf = next((os.path.join(tmp, f) for f in os.listdir(tmp) if f.endswith(".pdf")), None)
        if not pdf:
            return None
        subprocess.run([_PDFTOPPM, "-png", "-f", "1", "-l", "1", "-r", "72",
                        pdf, os.path.join(tmp, "prev")], timeout=25, capture_output=True)
        png = next((os.path.join(tmp, f) for f in os.listdir(tmp) if f.endswith(".png")), None)
        if not png:
            return None
        token = uuid.uuid4().hex
        dest = DOWNLOAD_DIR / f"prev_{token}.png"
        shutil.copy2(png, dest)
        with _TOKENS_LOCK:
            _PREVIEW_TOKENS[token] = {"path": str(dest), "created": time.time()}
        return f"/api/preview/{token}"
    except Exception:
        return None
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


# ── Generación de documento ───────────────────────────────────────────────────
def _generate_document(gen_data: dict) -> dict | None:
    tipo = (gen_data.get("tipo") or "").lower()
    try:
        if tipo == "word":
            from office.word_controller import generate_word
            path = generate_word(gen_data)
        elif tipo == "excel":
            from office.excel_controller import generate_excel
            res = generate_excel(gen_data)
            path = res[0] if isinstance(res, tuple) else res
        elif tipo == "powerpoint":
            from office.ppt_controller import generate_powerpoint
            path = generate_powerpoint(gen_data)
        else:
            return None
    except Exception as e:
        return {"error": f"No se pudo generar el documento: {e}"}

    if not path or not os.path.exists(path):
        return {"error": "El documento no se generó correctamente."}

    _sweep_expired_downloads()
    token = uuid.uuid4().hex
    filename = os.path.basename(path)
    dest = DOWNLOAD_DIR / f"{token}_{filename}"
    shutil.copy2(path, dest)
    with _TOKENS_LOCK:
        _DOWNLOAD_TOKENS[token] = {"path": str(dest), "created": time.time()}

    preview_url = _make_preview(str(dest))
    return {
        "tipo": tipo,
        "filename": filename,
        "download_url": f"/api/download/{token}",
        "preview_url": preview_url,
        "file_path": str(dest),
    }


app = FastAPI(title="WEP AI — Colombia", version="16.6")


# ── Modelos ───────────────────────────────────────────────────────────────────
class AuthRegister(BaseModel):
    name: str
    last_name: str = ""
    email: str
    password: str


class AuthLogin(BaseModel):
    email: str
    password: str


class ChatRequest(BaseModel):
    conversation_id: int | None = None
    message: str


def _plan_label(plan: str) -> str:
    return {"free": "Gratis", "pro": "Pro", "biz": "Empresas"}.get(plan, plan)


def _me_payload(uid: int) -> dict:
    plan = db.get_user_field(uid, "plan") or "free"
    name = db.get_user_field(uid, "name") or ""
    last = db.get_user_field(uid, "last_name") or ""
    email = db.get_user_field(uid, "email") or ""
    limits = db.get_plan_limits(plan)
    used = db.get_docs_used(uid)
    return {
        "id": uid, "name": name, "last_name": last, "email": email,
        "plan": plan, "plan_label": _plan_label(plan),
        "docs_used": used, "docs_limit": limits["docs_month"],  # -1 = ilimitado
        "corrections": limits["corrections"], "history_days": limits["history_days"],
    }


# ── Endpoints: salud ──────────────────────────────────────────────────────────
@app.get("/api/health")
async def health():
    try:
        import agent.brain  # noqa: F401
        engine_ok = True
    except Exception:
        engine_ok = False
    try:
        from agent.structural_alerts import has_pending_alerts
        pending = has_pending_alerts()
    except Exception:
        pending = False
    # v16.9.1 — diagnóstico de la base de datos: confirma que el proceso puede
    # REALMENTE escribir en ella (la causa #1 de "No se pudo crear la cuenta").
    db_info = {"path": None, "dir_writable": None, "write_test": None, "error": None}
    try:
        from storage import db as _dbmod
        db_path = _dbmod.DB_PATH
        db_dir = os.path.dirname(db_path)
        db_info["path"] = db_path
        db_info["dir_writable"] = os.access(db_dir, os.W_OK) if os.path.isdir(db_dir) else False
        # intento de escritura real en una tabla temporal
        with _dbmod._con() as _c:
            _c.execute("CREATE TABLE IF NOT EXISTS _healthcheck (ts TEXT)")
            _c.execute("INSERT INTO _healthcheck (ts) VALUES (datetime('now'))")
            _c.execute("DELETE FROM _healthcheck")
        db_info["write_test"] = "ok"
    except Exception as _e:
        db_info["write_test"] = "failed"
        db_info["error"] = f"{type(_e).__name__}: {_e}"
    return {
        "status": "ok", "version": "16.9.1", "country": "Colombia", "language": "es",
        "api_key_set": bool(os.environ.get("ANTHROPIC_API_KEY", "")),
        "engine_loaded": engine_ok,
        "preview_available": bool(_SOFFICE and _PDFTOPPM),
        "structural_alerts_pending": pending,
        "home_env": os.environ.get("HOME", "<unset>"),
        "db": db_info,
    }


# ── v16.8 — Alertas estructurales (cambios legales que requieren un dev) ──────
# Endpoint protegido por un token de admin (WEPAI_ADMIN_TOKEN). Permite al dueño
# revisar las alertas pendientes desde la web sin entrar al servidor.
@app.get("/api/admin/structural-alerts")
async def admin_structural_alerts(request: Request):
    admin_token = os.environ.get("WEPAI_ADMIN_TOKEN", "").strip()
    provided = request.headers.get("X-Admin-Token", "").strip()
    if not admin_token or not provided or provided != admin_token:
        raise HTTPException(status_code=403, detail="No autorizado.")
    try:
        from agent.structural_alerts import get_pending_alerts, developer_report
        return {"pending": get_pending_alerts(), "report": developer_report()}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error: {e}")


# ── Endpoints: autenticación ──────────────────────────────────────────────────
@app.post("/api/register")
async def register(body: AuthRegister):
    email = body.email.strip().lower()
    if not email or "@" not in email:
        raise HTTPException(status_code=400, detail="Correo no válido.")
    if len(body.password) < 8:
        raise HTTPException(status_code=400, detail="La contraseña debe tener al menos 8 caracteres.")
    if not body.name.strip():
        raise HTTPException(status_code=400, detail="Ingresa tu nombre.")
    uid, err = db.register_user(body.name.strip(), body.last_name.strip(), email, body.password, "free")
    if err == "email_exists":
        raise HTTPException(status_code=409, detail="Ya existe una cuenta con ese correo.")
    if err or not uid:
        # v16.9.1 — antes este 500 ocultaba la causa real (permisos de BD, disco
        # de solo lectura, HOME no escribible, etc.). Ahora la registramos y la
        # devolvemos de forma acotada para poder diagnosticar en producción.
        import sys as _sys
        print(f"[server] register_user failed: {err}", file=_sys.stderr)
        detail = "No se pudo crear la cuenta."
        if err:
            detail = f"No se pudo crear la cuenta ({err})."
        raise HTTPException(status_code=500, detail=detail)
    resp = JSONResponse({"ok": True, "user": _me_payload(uid)})
    _set_session_cookie(resp, uid)
    return resp


@app.post("/api/login")
async def login(body: AuthLogin):
    email = body.email.strip().lower()
    user = db.login_user(email, body.password)
    if not user:
        raise HTTPException(status_code=401, detail="Correo o contraseña incorrectos.")
    resp = JSONResponse({"ok": True, "user": _me_payload(user["id"])})
    _set_session_cookie(resp, user["id"])
    return resp


@app.post("/api/logout")
async def logout():
    resp = JSONResponse({"ok": True})
    resp.delete_cookie(SESSION_COOKIE, path="/")
    return resp


@app.get("/api/me")
async def me(request: Request):
    uid = _require_user(request)
    return _me_payload(uid)


# ── Endpoints: historial ──────────────────────────────────────────────────────
@app.get("/api/conversations")
async def list_conversations(request: Request):
    uid = _require_user(request)
    rows = db.get_recent_conversations(uid, limit=40)
    return {"conversations": [{"id": r[0], "title": r[1], "updated_at": r[2]} for r in rows]}


@app.get("/api/conversations/{conv_id}")
async def get_conversation(conv_id: int, request: Request):
    uid = _require_user(request)
    # Verificar propiedad de la conversación
    owner = db.get_user_field  # noqa
    own_ids = {r[0] for r in db.get_recent_conversations(uid, limit=1000)}
    if conv_id not in own_ids:
        raise HTTPException(status_code=404, detail="Conversación no encontrada.")
    messages = db.get_messages(conv_id)
    docs = db.get_documents(conv_id)
    documents = [{"title": d[0], "tipo": d[1], "created_at": d[3]} for d in docs]
    return {"id": conv_id, "messages": messages, "documents": documents}


# ── Endpoint: chat ────────────────────────────────────────────────────────────
@app.post("/api/chat")
async def chat(req: ChatRequest, request: Request):
    uid = _require_user(request)

    if not os.environ.get("ANTHROPIC_API_KEY", ""):
        raise HTTPException(status_code=500, detail="ANTHROPIC_API_KEY no configurada en el servidor.")

    client_ip = request.client.host if request.client else "unknown"
    if _rate_limited(client_ip):
        raise HTTPException(status_code=429, detail="Vas muy rápido. Espera unos segundos e intenta de nuevo.")

    text = (req.message or "").strip()
    if not text:
        raise HTTPException(status_code=400, detail="El mensaje está vacío.")

    # Resolver / crear conversación (verificando propiedad)
    conv_id = req.conversation_id
    own_ids = {r[0] for r in db.get_recent_conversations(uid, limit=1000)}
    title = None
    if conv_id and conv_id in own_ids:
        history = db.get_messages(conv_id)
    else:
        try:
            from agent.brain import generate_title
            title = generate_title(text)
        except Exception:
            title = (text[:48] + "…") if len(text) > 48 else text
        conv_id = db.new_conversation(uid, title or "Nueva conversación")
        history = []

    db.save_message(conv_id, "user", text)

    from agent.brain import chat as brain_chat
    try:
        result = brain_chat(history, text, ui_lang="es")
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"Error del motor: {e}")

    ai_text = result.get("text", "")
    response = {"text": ai_text, "document": None, "conversation_id": conv_id,
                "title": title, "limit_reached": False}

    if result.get("generate") and result.get("gen_data"):
        ok, reason = db.check_and_consume_doc(uid)
        if not ok and reason == "limit_reached":
            response["limit_reached"] = True
            response["text"] = (
                "Alcanzaste el límite de documentos de tu plan Gratis este mes. "
                "Mejora a Pro para generar documentos ilimitados."
            )
            ai_text = response["text"]
        else:
            doc = _generate_document(result["gen_data"])
            if doc and "error" not in doc:
                gd = result["gen_data"]
                db.save_document(conv_id, uid, gd.get("titulo", doc["filename"]),
                                 doc["tipo"], doc.get("file_path", ""))
                doc.pop("file_path", None)
                response["document"] = doc
            elif doc and "error" in doc:
                ai_text = ai_text + f"\n\n⚠️ {doc['error']}"
                response["text"] = ai_text

    db.save_message(conv_id, "assistant", ai_text)
    return JSONResponse(response)


# ── Endpoints: descarga y preview ─────────────────────────────────────────────
@app.get("/api/download/{token}")
async def download(token: str):
    _sweep_expired_downloads()
    with _TOKENS_LOCK:
        meta = _DOWNLOAD_TOKENS.get(token)
        path = meta.get("path") if meta else None
    if not path or not os.path.exists(path):
        raise HTTPException(status_code=404, detail="Documento no encontrado o expirado.")
    original_name = os.path.basename(path).split("_", 1)[-1]
    return FileResponse(path, filename=original_name, media_type="application/octet-stream")


@app.get("/api/preview/{token}")
async def preview(token: str):
    _sweep_expired_downloads()
    with _TOKENS_LOCK:
        meta = _PREVIEW_TOKENS.get(token)
        path = meta.get("path") if meta else None
    if not path or not os.path.exists(path):
        raise HTTPException(status_code=404, detail="Vista previa no disponible.")
    return FileResponse(path, media_type="image/png")


# ── Servir el frontend ────────────────────────────────────────────────────────
STATIC_DIR = Path(__file__).resolve().parent / "static"
app.mount("/", StaticFiles(directory=str(STATIC_DIR), html=True), name="static")


if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run("server:app", host="0.0.0.0", port=port, reload=False)
