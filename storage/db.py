"""
WEP AI — storage layer (SQLite).

Cambios en v14.1 (review fixes):
- Passwords: bcrypt en vez de SHA-256 sin sal.
- Compatibilidad hacia atrás: hashes SHA-256 antiguos se detectan en login
  y se actualizan a bcrypt automáticamente cuando el usuario inicia sesión.
- `check_and_consume_doc()`: operación atómica que reemplaza el par
  check_can_generate + increment_docs_used (race condition resuelta).
- Conexiones gestionadas con context manager (`with _con()`) — evita leaks
  y elimina la duplicación de open/close en cada función.
- Schema con migraciones versionadas (PRAGMA user_version) — ALTER TABLE
  centralizado, ya no se ejecuta en cada flujo de pago.
- Índices en todas las foreign keys — escalable a miles de filas.
"""

import os
import sqlite3
import hashlib
import logging
from contextlib import contextmanager
from datetime import datetime

import bcrypt

# v16.9.1 — DB_PATH configurable por entorno. En producción (systemd, Docker,
# paneles) el '~' puede resolver a un HOME no escribible y romper el registro con
# un 500 genérico. Define WEPAI_DB_PATH=/ruta/escribible/wepai.db para fijarla.
DB_PATH = os.environ.get("WEPAI_DB_PATH") or os.path.expanduser("~/.wepai/wepai.db")
SCHEMA_VERSION = 2

log = logging.getLogger("wepai.db")


# ── CONNECTION HELPERS ────────────────────────────────────────────────────────
@contextmanager
def _con():
    """Context manager — garantiza commit/close incluso si hay excepción."""
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    con = sqlite3.connect(DB_PATH, isolation_level=None)  # autocommit por defecto; transacciones explícitas
    con.execute("PRAGMA foreign_keys = ON")
    try:
        yield con
    finally:
        con.close()


def init_db():
    """Crea schema, aplica migraciones y crea índices."""
    with _con() as con:
        cur = con.cursor()
        cur.executescript("""
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT, last_name TEXT,
                email TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                plan TEXT DEFAULT 'free',
                docs_used_this_month INTEGER DEFAULT 0,
                month_reset TEXT,
                created_at TEXT
            );
            CREATE TABLE IF NOT EXISTS conversations (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                title TEXT,
                created_at TEXT,
                updated_at TEXT,
                FOREIGN KEY(user_id) REFERENCES users(id)
            );
            CREATE TABLE IF NOT EXISTS messages (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                conversation_id INTEGER NOT NULL,
                role TEXT, content TEXT, created_at TEXT,
                FOREIGN KEY(conversation_id) REFERENCES conversations(id)
            );
            CREATE TABLE IF NOT EXISTS documents (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                conversation_id INTEGER,
                user_id INTEGER NOT NULL,
                title TEXT, doc_type TEXT, file_path TEXT, created_at TEXT,
                FOREIGN KEY(conversation_id) REFERENCES conversations(id),
                FOREIGN KEY(user_id) REFERENCES users(id)
            );

            CREATE INDEX IF NOT EXISTS idx_conv_user      ON conversations(user_id);
            CREATE INDEX IF NOT EXISTS idx_conv_updated   ON conversations(updated_at);
            CREATE INDEX IF NOT EXISTS idx_msg_conv       ON messages(conversation_id);
            CREATE INDEX IF NOT EXISTS idx_doc_conv       ON documents(conversation_id);
            CREATE INDEX IF NOT EXISTS idx_doc_user       ON documents(user_id);
        """)
        _apply_migrations(cur)


def _apply_migrations(cur):
    """Migraciones versionadas con PRAGMA user_version."""
    current = cur.execute("PRAGMA user_version").fetchone()[0]

    if current < 1:
        # v1: añadir columna card_last4 (antes se hacía con ALTER TABLE en el flujo de pago)
        try:
            cur.execute("ALTER TABLE users ADD COLUMN card_last4 TEXT")
        except sqlite3.OperationalError:
            pass  # columna ya existe (instalación previa)

    if current < 2:
        # v2: reservado para futuras migraciones (ej. password algorithm tag)
        pass

    cur.execute(f"PRAGMA user_version = {SCHEMA_VERSION}")


# ── PASSWORD HASHING ──────────────────────────────────────────────────────────
def _hash_password(password: str) -> str:
    """Hash con bcrypt (cost 12 por defecto)."""
    return bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")


def _verify_password(password: str, stored_hash: str) -> tuple[bool, bool]:
    """
    Verifica un password contra el hash almacenado.
    Returns: (es_correcto, necesita_rehash)
    - Hashes bcrypt empiezan con '$2a$', '$2b$' o '$2y$'.
    - Cualquier otro formato se asume SHA-256 antiguo y se marca para rehash.
    """
    if not stored_hash:
        return False, False

    if stored_hash.startswith(("$2a$", "$2b$", "$2y$")):
        try:
            ok = bcrypt.checkpw(password.encode("utf-8"), stored_hash.encode("utf-8"))
            return ok, False
        except ValueError:
            return False, False

    # Legacy SHA-256 (sin sal). Verificar y marcar para upgrade.
    legacy = hashlib.sha256(password.encode("utf-8")).hexdigest()
    return (legacy == stored_hash, True)


# ── USER OPERATIONS ───────────────────────────────────────────────────────────
def register_user(name, last_name, email, password, plan="free"):
    pw_hash = _hash_password(password)
    now = datetime.now().isoformat()
    month = datetime.now().strftime("%Y-%m")
    try:
        with _con() as con:
            cur = con.cursor()
            cur.execute(
                """INSERT INTO users
                   (name,last_name,email,password_hash,plan,docs_used_this_month,month_reset,created_at)
                   VALUES (?,?,?,?,?,0,?,?)""",
                (name, last_name, email, pw_hash, plan, month, now),
            )
            return cur.lastrowid, None
    except sqlite3.IntegrityError:
        return None, "email_exists"
    except Exception as e:
        log.exception("register_user failed")
        return None, str(e)


def login_user(email, password):
    """Verifica credenciales. Migra hashes SHA-256 antiguos a bcrypt si match."""
    with _con() as con:
        cur = con.cursor()
        cur.execute(
            "SELECT id, name, last_name, email, plan, docs_used_this_month, "
            "       month_reset, password_hash "
            "FROM users WHERE email=?",
            (email,),
        )
        row = cur.fetchone()
        if not row:
            return None

        uid, name, last_name, em, plan, docs_used, month_reset, stored = row
        ok, needs_rehash = _verify_password(password, stored)
        if not ok:
            return None

        if needs_rehash:
            try:
                new_hash = _hash_password(password)
                cur.execute("UPDATE users SET password_hash=? WHERE id=?", (new_hash, uid))
                log.info("Upgraded legacy password hash for user %s", uid)
            except Exception:
                log.exception("Password rehash failed for user %s", uid)

        return {
            "id": uid, "name": name, "last_name": last_name, "email": em,
            "plan": plan, "docs_used": docs_used, "month_reset": month_reset,
        }


def update_user_plan(user_id, plan):
    with _con() as con:
        con.execute("UPDATE users SET plan=? WHERE id=?", (plan, user_id))


def save_card_last4(user_id, last4):
    """Persiste sólo los 4 últimos dígitos. Nunca PAN ni CVV."""
    if not last4 or len(last4) != 4 or not last4.isdigit():
        return
    with _con() as con:
        con.execute("UPDATE users SET card_last4=? WHERE id=?", (last4, user_id))


# ── PLAN LIMITS / USAGE ───────────────────────────────────────────────────────
def get_plan_limits(plan):
    return {
        "free": {"docs_month": 5,  "history_days": 7,   "corrections": False, "users": 1,  "support": False},
        "pro":  {"docs_month": -1, "history_days": 90,  "corrections": True,  "users": 1,  "support": True},
        "biz":  {"docs_month": -1, "history_days": 365, "corrections": True,  "users": 10, "support": True},
    }.get(plan, {"docs_month": 5, "history_days": 7, "corrections": False, "users": 1, "support": False})


def check_and_consume_doc(user_id) -> tuple[bool, str | None]:
    """
    Verifica el límite del plan y consume 1 doc en UNA transacción atómica.
    Reemplaza el par (check_can_generate + increment_docs_used) que tenía
    race condition.

    Returns:
        (True, None)              si se consumió correctamente
        (False, 'limit_reached')  si el plan free está al tope
        (False, 'no_user')        si no existe el usuario
    """
    cur_month = datetime.now().strftime("%Y-%m")
    with _con() as con:
        cur = con.cursor()
        # BEGIN IMMEDIATE adquiere lock de escritura — bloquea otros writers
        # hasta el COMMIT, así dos llamadas paralelas se serializan.
        cur.execute("BEGIN IMMEDIATE")
        try:
            cur.execute(
                "SELECT plan, docs_used_this_month, month_reset FROM users WHERE id=?",
                (user_id,),
            )
            row = cur.fetchone()
            if not row:
                cur.execute("ROLLBACK")
                return False, "no_user"

            plan, used, month_reset = row

            # Reset mensual
            if month_reset != cur_month:
                used = 0
                cur.execute(
                    "UPDATE users SET docs_used_this_month=0, month_reset=? WHERE id=?",
                    (cur_month, user_id),
                )

            limits = get_plan_limits(plan)
            limit = limits["docs_month"]

            if limit != -1 and used >= limit:
                cur.execute("ROLLBACK")
                return False, "limit_reached"

            cur.execute(
                "UPDATE users SET docs_used_this_month = docs_used_this_month + 1 WHERE id=?",
                (user_id,),
            )
            cur.execute("COMMIT")
            return True, None
        except Exception:
            cur.execute("ROLLBACK")
            log.exception("check_and_consume_doc failed")
            raise


def get_docs_used(user_id) -> int:
    """Cuántos documentos ha generado el usuario este mes."""
    cur_month = datetime.now().strftime("%Y-%m")
    with _con() as con:
        cur = con.cursor()
        cur.execute(
            "SELECT docs_used_this_month, month_reset FROM users WHERE id=?",
            (user_id,),
        )
        row = cur.fetchone()
        if not row:
            return 0
        used, month_reset = row
        if month_reset != cur_month:
            # Hacer el reset también aquí para que el counter del UI sea consistente.
            cur.execute(
                "UPDATE users SET docs_used_this_month=0, month_reset=? WHERE id=?",
                (cur_month, user_id),
            )
            return 0
        return used


def get_user_field(user_id, field: str):
    """v14.6 — Lee un campo arbitrario del usuario de forma segura.
    Maneja columnas que pueden no existir aún (como stripe_customer_id que
    crea el webhook server). Devuelve None si la columna o el usuario no existe.
    Whitelist de campos seguros para evitar SQL injection.
    """
    SAFE_FIELDS = {
        "stripe_customer_id", "stripe_subscription_id", "subscription_status",
        "plan", "email", "name", "last_name", "docs_used_this_month",
        "month_reset", "card_last4",
    }
    if field not in SAFE_FIELDS:
        return None
    with _con() as con:
        cur = con.cursor()
        # Verificar que la columna existe (puede no existir aún en DBs viejas)
        cur.execute("PRAGMA table_info(users)")
        cols = {row[1] for row in cur.fetchall()}
        if field not in cols:
            return None
        cur.execute(f"SELECT {field} FROM users WHERE id=?", (user_id,))
        row = cur.fetchone()
        return row[0] if row else None


# ── BACKWARD-COMPAT ALIASES ───────────────────────────────────────────────────
# El UI todavía llama check_can_generate + increment_docs_used. Mantenemos los
# nombres viejos como wrappers que delegan en check_and_consume_doc para no
# romper código externo, pero internamente ahora SÍ es atómico.
def check_can_generate(user_id) -> tuple[bool, str | None]:
    """DEPRECATED — usa check_and_consume_doc. Mantenido para compatibilidad."""
    return check_and_consume_doc(user_id)


def increment_docs_used(user_id):
    """DEPRECATED — el consumo ya se hace dentro de check_and_consume_doc.
    Mantenida como no-op para no romper a quien todavía la llama.
    """
    return  # already consumed atomically in check_and_consume_doc


# ── CONVERSATIONS / MESSAGES / DOCUMENTS ──────────────────────────────────────
def new_conversation(user_id, title="Nueva conversación"):
    now = datetime.now().isoformat()
    with _con() as con:
        cur = con.cursor()
        cur.execute(
            "INSERT INTO conversations (user_id,title,created_at,updated_at) VALUES (?,?,?,?)",
            (user_id, title, now, now),
        )
        return cur.lastrowid


def save_message(conversation_id, role, content):
    now = datetime.now().isoformat()
    with _con() as con:
        con.execute(
            "INSERT INTO messages (conversation_id,role,content,created_at) VALUES (?,?,?,?)",
            (conversation_id, role, content, now),
        )
        con.execute(
            "UPDATE conversations SET updated_at=? WHERE id=?",
            (now, conversation_id),
        )


def save_document(conversation_id, user_id, title, doc_type, file_path):
    now = datetime.now().isoformat()
    with _con() as con:
        con.execute(
            "INSERT INTO documents (conversation_id,user_id,title,doc_type,file_path,created_at) "
            "VALUES (?,?,?,?,?,?)",
            (conversation_id, user_id, title, doc_type, file_path, now),
        )


def get_recent_conversations(user_id, limit=20):
    with _con() as con:
        cur = con.cursor()
        cur.execute(
            "SELECT id,title,updated_at FROM conversations "
            "WHERE user_id=? ORDER BY updated_at DESC LIMIT ?",
            (user_id, limit),
        )
        return cur.fetchall()


def get_messages(conversation_id):
    with _con() as con:
        cur = con.cursor()
        cur.execute(
            "SELECT role,content FROM messages WHERE conversation_id=? ORDER BY id",
            (conversation_id,),
        )
        return [{"role": r[0], "content": r[1]} for r in cur.fetchall()]


def get_documents(conversation_id):
    with _con() as con:
        cur = con.cursor()
        cur.execute(
            "SELECT title,doc_type,file_path,created_at FROM documents "
            "WHERE conversation_id=? ORDER BY id DESC",
            (conversation_id,),
        )
        return cur.fetchall()


def update_conversation_title(conversation_id, title):
    with _con() as con:
        con.execute("UPDATE conversations SET title=? WHERE id=?", (title, conversation_id))
