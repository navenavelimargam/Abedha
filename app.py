from __future__ import annotations

import base64
import io
import json
import logging
import os
import re
import secrets
import sqlite3
import subprocess
import tempfile
import time
import uuid
from functools import wraps
from pathlib import Path
from typing import Any

import requests
from flask import Flask, jsonify, render_template, request, send_file, session
from werkzeug.utils import secure_filename

try:
    import fitz
except ImportError:
    fitz = None
try:
    import pandas as pd
except ImportError:
    pd = None
try:
    import openpyxl
except ImportError:
    openpyxl = None
try:
    import docx
except ImportError:
    docx = None
try:
    from pptx import Presentation
except ImportError:
    Presentation = None
try:
    from PIL import Image as PILImage
except ImportError:
    PILImage = None
try:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
except ImportError:
    plt = None
try:
    import pytesseract
except ImportError:
    pytesseract = None
try:
    import networkx as nx
except ImportError:
    nx = None

try:
    from kb_index import search_kb
except Exception:
    def search_kb(*args, **kwargs):
        return []

from local_store import (
    add_message as _store_add_message, add_team_member as _store_add_team_member, authenticate,
    get_chat_files, get_messages, get_user, init_db,
    join_team_by_code, list_chats as _store_list_chats, list_files as _store_list_files,
    list_teams, list_users, save_file, team_members,
    create_chat as _store_create_chat, create_team as _store_create_team,
)

# Compatibility layer for the different local_store.py revisions used by the prototype.
# Some revisions expose _db(), newer ones expose _conn(). Registration must work with
# either schema without changing the existing UI.
def _store_conn():
    import local_store as _ls
    conn_factory = getattr(_ls, "_conn", None) or getattr(_ls, "_db", None)
    if conn_factory is None:
        raise RuntimeError("local_store.py does not expose a SQLite connection helper.")
    return conn_factory()

def _user_table_columns(conn):
    return {row[1] for row in conn.execute("PRAGMA table_info(users)").fetchall()}

def _ensure_registration_schema():
    conn = _store_conn()
    try:
        cols = {row[1] for row in conn.execute("PRAGMA table_info(users)").fetchall()}
        if "name" not in cols:
            conn.execute("ALTER TABLE users ADD COLUMN name TEXT NOT NULL DEFAULT ''")
        if "email" not in cols:
            conn.execute("ALTER TABLE users ADD COLUMN email TEXT")
        if "status" not in cols:
            conn.execute("ALTER TABLE users ADD COLUMN status TEXT NOT NULL DEFAULT 'approved'")
        conn.execute("""CREATE TABLE IF NOT EXISTS registration_requests (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            company TEXT NOT NULL DEFAULT '',
            request_type TEXT NOT NULL DEFAULT 'employee',
            status TEXT NOT NULL DEFAULT 'pending',
            message TEXT NOT NULL DEFAULT '',
            created_at REAL NOT NULL,
            reviewed_at REAL,
            reviewed_by INTEGER,
            FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE
        )""")
        rcols = {row[1] for row in conn.execute("PRAGMA table_info(registration_requests)").fetchall()}
        if "company" not in rcols:
            conn.execute("ALTER TABLE registration_requests ADD COLUMN company TEXT NOT NULL DEFAULT ''")
        if "request_type" not in rcols:
            conn.execute("ALTER TABLE registration_requests ADD COLUMN request_type TEXT NOT NULL DEFAULT 'employee'")
        if "message" not in rcols:
            conn.execute("ALTER TABLE registration_requests ADD COLUMN message TEXT NOT NULL DEFAULT ''")
        if "reviewed_at" not in rcols:
            conn.execute("ALTER TABLE registration_requests ADD COLUMN reviewed_at REAL")
        if "reviewed_by" not in rcols:
            conn.execute("ALTER TABLE registration_requests ADD COLUMN reviewed_by INTEGER")
        conn.execute("""CREATE TABLE IF NOT EXISTS notifications (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            recipient_id INTEGER NOT NULL,
            kind TEXT NOT NULL DEFAULT '',
            title TEXT NOT NULL DEFAULT '',
            message TEXT NOT NULL DEFAULT '',
            related_user_id INTEGER,
            created_at REAL NOT NULL,
            read_at REAL
        )""")
        conn.commit()
    finally:
        conn.close()

# Compatibility: this local_store version uses employee_id/password and has no
# create_employee/create_admin helpers. Keep user creation inside app.py.
def _create_user_direct(name, employee_id, password, company, sector, role='Employee', is_admin=False, status='approved'):
    from local_store import _hash_password
    name = str(name or '').strip()
    employee_id = str(employee_id or '').strip()
    password = str(password or '')
    company = str(company or '').strip()
    sector = str(sector or '').strip() or 'General'
    role = str(role or '').strip() or 'Employee'
    status = str(status or '').strip() or 'approved'
    if not name:
        raise ValueError('Name is required.')
    if not employee_id:
        raise ValueError('Email is required.')
    if len(password) < 8:
        raise ValueError('Password must be at least 8 characters.')
    if not company:
        raise ValueError('Company is required.')

    _ensure_registration_schema()
    salt, digest = _hash_password(password)
    conn = _store_conn()
    try:
        cols = _user_table_columns(conn)
        # The UI calls this field email; the original prototype also used
        # employee_id as the login identifier. Store both so old/new schemas
        # remain compatible.
        values = {
            'employee_id': employee_id,
            'password_hash': digest,
            'salt': salt,
            'company': company,
            'sector': sector,
            'role': role,
            'is_admin': 1 if is_admin else 0,
            'created_at': time.time(),
            'name': name,
            'email': employee_id.lower(),
            'status': status,
        }
        ordered = [c for c in values if c in cols]
        placeholders = ','.join('?' for _ in ordered)
        sql = f"INSERT INTO users({','.join(ordered)}) VALUES({placeholders})"
        cur = conn.execute(sql, tuple(values[c] for c in ordered))
        uid = cur.lastrowid
        conn.commit()
        return uid
    finally:
        conn.close()

def _store_create_employee(name, employee_id, password, company, sector, role='Employee', status='approved'):
    return _create_user_direct(name, employee_id, password, company, sector, role, False, status=status)

def _store_create_admin(name, employee_id, password, company):
    return _create_user_direct(name, employee_id, password, company, 'General', 'Administrator', True, status='approved')

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
DATA_ROOT = Path(os.getenv("SOVEREIGN_DATA_ROOT", BASE_DIR / "data"))
UPLOAD_ROOT = DATA_ROOT / "uploads"
GENERATED_ROOT = DATA_ROOT / "generated"
KB_ROOT = BASE_DIR / "knowledge_base"
MAX_UPLOAD_BYTES = int(os.getenv("SOVEREIGN_MAX_UPLOAD_MB", "50")) * 1024 * 1024
MAX_CONTEXT_CHARS = int(os.getenv("SOVEREIGN_MAX_CONTEXT_CHARS", "9000"))
MAX_RETRIEVAL_CHUNKS = int(os.getenv("SOVEREIGN_MAX_RETRIEVAL_CHUNKS", "6"))
MAX_DOCUMENT_CHARS = int(os.getenv("SOVEREIGN_MAX_DOCUMENT_CHARS", "12000000"))
OLLAMA_RETRY_TIMEOUT = int(os.getenv("SOVEREIGN_OLLAMA_RETRY_TIMEOUT", "90"))
OLLAMA_BASE = os.getenv("OLLAMA_BASE", "http://127.0.0.1:11434").rstrip("/")
OLLAMA_TIMEOUT = int(os.getenv("OLLAMA_TIMEOUT", "180"))
OCR_FALLBACK_MAX_PAGES = int(os.getenv("OCR_FALLBACK_MAX_PAGES", "0")) or None
SESSION_COOKIE_SECURE = os.getenv("SESSION_COOKIE_SECURE", "0") == "1"

ALLOWED_EXTENSIONS = {
    "csv", "tsv", "xlsx", "xls", "ods",
    "pdf", "docx", "doc", "pptx", "ppt",
    "txt", "md", "json", "xml", "html", "log",
    "png", "jpg", "jpeg", "webp", "bmp", "tiff",
}

app = Flask(__name__, template_folder="templates", static_folder="static")
app.secret_key = os.getenv("SOVEREIGN_SESSION_SECRET") or secrets.token_hex(32)
app.config.update(
    MAX_CONTENT_LENGTH=MAX_UPLOAD_BYTES,
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=SESSION_COOKIE_SECURE,
)

UPLOAD_ROOT.mkdir(parents=True, exist_ok=True)
GENERATED_ROOT.mkdir(parents=True, exist_ok=True)
(BASE_DIR / "instance").mkdir(exist_ok=True)
init_db()

def _ensure_app_compat_tables():
    _ensure_registration_schema()


def create_registration_request(user_id):
    """Create a pending employee request and notify every approved admin in the same company."""
    conn=_store_conn()
    try:
        user=conn.execute("SELECT id,name,email,company,sector,role FROM users WHERE id=?",(int(user_id),)).fetchone()
        if not user: raise ValueError("User not found.")
        company=str(user["company"] or "")
        now=time.time()
        cur=conn.execute(
            "INSERT INTO registration_requests(user_id,company,status,request_type,message,created_at) VALUES(?,?,?,?,?,?)",
            (int(user_id),company,"pending","employee",
             "New employee registration requires administrator approval.",now),
        )
        conn.execute("UPDATE users SET status='pending' WHERE id=?",(int(user_id),))
        admins=conn.execute(
            "SELECT id FROM users WHERE company=? AND is_admin=1 AND status='approved'",
            (company,),
        ).fetchall()
        display_name=str(user["name"] or user["email"] or "New employee")
        for admin in admins:
            conn.execute(
                "INSERT INTO notifications(recipient_id,kind,title,message,related_user_id,created_at) VALUES(?,?,?,?,?,?)",
                (admin["id"],"employee_registration","Employee approval required",
                 f"{display_name} has registered and is waiting for your approval.",int(user_id),now),
            )
        conn.commit()
        return cur.lastrowid
    finally:
        conn.close()

def get_pending_requests(company):
    conn=_store_conn()
    try:
        rows=conn.execute("""SELECT r.id AS id,r.id AS request_id,r.user_id,r.company,r.status,r.created_at,
            u.employee_id,u.name,u.email,u.company AS user_company,u.sector,u.role,u.is_admin
            FROM registration_requests r JOIN users u ON u.id=r.user_id
            WHERE r.company=? AND r.status='pending' AND COALESCE(u.is_admin,0)=0 AND COALESCE(u.status,'pending')='pending'
            ORDER BY r.created_at ASC""",(str(company or ""),)).fetchall()
        return [dict(x) for x in rows]
    finally: conn.close()

def approve_registration(request_id,reviewer_id,decision):
    d=str(decision or "").strip().lower()
    if d not in {"approve","approved","reject","rejected","deny","denied"}: raise ValueError("Decision must be approve or reject.")
    new_status="approved" if d in {"approve","approved"} else "rejected"
    conn=_store_conn()
    try:
        row=conn.execute("SELECT r.*,u.company AS user_company FROM registration_requests r JOIN users u ON u.id=r.user_id WHERE r.id=?",(int(request_id),)).fetchone()
        reviewer=conn.execute("SELECT company,is_admin FROM users WHERE id=?",(int(reviewer_id),)).fetchone()
        if not row or row["status"]!="pending" or not reviewer or not reviewer["is_admin"] or str(reviewer["company"])!=str(row["company"]): return None
        now=time.time()
        conn.execute("UPDATE registration_requests SET status=?,reviewed_at=?,reviewed_by=? WHERE id=?",(new_status,now,int(reviewer_id),int(request_id)))
        conn.execute("UPDATE users SET status=? WHERE id=?",(new_status,int(row["user_id"])))
        conn.execute(
            "INSERT INTO notifications(recipient_id,kind,title,message,related_user_id,created_at) VALUES(?,?,?,?,?,?)",
            (int(row["user_id"]),"registration_decision",
             "Registration approved" if new_status=="approved" else "Registration rejected",
             ("Your employee registration was approved. You can now sign in and use the workbench."
              if new_status=="approved" else
              "Your employee registration was not approved by the administrator."),
             int(row["user_id"]),now),
        )
        conn.commit()
        out=conn.execute("SELECT * FROM users WHERE id=?",(int(row["user_id"]),)).fetchone(); return dict(out) if out else None
    finally: conn.close()

logging.basicConfig(
    level=os.getenv("SOVEREIGN_LOG_LEVEL", "INFO"),
    format="%(asctime)s %(levelname)s %(name)s %(message)s",
)
logger = logging.getLogger("sovereign")


# ---------------------------------------------------------------------------
# Local-only network policy
# ---------------------------------------------------------------------------
def _validate_local_ollama_url(value: str) -> str:
    """Fail closed: only loopback Ollama endpoints are accepted."""
    from urllib.parse import urlparse
    parsed = urlparse(value)
    if parsed.scheme not in {"http", "https"}:
        raise RuntimeError("OLLAMA_BASE must use HTTP(S) to a loopback address.")
    host = (parsed.hostname or "").lower()
    if host not in {"127.0.0.1", "localhost", "::1"}:
        raise RuntimeError(
            "External Ollama endpoints are blocked. Configure OLLAMA_BASE to "
            "localhost/127.0.0.1 only."
        )
    return value.rstrip("/")


try:
    OLLAMA_BASE = _validate_local_ollama_url(OLLAMA_BASE)
except RuntimeError as exc:
    logger.critical(str(exc))
    # Fail closed instead of silently falling back to a remote endpoint.
    raise

OLLAMA_GENERATE_URL = f"{OLLAMA_BASE}/api/generate"
OLLAMA_TAGS_URL = f"{OLLAMA_BASE}/api/tags"
OLLAMA_SHOW_URL = f"{OLLAMA_BASE}/api/show"


# ---------------------------------------------------------------------------
# Compatibility adapter for the existing local_store.py
# ---------------------------------------------------------------------------
def _chat_for_user(chat_id, user_id):
    try: cid=int(chat_id); uid=int(user_id)
    except (TypeError,ValueError): return None
    for chat in _store_list_chats(uid) or []:
        if int(chat.get("id"))==cid: return dict(chat)
    return None

def _user_can_access_chat(user_id,chat_id): return _chat_for_user(chat_id,user_id) is not None

def _get_file_compat(file_id):
    from local_store import _conn
    conn=_conn()
    try:
        row=conn.execute("SELECT * FROM files WHERE id=?", (int(file_id),)).fetchone()
        return dict(row) if row else None
    finally:
        conn.close()

def _get_team_compat(team_id):
    from local_store import _conn
    conn=_conn()
    try:
        row=conn.execute("SELECT * FROM teams WHERE id=?", (int(team_id),)).fetchone()
        return dict(row) if row else None
    finally:
        conn.close()

get_file=_get_file_compat
get_team=_get_team_compat

def _user_can_access_file(user_id,file_id):
    rec=_get_file_compat(file_id)
    if not rec: return False
    if str(rec.get("owner_id"))==str(user_id): return True
    if rec.get("team_id") is not None: return team_members(rec["team_id"],user_id) is not None
    return bool(rec.get("chat_id") and _user_can_access_chat(user_id,rec["chat_id"]))

def _list_chats_for_user(user_id): return _store_list_chats(user_id) or []
def _list_files_for_chat(user_id,chat_id): return get_chat_files(chat_id,user_id) or []
def _list_messages(chat_id,user_id=None):
    """Return messages with the REAL sender identity.

    Team-chat viewers must never become the displayed sender.  The sender is
    always resolved from messages.sender_id -> users.id, with the stored
    sender_name only used as a fallback.  This keeps the existing frontend
    unchanged while making every member see who actually asked the question.
    """
    viewer_id = user_id or session.get("user_id")
    messages = get_messages(chat_id, viewer_id) or []
    out = []
    for msg in messages:
        item = dict(msg)
        sender_id = item.get("sender_id")
        if item.get("sender_type") == "assistant":
            item["sender_name"] = item.get("sender_name") or "Orion"
            item["sender_email"] = item.get("sender_email") or ""
        elif sender_id is not None:
            try:
                sender = get_user(int(sender_id))
            except Exception:
                sender = None
            if sender:
                item["sender_name"] = str(sender.get("name") or "").strip() or str(sender.get("email") or "").split("@",1)[0] or "User"
                item["sender_email"] = str(sender.get("email") or "")
                item["sender_role"] = sender.get("role") or "Employee"
            else:
                item["sender_name"] = item.get("sender_name") or "User"
        else:
            item["sender_name"] = item.get("sender_name") or "User"
        out.append(item)
    return out
def _create_chat_compat(owner_id,title="New chat",team_id=None):
    cid=_store_create_chat(owner_id,title,team_id)
    return _chat_for_user(cid,owner_id) or {"id":cid,"owner_id":owner_id,"team_id":team_id,"title":title}
def _add_team_member_compat(team_id,*args):
    if len(args)==1: member_id=args[0]
    elif len(args)>=2: member_id=args[-1]
    else: return False
    _store_add_team_member(int(team_id),int(member_id)); return team_members(int(team_id),int(member_id)) is not None
def _create_user_compat(username,password,company,sector="",role="Employee",is_admin=False):
    email=str(username or "").strip().lower(); name=email.split("@",1)[0] if "@" in email else email
    uid=_store_create_admin(name,email,password,company) if is_admin else _store_create_employee(name,email,password,company,sector,role=role,status="approved")
    return get_user(uid) or {}
def _count_users():
    conn=_store_conn()
    try:
        row=conn.execute("SELECT COUNT(*) AS n FROM users").fetchone(); return int(row["n"] if row else 0)
    finally: conn.close()
def _touch_chat(chat_id,owner_id,title=None):
    if not title:
        return
    try:
        from local_store import _conn
        import time as _time
        conn=_conn()
        try:
            conn.execute("UPDATE chats SET title=?, updated_at=? WHERE id=? AND owner_id=?",
                         (str(title).strip() or "New Chat", _time.time(), int(chat_id), int(owner_id)))
            conn.commit()
        finally:
            conn.close()
    except Exception:
        logger.debug("Chat title update skipped",exc_info=True)
add_message=_store_add_message; add_team_member=_add_team_member_compat; create_chat=_create_chat_compat
create_team=lambda name,company,owner_id:_store_create_team(company,name,owner_id)
is_team_member=lambda user_id,team_id: team_members(team_id,user_id) is not None
list_chats_for_user=_list_chats_for_user; list_files_for_chat=_list_files_for_chat; list_messages=_list_messages
user_can_access_chat=_user_can_access_chat; user_can_access_file=_user_can_access_file
touch_chat=lambda chat_id,title=None:_touch_chat(int(chat_id),int(session.get("user_id")),title)
def get_chat(chat_id): return _chat_for_user(chat_id,session.get("user_id"))

# ---------------------------------------------------------------------------
# Authentication / authorization
# ---------------------------------------------------------------------------
def current_user() -> dict[str, Any] | None:
    uid = session.get("user_id")
    return get_user(uid) if uid else None


def login_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        user = current_user()
        if not user:
            return jsonify({"status": "error", "message": "Authentication required."}), 401
        if str(user.get("status") or "approved").lower() != "approved":
            return jsonify({"status": "error", "message": "Account pending administrator approval."}), 403
        return fn(*args, **kwargs)
    return wrapper


def admin_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        user = current_user()
        if not user:
            return jsonify({"status": "error", "message": "Authentication required."}), 401
        if not bool(user.get("is_admin")):
            return jsonify({"status": "error", "message": "Administrator access required."}), 403
        return fn(*args, **kwargs)
    return wrapper


def _authorized_chat(user_id: str, chat_id: str | None) -> dict[str, Any] | None:
    if not chat_id or not user_can_access_chat(user_id, chat_id):
        return None
    return get_chat(chat_id)


# ---------------------------------------------------------------------------
# Dynamic Ollama model registry
# ---------------------------------------------------------------------------
def _ollama_get(path: str, timeout: int = 5) -> dict[str, Any]:
    response = requests.get(f"{OLLAMA_BASE}{path}", timeout=timeout)
    response.raise_for_status()
    return response.json()


def get_model_registry() -> list[dict[str, Any]]:
    """Discover installed models at runtime and inspect their metadata.

    No model name is required or embedded in routing logic. Capability values
    come from Ollama's metadata when available; lightweight name/template
    signals are used only as fallbacks for capabilities Ollama did not expose.
    """
    try:
        tags = _ollama_get("/api/tags")
    except Exception as exc:
        logger.warning("Ollama discovery failed: %s", exc)
        return []

    registry = []
    for item in tags.get("models", []):
        name = str(item.get("name") or item.get("model") or "").strip()
        if not name:
            continue
        record: dict[str, Any] = {
            "name": name,
            "size": int(item.get("size") or 0),
            "family": item.get("details", {}).get("family"),
            "parameter_size": item.get("details", {}).get("parameter_size"),
            "quantization": item.get("details", {}).get("quantization_level"),
            "capabilities": [],
            "vision": False,
            "text": True,
            "code": False,
            "embedding": False,
            "details": item.get("details") or {},
        }
        try:
            shown = None
            response = requests.post(
                OLLAMA_SHOW_URL, json={"name": name}, timeout=5
            )
            if response.ok:
                shown = response.json()
                record["details"].update(shown.get("details") or {})
                caps = shown.get("capabilities") or []
                if isinstance(caps, list):
                    record["capabilities"] = [str(x) for x in caps]
                model_info = shown.get("model_info") or {}
                record["model_info"] = model_info
                template = str(shown.get("template") or "").lower()
                architecture = str(record["details"].get("family") or "").lower()
                name_lower = name.lower()
                cap_text = " ".join(record["capabilities"]).lower()
                record["vision"] = any(x in cap_text for x in ("vision", "image")) or any(
                    x in template + " " + name_lower + " " + architecture
                    for x in ("vision", "vl")
                )
                record["embedding"] = "embedding" in cap_text or "embed" in name_lower
                record["code"] = any(
                    x in (name_lower + " " + architecture + " " + template)
                    for x in ("coder", "code")
                )
        except Exception as exc:
            logger.debug("Model metadata unavailable for %s: %s", name, exc)

        # Models without explicit text capability are still text models unless
        # their metadata clearly identifies them as embeddings-only.
        record["text"] = not record["embedding"] or "generate" in record["capabilities"]
        registry.append(record)

    # Prefer smaller models when otherwise equivalent on constrained hardware.
    registry.sort(key=lambda x: (x["size"] if x["size"] else 10**30, x["name"]))
    return registry


def _env_preference(capability: str) -> list[str]:
    key = f"SOVEREIGN_{capability.upper()}_MODEL"
    raw = os.getenv(key, "").strip()
    return [x.strip() for x in raw.split(",") if x.strip()]


def _preference_names() -> list[str]:
    raw = os.getenv("SOVEREIGN_MODEL_PREFERENCES", "")
    return [x.strip() for x in raw.split(",") if x.strip()]


def resolve_model_for_task(task_type: str, vision: bool = False) -> tuple[str | None, str | None]:
    registry = get_model_registry()
    if not registry:
        return None, "Ollama is unreachable or no local models are installed."

    capability = "vision" if vision else (
        "code" if task_type == "code" else "embedding" if task_type == "embedding" else "text"
    )
    candidates = [m for m in registry if m.get(capability, False)]

    # Explicit preferences are optional and only accepted when the model was
    # discovered locally.
    pref = _env_preference(capability)
    all_pref = _preference_names()
    for desired in pref + all_pref:
        exact = next((m for m in candidates if m["name"] == desired), None)
        if exact:
            return exact["name"], None

    if not candidates:
        if vision:
            return None, "No locally installed vision-capable Ollama model was detected."
        if capability == "code":
            candidates = [m for m in registry if m.get("text")]
        else:
            candidates = [m for m in registry if m.get("text")]
    if not candidates:
        return None, f"No local model supports the required capability: {capability}."

    # Capability first; size second. This makes the result dynamic across
    # installations and avoids assuming any specific model name.
    candidates.sort(key=lambda m: (
        0 if m.get(capability) else 1,
        m.get("size") or 10**30,
        m["name"],
    ))
    return candidates[0]["name"], None


def model_status() -> dict[str, Any]:
    registry = get_model_registry()
    return {
        "local_only": True,
        "ollama_reachable": bool(registry),
        "models": [
            {
                "name": m["name"],
                "size": m["size"],
                "family": m.get("family"),
                "parameter_size": m.get("parameter_size"),
                "quantization": m.get("quantization"),
                "capabilities": m.get("capabilities", []),
                "vision": m.get("vision", False),
                "text": m.get("text", False),
                "code": m.get("code", False),
                "embedding": m.get("embedding", False),
            }
            for m in registry
        ],
    }


def query_ollama(
    prompt: str,
    images: list[str] | None = None,
    vision: bool = False,
    task_type: str = "general",
    model: str | None = None,
    timeout: int = OLLAMA_TIMEOUT,
) -> tuple[str | None, str | None, str | None]:
    if model is None:
        model, err = resolve_model_for_task(task_type, vision=vision)
        if err:
            return None, err, None

    # A model must be present in the runtime registry; never accept an
    # arbitrary model string supplied by a client.
    installed = {m["name"] for m in get_model_registry()}
    if model not in installed:
        return None, "Requested model is not installed locally.", None

    keep_alive=os.getenv("SOVEREIGN_OLLAMA_KEEP_ALIVE","10m").strip()
    if not re.fullmatch(r"(?:0|[0-9]+(?:ns|us|µs|ms|s|m|h))",keep_alive):
        logger.warning("Invalid SOVEREIGN_OLLAMA_KEEP_ALIVE=%r; using 10m",keep_alive); keep_alive="10m"
    payload: dict[str, Any] = {"model":model,"prompt":prompt,"stream":False,"keep_alive":keep_alive}
    if images and vision:
        payload["images"] = images

    try:
        logger.info("Local model call task=%s vision=%s model=%s", task_type, vision, model)
        response = requests.post(OLLAMA_GENERATE_URL, json=payload, timeout=timeout)
        if response.ok:
            return response.json().get("response", ""), None, model
        return None, f"Ollama returned HTTP {response.status_code}.", model
    except requests.RequestException as exc:
        logger.warning("Ollama request failed: %s", exc)
        return None, "Ollama timeout or connection issue.", model


# ---------------------------------------------------------------------------
# Generic task/capability classification
# ---------------------------------------------------------------------------
def classify_task(question: str) -> str:
    q = (question or "").lower()
    if any(x in q for x in (
        "diagram", "flowchart", "workflow", "architecture", "sequence diagram",
        "network diagram", "process flow", "web diagram", "draw a diagram",
    )):
        return "diagram"
    if any(x in q for x in (
        "code", "script", "program", "debug", "function", "class ", "regex",
        "sql", "python", "javascript", "typescript", "compile", "stack trace",
    )):
        return "code"
    if any(x in q for x in (
        "chart", "graph", "plot", "trend", "visualize", "visualise",
    )):
        return "data"
    return "text"


def detect_chart_need(question: str) -> bool:
    q = (question or "").lower()
    if detect_diagram_need(q):
        return False
    return any(x in q for x in (
        "chart", "graph", "plot", "trend", "visualize", "visualise",
        "bar chart", "line chart", "pie chart", "scatter", "area chart",
        "compare", "breakdown",
    ))


def detect_diagram_need(question: str) -> bool:
    q = (question or "").lower()
    return any(x in q for x in (
        "diagram", "flowchart", "flow chart", "workflow", "architecture",
        "process flow", "process diagram", "system diagram", "network diagram",
        "block diagram", "sequence diagram", "mind map", "web diagram",
        "user flow", "screen flow", "page flow", "app flow",
    ))


# ---------------------------------------------------------------------------
# File handling / extraction
# ---------------------------------------------------------------------------
def allowed_file(filename: str) -> bool:
    return Path(filename).suffix.lower().lstrip(".") in ALLOWED_EXTENSIONS


def _safe_upload_path(user_id: str, chat_id: str, original_name: str) -> Path:
    ext = Path(secure_filename(original_name)).suffix.lower()
    target_dir = UPLOAD_ROOT / str(user_id) / str(chat_id)
    target_dir.mkdir(parents=True, exist_ok=True)
    return target_dir / f"{uuid.uuid4().hex}{ext}"


def _read_text_file(path: Path) -> str:
    return path.read_text(encoding="utf-8", errors="replace")[:MAX_CONTEXT_CHARS]


def _safe_text_from_pdf(path: Path) -> str:
    """Extract a PDF locally without sending every page to Ollama.

    Text PDFs are extracted page-by-page and retained with page markers so the
    retrieval layer can find the relevant passages later. Scanned PDFs use
    optional local Tesseract OCR; they never trigger one Ollama vision request
    per page. This is critical for large documents (hundreds of pages).
    """
    if fitz is None:
        raise RuntimeError("PyMuPDF is not installed.")

    doc = fitz.open(path)
    try:
        page_parts = []
        text_chars = 0
        pages_with_text = 0

        for i in range(doc.page_count):
            page_text = (doc.load_page(i).get_text("text") or "").strip()
            if page_text:
                pages_with_text += 1
            part = f"[PDF page {i + 1}]\n{page_text}"
            page_parts.append(part)
            text_chars += len(page_text)
            if text_chars >= MAX_DOCUMENT_CHARS:
                page_parts.append(
                    f"[Document extraction limit reached at page {i + 1} of {doc.page_count}.]"
                )
                break

        raw = "\n\n".join(page_parts)

        # A normal text PDF: return the complete locally extracted document.
        if text_chars >= 200 and pages_with_text >= max(1, min(doc.page_count, 3) // 2):
            return raw

        # Scanned/image-only PDF: optional LOCAL OCR only.
        if pytesseract is None:
            return (
                raw
                + "\n\n[Scanned PDF detected. Local Tesseract OCR is not installed; "
                  "no remote/cloud OCR or per-page Ollama calls were made.]"
            )

        ocr_parts = []
        max_pages = min(doc.page_count, int(os.getenv("SOVEREIGN_OCR_MAX_PAGES", "500")))
        for i in range(max_pages):
            page = doc.load_page(i)
            pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5), alpha=False)
            try:
                from PIL import Image as _Image
                image = _Image.open(io.BytesIO(pix.tobytes("png")))
                ocr_text = pytesseract.image_to_string(image).strip()
            except Exception as exc:
                ocr_text = f"[Local OCR failed on page {i + 1}: {exc}]"
            ocr_parts.append(f"[PDF page {i + 1} — local OCR]\n{ocr_text}")
            if sum(len(x) for x in ocr_parts) >= MAX_DOCUMENT_CHARS:
                break

        if max_pages < doc.page_count:
            ocr_parts.append(
                f"[OCR limit: first {max_pages} of {doc.page_count} scanned pages were processed locally.]"
            )
        return "\n\n".join(ocr_parts)
    finally:
        doc.close()


def _safe_text_from_docx(path: Path) -> str:
    if docx is None:
        raise RuntimeError("python-docx is not installed.")
    d = docx.Document(path)
    parts = [p.text.strip() for p in d.paragraphs if p.text.strip()]
    for table in d.tables:
        for row in table.rows:
            vals = [cell.text.strip() for cell in row.cells]
            if any(vals):
                parts.append(" | ".join(vals))
    return "\n".join(parts)


def _convert_legacy_with_libreoffice(path: Path) -> Path:
    """Convert legacy DOC/PPT locally. Returns a temporary converted file."""
    import shutil
    executable = shutil.which("libreoffice") or shutil.which("soffice")
    if not executable:
        raise RuntimeError(
            "LibreOffice is required to read legacy DOC/PPT files locally. "
            "Convert the file to DOCX/PPTX or install LibreOffice."
        )
    out_dir = Path(tempfile.mkdtemp(prefix="sovereign_convert_"))
    subprocess.run(
        [executable, "--headless", "--convert-to", "docx" if path.suffix.lower() == ".doc" else "pptx",
         "--outdir", str(out_dir), str(path)],
        check=True, timeout=120, capture_output=True,
    )
    converted = out_dir / (path.stem + (".docx" if path.suffix.lower() == ".doc" else ".pptx"))
    if not converted.exists():
        raise RuntimeError("LibreOffice conversion did not produce a readable local file.")
    return converted


def _safe_text_from_pptx(path: Path) -> str:
    if Presentation is None:
        raise RuntimeError("python-pptx is not installed.")
    prs = Presentation(path)
    slides = []
    for idx, slide in enumerate(prs.slides, 1):
        texts = []
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text.strip():
                texts.append(shape.text.strip())
        if texts:
            slides.append(f"[PPTX slide {idx}]\n" + "\n".join(texts))
    return "\n\n".join(slides)


def _smart_read_excel(path: Path, query: str = ""):
    if pd is None:
        raise RuntimeError("pandas is not installed.")
    xls = pd.ExcelFile(path)
    if not xls.sheet_names:
        raise RuntimeError("Workbook has no sheets.")
    q = (query or "").casefold()

    # Choose a sheet by actual sheet name when explicitly mentioned; otherwise
    # inspect the first non-empty sheet. No domain assumptions are used.
    selected = next((s for s in xls.sheet_names if s.casefold() in q), None)
    if selected is None:
        for s in xls.sheet_names:
            preview = pd.read_excel(path, sheet_name=s, header=None, nrows=20)
            if not preview.dropna(how="all").empty:
                selected = s
                break
    selected = selected or xls.sheet_names[0]

    raw = pd.read_excel(path, sheet_name=selected, header=None)
    if raw.empty:
        return raw, selected, xls.sheet_names

    # Find a likely header row using text density and uniqueness, not a fixed
    # domain schema.
    candidates = []
    for i in range(min(len(raw), 50)):
        row = raw.iloc[i]
        nonempty = row.dropna().astype(str).str.strip()
        if len(nonempty) < 1:
            continue
        unique_ratio = nonempty.nunique() / max(len(nonempty), 1)
        numeric_count = pd.to_numeric(row, errors="coerce").notna().sum()
        candidates.append((i, len(nonempty), unique_ratio, numeric_count))
    header_row = max(
        candidates,
        key=lambda x: (x[1], x[2], -x[3]),
        default=(0, 0, 0, 0),
    )[0]

    df = pd.read_excel(path, sheet_name=selected, header=header_row)
    df = df.dropna(axis=1, how="all").dropna(how="all").reset_index(drop=True)
    df.columns = [str(c).strip() if str(c).strip() else f"Column {i+1}"
                  for i, c in enumerate(df.columns)]
    return df, selected, xls.sheet_names


def _safe_text_from_excel(path: Path, query: str = "") -> str:
    if pd is None:
        raise RuntimeError("pandas is not installed.")
    xls = pd.ExcelFile(path)
    parts = [f"[Workbook] sheets={xls.sheet_names}"]
    for sheet in xls.sheet_names:
        df, _, _ = _smart_read_excel(path, query=sheet if sheet in (query or "") else "")
        parts.append(
            f"[Sheet: {sheet}] rows={len(df)} columns={list(df.columns)}\n"
            f"{df.to_string(index=False, max_rows=300)}"
        )
    return "\n\n".join(parts)[:MAX_CONTEXT_CHARS]


def _safe_text_from_tabular(path: Path) -> str:
    if pd is None:
        raise RuntimeError("pandas is not installed.")
    if path.suffix.lower() == ".tsv":
        df = pd.read_csv(path, sep="\t")
    else:
        df = pd.read_csv(path)
    df.columns = [str(c).strip() for c in df.columns]
    return f"[CSV/TSV] rows={len(df)} columns={list(df.columns)}\n{df.to_string(index=False, max_rows=500)}"[:MAX_CONTEXT_CHARS]


def _safe_image(path: Path) -> tuple[str, str]:
    if PILImage is None:
        raise RuntimeError("Pillow is not installed.")
    image = PILImage.open(path)
    if image.mode not in ("RGB", "L"):
        image = image.convert("RGB")
    image.thumbnail((1800, 1800))
    buf = io.BytesIO()
    image.save(buf, format="JPEG", quality=90)
    return base64.b64encode(buf.getvalue()).decode(), "image/jpeg"


def extract_document_text(file_path: str, filename: str, query: str = "") -> tuple[str, bool, str | None]:
    path = Path(file_path)
    ext = path.suffix.lower()
    if ext == ".pdf":
        return _safe_text_from_pdf(path), False, None
    if ext == ".docx":
        return _safe_text_from_docx(path), False, None
    if ext == ".doc":
        converted = _convert_legacy_with_libreoffice(path)
        return _safe_text_from_docx(converted), False, None
    if ext == ".pptx":
        return _safe_text_from_pptx(path), False, None
    if ext == ".ppt":
        converted = _convert_legacy_with_libreoffice(path)
        return _safe_text_from_pptx(converted), False, None
    if ext in {".xlsx", ".xls", ".ods"}:
        return _safe_text_from_excel(path, query=query), False, None
    if ext in {".csv", ".tsv"}:
        return _safe_text_from_tabular(path), False, None
    if ext in {".md", ".txt", ".json", ".xml", ".html", ".log"}:
        text = _read_text_file(path)
        if ext == ".html":
            text = re.sub(r"<[^>]+>", " ", text)
            text = re.sub(r"\s+", " ", text)
        return text, False, None
    if ext in {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tiff"}:
        b64, _ = _safe_image(path)
        prompt = (
            "Transcribe only visible text and describe visible structure. "
            "Do not invent labels, numbers, connections, or meanings. Mark unreadable "
            "content as [unreadable]. For handwriting provide readable text, uncertain "
            "text, and visible structure separately."
        )
        transcript, err, _ = query_ollama(prompt, images=[b64], vision=True, task_type="vision")
        return transcript or f"[Visual transcription unavailable: {err}]", True, b64
    raise RuntimeError(f"Unsupported file type: {ext}")


# ---------------------------------------------------------------------------
# Document-intent helpers
# ---------------------------------------------------------------------------
def is_whole_document_analysis_request(query: str) -> bool:
    """Return True for broad requests that ask Orion to analyze the file itself.

    These requests must not be treated as literal keyword searches.  Examples:
    "analyse the file", "analyze this document", "review the PDF", "summarize
    the uploaded report", and "give me an overview" all mean document-level
    analysis rather than a search for those exact words.
    """
    q = re.sub(r"\s+", " ", (query or "").strip().casefold())
    if not q:
        return True
    # Remove harmless references to the uploaded object and inspect the intent.
    normalized = re.sub(
        r"\b(this|the|uploaded|attached|provided|given|following|current)\b",
        " ", q,
    )
    normalized = re.sub(
        r"\b(file|document|pdf|report|paper|spreadsheet|workbook|source|sources)\b",
        " ", normalized,
    )
    normalized = re.sub(r"[^a-z0-9]+", " ", normalized).strip()
    broad_verbs = (
        r"analyse", r"analyze", r"analysis", r"review", r"summarize",
        r"summarise", r"overview", r"explain", r"assess", r"evaluate",
        r"digest", r"study", r"understand", r"give me a summary",
    )
    if any(re.search(rf"\b{v}\b", normalized) for v in broad_verbs):
        # Do not classify a targeted question such as "analyze revenue for AAPL"
        # as whole-document analysis.
        targeted = re.search(
            r"\b(for|about|regarding|on|of)\b\s+\S+", normalized
        )
        return not bool(targeted)
    # Common exact broad phrasings after file-word removal.
    return normalized in {"", "the"}


def _document_page_count(text: str) -> int:
    return len(re.findall(r"\[PDF page \d+", text or ""))


def build_whole_document_context(document_text: str, query: str) -> str:
    """Build distributed evidence for a whole-document analysis.

    The complete document remains local.  We combine the beginning, end,
    evenly-distributed sections, and the strongest query-relevant chunks. This
    avoids sending hundreds of pages to Ollama while preventing the old failure
    mode where a broad query retrieves only the first few keyword matches.
    """
    if not document_text:
        return ""
    chunks = _chunk_text(document_text, size=2300, overlap=250)
    if not chunks:
        return document_text[:MAX_CONTEXT_CHARS]

    selected: list[int] = []
    # Beginning and end establish identity, purpose, conclusions and appendices.
    selected.extend([0, min(1, len(chunks) - 1), max(0, len(chunks) - 2), len(chunks) - 1])

    # Sample the entire document rather than clustering on one keyword.
    target_distributed = min(10, len(chunks))
    if target_distributed > 1:
        for i in range(target_distributed):
            idx = round(i * (len(chunks) - 1) / (target_distributed - 1))
            selected.append(idx)

    # Add query-relevant chunks as a separate signal.
    try:
        from sklearn.feature_extraction.text import TfidfVectorizer
        from sklearn.metrics.pairwise import cosine_similarity
        vectorizer = TfidfVectorizer(stop_words="english", max_features=16000, ngram_range=(1, 2))
        matrix = vectorizer.fit_transform(chunks)
        scores = cosine_similarity(vectorizer.transform([query or "overview"]), matrix).flatten()
        selected.extend(sorted(range(len(chunks)), key=lambda i: float(scores[i]), reverse=True)[:6])
    except Exception:
        q_tokens = set(re.findall(r"[a-zA-Z0-9]{2,}", (query or "").casefold()))
        scored = []
        for i, chunk in enumerate(chunks):
            tokens = set(re.findall(r"[a-zA-Z0-9]{2,}", chunk.casefold()))
            scored.append((len(q_tokens & tokens), i))
        selected.extend(i for _, i in sorted(scored, reverse=True)[:6])

    selected = sorted(set(i for i in selected if 0 <= i < len(chunks)))
    # Prefer one compact representative from each region when the context cap is
    # tight.  Keep source/page markers intact.
    pieces = []
    budget = min(MAX_CONTEXT_CHARS, 12000)
    used = 0
    for i in selected:
        piece = chunks[i].strip()
        if not piece:
            continue
        if used + len(piece) > budget:
            remaining = budget - used
            if remaining > 500:
                pieces.append(piece[:remaining])
            break
        pieces.append(piece)
        used += len(piece)
    return "\n\n[... distributed local evidence gap ...]\n\n".join(pieces)


def build_local_document_overview(document_text: str, filename: str) -> str:
    """Deterministic fallback when the local model is unavailable."""
    text = document_text or ""
    pages = _document_page_count(text)
    chars = len(text)
    headings = []
    for line in text.splitlines():
        clean = re.sub(r"\s+", " ", line).strip()
        if not clean or len(clean) > 140:
            continue
        if re.match(r"^(?:[A-Z][A-Z0-9 &,:()'\-]{5,}|\d+[.)]\s+.+)$", clean):
            if clean not in headings:
                headings.append(clean)
        if len(headings) >= 20:
            break
    keywords = {
        "financial statements": len(re.findall(r"\bfinancial statements?\b", text, re.I)),
        "annual general meeting": len(re.findall(r"\bannual general meeting\b|\bAGM\b", text, re.I)),
        "board/directors": len(re.findall(r"\bboard of directors?\b|\bdirectors?\b", text, re.I)),
        "auditor/audit": len(re.findall(r"\baudit(?:or|ed|ing)?\b", text, re.I)),
        "corporate governance": len(re.findall(r"\bcorporate governance\b", text, re.I)),
        "CSR": len(re.findall(r"\bCSR\b|corporate social responsibility", text, re.I)),
        "risk": len(re.findall(r"\brisk(?:s)?\b", text, re.I)),
    }
    lines = [
        f"Local document analysis for: {filename}",
        f"Extracted locally: approximately {pages or 'unknown'} PDF pages, {chars:,} characters.",
        "No cloud service was used.",
        "Detected topic signals: " + ", ".join(f"{k} ({v})" for k, v in keywords.items() if v) or "No common topic signals detected.",
    ]
    if headings:
        lines.append("Representative headings/sections detected locally:")
        lines.extend(f"- {h}" for h in headings[:20])
    evidence = retrieve_relevant_passages(document_text, "company overview financial performance directors governance audit AGM conclusions", top_k=6)
    if evidence:
        lines.append("Representative local evidence:")
        lines.append(evidence[:5000])
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Local retrieval
# ---------------------------------------------------------------------------
def _chunk_text(text: str, size: int = 1800, overlap: int = 250) -> list[str]:
    chunks, start = [], 0
    while start < len(text):
        chunk = text[start:start + size]
        if chunk.strip():
            chunks.append(chunk)
        start += max(1, size - overlap)
    return chunks


def retrieve_relevant_passages(document_text: str, query: str, top_k: int | None = None) -> str:
    """Retrieve a small evidence set from potentially huge local documents."""
    if not document_text:
        return ""
    limit = top_k or MAX_RETRIEVAL_CHUNKS
    if len(document_text) <= MAX_CONTEXT_CHARS:
        return document_text

    chunks = _chunk_text(document_text, size=2200, overlap=300)
    if not chunks:
        return document_text[:MAX_CONTEXT_CHARS]

    try:
        from sklearn.feature_extraction.text import TfidfVectorizer
        from sklearn.metrics.pairwise import cosine_similarity

        vectorizer = TfidfVectorizer(
            stop_words="english",
            max_features=12000,
            ngram_range=(1, 2),
        )
        matrix = vectorizer.fit_transform(chunks)
        scores = cosine_similarity(vectorizer.transform([query or ""]), matrix).flatten()
        ranked = sorted(range(len(chunks)), key=lambda i: float(scores[i]), reverse=True)[:limit]
    except Exception:
        # Deterministic fallback: keyword overlap without an ML dependency.
        q_tokens = set(re.findall(r"[a-zA-Z0-9]{2,}", (query or "").casefold()))
        scored = []
        for i, chunk in enumerate(chunks):
            tokens = set(re.findall(r"[a-zA-Z0-9]{2,}", chunk.casefold()))
            scored.append((len(q_tokens & tokens), i))
        ranked = [i for _, i in sorted(scored, reverse=True)[:limit]]

    # Keep document structure: beginning/end plus the best evidence chunks.
    for idx in (0, len(chunks) - 1):
        if idx not in ranked:
            ranked.append(idx)

    selected = [chunks[i] for i in sorted(set(ranked))]
    result = "\n\n[... local retrieval gap ...]\n\n".join(selected)
    return result[:MAX_CONTEXT_CHARS]


# ---------------------------------------------------------------------------
# Deterministic tabular reasoning
# ---------------------------------------------------------------------------
def _load_dataframe(path: Path, query: str = ""):
    if pd is None:
        raise RuntimeError("pandas is not installed.")
    ext = path.suffix.lower()
    if ext == ".csv":
        return pd.read_csv(path)
    if ext == ".tsv":
        return pd.read_csv(path, sep="\t")
    if ext in {".xlsx", ".xls", ".ods"}:
        return _smart_read_excel(path, query=query)[0]
    raise ValueError("Not a supported tabular file.")


def _column_candidates(columns: list[str], query: str) -> list[str]:
    q = re.sub(r"[^a-z0-9]+", " ", (query or "").casefold()).strip()
    q_tokens = set(q.split())
    scored = []
    for col in columns:
        c = re.sub(r"[^a-z0-9]+", " ", str(col).casefold()).strip()
        c_tokens = set(c.split())
        exact = 1 if c and c in q else 0
        overlap = len(c_tokens & q_tokens)
        scored.append((exact, overlap, -abs(len(c_tokens) - len(q_tokens)), col))
    scored.sort(reverse=True)
    best = [x[3] for x in scored if x[0] or x[1] > 0]
    return best


def build_tabular_schema_prompt(columns: list[str], query: str) -> str:
    return (
        "Map the user's request to the ACTUAL dataframe schema below. Return ONLY JSON.\n"
        '{"operation":"sum|average|median|min|max|count|difference|ratio|percentage_change|'
        'trend|comparison|group_by|sort|filter|select|none",'
        '"column":null,"column_b":null,"group_by":null,"filter_column":null,'
        '"filter_value":null,"period_column":null,"ascending":true}\n'
        "Rules: every non-null column field MUST exactly equal one of the supplied "
        "columns. Never invent a column. If the request is ambiguous, use operation "
        "'none'. Do not calculate any values.\n\n"
        f"Columns: {json.dumps([str(c) for c in columns], ensure_ascii=False)}\n"
        f"User request: {query}"
    )


def _parse_json_object(text: str) -> dict | None:
    if not text:
        return None
    text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text.strip(), flags=re.I)
    start = text.find("{")
    if start < 0:
        return None
    depth, quoted, escaped = 0, False, False
    for i in range(start, len(text)):
        ch = text[i]
        if ch == '"' and not escaped:
            quoted = not quoted
        escaped = (ch == "\\" and not escaped)
        if quoted:
            continue
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                try:
                    obj = json.loads(text[start:i + 1])
                    return obj if isinstance(obj, dict) else None
                except json.JSONDecodeError:
                    return None
    return None


def _validate_schema_spec(spec: dict, columns: list[str]) -> bool:
    allowed = set(map(str, columns))
    for key in ("column", "column_b", "group_by", "filter_column", "period_column"):
        value = spec.get(key)
        if value is not None and str(value) not in allowed:
            return False
    return spec.get("operation") in {
        "sum", "average", "median", "min", "max", "count", "difference",
        "ratio", "percentage_change", "trend", "comparison", "group_by",
        "sort", "filter", "select", "none",
    }


def _find_entity_filter(df, query: str) -> tuple[str | None, Any | None, bool]:
    """Find an entity/category directly from categorical values.

    Exact normalized matching is preferred. Fuzzy matching is only accepted
    when exactly one categorical value is clearly closest.
    """
    q = re.sub(r"[^a-z0-9]+", " ", (query or "").casefold()).strip()
    candidates = []
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            continue
        vals = [v for v in df[col].dropna().astype(str).unique() if v.strip()]
        for value in vals:
            norm = re.sub(r"[^a-z0-9]+", " ", value.casefold()).strip()
            if norm and (norm in q or q in norm):
                candidates.append((col, value, 1.0))
    if len(candidates) == 1:
        return candidates[0]
    if len(candidates) > 1:
        # Only accept the most specific exact containment if it is unique.
        candidates.sort(key=lambda x: len(str(x[1])), reverse=True)
        if len(candidates) == 1 or len(str(candidates[0][1])) > len(str(candidates[1][1])) + 2:
            return candidates[0]
        return None, None, True
    return None, None, False


def deterministic_dataframe_answer(df, query: str, spec: dict) -> str | None:
    op = spec.get("operation")
    if op in {None, "none"}:
        return None
    col = spec.get("column")
    col_b = spec.get("column_b")
    group_by = spec.get("group_by")
    filter_col = spec.get("filter_column")
    filter_value = spec.get("filter_value")
    period_col = spec.get("period_column")

    if op not in {"count", "group_by", "sort", "select", "filter"} and not col:
        return None
    work = df.copy()

    if filter_col:
        if filter_col not in work.columns:
            return None
        work = work[work[filter_col].astype(str) == str(filter_value)]
    else:
        entity_col, entity_value, ambiguous = _find_entity_filter(work, query)
        if ambiguous:
            return "AMBIGUOUS_ENTITY"
        if entity_col and entity_value:
            work = work[work[entity_col].astype(str) == str(entity_value)]

    if work.empty:
        return "No matching rows were found in the uploaded data."

    if op == "count":
        return f"Count of matching rows in the uploaded data: {len(work)}."

    if op in {"sum", "average", "median", "min", "max"}:
        numeric = pd.to_numeric(work[col], errors="coerce").dropna()
        if numeric.empty:
            return f"Column '{col}' contains no numeric values that can be calculated."
        funcs = {
            "sum": numeric.sum, "average": numeric.mean, "median": numeric.median,
            "min": numeric.min, "max": numeric.max,
        }
        value = funcs[op]()
        return f"{op.capitalize()} of '{col}' using the uploaded data: {value}. Actual column used: '{col}'."

    if op == "difference":
        if not col_b:
            return None
        a = pd.to_numeric(work[col], errors="coerce").dropna()
        b = pd.to_numeric(work[col_b], errors="coerce").dropna()
        if len(a) != len(b) or len(a) == 0:
            return None
        value = a.reset_index(drop=True) - b.reset_index(drop=True)
        return f"Difference between '{col}' and '{col_b}' (row-aligned): {value.tolist()}."

    if op == "ratio":
        if not col_b:
            return None
        a = pd.to_numeric(work[col], errors="coerce")
        b = pd.to_numeric(work[col_b], errors="coerce").replace(0, pd.NA)
        ratio = (a / b).dropna()
        return f"Ratio '{col}' / '{col_b}' (row-aligned): {ratio.tolist()}."

    if op == "percentage_change":
        if not col_b:
            return None
        a = pd.to_numeric(work[col], errors="coerce")
        b = pd.to_numeric(work[col_b], errors="coerce").replace(0, pd.NA)
        pct = ((a - b) / b * 100).dropna()
        return f"Percentage change from '{col_b}' to '{col}': {pct.tolist()}%."

    if op in {"group_by", "comparison"}:
        if not group_by:
            return None
        grouped = work.groupby(group_by, dropna=False)[col].agg(["count", "sum", "mean"]).reset_index()
        return (
            f"Grouped by actual column '{group_by}', using '{col}':\n"
            + grouped.to_string(index=False)
        )

    if op == "sort":
        ascending = bool(spec.get("ascending", True))
        result = work.sort_values(col, ascending=ascending)
        return f"Sorted by actual column '{col}' ({'ascending' if ascending else 'descending'}):\n{result.to_string(index=False)}"

    if op in {"select", "filter"}:
        cols = [x for x in (col, col_b, group_by) if x and x in work.columns]
        if not cols:
            return None
        return f"Selected actual columns {cols}:\n{work[cols].to_string(index=False)}"

    return None


def answer_from_dataframe(path: Path, query: str) -> tuple[str | None, dict | None]:
    """Answer tabular questions locally without calling Ollama.

    Uploaded CSV/Excel questions must remain deterministic and available even
    when Ollama is slow, stopped, or unavailable. The schema and entity/metric
    selection are resolved from the actual dataframe columns and values.
    """
    try:
        df = _load_dataframe(path, query)
    except Exception:
        return None, None
    if df.empty:
        return "The uploaded table is empty.", None
    df.columns = [str(c).strip() for c in df.columns]

    q = (query or "").casefold()

    # Resolve an explicit entity/ticker first. This prevents the entire
    # multi-company dataset from being used for a company-specific request.
    entity_col, entity_value, ambiguous = _find_entity_filter(df, query)
    if ambiguous:
        # If the query contains a recognizable ticker/company token, prefer an
        # exact match over the generic ambiguity guard.
        explicit = None
        for candidate_col in [c for c in df.columns if str(c).casefold() in {"company", "ticker", "symbol"}]:
            vals = df[candidate_col].dropna().astype(str).str.strip()
            for value in vals.unique():
                if re.search(rf"(?<![a-z0-9]){re.escape(str(value).casefold())}(?![a-z0-9])", q):
                    explicit = (candidate_col, value)
                    break
            if explicit:
                break
        if explicit:
            entity_col, entity_value = explicit
            ambiguous = False

    work = df.copy()
    if not ambiguous and entity_col and entity_value is not None:
        work = work[work[entity_col].astype(str).str.strip().str.casefold() == str(entity_value).strip().casefold()].copy()

    if ambiguous:
        return "The uploaded data contains multiple possible entity/category matches. Please specify which value you mean.", None
    if work.empty:
        return f"No matching rows were found for {entity_value or 'the requested entity'}.", None

    numeric = [c for c in work.columns if pd.api.types.is_numeric_dtype(work[c])]
    if not numeric:
        return None, None

    # Exact semantic metric matching for common financial questions.
    metric_aliases = {
        "revenue": ["revenue"],
        "net income": ["net income", "netincome"],
        "gross profit": ["gross profit"],
        "ebitda": ["ebitda"],
        "net profit margin": ["net profit margin"],
        "market cap": ["market cap", "market cap(in b usd)", "market cap in b usd"],
    }
    selected = []
    for canonical, aliases in metric_aliases.items():
        if any(re.search(rf"(?<![a-z0-9]){re.escape(alias)}(?![a-z0-9])", q) for alias in aliases):
            for col in numeric:
                cl = str(col).casefold().strip()
                if cl == canonical or cl in aliases:
                    selected.append(col)
                    break

    # Support "X and Y" / "compare X vs Y" even if a metric has unusual
    # capitalization in the source.
    if len(selected) < 2:
        for col in numeric:
            cl = str(col).casefold()
            if re.search(rf"(?<![a-z0-9]){re.escape(cl)}(?![a-z0-9])", q) and col not in selected:
                selected.append(col)

    is_compare = bool(re.search(r"\b(compare|versus|vs\.?|against)\b", q))
    if is_compare and len(selected) >= 2:
        selected = selected[:3]
        period_col = next(
            (c for c in work.columns if str(c).casefold().strip() in {"year", "date", "period"}),
            None,
        )
        if period_col:
            try:
                ordered = work.sort_values(period_col)
            except Exception:
                ordered = work
            lines = [
                f"{str(entity_value).strip() if entity_value is not None else 'Uploaded data'} — comparison of "
                + " and ".join(str(c) for c in selected) + ":"
            ]
            for _, row in ordered.iterrows():
                vals = []
                for col in selected:
                    val = pd.to_numeric(pd.Series([row[col]]), errors="coerce").iloc[0]
                    vals.append(f"{col}={val:g}" if pd.notna(val) else f"{col}=N/A")
                lines.append(f"{row[period_col]}: " + ", ".join(vals))
            spec = {
                "operation": "comparison",
                "columns": selected,
                "period_column": period_col,
                "entity_column": entity_col,
                "entity_value": entity_value,
            }
            return "\n".join(lines), spec

    # Simple deterministic aggregations.
    operation = (
        "average" if re.search(r"\b(average|mean|avg)\b", q) else
        "sum" if re.search(r"\b(sum|total)\b", q) else
        "median" if "median" in q else
        "min" if re.search(r"\b(min|minimum|lowest)\b", q) else
        "max" if re.search(r"\b(max|maximum|highest|greatest)\b", q) else
        "count" if re.search(r"\b(count|how many|number of rows)\b", q) else None
    )
    candidates = _column_candidates(list(work.columns), query)
    if operation and len(candidates) == 1:
        spec = {
            "operation": operation,
            "column": candidates[0],
            "entity_column": entity_col,
            "entity_value": entity_value,
        }
        result = deterministic_dataframe_answer(work, query, spec)
        return result, spec
    if operation == "count":
        spec = {"operation": "count", "entity_column": entity_col, "entity_value": entity_value}
        return deterministic_dataframe_answer(work, query, spec), spec

    # For a direct metric request, return the selected column(s) rather than
    # falling through to Ollama.
    if selected:
        period_col = next(
            (c for c in work.columns if str(c).casefold().strip() in {"year", "date", "period"}),
            None,
        )
        if period_col:
            ordered = work.sort_values(period_col)
            lines = [f"{str(entity_value).strip() if entity_value is not None else 'Uploaded data'}:"]
            for _, row in ordered.iterrows():
                vals = []
                for col in selected:
                    val = pd.to_numeric(pd.Series([row[col]]), errors="coerce").iloc[0]
                    vals.append(f"{col}={val:g}" if pd.notna(val) else f"{col}=N/A")
                lines.append(f"{row[period_col]}: " + ", ".join(vals))
            return "\n".join(lines), {
                "operation": "select",
                "columns": selected,
                "period_column": period_col,
                "entity_column": entity_col,
                "entity_value": entity_value,
            }

    return None, None


# ---------------------------------------------------------------------------
# Charts: validated schema + deterministic rendering
# ---------------------------------------------------------------------------
def build_chart_spec_prompt(columns: list[str], query: str) -> str:
    return (
        "Choose a chart from the ACTUAL dataframe columns. Return ONLY JSON:\n"
        '{"chart_type":"bar|line|scatter|pie|area|step",'
        '"category":"actual column or null","series":["actual numeric column"],'
        '"title":"short title"}\n'
        "Every column must exactly match the supplied list. Never invent data. "
        "Use line/step when a period-like category is available; scatter requires "
        "numeric category; pie requires a composition-like category plus numeric series. "
        "If the request is not chartable, return chart_type 'none'.\n\n"
        f"Columns: {json.dumps(columns, ensure_ascii=False)}\nRequest: {query}"
    )


def _chart_requested_entities(df, query: str, entity_col: str | None) -> list[str]:
    """Return company/ticker values explicitly mentioned in the request."""
    if not entity_col or entity_col not in df.columns:
        return []
    q = (query or "").casefold()
    values = []
    seen = set()
    for raw in df[entity_col].dropna().astype(str).str.strip().unique():
        if not raw:
            continue
        if re.search(rf"(?<![a-z0-9]){re.escape(raw.casefold())}(?![a-z0-9])", q):
            key = raw.casefold()
            if key not in seen:
                values.append(raw); seen.add(key)
    return values


def _requested_chart_metrics(df, query: str) -> list[str]:
    """Map natural-language financial metrics to exact dataframe columns."""
    q = (query or "").casefold()
    numeric = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
    aliases = {
        "revenue": ["revenue"],
        "net income": ["net income", "netincome", "net profit", "profit"],
        "gross profit": ["gross profit"],
        "ebitda": ["ebitda"],
        "market cap": ["market cap", "market cap(in b usd)", "market cap in b usd"],
        "cash flow from operating": ["cash flow from operating"],
        "cash flow from investing": ["cash flow from investing"],
        "cash flow from financial activities": ["cash flow from financial activities"],
        "net profit margin": ["net profit margin"],
        "number of employees": ["number of employees", "employees", "employee count"],
    }
    selected = []
    for canonical, words in aliases.items():
        if any(re.search(rf"(?<![a-z0-9]){re.escape(w)}(?![a-z0-9])", q) for w in words):
            for col in numeric:
                cl = str(col).casefold().strip()
                if cl == canonical or cl in words:
                    if col not in selected:
                        selected.append(col)
                    break
    return selected


def _resolve_chart_spec(df, query: str) -> dict | None:
    """Build a deterministic chart from the ACTUAL uploaded financial table.

    Important rules for the supplied Financial Statements CSV:
    * Never plot unrelated numeric columns such as employees with cash flow.
    * Explicit company/ticker names are filtered before plotting.
    * A request for two-company comparison without named companies uses the
      first two companies in the uploaded data, rather than mixing all rows.
    * A multi-company trend without a company specified is aggregated by Year,
      preventing repeated years from producing a misleading spaghetti chart.
    * "net profit" maps to the actual "Net Income" column in this dataset.
    """
    if pd is None or df is None or df.empty:
        return None
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    q = (query or "").casefold()

    entity_col = next((c for c in df.columns if str(c).casefold().strip() in {"company", "ticker", "symbol"}), None)
    requested_entities = _chart_requested_entities(df, query, entity_col)
    is_compare = bool(re.search(r"\b(compare|comparison|versus|vs\.?|against|two companies|2 companies)\b", q))
    is_trend = bool(re.search(r"\b(trend|over time|time series|timeline|historical|history)\b", q))

    # Exact financial metrics requested by the user.
    selected = _requested_chart_metrics(df, query)

    # Handle the common generic request "financial performance of two companies".
    if is_compare and entity_col and (not requested_entities or len(requested_entities) >= 2):
        if requested_entities:
            comparison_entities = requested_entities[:2]
        else:
            comparison_entities = [str(v).strip() for v in df[entity_col].dropna().astype(str).str.strip().unique()[:2]]
        if len(comparison_entities) < 2:
            return None

        # If no metric was specified, Revenue is the safest common financial
        # performance measure in this dataset. Do not mix it with employee count.
        if not selected:
            selected = [next((c for c in df.columns if str(c).casefold().strip() == "revenue"), None)]
            selected = [c for c in selected if c]
        selected = selected[:2]
        period_col = next((c for c in df.columns if str(c).casefold().strip() in {"year", "date", "period"}), None)
        if period_col:
            title = f"{' vs '.join(comparison_entities)} — " + " vs ".join(str(c) for c in selected)
            return {
                "chart_type": "line",
                "category": period_col,
                "series": selected,
                "title": title,
                "entity_column": entity_col,
                "entity_values": comparison_entities,
                "entity_value": None,
                "compare_entities": comparison_entities,
                "aggregate_by_period": False,
            }

    # If a specific company is mentioned, filter it. If two are mentioned for a
    # comparison, render each company as a separate series of the same metric.
    if requested_entities and entity_col:
        period_col = next((c for c in df.columns if str(c).casefold().strip() in {"year", "date", "period"}), None)
        if len(requested_entities) >= 2 and is_compare:
            if not selected:
                selected = [next((c for c in df.columns if str(c).casefold().strip() == "revenue"), None)]
                selected = [c for c in selected if c]
            selected = selected[:2]
            return {
                "chart_type": "line",
                "category": period_col,
                "series": selected,
                "title": f"{' vs '.join(requested_entities[:2])} — " + " vs ".join(str(c) for c in selected),
                "entity_column": entity_col,
                "entity_values": requested_entities[:2],
                "entity_value": None,
                "compare_entities": requested_entities[:2],
                "aggregate_by_period": False,
            }
        entity_value = requested_entities[0]
        filtered = df[df[entity_col].astype(str).str.strip().str.casefold() == str(entity_value).casefold()].copy()
    else:
        entity_value = None
        filtered = df.copy()

    numeric = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
    if not numeric:
        return None

    # If the user explicitly asked for revenue/net profit, use ONLY those actual columns.
    if not selected:
        candidates = _column_candidates([str(c) for c in df.columns], query)
        selected = [c for c in candidates if c in numeric]
    selected = selected[:3]
    if not selected:
        return None

    period_col = next((c for c in df.columns if str(c).casefold().strip() in {"year", "date", "period"}), None)
    if period_col:
        # For a multi-company trend, aggregate by period. This removes repeated
        # years and makes the chart genuinely "over time".
        if entity_value is None and len(filtered[entity_col].dropna().unique()) > 1 if entity_col else False:
            chart_type = "line"
            title = "Total " + " and ".join(str(c) for c in selected) + " over time"
            return {
                "chart_type": chart_type,
                "category": period_col,
                "series": selected,
                "title": title,
                "entity_column": entity_col,
                "entity_value": None,
                "entity_values": [],
                "compare_entities": [],
                "aggregate_by_period": True,
            }
        chart_type = "line" if len(filtered) > 1 else "bar"
        if is_trend or len(selected) > 1:
            chart_type = "line"
        title = " and ".join(str(c) for c in selected) + " over time"
        if entity_value:
            title = f"{entity_value} — {title}"
        return {
            "chart_type": chart_type,
            "category": period_col,
            "series": selected,
            "title": title,
            "entity_column": entity_col,
            "entity_value": entity_value,
            "entity_values": [],
            "compare_entities": [],
            "aggregate_by_period": False,
        }

    nonnumeric = [c for c in df.columns if c not in numeric]
    category = entity_col if entity_col in df.columns else (nonnumeric[0] if nonnumeric else df.columns[0])
    chart_type = "pie" if "pie" in q else "bar"
    return {
        "chart_type": chart_type,
        "category": category,
        "series": selected,
        "title": " vs ".join(str(c) for c in selected),
        "entity_column": entity_col,
        "entity_value": entity_value,
        "entity_values": [],
        "compare_entities": [],
        "aggregate_by_period": False,
    }


def render_chart(df, spec: dict, out_path: Path) -> None:
    if plt is None:
        raise RuntimeError("Matplotlib is not installed.")
    if pd is None:
        raise RuntimeError("Pandas is not installed.")

    chart_type = spec["chart_type"]
    category = spec.get("category")
    series = spec.get("series") or []
    if not series:
        raise ValueError("No numeric series selected.")

    data = df.copy()
    data.columns = [str(c).strip() for c in data.columns]
    entity_column = spec.get("entity_column")
    entity_value = spec.get("entity_value")
    entity_values = [str(x).strip() for x in (spec.get("entity_values") or []) if str(x).strip()]

    # Filter explicitly selected company/entities before plotting.
    if entity_column and entity_column in data.columns:
        if entity_values:
            wanted = {x.casefold() for x in entity_values}
            data = data[data[entity_column].astype(str).str.strip().str.casefold().isin(wanted)].copy()
        elif entity_value:
            data = data[data[entity_column].astype(str).str.strip().str.casefold() == str(entity_value).strip().casefold()].copy()

    for s in series:
        if s not in data.columns:
            raise ValueError(f"Chart column '{s}' is not present in the uploaded data.")
        data[s] = pd.to_numeric(data[s], errors="coerce")

    # For a multi-company trend with no company specified, aggregate numeric
    # financial metrics by year. Never add unrelated columns.
    aggregate = bool(spec.get("aggregate_by_period"))
    if aggregate and category in data.columns:
        work = data[[category] + series].copy()
        work[category] = pd.to_numeric(work[category], errors="coerce")
        for s in series:
            work[s] = pd.to_numeric(work[s], errors="coerce")
        data = work.groupby(category, as_index=False)[series].sum(min_count=1)

    data = data.dropna(subset=series, how="all")
    if category in data.columns and str(category).casefold() in {"year", "date", "period"}:
        if str(category).casefold() == "date":
            try:
                data[category] = pd.to_datetime(data[category], errors="coerce")
            except Exception:
                pass
        else:
            data[category] = pd.to_numeric(data[category], errors="coerce")
        data = data.sort_values(category)

    # Comparison of multiple companies: one line per company for the selected
    # metric(s). If multiple metrics were requested, each company/metric gets a
    # distinct legend entry; all values still come from actual CSV columns.
    compare_entities = [str(x).strip() for x in (spec.get("compare_entities") or []) if str(x).strip()]
    is_company_comparison = bool(compare_entities and entity_column and category in data.columns)

    fig, ax = plt.subplots(figsize=(9, 5.5), dpi=170)
    if is_company_comparison:
        x_values = sorted(data[category].dropna().unique().tolist())
        for company in compare_entities:
            company_data = data[data[entity_column].astype(str).str.strip().str.casefold() == company.casefold()].copy()
            company_data = company_data.sort_values(category)
            if company_data.empty:
                continue
            for s in series:
                ax.plot(company_data[category].tolist(), company_data[s].tolist(), marker="o", label=f"{company} — {s}")
        ax.set_xlabel(str(category))
        ax.set_ylabel(", ".join(map(str, series)))
        ax.set_title(spec.get("title") or "Company financial comparison")
        ax.legend()
        ax.grid(axis="y", alpha=0.2)
    else:
        categories = data[category].astype(str).tolist() if category in data.columns else [str(i) for i in range(len(data))]
        x = list(range(len(data)))
        if chart_type == "pie":
            if len(series) != 1 or not category:
                raise ValueError("Pie chart requires one numeric series and one category.")
            vals = data[series[0]].fillna(0).tolist()
            ax.pie(vals, labels=categories, autopct="%1.1f%%")
            ax.set_title(spec.get("title") or "Chart")
        elif chart_type == "scatter":
            if not category or not pd.api.types.is_numeric_dtype(data[category]):
                raise ValueError("Scatter chart requires a numeric category column.")
            for s in series:
                ax.scatter(data[category].tolist(), data[s].tolist(), label=str(s))
            ax.set_xlabel(str(category)); ax.set_ylabel(", ".join(map(str, series)))
            ax.set_title(spec.get("title") or "Chart"); ax.legend()
        else:
            for s in series:
                vals = data[s].tolist()
                if chart_type == "line":
                    # Use actual period values on the x-axis, avoiding duplicate
                    # company rows when aggregation is enabled.
                    if category in data.columns and str(category).casefold() in {"year", "date", "period"}:
                        ax.plot(data[category].tolist(), vals, marker="o", label=str(s))
                    else:
                        ax.plot(x, vals, marker="o", label=str(s))
                elif chart_type == "step":
                    ax.step(x, vals, where="mid", label=str(s))
                elif chart_type == "area":
                    ax.fill_between(x, vals, alpha=0.35, label=str(s)); ax.plot(x, vals)
                else:
                    width = 0.8 / len(series)
                    offset = (series.index(s) - (len(series) - 1) / 2) * width
                    ax.bar([i + offset for i in x], vals, width=width, label=str(s))
            if not (category in data.columns and str(category).casefold() in {"year", "date", "period"} and chart_type == "line"):
                ax.set_xticks(x)
                ax.set_xticklabels(categories, rotation=25, ha="right")
            elif len(data) > 20:
                # Keep long financial histories readable.
                step = max(1, len(data) // 12)
                ticks = list(range(0, len(data), step))
                ax.set_xticks(data[category].iloc[ticks].tolist())
                ax.tick_params(axis="x", rotation=25)
            ax.set_title(spec.get("title") or "Chart")
            ax.set_xlabel(str(category or "Rows"))
            ax.set_ylabel(", ".join(map(str, series)))
            if len(series) > 1 or is_company_comparison:
                ax.legend()
            ax.grid(axis="y", alpha=0.2)

    fig.tight_layout()
    fig.savefig(out_path, bbox_inches="tight")
    plt.close(fig)


# ---------------------------------------------------------------------------
# Diagrams: local model JSON -> strict validation -> NetworkX renderer
# ---------------------------------------------------------------------------
def build_diagram_prompt(query: str, context: str) -> str:
    return (
        "Create a structural diagram description from the request and source "
        "context. Return ONLY JSON:\n"
        '{"title":"...","nodes":[{"id":"...","label":"...","type":"generic"}],'
        '"edges":[{"source":"...","target":"...","label":"..."}]}\n'
        "Rules: 2-30 unique nodes, non-empty short labels, every edge endpoint "
        "must reference a node, no invented source-specific entities, and only "
        "use components implied by the request/context. For a web diagram, "
        "model generic browser/frontend/backend/database/API components only when "
        "the request/context supports them.\n\n"
        f"Request: {query}\nContext:\n{context[:10000]}"
    )


def validate_graph(graph: dict) -> bool:
    if not isinstance(graph, dict) or not isinstance(graph.get("nodes"), list):
        return False
    ids = []
    for node in graph["nodes"]:
        if not isinstance(node, dict):
            return False
        nid = str(node.get("id", "")).strip()
        label = str(node.get("label", "")).strip()
        if not nid or not label or nid in ids:
            return False
        ids.append(nid)
    if not 2 <= len(ids) <= 30:
        return False
    for edge in graph.get("edges", []):
        if not isinstance(edge, dict):
            return False
        if str(edge.get("source", "")).strip() not in ids or str(edge.get("target", "")).strip() not in ids:
            return False
    return True


def render_diagram(graph: dict, out_path: Path) -> None:
    if nx is None or plt is None:
        raise RuntimeError("NetworkX and Matplotlib are required for diagrams.")
    G = nx.DiGraph()
    labels = {}
    for n in graph["nodes"]:
        nid = str(n["id"]).strip()
        G.add_node(nid)
        labels[nid] = str(n["label"])[:80]
    edge_labels = {}
    for e in graph.get("edges", []):
        src, dst = str(e["source"]).strip(), str(e["target"]).strip()
        G.add_edge(src, dst)
        if e.get("label"):
            edge_labels[(src, dst)] = str(e["label"])[:40]
    try:
        layers = list(nx.topological_generations(G))
        pos = {}
        for li, layer in enumerate(layers):
            for i, node in enumerate(layer):
                pos[node] = (li * 2.8, -i * 1.6 + (len(layer) - 1) * 0.8)
    except Exception:
        pos = nx.spring_layout(G, seed=7)
    fig, ax = plt.subplots(figsize=(11, 7), dpi=180)
    nx.draw_networkx_edges(G, pos, ax=ax, arrows=True, arrowsize=18, width=1.5, node_size=3200)
    nx.draw_networkx_nodes(G, pos, ax=ax, node_size=3200, node_shape="s")
    nx.draw_networkx_labels(G, pos, labels=labels, ax=ax, font_size=8, font_weight="bold")
    if edge_labels:
        nx.draw_networkx_edge_labels(G, pos, edge_labels=edge_labels, ax=ax, font_size=7)
    ax.set_title(str(graph.get("title") or "Diagram"))
    ax.axis("off")
    fig.tight_layout()
    fig.savefig(out_path, bbox_inches="tight")
    plt.close(fig)


# ---------------------------------------------------------------------------
# API: authentication and chat persistence
# ---------------------------------------------------------------------------
_ORIGINAL_INDEX_B64 = 'PCFET0NUWVBFIGh0bWw+Cgo8aHRtbCBsYW5nPSJlbiI+CjxoZWFkPgo8bWV0YSBjaGFyc2V0PSJVVEYtOCI+CjxtZXRhIG5hbWU9InZpZXdwb3J0IiBjb250ZW50PSJ3aWR0aD1kZXZpY2Utd2lkdGgsIGluaXRpYWwtc2NhbGU9MS4wIj4KPHRpdGxlPk9yaW9uIOKAlCBTb3ZlcmVpZ24gV29ya2JlbmNoPC90aXRsZT4KPHN0eWxlPgo6cm9vdHstLWJnOiMwNTA1MDU7LS1iZzI6IzBhMGEwYTstLXBhbmVsOiMxMTE7LS1wYW5lbDI6IzE3MTcxNzstLWxpbmU6IzI5MjkyOTstLXRleHQ6I2Y1ZjVmNTstLW11dGVkOiM5MjkyOTI7LS1waW5rOiNlYzQ4OTk7LS1wdXJwbGU6I2E4NTVmNzstLWdyZWVuOiMzNGQzOTk7LS15ZWxsb3c6I2ZhY2MxNTstLXJlZDojZjg3MTcxOy0tZ3JhZGllbnQ6bGluZWFyLWdyYWRpZW50KDEzNWRlZywjZWM0ODk5LCNhODU1ZjcpfQpib2R5LmxpZ2h0ey0tYmc6I2Y3ZjdmODstLWJnMjojZmZmOy0tcGFuZWw6I2ZmZjstLXBhbmVsMjojZjJmMmYyOy0tbGluZTojZGRkOy0tdGV4dDojMTgxODFiOy0tbXV0ZWQ6IzY2Nn0KKntib3gtc2l6aW5nOmJvcmRlci1ib3h9aHRtbCxib2R5e21hcmdpbjowO3dpZHRoOjEwMCU7aGVpZ2h0OjEwMCU7b3ZlcmZsb3c6aGlkZGVuO2ZvbnQtZmFtaWx5OkludGVyLHVpLXNhbnMtc2VyaWYsc3lzdGVtLXVpLC1hcHBsZS1zeXN0ZW0sIlNlZ29lIFVJIixzYW5zLXNlcmlmO2JhY2tncm91bmQ6dmFyKC0tYmcpO2NvbG9yOnZhcigtLXRleHQpfWJ1dHRvbixpbnB1dHtmb250OmluaGVyaXR9YnV0dG9ue2N1cnNvcjpwb2ludGVyfQoKOnJvb3R7LS1iZzojMDUwNTA1Oy0tYmcyOiMwYTBhMGE7LS1wYW5lbDojMTExOy0tcGFuZWwyOiMxNzE3MTc7LS1saW5lOiMyOTI5Mjk7LS10ZXh0OiNmNWY1ZjU7LS1tdXRlZDojOTI5MjkyOy0tcGluazojZWM0ODk5Oy0tcHVycGxlOiNhODU1Zjc7LS1ncmVlbjojMzRkMzk5Oy0teWVsbG93OiNmYWNjMTU7LS1yZWQ6I2Y4NzE3MTstLWdyYWRpZW50OmxpbmVhci1ncmFkaWVudCgxMzVkZWcsI2VjNDg5OSwjYTg1NWY3KX0KYm9keS5saWdodHstLWJnOiNmN2Y3Zjg7LS1iZzI6I2ZmZjstLXBhbmVsOiNmZmY7LS1wYW5lbDI6I2YyZjJmMjstLWxpbmU6I2RkZDstLXRleHQ6IzE4MTgxYjstLW11dGVkOiM2NjZ9Cip7Ym94LXNpemluZzpib3JkZXItYm94fWh0bWwsYm9keXttYXJnaW46MDt3aWR0aDoxMDAlO2hlaWdodDoxMDAlO292ZXJmbG93OmhpZGRlbjtmb250LWZhbWlseTpJbnRlcix1aS1zYW5zLXNlcmlmLHN5c3RlbS11aSwtYXBwbGUtc3lzdGVtLCJTZWdvZSBVSSIsc2Fucy1zZXJpZjtiYWNrZ3JvdW5kOnZhcigtLWJnKTtjb2xvcjp2YXIoLS10ZXh0KX1idXR0b24saW5wdXR7Zm9udDppbmhlcml0fWJ1dHRvbntjdXJzb3I6cG9pbnRlcn0KLmhpZGRlbntkaXNwbGF5Om5vbmUhaW1wb3J0YW50fS5tdXRlZHtjb2xvcjp2YXIoLS1tdXRlZCl9CiNsb2dpbntwb3NpdGlvbjpmaXhlZDtpbnNldDowO3otaW5kZXg6MjAwO2Rpc3BsYXk6ZmxleDthbGlnbi1pdGVtczpjZW50ZXI7anVzdGlmeS1jb250ZW50OmNlbnRlcjtiYWNrZ3JvdW5kOnJhZGlhbC1ncmFkaWVudChjaXJjbGUgYXQgMjUlIDIwJSxyZ2JhKDIzNiw3MiwxNTMsLjEyKSx0cmFuc3BhcmVudCA0MCUpLHJhZGlhbC1ncmFkaWVudChjaXJjbGUgYXQgODAlIDc1JSxyZ2JhKDE2OCw4NSwyNDcsLjE0KSx0cmFuc3BhcmVudCA0NSUpLCMwNTA1MDV9Ci5sb2dpbi1jYXJke3dpZHRoOm1pbig0MjBweCw5MnZ3KTtiYWNrZ3JvdW5kOiMwYzBjMGM7Ym9yZGVyOjFweCBzb2xpZCAjMjkyOTI5O2JvcmRlci1yYWRpdXM6MjBweDtwYWRkaW5nOjMwcHg7Ym94LXNoYWRvdzowIDMwcHggMTAwcHggcmdiYSgwLDAsMCwuNil9Ci5sb2dpbi1sb2dve2ZvbnQtc2l6ZToyM3B4O2ZvbnQtd2VpZ2h0OjgwMDtiYWNrZ3JvdW5kOnZhcigtLWdyYWRpZW50KTstd2Via2l0LWJhY2tncm91bmQtY2xpcDp0ZXh0O2JhY2tncm91bmQtY2xpcDp0ZXh0O2NvbG9yOnRyYW5zcGFyZW50fS5sb2dpbi1zdWJ7Zm9udC1zaXplOjEzcHg7Y29sb3I6Izk5OTttYXJnaW46OHB4IDAgMjVweDtsaW5lLWhlaWdodDoxLjV9LmZpZWxke21hcmdpbjoxM3B4IDB9LmZpZWxkIGxhYmVse2Rpc3BsYXk6YmxvY2s7Zm9udC1zaXplOjExcHg7dGV4dC10cmFuc2Zvcm06dXBwZXJjYXNlO2xldHRlci1zcGFjaW5nOi4wOGVtO2NvbG9yOiM4ODg7bWFyZ2luLWJvdHRvbTo3cHh9LmZpZWxkIGlucHV0e3dpZHRoOjEwMCU7cGFkZGluZzoxMnB4IDEzcHg7Ym9yZGVyLXJhZGl1czoxMHB4O2JvcmRlcjoxcHggc29saWQgIzMwMzAzMDtiYWNrZ3JvdW5kOiMxNTE1MTU7Y29sb3I6I2ZmZjtvdXRsaW5lOm5vbmV9LmZpZWxkIGlucHV0OmZvY3Vze2JvcmRlci1jb2xvcjp2YXIoLS1wdXJwbGUpfS5wcmltYXJ5e3dpZHRoOjEwMCU7Ym9yZGVyOjA7Ym9yZGVyLXJhZGl1czoxMHB4O3BhZGRpbmc6MTJweDtiYWNrZ3JvdW5kOnZhcigtLWdyYWRpZW50KTtjb2xvcjojMTkwOTE0O2ZvbnQtd2VpZ2h0OjgwMDttYXJnaW4tdG9wOjhweH0uZXJyb3J7Y29sb3I6dmFyKC0tcmVkKTtmb250LXNpemU6MTJweDttYXJnaW4tdG9wOjEycHg7bWluLWhlaWdodDoxOHB4fS5zZWN1cmUtbm90ZXttYXJnaW4tdG9wOjIwcHg7Ym9yZGVyOjFweCBzb2xpZCByZ2JhKDUyLDIxMSwxNTMsLjI1KTtiYWNrZ3JvdW5kOnJnYmEoNTIsMjExLDE1MywuMDYpO3BhZGRpbmc6MTBweDtib3JkZXItcmFkaXVzOjEwcHg7Zm9udC1zaXplOjExcHg7Y29sb3I6dmFyKC0tZ3JlZW4pfQoKLyogRVhBQ1QgV0hJVEUgLyBQVVJQTEUgLyBCTFVFIFNVQlRMRSBHUkFESUVOVCBGT1IgUk9CT1QgU1BMQVNIICovCiNzcGxhc2h7cG9zaXRpb246Zml4ZWQ7aW5zZXQ6MDt6LWluZGV4OjMwMDtkaXNwbGF5OmZsZXg7ZmxleC1kaXJlY3Rpb246Y29sdW1uO2FsaWduLWl0ZW1zOmNlbnRlcjtqdXN0aWZ5LWNvbnRlbnQ6Y2VudGVyO2JhY2tncm91bmQ6cmFkaWFsLWdyYWRpZW50KGNpcmNsZSBhdCA1MCUgMjAlLCByZ2JhKDE2OCw4NSwyNDcsMC4xOCkgMCUsIHJnYmEoMTQsMjAsMzgsMC45MikgNTAlLCAjMDMwNzEyIDEwMCUpO292ZXJmbG93OmhpZGRlbjt0cmFuc2l0aW9uOm9wYWNpdHkgLjhzIGVhc2UsdmlzaWJpbGl0eSAuOHMgZWFzZX0KI3NwbGFzaC5oaWRle29wYWNpdHk6MDt2aXNpYmlsaXR5OmhpZGRlbjtwb2ludGVyLWV2ZW50czpub25lfQouc3BsYXNoLXN0YXJze3Bvc2l0aW9uOmFic29sdXRlO2luc2V0OjA7b3BhY2l0eTouNTtwb2ludGVyLWV2ZW50czpub25lfS5zcGxhc2gtc3RhcnMgc3Bhbntwb3NpdGlvbjphYnNvbHV0ZTt3aWR0aDoycHg7aGVpZ2h0OjJweDtib3JkZXItcmFkaXVzOjUwJTtiYWNrZ3JvdW5kOiNjZmUwZmY7b3BhY2l0eTowO2FuaW1hdGlvbjp0d2lua2xlIDMuNHMgZWFzZS1pbi1vdXQgaW5maW5pdGV9QGtleWZyYW1lcyB0d2lua2xlezAlLDEwMCV7b3BhY2l0eTowfTUwJXtvcGFjaXR5Oi45fX0KLnJvYm90LXN0YWdle3Bvc2l0aW9uOnJlbGF0aXZlO3dpZHRoOjI4MHB4O2hlaWdodDoyODBweDthbmltYXRpb246cm9ib3RGbG9hdCAzLjZzIGVhc2UtaW4tb3V0IGluZmluaXRlfS5yb2JvdC1nbG93e3Bvc2l0aW9uOmFic29sdXRlO2luc2V0Oi0zMHB4O2JhY2tncm91bmQ6cmFkaWFsLWdyYWRpZW50KGNpcmNsZSxyZ2JhKDE2OCw4NSwyNDcsMC4yNSkgMCUsdHJhbnNwYXJlbnQgNzAlKTtmaWx0ZXI6Ymx1cig2cHgpO2FuaW1hdGlvbjpnbG93UHVsc2UgMy42cyBlYXNlLWluLW91dCBpbmZpbml0ZX0ucm9ib3Qtc2hhZG93e3Bvc2l0aW9uOmFic29sdXRlO2xlZnQ6NTAlO2JvdHRvbTo2cHg7d2lkdGg6MTIwcHg7aGVpZ2h0OjE4cHg7dHJhbnNmb3JtOnRyYW5zbGF0ZVgoLTUwJSk7YmFja2dyb3VuZDpyYWRpYWwtZ3JhZGllbnQoZWxsaXBzZSBhdCBjZW50ZXIscmdiYSgxNjgsODUsMjQ3LDAuMykgMCUsdHJhbnNwYXJlbnQgNzIlKTtib3JkZXItcmFkaXVzOjUwJTthbmltYXRpb246c2hhZG93UHVsc2UgMy42cyBlYXNlLWluLW91dCBpbmZpbml0ZX1Aa2V5ZnJhbWVzIHJvYm90RmxvYXR7MCUsMTAwJXt0cmFuc2Zvcm06dHJhbnNsYXRlWSgwKX01MCV7dHJhbnNmb3JtOnRyYW5zbGF0ZVkoLTE0cHgpfX1Aa2V5ZnJhbWVzIHNoYWRvd1B1bHNlezAlLDEwMCV7dHJhbnNmb3JtOnRyYW5zbGF0ZVgoLTUwJSkgc2NhbGUoMSk7b3BhY2l0eTouNTV9NTAle3RyYW5zZm9ybTp0cmFuc2xhdGVYKC01MCUpIHNjYWxlKC43OCk7b3BhY2l0eTouM319QGtleWZyYW1lcyBnbG93UHVsc2V7MCUsMTAwJXtvcGFjaXR5Oi41fTUwJXtvcGFjaXR5OjF9fQojYXJtUmlnaHR7dHJhbnNmb3JtLW9yaWdpbjoxNTBweCAxNTBweDthbmltYXRpb246d2F2ZSAyLjRzIGVhc2UtaW4tb3V0IGluZmluaXRlO2FuaW1hdGlvbi1kZWxheTouNnN9QGtleWZyYW1lcyB3YXZlezAlLDYwJSwxMDAle3RyYW5zZm9ybTpyb3RhdGUoMCl9MTAle3RyYW5zZm9ybTpyb3RhdGUoLTE4ZGVnKX0yMCV7dHJhbnNmb3JtOnJvdGF0ZSg2ZGVnKX0zMCV7dHJhbnNmb3JtOnJvdGF0ZSgtMTRkZWcpfTQwJXt0cmFuc2Zvcm06cm90YXRlKDRkZWcpfTUwJXt0cmFuc2Zvcm06cm90YXRlKC04ZGVnKX19I2V5ZUwsI2V5ZVJ7YW5pbWF0aW9uOmJsaW5rIDQuMnMgZWFzZS1pbi1vdXQgaW5maW5pdGU7dHJhbnNmb3JtLW9yaWdpbjpjZW50ZXJ9I2V5ZVJ7YW5pbWF0aW9uLWRlbGF5Oi4wNnN9QGtleWZyYW1lcyBibGlua3swJSw5MiUsMTAwJXt0cmFuc2Zvcm06c2NhbGVZKDEpfTk1JXt0cmFuc2Zvcm06c2NhbGVZKC4xMil9fSN2aXNvckdsb3d7YW5pbWF0aW9uOnZpc29yUHVsc2UgMi42cyBlYXNlLWluLW91dCBpbmZpbml0ZX1Aa2V5ZnJhbWVzIHZpc29yUHVsc2V7MCUsMTAwJXtvcGFjaXR5Oi41NX01MCV7b3BhY2l0eToxfX0jYW50ZW5uYVRpcHthbmltYXRpb246YW50ZW5uYVB1bHNlIDEuOHMgZWFzZS1pbi1vdXQgaW5maW5pdGV9QGtleWZyYW1lcyBhbnRlbm5hUHVsc2V7MCUsMTAwJXtvcGFjaXR5Oi40fTUwJXtvcGFjaXR5OjF9fQouc3BsYXNoLXRleHR7cG9zaXRpb246cmVsYXRpdmU7ei1pbmRleDoyO3RleHQtYWxpZ246Y2VudGVyO21hcmdpbi10b3A6MjJweDtvcGFjaXR5OjA7YW5pbWF0aW9uOnRleHRJbiAxcyBlYXNlIGZvcndhcmRzO2FuaW1hdGlvbi1kZWxheTouOXN9QGtleWZyYW1lcyB0ZXh0SW57ZnJvbXtvcGFjaXR5OjA7dHJhbnNmb3JtOnRyYW5zbGF0ZVkoMTRweCl9dG97b3BhY2l0eToxO3RyYW5zZm9ybTp0cmFuc2xhdGVZKDApfX0uc3BsYXNoLXdlbGNvbWV7Zm9udC1zaXplOjEzcHg7bGV0dGVyLXNwYWNpbmc6Mi41cHg7Y29sb3I6dmFyKC0tbXV0ZWQpO3RleHQtdHJhbnNmb3JtOnVwcGVyY2FzZTttYXJnaW4tYm90dG9tOjhweH0uc3BsYXNoLWJyYW5ke2ZvbnQtc2l6ZTozMHB4O2ZvbnQtd2VpZ2h0OjgwMDtsZXR0ZXItc3BhY2luZzouNXB4fS5zcGxhc2gtYnJhbmQgc3Bhbntjb2xvcjp2YXIoLS1hY2NlbnQsI2E4NTVmNyl9LnNwbGFzaC1xdW90ZXttYXJnaW4tdG9wOjE2cHg7bWF4LXdpZHRoOjUyMHB4O3BhZGRpbmc6MCAyNHB4O2ZvbnQtc2l6ZToxNHB4O2xpbmUtaGVpZ2h0OjEuNjtjb2xvcjojZTJlOGYwO2ZvbnQtc3R5bGU6aXRhbGljO29wYWNpdHk6MDthbmltYXRpb246dGV4dEluIDFzIGVhc2UgZm9yd2FyZHM7YW5pbWF0aW9uLWRlbGF5OjEuNnN9LnNwbGFzaC1xdW90ZTo6YmVmb3JlLC5zcGxhc2gtcXVvdGU6OmFmdGVye2NvbnRlbnQ6JyInO2NvbG9yOnZhcigtLWFjY2VudCwjYTg1NWY3KTtmb250LXN0eWxlOm5vcm1hbH0uc3BsYXNoLWVudGVye21hcmdpbi10b3A6MjZweDtvcGFjaXR5OjA7YW5pbWF0aW9uOnRleHRJbiAxcyBlYXNlIGZvcndhcmRzO2FuaW1hdGlvbi1kZWxheToyLjNzfS5zcGxhc2gtZW50ZXIgYnV0dG9ue3BhZGRpbmc6MTJweCAyNnB4O2JvcmRlcjoxcHggc29saWQgI2E4NTVmNztib3JkZXItcmFkaXVzOjEwcHg7YmFja2dyb3VuZDp2YXIoLS1ncmFkaWVudCk7Y29sb3I6I2ZmZjtmb250LXdlaWdodDo4MDA7Zm9udC1zaXplOjEzcHg7bGV0dGVyLXNwYWNpbmc6LjRweH0uc3BsYXNoLWVudGVyIGJ1dHRvbjpob3ZlcntvcGFjaXR5Oi45fS5zcGxhc2gtc2tpcHtwb3NpdGlvbjphYnNvbHV0ZTt0b3A6MjBweDtyaWdodDoyMnB4O3otaW5kZXg6Mztjb2xvcjojOTFhMGI1O2ZvbnQtc2l6ZToxMnB4O2JhY2tncm91bmQ6dHJhbnNwYXJlbnQ7Ym9yZGVyOjFweCBzb2xpZCAjMjYzMjQ2O3BhZGRpbmc6N3B4IDEycHg7Ym9yZGVyLXJhZGl1czo4cHh9LnNwbGFzaC1za2lwOmhvdmVye2NvbG9yOiNlOGVkZjU7Ym9yZGVyLWNvbG9yOiMzMzQxNTl9CgovKiBSRVNUIE9GIFVJIFNUWUxJTkcgKi8KI2FwcHtkaXNwbGF5OmdyaWQ7Z3JpZC10ZW1wbGF0ZS1jb2x1bW5zOjI4MHB4IG1pbm1heCgwLDFmcik7aGVpZ2h0OjEwMHZofQouc2lkZWJhcnttaW4taGVpZ2h0OjA7b3ZlcmZsb3c6aGlkZGVuO2JhY2tncm91bmQ6dmFyKC0tYmcyKTtib3JkZXItcmlnaHQ6MXB4IHNvbGlkIHZhcigtLWxpbmUpO2Rpc3BsYXk6ZmxleDtmbGV4LWRpcmVjdGlvbjpjb2x1bW47cGFkZGluZzoxNnB4O2dhcDoxMnB4fS5icmFuZHtkaXNwbGF5OmZsZXg7YWxpZ24taXRlbXM6Y2VudGVyO2p1c3RpZnktY29udGVudDpzcGFjZS1iZXR3ZWVuO2ZvbnQtd2VpZ2h0OjgwMDtmb250LXNpemU6MTZweH0uYnJhbmQtbGVmdHtkaXNwbGF5OmZsZXg7YWxpZ24taXRlbXM6Y2VudGVyO2dhcDo5cHh9LmJyYW5kLWRvdHt3aWR0aDo5cHg7aGVpZ2h0OjlweDtib3JkZXItcmFkaXVzOjUwJTtiYWNrZ3JvdW5kOnZhcigtLWdyYWRpZW50KTtib3gtc2hhZG93OjAgMCAxMnB4IHJnYmEoMjM2LDcyLDE1MywuNyl9Lmljb24tYnRue3dpZHRoOjM0cHg7aGVpZ2h0OjM0cHg7Ym9yZGVyLXJhZGl1czo5cHg7Ym9yZGVyOjFweCBzb2xpZCB2YXIoLS1saW5lKTtiYWNrZ3JvdW5kOnZhcigtLXBhbmVsKTtjb2xvcjp2YXIoLS1tdXRlZCl9Lmljb24tYnRuOmhvdmVye2JvcmRlci1jb2xvcjp2YXIoLS1waW5rKTtjb2xvcjp2YXIoLS1waW5rKX0KLnNlY3VyZXtmb250LXNpemU6MTBweDt0ZXh0LWFsaWduOmNlbnRlcjtwYWRkaW5nOjdweCA4cHg7Ym9yZGVyOjFweCBzb2xpZCByZ2JhKDUyLDIxMSwxNTMsLjI1KTtiYWNrZ3JvdW5kOnJnYmEoNTIsMjExLDE1MywuMDYpO2NvbG9yOnZhcigtLWdyZWVuKTtib3JkZXItcmFkaXVzOjhweDtsZXR0ZXItc3BhY2luZzouMDVlbX0ucHJvZmlsZXtib3JkZXI6MXB4IHNvbGlkIHZhcigtLWxpbmUpO2JhY2tncm91bmQ6dmFyKC0tcGFuZWwpO2JvcmRlci1yYWRpdXM6MTBweDtwYWRkaW5nOjEwcHh9LnByb2ZpbGUgLm5hbWV7Zm9udC13ZWlnaHQ6NzAwO2ZvbnQtc2l6ZToxM3B4fS5wcm9maWxlIC5tZXRhe2ZvbnQtc2l6ZToxMXB4O2NvbG9yOnZhcigtLW11dGVkKTtsaW5lLWhlaWdodDoxLjY7bWFyZ2luLXRvcDo0cHh9Lm5ldy1jaGF0e2JvcmRlcjoxcHggc29saWQgdmFyKC0tbGluZSk7YmFja2dyb3VuZDp0cmFuc3BhcmVudDtjb2xvcjp2YXIoLS10ZXh0KTtib3JkZXItcmFkaXVzOjlweDtwYWRkaW5nOjEwcHg7Zm9udC13ZWlnaHQ6NzAwfS5uZXctY2hhdDpob3Zlcntib3JkZXItY29sb3I6dmFyKC0tcGluayk7YmFja2dyb3VuZDpyZ2JhKDIzNiw3MiwxNTMsLjA2KX0uc2VjdGlvbi10aXRsZXtmb250LXNpemU6MTBweDt0ZXh0LXRyYW5zZm9ybTp1cHBlcmNhc2U7bGV0dGVyLXNwYWNpbmc6LjFlbTtjb2xvcjp2YXIoLS1tdXRlZCk7cGFkZGluZzo0cHggNHB4IDJweH0KLmNoYXQtbGlzdCwudGVhbS1saXN0e292ZXJmbG93LXg6aGlkZGVuO292ZXJmbG93LXk6YXV0bzttaW4taGVpZ2h0OjA7cGFkZGluZzowIDJweDtzY3JvbGxiYXItd2lkdGg6dGhpbn0uY2hhdC1saXN0e2ZsZXg6MSAxIGF1dG87bWluLWhlaWdodDoxMjBweH0udGVhbS1saXN0e2ZsZXg6MCAxIDE1MHB4O21heC1oZWlnaHQ6MTUwcHh9LmNoYXQtaXRlbXtkaXNwbGF5OmZsZXg7YWxpZ24taXRlbXM6Y2VudGVyO2dhcDo4cHg7d2lkdGg6MTAwJTtib3JkZXI6MXB4IHNvbGlkIHRyYW5zcGFyZW50O2JhY2tncm91bmQ6dHJhbnNwYXJlbnQ7Y29sb3I6dmFyKC0tdGV4dCk7cGFkZGluZzo5cHggOXB4O2JvcmRlci1yYWRpdXM6OHB4O3RleHQtYWxpZ246bGVmdDttYXJnaW4tYm90dG9tOjNweH0uY2hhdC1pdGVtOmhvdmVye2JhY2tncm91bmQ6dmFyKC0tcGFuZWwpO2JvcmRlci1jb2xvcjp2YXIoLS1saW5lKX0uY2hhdC1pdGVtLmFjdGl2ZXtiYWNrZ3JvdW5kOnJnYmEoMTY4LDg1LDI0NywuMTIpO2JvcmRlci1jb2xvcjpyZ2JhKDE2OCw4NSwyNDcsLjM1KX0uY2hhdC1pdGVtIC5kb3R7d2lkdGg6NnB4O2hlaWdodDo2cHg7Ym9yZGVyLXJhZGl1czo1MCU7YmFja2dyb3VuZDojNTU1O2ZsZXg6bm9uZX0uY2hhdC1pdGVtLmFjdGl2ZSAuZG90e2JhY2tncm91bmQ6dmFyKC0tcHVycGxlKX0uY2hhdC1pdGVtIC50aXRsZXtvdmVyZmxvdzpoaWRkZW47dGV4dC1vdmVyZmxvdzplbGxpcHNpczt3aGl0ZS1zcGFjZTpub3dyYXA7Zm9udC1zaXplOjEycHh9LmNoYXQtaXRlbSAudGVhbS10YWd7bWFyZ2luLWxlZnQ6YXV0bztmb250LXNpemU6OXB4O2NvbG9yOnZhcigtLWdyZWVuKX0KLmVtcHR5LWxpc3R7Zm9udC1zaXplOjExcHg7Y29sb3I6IzY2NjtwYWRkaW5nOjhweH0KLnNpZGViYXItbmF2e2Rpc3BsYXk6Z3JpZDtncmlkLXRlbXBsYXRlLWNvbHVtbnM6MWZyIDFmcjtnYXA6NnB4fQouc2lkZS1uYXYtYnRue2JvcmRlcjoxcHggc29saWQgdmFyKC0tbGluZSk7YmFja2dyb3VuZDp0cmFuc3BhcmVudDtjb2xvcjp2YXIoLS1tdXRlZCk7Ym9yZGVyLXJhZGl1czo5cHg7cGFkZGluZzo5cHggOHB4O3RleHQtYWxpZ246bGVmdDtmb250LXNpemU6MTFweH0KLnNpZGUtbmF2LWJ0bjpob3Zlcntib3JkZXItY29sb3I6dmFyKC0tcHVycGxlKTtjb2xvcjp2YXIoLS10ZXh0KTtiYWNrZ3JvdW5kOnJnYmEoMTY4LDg1LDI0NywuMDYpfQouZmlsZS1jYXJkc3tkaXNwbGF5OmZsZXg7ZmxleC13cmFwOndyYXA7Z2FwOjdweDttYXJnaW4tYm90dG9tOjEwcHh9Ci5maWxlLWNhcmR7ZGlzcGxheTpmbGV4O2FsaWduLWl0ZW1zOmNlbnRlcjtnYXA6OXB4O21pbi13aWR0aDoyMTBweDttYXgtd2lkdGg6MzIwcHg7cGFkZGluZzo4cHggMTBweDtib3JkZXI6MXB4IHNvbGlkIHZhcigtLWxpbmUpO2JhY2tncm91bmQ6dmFyKC0tcGFuZWwyKTtib3JkZXItcmFkaXVzOjEwcHg7Y29sb3I6dmFyKC0tdGV4dCk7dGV4dC1kZWNvcmF0aW9uOm5vbmV9Ci5maWxlLWNhcmQ6aG92ZXJ7Ym9yZGVyLWNvbG9yOnZhcigtLXB1cnBsZSl9Ci5maWxlLWljb257d2lkdGg6MzBweDtoZWlnaHQ6MzBweDtib3JkZXItcmFkaXVzOjdweDtkaXNwbGF5OmZsZXg7YWxpZ24taXRlbXM6Y2VudGVyO2p1c3RpZnktY29udGVudDpjZW50ZXI7YmFja2dyb3VuZDpyZ2JhKDE2OCw4NSwyNDcsLjEyKTtmb250LXNpemU6MTVweDtmbGV4Om5vbmV9Ci5maWxlLWNhcmQtbmFtZXtmb250LXNpemU6MTFweDtmb250LXdlaWdodDo3MDA7b3ZlcmZsb3c6aGlkZGVuO3RleHQtb3ZlcmZsb3c6ZWxsaXBzaXM7d2hpdGUtc3BhY2U6bm93cmFwfQouZmlsZS1jYXJkLW1ldGF7Zm9udC1zaXplOjlweDtjb2xvcjp2YXIoLS1tdXRlZCk7bWFyZ2luLXRvcDoycHh9Ci5zZW5kZXItbGFiZWx7Zm9udC1zaXplOjExcHg7Zm9udC13ZWlnaHQ6ODAwO21hcmdpbi1ib3R0b206NXB4O2NvbG9yOnZhcigtLXRleHQpfQouc2VuZGVyLWVtYWlse2ZvbnQtc2l6ZTo5cHg7Y29sb3I6dmFyKC0tbXV0ZWQpO2ZvbnQtd2VpZ2h0OjUwMDttYXJnaW4tbGVmdDo1cHh9Ci5tZXNzYWdlLnVzZXIgLnNlbmRlci1sYWJlbHt0ZXh0LWFsaWduOnJpZ2h0fQoudXNlci1jb250ZW50e3doaXRlLXNwYWNlOnByZS13cmFwfQouYXNzaXN0YW50LWJvZHl7d2hpdGUtc3BhY2U6cHJlLXdyYXB9Ci5saWJyYXJ5LWxpc3R7bWF4LWhlaWdodDo0MzBweDtvdmVyZmxvdzphdXRvO2Rpc3BsYXk6Z3JpZDtnYXA6N3B4fQoubGlicmFyeS1pdGVte2Rpc3BsYXk6ZmxleDtqdXN0aWZ5LWNvbnRlbnQ6c3BhY2UtYmV0d2VlbjtnYXA6MTBweDthbGlnbi1pdGVtczpjZW50ZXI7Ym9yZGVyOjFweCBzb2xpZCB2YXIoLS1saW5lKTtiYWNrZ3JvdW5kOnZhcigtLXBhbmVsMik7cGFkZGluZzoxMHB4O2JvcmRlci1yYWRpdXM6OXB4fQoubGlicmFyeS1pdGVtIC5uYW1le2ZvbnQtc2l6ZToxMnB4O2ZvbnQtd2VpZ2h0OjcwMH0ubGlicmFyeS1pdGVtIC5tZXRhe2ZvbnQtc2l6ZTo5cHg7Y29sb3I6dmFyKC0tbXV0ZWQpO21hcmdpbi10b3A6M3B4fQouc2V0dGluZ3Mtcm93e2Rpc3BsYXk6ZmxleDtqdXN0aWZ5LWNvbnRlbnQ6c3BhY2UtYmV0d2VlbjtnYXA6MTJweDtwYWRkaW5nOjEwcHggMDtib3JkZXItYm90dG9tOjFweCBzb2xpZCB2YXIoLS1saW5lKTtmb250LXNpemU6MTJweH0uc2V0dGluZ3Mtcm93Omxhc3QtY2hpbGR7Ym9yZGVyLWJvdHRvbTowfS5zZXR0aW5ncy12YWx1ZXtjb2xvcjp2YXIoLS1ncmVlbik7dGV4dC1hbGlnbjpyaWdodDt3b3JkLWJyZWFrOmJyZWFrLXdvcmR9Cgouc2lkZWJhci1ib3R0b217Ym9yZGVyLXRvcDoxcHggc29saWQgdmFyKC0tbGluZSk7cGFkZGluZy10b3A6MTBweH0uc3RhdHVzLXJvd3tkaXNwbGF5OmZsZXg7anVzdGlmeS1jb250ZW50OnNwYWNlLWJldHdlZW47YWxpZ24taXRlbXM6Y2VudGVyO2ZvbnQtc2l6ZToxMXB4O2NvbG9yOnZhcigtLW11dGVkKTttYXJnaW46NXB4IDB9LnN0YXR1cy1va3tjb2xvcjp2YXIoLS1ncmVlbil9LmxvZ291dHt3aWR0aDoxMDAlO21hcmdpbi10b3A6OHB4O2JvcmRlcjoxcHggc29saWQgdmFyKC0tbGluZSk7YmFja2dyb3VuZDp0cmFuc3BhcmVudDtjb2xvcjp2YXIoLS1tdXRlZCk7Ym9yZGVyLXJhZGl1czo4cHg7cGFkZGluZzo4cHh9LmxvZ291dDpob3Zlcntib3JkZXItY29sb3I6dmFyKC0tcmVkKTtjb2xvcjp2YXIoLS1yZWQpfQptYWlue21pbi13aWR0aDowO2hlaWdodDoxMDB2aDtkaXNwbGF5OmZsZXg7ZmxleC1kaXJlY3Rpb246Y29sdW1uO2JhY2tncm91bmQ6dmFyKC0tYmcpO3Bvc2l0aW9uOnJlbGF0aXZlfS5oZWFkZXJ7aGVpZ2h0OjYycHg7ZmxleDpub25lO2JvcmRlci1ib3R0b206MXB4IHNvbGlkIHZhcigtLWxpbmUpO2Rpc3BsYXk6ZmxleDthbGlnbi1pdGVtczpjZW50ZXI7anVzdGlmeS1jb250ZW50OnNwYWNlLWJldHdlZW47cGFkZGluZzowIDIycHh9LmhlYWRlci10aXRsZXtmb250LXdlaWdodDo4MDB9LmhlYWRlci10aXRsZSBzcGFue2JhY2tncm91bmQ6dmFyKC0tZ3JhZGllbnQpOy13ZWJraXQtYmFja2dyb3VuZC1jbGlwOnRleHQ7YmFja2dyb3VuZC1jbGlwOnRleHQ7Y29sb3I6dHJhbnNwYXJlbnR9LmhlYWRlci1yaWdodHtkaXNwbGF5OmZsZXg7YWxpZ24taXRlbXM6Y2VudGVyO2dhcDoxMnB4O2ZvbnQtc2l6ZToxMXB4O2NvbG9yOnZhcigtLW11dGVkKX0ubW9kZS1waWxse2JvcmRlcjoxcHggc29saWQgcmdiYSg1MiwyMTEsMTUzLC4yNSk7YmFja2dyb3VuZDpyZ2JhKDUyLDIxMSwxNTMsLjA1KTtjb2xvcjp2YXIoLS1ncmVlbik7cGFkZGluZzo2cHggOXB4O2JvcmRlci1yYWRpdXM6MjBweH0KLndvcmtzcGFjZXtmbGV4OjE7bWluLWhlaWdodDowO292ZXJmbG93LXk6YXV0bztvdmVyZmxvdy14OmhpZGRlbjtwYWRkaW5nOjI0cHggMjVweCAxNTBweDtzY3JvbGwtYmVoYXZpb3I6YXV0b30uaGVyb3ttYXgtd2lkdGg6ODUwcHg7bWFyZ2luOjEwdmggYXV0byAwO3RleHQtYWxpZ246Y2VudGVyfS5oZXJvIGgxe2ZvbnQtc2l6ZTpjbGFtcCgzMHB4LDV2dyw0OHB4KTttYXJnaW46MDtiYWNrZ3JvdW5kOnZhcigtLWdyYWRpZW50KTstd2Via2l0LWJhY2tncm91bmQtY2xpcDp0ZXh0O2JhY2tncm91bmQtY2xpcDp0ZXh0O2NvbG9yOnRyYW5zcGFyZW50fS5oZXJvIHB7Y29sb3I6dmFyKC0tbXV0ZWQpO2xpbmUtaGVpZ2h0OjEuNjtmb250LXNpemU6MTRweH0uaGludC1ncmlke2Rpc3BsYXk6Z3JpZDtncmlkLXRlbXBsYXRlLWNvbHVtbnM6cmVwZWF0KDMsMWZyKTtnYXA6MTBweDttYXJnaW4tdG9wOjI1cHh9LmhpbnR7Ym9yZGVyOjFweCBzb2xpZCB2YXIoLS1saW5lKTtiYWNrZ3JvdW5kOnZhcigtLXBhbmVsKTtwYWRkaW5nOjE0cHg7Ym9yZGVyLXJhZGl1czoxMnB4O3RleHQtYWxpZ246bGVmdDtmb250LXNpemU6MTJweDtjb2xvcjp2YXIoLS1tdXRlZCl9LmhpbnQgc3Ryb25ne2Rpc3BsYXk6YmxvY2s7Y29sb3I6dmFyKC0tdGV4dCk7bWFyZ2luLWJvdHRvbTo1cHh9Ci5tZXNzYWdle2Rpc3BsYXk6ZmxleDttYXJnaW46MCBhdXRvIDE2cHg7bWF4LXdpZHRoOjkwMHB4O3dpZHRoOjEwMCV9Lm1lc3NhZ2UudXNlcntqdXN0aWZ5LWNvbnRlbnQ6ZmxleC1lbmR9LmJ1YmJsZXttYXgtd2lkdGg6NzglO3BhZGRpbmc6MTNweCAxNXB4O2JvcmRlcjoxcHggc29saWQgdmFyKC0tbGluZSk7Ym9yZGVyLXJhZGl1czoxNHB4O2JhY2tncm91bmQ6dmFyKC0tcGFuZWwpO2ZvbnQtc2l6ZToxNHB4O2xpbmUtaGVpZ2h0OjEuNjt3aGl0ZS1zcGFjZTpwcmUtd3JhcDtvdmVyZmxvdy13cmFwOmFueXdoZXJlfS5tZXNzYWdlLnVzZXIgLmJ1YmJsZXtiYWNrZ3JvdW5kOnZhcigtLWdyYWRpZW50KTtjb2xvcjojMTkwOTE0O2JvcmRlcjowfS5tZXNzYWdlLmFzc2lzdGFudCAuYnViYmxle2JhY2tncm91bmQ6dmFyKC0tcGFuZWwpfS5tZXNzYWdlLWxhYmVse2ZvbnQtc2l6ZToxMHB4O3RleHQtdHJhbnNmb3JtOnVwcGVyY2FzZTtsZXR0ZXItc3BhY2luZzouMDhlbTtjb2xvcjp2YXIoLS1tdXRlZCk7bWFyZ2luLWJvdHRvbTo1cHh9LmFzc2lzdGFudC1ib2R5e3doaXRlLXNwYWNlOnByZS13cmFwfS5tZXRhe2Rpc3BsYXk6ZmxleDtmbGV4LXdyYXA6d3JhcDtnYXA6OHB4IDE0cHg7Ym9yZGVyLXRvcDoxcHggc29saWQgdmFyKC0tbGluZSk7bWFyZ2luLXRvcDoxMnB4O3BhZGRpbmctdG9wOjlweDtmb250LXNpemU6MTBweDtjb2xvcjp2YXIoLS1tdXRlZCl9Lm1ldGEgYntjb2xvcjp2YXIoLS1waW5rKX0uZmlsZXMtbGluZXtmb250LXNpemU6MTFweDtjb2xvcjp2YXIoLS1ncmVlbik7bWFyZ2luLXRvcDo5cHh9LmFydGlmYWN0e2Rpc3BsYXk6ZmxleDtmbGV4LXdyYXA6d3JhcDtnYXA6OHB4O21hcmdpbi10b3A6MTJweH0uYXJ0aWZhY3QgYXtib3JkZXI6MXB4IHNvbGlkIHZhcigtLXB1cnBsZSk7YmFja2dyb3VuZDpyZ2JhKDE2OCw4NSwyNDcsLjEpO2NvbG9yOiNkOGI0ZmU7dGV4dC1kZWNvcmF0aW9uOm5vbmU7cGFkZGluZzo3cHggMTBweDtib3JkZXItcmFkaXVzOjhweDtmb250LXNpemU6MTFweH0uYXJ0aWZhY3QgaW1ne21heC13aWR0aDoxMDAlO2JvcmRlci1yYWRpdXM6MTBweDtib3JkZXI6MXB4IHNvbGlkIHZhcigtLWxpbmUpO21hcmdpbi10b3A6OHB4fWltZy5hcnRpZmFjdHtkaXNwbGF5OmJsb2NrO3dpZHRoOmF1dG87aGVpZ2h0OmF1dG87bWF4LXdpZHRoOjMyMHB4O21heC1oZWlnaHQ6MjQwcHg7b2JqZWN0LWZpdDpjb250YWluO2JvcmRlci1yYWRpdXM6MTBweDtib3JkZXI6MXB4IHNvbGlkIHZhcigtLWxpbmUpO21hcmdpbi10b3A6OHB4O2N1cnNvcjp6b29tLWluO2JhY2tncm91bmQ6dmFyKC0tcGFuZWwyKX0uZGFuZ2VyLW5vdGV7Y29sb3I6dmFyKC0tcmVkKTtmb250LXNpemU6MTFweH0KLmNvbXBvc2VyLXdyYXB7cG9zaXRpb246YWJzb2x1dGU7Ym90dG9tOjA7bGVmdDowO3JpZ2h0OjA7cGFkZGluZzoxMnB4IDIycHggMjBweDtiYWNrZ3JvdW5kOmxpbmVhci1ncmFkaWVudCh0byB0b3AsdmFyKC0tYmcpIDc4JSx0cmFuc3BhcmVudCl9LmF0dGFjaG1lbnR7bWF4LXdpZHRoOjkwMHB4O21hcmdpbjowIGF1dG8gOHB4O2JvcmRlcjoxcHggc29saWQgcmdiYSg1MiwyMTEsMTUzLC4zKTtiYWNrZ3JvdW5kOnZhcigtLXBhbmVsKTtib3JkZXItcmFkaXVzOjEwcHg7cGFkZGluZzo4cHggMTFweDtmb250LXNpemU6MTFweDtjb2xvcjp2YXIoLS1ncmVlbik7ZGlzcGxheTpub25lO2FsaWduLWl0ZW1zOmNlbnRlcjtqdXN0aWZ5LWNvbnRlbnQ6c3BhY2UtYmV0d2Vlbn0uYXR0YWNobWVudCBidXR0b257Ym9yZGVyOjA7YmFja2dyb3VuZDp0cmFuc3BhcmVudDtjb2xvcjp2YXIoLS1yZWQpO2ZvbnQtc2l6ZToxOHB4fS5jb21wb3NlcnttYXgtd2lkdGg6OTAwcHg7bWFyZ2luOmF1dG87ZGlzcGxheTpmbGV4O2FsaWduLWl0ZW1zOmNlbnRlcjtnYXA6OXB4O2JhY2tncm91bmQ6dmFyKC0tcGFuZWwpO2JvcmRlcjoxcHggc29saWQgdmFyKC0tbGluZSk7Ym9yZGVyLXJhZGl1czoyOHB4O3BhZGRpbmc6N3B4IDhweCA3cHggMTNweH0uY29tcG9zZXI6Zm9jdXMtd2l0aGlue2JvcmRlci1jb2xvcjp2YXIoLS1wdXJwbGUpfS51cGxvYWR7Zm9udC1zaXplOjE4cHg7Y29sb3I6dmFyKC0tbXV0ZWQpO2N1cnNvcjpwb2ludGVyfS51cGxvYWQgaW5wdXR7ZGlzcGxheTpub25lfS5jb21wb3NlciBpbnB1dFt0eXBlPXRleHRde2ZsZXg6MTttaW4td2lkdGg6MDtib3JkZXI6MDtvdXRsaW5lOjA7YmFja2dyb3VuZDp0cmFuc3BhcmVudDtjb2xvcjp2YXIoLS10ZXh0KTtwYWRkaW5nOjlweCA0cHg7Zm9udC1zaXplOjE0cHh9LnNlbmR7d2lkdGg6MzlweDtoZWlnaHQ6MzlweDtib3JkZXI6MDtib3JkZXItcmFkaXVzOjUwJTtiYWNrZ3JvdW5kOnZhcigtLWdyYWRpZW50KTtmb250LXdlaWdodDo5MDA7Y29sb3I6IzE5MDkxNH0uY29tcG9zZXIuZGlzYWJsZWR7b3BhY2l0eTouNDU7cG9pbnRlci1ldmVudHM6bm9uZX0uY29tcG9zZXItbm90ZXttYXgtd2lkdGg6OTAwcHg7bWFyZ2luOjZweCBhdXRvIDA7dGV4dC1hbGlnbjpjZW50ZXI7Zm9udC1zaXplOjEwcHg7Y29sb3I6IzY2Nn0KLm1vZGFse3Bvc2l0aW9uOmZpeGVkO2luc2V0OjA7ei1pbmRleDoxMjA7YmFja2dyb3VuZDpyZ2JhKDAsMCwwLC43KTtkaXNwbGF5OmZsZXg7YWxpZ24taXRlbXM6Y2VudGVyO2p1c3RpZnktY29udGVudDpjZW50ZXJ9Lm1vZGFsLWNhcmR7d2lkdGg6bWluKDUwMHB4LDkydncpO2JhY2tncm91bmQ6dmFyKC0tcGFuZWwpO2JvcmRlcjoxcHggc29saWQgdmFyKC0tbGluZSk7Ym9yZGVyLXJhZGl1czoxNnB4O3BhZGRpbmc6MjBweH0ubW9kYWwtaGVhZHtkaXNwbGF5OmZsZXg7anVzdGlmeS1jb250ZW50OnNwYWNlLWJldHdlZW47YWxpZ24taXRlbXM6Y2VudGVyfS5tb2RhbC1oZWFkIGgze21hcmdpbjowfS5jbG9zZXtib3JkZXI6MDtiYWNrZ3JvdW5kOnRyYW5zcGFyZW50O2NvbG9yOnZhcigtLW11dGVkKTtmb250LXNpemU6MjJweH0uYWRtaW4tZ3JpZHtkaXNwbGF5OmdyaWQ7Z3JpZC10ZW1wbGF0ZS1jb2x1bW5zOjFmciAxZnI7Z2FwOjEwcHh9LnJlZ2lzdGVyLXN3aXRjaHtkaXNwbGF5OmZsZXg7Z2FwOjdweDttYXJnaW46MTVweCAwfS5yZWdpc3Rlci1zd2l0Y2ggYnV0dG9ue2ZsZXg6MTtib3JkZXI6MXB4IHNvbGlkICMzMDMwMzA7YmFja2dyb3VuZDojMTUxNTE1O2NvbG9yOiNhYWE7Ym9yZGVyLXJhZGl1czo5cHg7cGFkZGluZzo5cHh9LnJlZ2lzdGVyLXN3aXRjaCBidXR0b24uYWN0aXZle2JvcmRlci1jb2xvcjp2YXIoLS1wdXJwbGUpO2NvbG9yOiNmZmY7YmFja2dyb3VuZDpyZ2JhKDE2OCw4NSwyNDcsLjEyKX0ucmVxdWVzdC1jYXJke2JvcmRlcjoxcHggc29saWQgdmFyKC0tbGluZSk7Ym9yZGVyLXJhZGl1czoxMHB4O3BhZGRpbmc6MTBweDttYXJnaW46OHB4IDA7YmFja2dyb3VuZDp2YXIoLS1wYW5lbDIpfS5yZXF1ZXN0LWNhcmQgc3Ryb25ne2Rpc3BsYXk6YmxvY2s7bWFyZ2luLWJvdHRvbTo0cHh9LnJlcXVlc3QtYWN0aW9uc3tkaXNwbGF5OmZsZXg7Z2FwOjdweDttYXJnaW4tdG9wOjhweH0ucmVxdWVzdC1hY3Rpb25zIGJ1dHRvbntmbGV4OjF9Lm1lbWJlci1saXN0e21heC1oZWlnaHQ6MjIwcHg7b3ZlcmZsb3c6YXV0bztib3JkZXI6MXB4IHNvbGlkIHZhcigtLWxpbmUpO2JvcmRlci1yYWRpdXM6MTBweDtwYWRkaW5nOjZweH0ubWVtYmVyLXJvd3tkaXNwbGF5OmZsZXg7YWxpZ24taXRlbXM6Y2VudGVyO2p1c3RpZnktY29udGVudDpzcGFjZS1iZXR3ZWVuO2dhcDo4cHg7cGFkZGluZzo4cHg7Ym9yZGVyLWJvdHRvbToxcHggc29saWQgdmFyKC0tbGluZSk7Zm9udC1zaXplOjEycHh9Lm1lbWJlci1yb3c6bGFzdC1jaGlsZHtib3JkZXItYm90dG9tOjB9LnNtYWxsLWJ0bntib3JkZXI6MXB4IHNvbGlkIHZhcigtLWxpbmUpO2JhY2tncm91bmQ6dHJhbnNwYXJlbnQ7Y29sb3I6dmFyKC0tdGV4dCk7Ym9yZGVyLXJhZGl1czo3cHg7cGFkZGluZzo2cHggOXB4O2ZvbnQtc2l6ZToxMXB4fS5hZG1pbi1hY3Rpb25ze21hcmdpbi10b3A6MTRweDtkaXNwbGF5OmZsZXg7Z2FwOjhweH0uYWRtaW4tYWN0aW9ucyBidXR0b257ZmxleDoxfS50ZWFtLXNlY3Rpb257bWFyZ2luLXRvcDoxNHB4O3BhZGRpbmctdG9wOjE0cHg7Ym9yZGVyLXRvcDoxcHggc29saWQgdmFyKC0tbGluZSl9LnRlYW0tc2VjdGlvbiBoNHttYXJnaW46MCAwIDhweDtmb250LXNpemU6MTJweH0udGVhbS1jb2Rle2ZvbnQtZmFtaWx5OnVpLW1vbm9zcGFjZSxtb25vc3BhY2U7Zm9udC1zaXplOjEycHg7Y29sb3I6dmFyKC0tZ3JlZW4pO2JhY2tncm91bmQ6dmFyKC0tcGFuZWwyKTtwYWRkaW5nOjhweDtib3JkZXItcmFkaXVzOjhweDtib3JkZXI6MXB4IHNvbGlkIHZhcigtLWxpbmUpO3dvcmQtYnJlYWs6YnJlYWstYWxsfS5tZW1iZXItY2hlY2t7ZGlzcGxheTpmbGV4O2FsaWduLWl0ZW1zOmNlbnRlcjtnYXA6OHB4O3BhZGRpbmc6N3B4O2JvcmRlci1ib3R0b206MXB4IHNvbGlkIHZhcigtLWxpbmUpO2ZvbnQtc2l6ZToxMXB4fS5tZW1iZXItY2hlY2s6bGFzdC1jaGlsZHtib3JkZXItYm90dG9tOjB9Lm1lbWJlci1jaGVjayBpbnB1dHthY2NlbnQtY29sb3I6I2E4NTVmN30udGVhbS1hY3Rpb25ze2Rpc3BsYXk6ZmxleDtnYXA6OHB4O21hcmdpbi10b3A6OXB4fS50ZWFtLWFjdGlvbnMgYnV0dG9ue2ZsZXg6MX0KQG1lZGlhKG1heC13aWR0aDo4NTBweCl7I2FwcHtncmlkLXRlbXBsYXRlLWNvbHVtbnM6MjIwcHggbWlubWF4KDAsMWZyKX0uaGludC1ncmlke2dyaWQtdGVtcGxhdGUtY29sdW1uczoxZnJ9LmJ1YmJsZXttYXgtd2lkdGg6OTAlfS5oZWFkZXItcmlnaHQgLm1vZGUtcGlsbHtkaXNwbGF5Om5vbmV9fQpAbWVkaWEobWF4LXdpZHRoOjY1MHB4KXsucm9ib3Qtc3RhZ2V7d2lkdGg6MjIwcHg7aGVpZ2h0OjIyMHB4fS5yb2JvdC1zdGFnZSBzdmd7d2lkdGg6MjIwcHg7aGVpZ2h0OjIyMHB4fS5zcGxhc2gtYnJhbmR7Zm9udC1zaXplOjI0cHh9LnNwbGFzaC1xdW90ZXtmb250LXNpemU6MTNweH19CkBtZWRpYShtYXgtd2lkdGg6NjUwcHgpeyNhcHB7Z3JpZC10ZW1wbGF0ZS1jb2x1bW5zOjFmcn0uc2lkZWJhcntkaXNwbGF5Om5vbmV9LmhlYWRlcntwYWRkaW5nOjAgMTRweH0ud29ya3NwYWNle3BhZGRpbmc6MThweCAxMnB4IDE0NXB4fS5jb21wb3Nlci13cmFwe3BhZGRpbmc6MTBweCAxMHB4IDE1cHh9fQoKLyogR0xPQkFMIFRIRU1FIENPTlRST0wg4oCUIHZpc2libGUgYmVmb3JlIGxvZ2luIGFuZCBvbiBzcGxhc2ggKi8KLmdsb2JhbC10aGVtZXtwb3NpdGlvbjpmaXhlZDt0b3A6MThweDtyaWdodDoyMHB4O3otaW5kZXg6NTAwO3dpZHRoOjQycHg7aGVpZ2h0OjQycHg7Ym9yZGVyLXJhZGl1czoxMnB4O2JvcmRlcjoxcHggc29saWQgdmFyKC0tbGluZSk7YmFja2dyb3VuZDpyZ2JhKDI1NSwyNTUsMjU1LC4xMCk7YmFja2Ryb3AtZmlsdGVyOmJsdXIoMTRweCk7Y29sb3I6dmFyKC0tdGV4dCk7Zm9udC1zaXplOjE4cHg7Y3Vyc29yOnBvaW50ZXI7Ym94LXNoYWRvdzowIDhweCAzMHB4IHJnYmEoMCwwLDAsLjE4KX0KLmdsb2JhbC10aGVtZTpob3Zlcntib3JkZXItY29sb3I6dmFyKC0tcHVycGxlKTt0cmFuc2Zvcm06dHJhbnNsYXRlWSgtMXB4KX0KLnNwbGFzaC1jb250cm9sc3twb3NpdGlvbjphYnNvbHV0ZTt0b3A6MDtyaWdodDowO3otaW5kZXg6NX0KLnNwbGFzaC1jb250cm9scyAuZ2xvYmFsLXRoZW1le3JpZ2h0Ojc0cHh9CmJvZHkubGlnaHQgI2xvZ2lue2JhY2tncm91bmQ6cmFkaWFsLWdyYWRpZW50KGNpcmNsZSBhdCAyMCUgMjAlLHJnYmEoMTY4LDg1LDI0NywuMjApLHRyYW5zcGFyZW50IDQwJSkscmFkaWFsLWdyYWRpZW50KGNpcmNsZSBhdCA4MCUgODAlLHJnYmEoOTYsMTY1LDI1MCwuMTYpLHRyYW5zcGFyZW50IDQyJSksbGluZWFyLWdyYWRpZW50KDEzNWRlZywjZmZmIDAlLCNmYWY3ZmYgNDUlLCNmM2U4ZmYgMTAwJSl9CmJvZHkubGlnaHQgLmxvZ2luLWNhcmR7YmFja2dyb3VuZDpyZ2JhKDI1NSwyNTUsMjU1LC45NCk7Ym9yZGVyLWNvbG9yOiNkZGQwZjU7Ym94LXNoYWRvdzowIDMwcHggMTAwcHggcmdiYSg5MSwzMywxODIsLjE1KX0KYm9keS5saWdodCAuZmllbGQgaW5wdXR7YmFja2dyb3VuZDojZmZmO2NvbG9yOiMxODE4MWI7Ym9yZGVyLWNvbG9yOiNkZGQwZjV9CmJvZHkubGlnaHQgLnJlZ2lzdGVyLXN3aXRjaCBidXR0b257YmFja2dyb3VuZDojZmFmN2ZmO2NvbG9yOiM2YjVhN2Q7Ym9yZGVyLWNvbG9yOiNkZGQwZjV9CmJvZHkubGlnaHQgLnJlZ2lzdGVyLXN3aXRjaCBidXR0b24uYWN0aXZle2JhY2tncm91bmQ6cmdiYSgxNjgsODUsMjQ3LC4xMCk7Y29sb3I6IzViMjFiNjtib3JkZXItY29sb3I6cmdiYSgxNjgsODUsMjQ3LC4zNSl9CmJvZHkubGlnaHQgI3NwbGFzaHtiYWNrZ3JvdW5kOnJhZGlhbC1ncmFkaWVudChjaXJjbGUgYXQgNTAlIDE4JSxyZ2JhKDE2OCw4NSwyNDcsLjIyKSx0cmFuc3BhcmVudCA0MiUpLGxpbmVhci1ncmFkaWVudCgxMzVkZWcsI2ZmZiAwJSwjZmFmN2ZmIDUyJSwjZWRlOWZlIDEwMCUpfQpib2R5LmxpZ2h0IC5zcGxhc2gtd2VsY29tZSxib2R5LmxpZ2h0IC5zcGxhc2gtcXVvdGV7Y29sb3I6IzZiNWE3ZH0KYm9keS5saWdodCAuc3BsYXNoLWJyYW5ke2NvbG9yOiMyZjE3NGV9CmJvZHkubGlnaHQgLnNwbGFzaC1za2lwe2NvbG9yOiM2YjVhN2Q7Ym9yZGVyLWNvbG9yOiNkZGQwZjV9Ci5hcnRpZmFjdC1hY3Rpb25ze2Rpc3BsYXk6ZmxleDtmbGV4LXdyYXA6d3JhcDtnYXA6N3B4O21hcmdpbi10b3A6MTBweH0KLmFydGlmYWN0LWFjdGlvbnMgYSwuYXJ0aWZhY3QtYWN0aW9ucyBidXR0b257Ym9yZGVyOjFweCBzb2xpZCB2YXIoLS1wdXJwbGUpO2JhY2tncm91bmQ6cmdiYSgxNjgsODUsMjQ3LC4xMCk7Y29sb3I6I2Q4YjRmZTt0ZXh0LWRlY29yYXRpb246bm9uZTtwYWRkaW5nOjdweCAxMHB4O2JvcmRlci1yYWRpdXM6OHB4O2ZvbnQtc2l6ZToxMXB4O2N1cnNvcjpwb2ludGVyfQpib2R5LmxpZ2h0IC5hcnRpZmFjdC1hY3Rpb25zIGEsYm9keS5saWdodCAuYXJ0aWZhY3QtYWN0aW9ucyBidXR0b257Y29sb3I6IzViMjFiNjtiYWNrZ3JvdW5kOnJnYmEoMTI0LDU4LDIzNywuMDcpfQouYW5zd2VyLXRvb2xze2Rpc3BsYXk6ZmxleDtmbGV4LXdyYXA6d3JhcDtnYXA6N3B4O21hcmdpbi10b3A6MTBweH0KLmFuc3dlci10b29scyBidXR0b257Ym9yZGVyOjFweCBzb2xpZCB2YXIoLS1saW5lKTtiYWNrZ3JvdW5kOnZhcigtLXBhbmVsMik7Y29sb3I6dmFyKC0tdGV4dCk7Ym9yZGVyLXJhZGl1czo4cHg7cGFkZGluZzo3cHggMTBweDtmb250LXNpemU6MTFweH0KLmFuc3dlci10b29scyBidXR0b246aG92ZXJ7Ym9yZGVyLWNvbG9yOnZhcigtLXB1cnBsZSl9Ci5hcnRpZmFjdCBpbWd7d2lkdGg6YXV0bzttYXgtd2lkdGg6MzIwcHg7bWF4LWhlaWdodDoyNDBweDtvYmplY3QtZml0OmNvbnRhaW59CkBtZWRpYShtYXgtd2lkdGg6NjUwcHgpey5nbG9iYWwtdGhlbWV7dG9wOjEycHg7cmlnaHQ6MTJweDt3aWR0aDozOHB4O2hlaWdodDozOHB4fS5zcGxhc2gtY29udHJvbHMgLmdsb2JhbC10aGVtZXtyaWdodDo2MnB4fS5idWJibGV7bWF4LXdpZHRoOjk0JX19Ci8qIFBST0ZFU1NJT05BTCBPUklPTiBMSUdIVCBUSEVNRSAqLwpib2R5LmxpZ2h0ewogIC0tYmc6I2Y4ZjVmZjstLWJnMjpyZ2JhKDI1NSwyNTUsMjU1LC43OCk7LS1wYW5lbDpyZ2JhKDI1NSwyNTUsMjU1LC44OCk7LS1wYW5lbDI6cmdiYSgyNDYsMjQwLDI1NSwuODIpOwogIC0tbGluZTpyZ2JhKDkxLDMzLDE4MiwuMTQpOy0tdGV4dDojMjQxNDNiOy0tbXV0ZWQ6IzZkNWM3ZDsKICBiYWNrZ3JvdW5kOmxpbmVhci1ncmFkaWVudCgxMzVkZWcsI2ZmZmZmZiAwJSwjZmJmN2ZmIDI4JSwjZjJlOGZmIDU4JSwjZWFkY2ZmIDc4JSwjZmZmZmZmIDEwMCUpOwp9CmJvZHkubGlnaHQ6OmJlZm9yZXtjb250ZW50OiIiO3Bvc2l0aW9uOmZpeGVkO2luc2V0OjA7cG9pbnRlci1ldmVudHM6bm9uZTt6LWluZGV4Oi0xO2JhY2tncm91bmQ6CiAgcmFkaWFsLWdyYWRpZW50KGNpcmNsZSBhdCAxMiUgMTglLHJnYmEoMTY4LDg1LDI0NywuMTgpLHRyYW5zcGFyZW50IDMyJSksCiAgcmFkaWFsLWdyYWRpZW50KGNpcmNsZSBhdCA4OCUgNzglLHJnYmEoMTI0LDU4LDIzNywuMTMpLHRyYW5zcGFyZW50IDMwJSksCiAgbGluZWFyLWdyYWRpZW50KDEzNWRlZyxyZ2JhKDI1NSwyNTUsMjU1LC41NSkscmdiYSgyMzMsMjEzLDI1NSwuMTgpKTt9CmJvZHkubGlnaHQgI2FwcCxib2R5LmxpZ2h0IG1haW57YmFja2dyb3VuZDp0cmFuc3BhcmVudH0KYm9keS5saWdodCAuc2lkZWJhcntiYWNrZ3JvdW5kOnJnYmEoMjU1LDI1NSwyNTUsLjY2KTtiYWNrZHJvcC1maWx0ZXI6Ymx1cigyMnB4KTtib3JkZXItcmlnaHQtY29sb3I6cmdiYSg5MSwzMywxODIsLjEzKX0KYm9keS5saWdodCAuaGVhZGVye2JhY2tncm91bmQ6cmdiYSgyNTUsMjU1LDI1NSwuNDIpO2JhY2tkcm9wLWZpbHRlcjpibHVyKDE4cHgpO2JvcmRlci1ib3R0b20tY29sb3I6cmdiYSg5MSwzMywxODIsLjEzKX0KYm9keS5saWdodCAud29ya3NwYWNle2JhY2tncm91bmQ6dHJhbnNwYXJlbnR9CmJvZHkubGlnaHQgLm1lc3NhZ2UuYXNzaXN0YW50IC5idWJibGV7YmFja2dyb3VuZDpsaW5lYXItZ3JhZGllbnQoMTQ1ZGVnLCMyYjEyNDggMCUsIzRjMWQ3OCA1OCUsIzM1MTI1YiAxMDAlKTtjb2xvcjojZmZmO2JvcmRlci1jb2xvcjpyZ2JhKDI1NSwyNTUsMjU1LC4xMik7Ym94LXNoYWRvdzowIDE2cHggNDJweCByZ2JhKDc2LDI5LDEyMCwuMTgpfQpib2R5LmxpZ2h0IC5tZXNzYWdlLmFzc2lzdGFudCAuYnViYmxlIC5zZW5kZXItbGFiZWx7Y29sb3I6I2ZmZn0KYm9keS5saWdodCAubWVzc2FnZS5hc3Npc3RhbnQgLmJ1YmJsZSAuYXNzaXN0YW50LWJvZHl7Y29sb3I6I2Y4ZjJmZn0KYm9keS5saWdodCAubWVzc2FnZS5hc3Npc3RhbnQgLmJ1YmJsZSAubWV0YXtib3JkZXItdG9wLWNvbG9yOnJnYmEoMjU1LDI1NSwyNTUsLjE2KTtjb2xvcjojZDljN2U4fQpib2R5LmxpZ2h0IC5tZXNzYWdlLmFzc2lzdGFudCAuYnViYmxlIC5tZXRhIGJ7Y29sb3I6I2YwYWJmZn0KYm9keS5saWdodCAubWVzc2FnZS5hc3Npc3RhbnQgLmJ1YmJsZSAuZmlsZXMtbGluZXtjb2xvcjojODZlZmFjfQpib2R5LmxpZ2h0IC5jb21wb3Nlci13cmFwe2JhY2tncm91bmQ6bGluZWFyLWdyYWRpZW50KHRvIHRvcCxyZ2JhKDI0OCwyNDUsMjU1LC45NikgNzIlLHRyYW5zcGFyZW50KX0KYm9keS5saWdodCAuY29tcG9zZXJ7YmFja2dyb3VuZDpyZ2JhKDI1NSwyNTUsMjU1LC45MCk7Ym9yZGVyLWNvbG9yOnJnYmEoOTEsMzMsMTgyLC4xOCk7Ym94LXNoYWRvdzowIDE0cHggMzhweCByZ2JhKDkxLDMzLDE4MiwuMTIpO2JhY2tkcm9wLWZpbHRlcjpibHVyKDE4cHgpfQpib2R5LmxpZ2h0IC5oaW50LGJvZHkubGlnaHQgLmxpYnJhcnktaXRlbSxib2R5LmxpZ2h0IC5wcm9maWxlLGJvZHkubGlnaHQgLm1vZGFsLWNhcmR7YmFja2dyb3VuZDpyZ2JhKDI1NSwyNTUsMjU1LC43OCk7YmFja2Ryb3AtZmlsdGVyOmJsdXIoMThweCl9CmJvZHkubGlnaHQgLnN1Z2dlc3Rpb24tY2FyZHtiYWNrZ3JvdW5kOnJnYmEoMjU1LDI1NSwyNTUsLjY2KTtib3JkZXItY29sb3I6cmdiYSg5MSwzMywxODIsLjE0KX0KYm9keS5saWdodCAuc3VnZ2VzdGlvbi1jYXJkOmhvdmVye2JhY2tncm91bmQ6cmdiYSgyNTUsMjU1LDI1NSwuOSk7Ym9yZGVyLWNvbG9yOnJnYmEoMTI0LDU4LDIzNywuMzUpO3RyYW5zZm9ybTp0cmFuc2xhdGVZKC0ycHgpfQpib2R5LmxpZ2h0IC5hcnRpZmFjdC1hY3Rpb25zIGEsYm9keS5saWdodCAuYXJ0aWZhY3QtYWN0aW9ucyBidXR0b257YmFja2dyb3VuZDpsaW5lYXItZ3JhZGllbnQoMTM1ZGVnLHJnYmEoMTI0LDU4LDIzNywuMTMpLHJnYmEoMTY4LDg1LDI0NywuMDgpKTtib3JkZXItY29sb3I6cmdiYSgxMjQsNTgsMjM3LC4yOCk7Y29sb3I6IzViMjFiNn0KYm9keS5saWdodCAuc2VuZHtjb2xvcjojZmZmfQpib2R5LmxpZ2h0IC5oZXJvIGgxe2ZpbHRlcjpkcm9wLXNoYWRvdygwIDhweCAyNHB4IHJnYmEoMTI0LDU4LDIzNywuMTIpKX0KCgovKiBPUklPTiBQT0xJU0hFRCBMSUdIVCBDSEFUICovCmJvZHkubGlnaHQgLm1lc3NhZ2UudXNlciAuYnViYmxle2JhY2tncm91bmQ6bGluZWFyLWdyYWRpZW50KDEzNWRlZywjMzUxMTVmIDAlLCM1NDIwODAgNTUlLCM3YzNhZWQgMTAwJSk7Y29sb3I6I2ZmZjtib3gtc2hhZG93OjAgMTJweCAzMHB4IHJnYmEoNzYsMjksMTIwLC4xNil9CmJvZHkubGlnaHQgLm1lc3NhZ2UudXNlciAuYnViYmxlIC5zZW5kZXItbGFiZWx7Y29sb3I6I2Y1ZDBmZX0KYm9keS5saWdodCAuYXJ0aWZhY3QtYWN0aW9uc3twYWRkaW5nOjhweDtib3JkZXItcmFkaXVzOjEycHg7YmFja2dyb3VuZDpyZ2JhKDI1NSwyNTUsMjU1LC4xMCk7Ym9yZGVyOjFweCBzb2xpZCByZ2JhKDI1NSwyNTUsMjU1LC4xMil9CmJvZHkubGlnaHQgLmFydGlmYWN0LWFjdGlvbnMgYnV0dG9ue2JhY2tncm91bmQ6cmdiYSgyNTUsMjU1LDI1NSwuMTEpO2NvbG9yOiNmZmY7Ym9yZGVyLWNvbG9yOnJnYmEoMjU1LDI1NSwyNTUsLjI2KTtmb250LXdlaWdodDo2MDA7Ym94LXNoYWRvdzowIDVweCAxNnB4IHJnYmEoMCwwLDAsLjA4KX0KYm9keS5saWdodCAuYXJ0aWZhY3QtYWN0aW9ucyBidXR0b246aG92ZXJ7YmFja2dyb3VuZDpyZ2JhKDI1NSwyNTUsMjU1LC4xOSk7Ym9yZGVyLWNvbG9yOnJnYmEoMjU1LDI1NSwyNTUsLjUpO3RyYW5zZm9ybTp0cmFuc2xhdGVZKC0xcHgpfQpib2R5LmxpZ2h0IC5tZXNzYWdlLmFzc2lzdGFudCAuYXJ0aWZhY3R7cGFkZGluZzoxMHB4O2JvcmRlci1yYWRpdXM6MTRweDtiYWNrZ3JvdW5kOnJnYmEoMjU1LDI1NSwyNTUsLjA3KTtib3JkZXI6MXB4IHNvbGlkIHJnYmEoMjU1LDI1NSwyNTUsLjEwKX0KYm9keS5saWdodCAubWVzc2FnZS5hc3Npc3RhbnQgaW1nLmFydGlmYWN0e21heC13aWR0aDptaW4oNjgwcHgsMTAwJSk7bWF4LWhlaWdodDo0NjBweDtiYWNrZ3JvdW5kOiNmZmY7Ym9yZGVyLWNvbG9yOnJnYmEoMjU1LDI1NSwyNTUsLjIpO2JveC1zaGFkb3c6MCAxNHB4IDM2cHggcmdiYSgwLDAsMCwuMTgpfQpib2R5LmxpZ2h0IC5jb21wb3NlciBpbnB1dFt0eXBlPXRleHRdOjpwbGFjZWhvbGRlcntjb2xvcjojOGE3ODk5fQpib2R5LmxpZ2h0IC5jb21wb3Nlci1ub3Rle2NvbG9yOiM3NzY2ODZ9CgovKiBEb3dubG9hZCBjb250cm9scyAqLwouYW5zd2VyLXRvb2xzIC5kb3dubG9hZC1kb2N4e2JvcmRlci1jb2xvcjpyZ2JhKDE2OCw4NSwyNDcsLjU1KTtiYWNrZ3JvdW5kOnJnYmEoMTY4LDg1LDI0NywuMTIpfQouZG93bmxvYWQtYnVzeXtvcGFjaXR5Oi42NSFpbXBvcnRhbnQ7cG9pbnRlci1ldmVudHM6bm9uZX0KPC9zdHlsZT4KPC9oZWFkPgo8Ym9keT4KPGRpdiBpZD0ibG9naW4iIGNsYXNzPSJoaWRkZW4iPgogIDxidXR0b24gY2xhc3M9Imdsb2JhbC10aGVtZSIgaWQ9ImxvZ2luVGhlbWVCdG4iIG9uY2xpY2s9InRvZ2dsZVRoZW1lKCkiIHRpdGxlPSJUb2dnbGUgbGlnaHQvZGFyayB0aGVtZSI+8J+MmTwvYnV0dG9uPgogIDxkaXYgY2xhc3M9ImxvZ2luLWNhcmQiPgogICAgPGRpdiBjbGFzcz0ibG9naW4tbG9nbyI+T1JJT04g4oCUIFNPVkVSRUlHTiBXT1JLQkVOQ0g8L2Rpdj4KICAgIDxkaXYgY2xhc3M9ImxvZ2luLXN1YiI+U2VjdXJlIGxvY2FsIGVudGVycHJpc2Ugd29ya3NwYWNlLiBBY2NvdW50cywgYXBwcm92YWxzLCBub3RpZmljYXRpb25zIGFuZCBkb2N1bWVudCBwcm9jZXNzaW5nIGFyZSBoYW5kbGVkIGJ5IHRoZSBsb2NhbCBhcHBsaWNhdGlvbi48L2Rpdj4KICAgIDxkaXYgY2xhc3M9InJlZ2lzdGVyLXN3aXRjaCI+CiAgICAgIDxidXR0b24gaWQ9ImxvZ2luVGFiIiBjbGFzcz0iYWN0aXZlIiBvbmNsaWNrPSJzaG93QXV0aE1vZGUoJ2xvZ2luJykiPlNpZ24gaW48L2J1dHRvbj4KICAgICAgPGJ1dHRvbiBpZD0iZW1wbG95ZWVUYWIiIG9uY2xpY2s9InNob3dBdXRoTW9kZSgnZW1wbG95ZWUnKSI+RW1wbG95ZWUgcmVnaXN0ZXI8L2J1dHRvbj4KICAgICAgPGJ1dHRvbiBpZD0iYWRtaW5UYWIiIG9uY2xpY2s9InNob3dBdXRoTW9kZSgnYWRtaW4nKSI+QWRtaW4gcmVnaXN0ZXI8L2J1dHRvbj4KICAgIDwvZGl2PgogICAgPGZvcm0gaWQ9ImxvZ2luRm9ybSIgb25zdWJtaXQ9ImxvZ2luKGV2ZW50KSI+CiAgICAgIDxkaXYgY2xhc3M9ImZpZWxkIj48bGFiZWw+RW1haWw8L2xhYmVsPjxpbnB1dCBpZD0ibG9naW5FbWFpbCIgdHlwZT0iZW1haWwiIGF1dG9jb21wbGV0ZT0idXNlcm5hbWUiIHBsYWNlaG9sZGVyPSJFbnRlciB5b3VyIGVtYWlsIiByZXF1aXJlZD48L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0iZmllbGQiPjxsYWJlbD5QYXNzd29yZDwvbGFiZWw+PGlucHV0IGlkPSJsb2dpblBhc3N3b3JkIiB0eXBlPSJwYXNzd29yZCIgYXV0b2NvbXBsZXRlPSJjdXJyZW50LXBhc3N3b3JkIiBwbGFjZWhvbGRlcj0iRW50ZXIgcGFzc3dvcmQiIHJlcXVpcmVkPjwvZGl2PgogICAgICA8YnV0dG9uIGNsYXNzPSJwcmltYXJ5IiB0eXBlPSJzdWJtaXQiPlNpZ24gaW48L2J1dHRvbj4KICAgICAgPGRpdiBjbGFzcz0iZXJyb3IiIGlkPSJsb2dpbkVycm9yIj48L2Rpdj4KICAgIDwvZm9ybT4KICAgIDxmb3JtIGlkPSJyZWdpc3RlckZvcm0iIGNsYXNzPSJoaWRkZW4iIG9uc3VibWl0PSJyZWdpc3RlckFjY291bnQoZXZlbnQpIj4KICAgICAgPGRpdiBjbGFzcz0iZmllbGQiPjxsYWJlbD5OYW1lPC9sYWJlbD48aW5wdXQgaWQ9InJlZ05hbWUiIHJlcXVpcmVkIHBsYWNlaG9sZGVyPSJGdWxsIG5hbWUiPjwvZGl2PgogICAgICA8ZGl2IGNsYXNzPSJmaWVsZCI+PGxhYmVsPkVtYWlsPC9sYWJlbD48aW5wdXQgaWQ9InJlZ0VtYWlsIiB0eXBlPSJlbWFpbCIgcmVxdWlyZWQgcGxhY2Vob2xkZXI9IldvcmsgZW1haWwiPjwvZGl2PgogICAgICA8ZGl2IGNsYXNzPSJmaWVsZCI+PGxhYmVsPlBhc3N3b3JkPC9sYWJlbD48aW5wdXQgaWQ9InJlZ1Bhc3N3b3JkIiB0eXBlPSJwYXNzd29yZCIgcmVxdWlyZWQgcGxhY2Vob2xkZXI9IkNyZWF0ZSBwYXNzd29yZCI+PC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9ImZpZWxkIj48bGFiZWw+Q29tcGFueTwvbGFiZWw+PGlucHV0IGlkPSJyZWdDb21wYW55IiByZXF1aXJlZCBwbGFjZWhvbGRlcj0iT3JnYW5pemF0aW9uIC8gY29tcGFueSI+PC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9ImZpZWxkIiBpZD0icmVnU2VjdG9yV3JhcCI+PGxhYmVsPlNlY3RvcjwvbGFiZWw+PGlucHV0IGlkPSJyZWdTZWN0b3IiIHBsYWNlaG9sZGVyPSJFbmdpbmVlcmluZywgRmluYW5jZSwgSVTigKYiPjwvZGl2PgogICAgICA8YnV0dG9uIGNsYXNzPSJwcmltYXJ5IiB0eXBlPSJzdWJtaXQiIGlkPSJyZWdpc3RlclN1Ym1pdCI+UmVnaXN0ZXIgZW1wbG95ZWU8L2J1dHRvbj4KICAgICAgPGRpdiBjbGFzcz0iZXJyb3IiIGlkPSJyZWdpc3RlckVycm9yIj48L2Rpdj4KICAgIDwvZm9ybT4KICAgIDxkaXYgY2xhc3M9InNlY3VyZS1ub3RlIj7il48gTE9DQUwgLyBBSVItR0FQUEVEIMK3IFJlZ2lzdHJhdGlvbiByZXF1ZXN0cyBhbmQgYXBwcm92YWwgbm90aWZpY2F0aW9ucyByZW1haW4gaW5zaWRlIHRoZSBvcmdhbml6YXRpb24ncyBsb2NhbCBiYWNrZW5kLjwvZGl2PgogIDwvZGl2Pgo8L2Rpdj4KCjxkaXYgaWQ9InNwbGFzaCI+CiAgPGRpdiBjbGFzcz0ic3BsYXNoLWNvbnRyb2xzIj48YnV0dG9uIGNsYXNzPSJnbG9iYWwtdGhlbWUiIGlkPSJzcGxhc2hUaGVtZUJ0biIgb25jbGljaz0idG9nZ2xlVGhlbWUoKSIgdGl0bGU9IlRvZ2dsZSBsaWdodC9kYXJrIHRoZW1lIj7wn4yZPC9idXR0b24+PGJ1dHRvbiBjbGFzcz0ic3BsYXNoLXNraXAiIG9uY2xpY2s9ImxlYXZlU3BsYXNoKGV2ZW50KSI+U2tpcDwvYnV0dG9uPjwvZGl2PgogIDxkaXYgY2xhc3M9InNwbGFzaC1zdGFycyIgaWQ9InNwbGFzaFN0YXJzIj48L2Rpdj4KICA8ZGl2IGNsYXNzPSJyb2JvdC1zdGFnZSI+CiAgICA8ZGl2IGNsYXNzPSJyb2JvdC1nbG93Ij48L2Rpdj48ZGl2IGNsYXNzPSJyb2JvdC1zaGFkb3ciPjwvZGl2PgogICAgPHN2ZyB2aWV3Qm94PSIwIDAgMzAwIDMwMCIgd2lkdGg9IjI4MCIgaGVpZ2h0PSIyODAiIGFyaWEtbGFiZWw9Ik9yaW9uIHJvYm90Ij4KICAgICAgPGRlZnM+CiAgICAgICAgPGxpbmVhckdyYWRpZW50IGlkPSJib2R5R3JhZCIgeDE9IjAiIHkxPSIwIiB4Mj0iMCIgeTI9IjEiPgogICAgICAgICAgPHN0b3Agb2Zmc2V0PSIwJSIgc3RvcC1jb2xvcj0iI2ZmZmZmZiIvPgogICAgICAgICAgPHN0b3Agb2Zmc2V0PSIxMDAlIiBzdG9wLWNvbG9yPSIjZWNlM2ZiIi8+CiAgICAgICAgPC9saW5lYXJHcmFkaWVudD4KICAgICAgICA8bGluZWFyR3JhZGllbnQgaWQ9ImVhckdyYWQiIHgxPSIwIiB5MT0iMCIgeDI9IjEiIHkyPSIxIj4KICAgICAgICAgIDxzdG9wIG9mZnNldD0iMCUiIHN0b3AtY29sb3I9IiNjNGI1ZmQiLz4KICAgICAgICAgIDxzdG9wIG9mZnNldD0iMTAwJSIgc3RvcC1jb2xvcj0iI2Y1YjhkYyIvPgogICAgICAgIDwvbGluZWFyR3JhZGllbnQ+CiAgICAgIDwvZGVmcz4KICAgICAgPGxpbmUgeDE9IjE1MCIgeTE9IjQ2IiB4Mj0iMTUwIiB5Mj0iMjQiIHN0cm9rZT0iI2M5YjhmYiIgc3Ryb2tlLXdpZHRoPSI0IiBzdHJva2UtbGluZWNhcD0icm91bmQiLz4KICAgICAgPGNpcmNsZSBpZD0iYW50ZW5uYVRpcCIgY3g9IjE1MCIgY3k9IjIwIiByPSI1IiBmaWxsPSIjZTlkOGZmIi8+CiAgICAgIDxjaXJjbGUgY3g9IjgyIiBjeT0iMTE4IiByPSIyMiIgZmlsbD0idXJsKCNlYXJHcmFkKSIvPgogICAgICA8Y2lyY2xlIGN4PSIyMTgiIGN5PSIxMTgiIHI9IjIyIiBmaWxsPSJ1cmwoI2VhckdyYWQpIi8+CiAgICAgIDxjaXJjbGUgY3g9IjgyIiBjeT0iMTE4IiByPSIxMCIgZmlsbD0iI2ZmZiIgb3BhY2l0eT0iLjU1Ii8+CiAgICAgIDxjaXJjbGUgY3g9IjIxOCIgY3k9IjExOCIgcj0iMTAiIGZpbGw9IiNmZmYiIG9wYWNpdHk9Ii41NSIvPgogICAgICA8cmVjdCB4PSI5MiIgeT0iNTgiIHdpZHRoPSIxMTYiIGhlaWdodD0iMTA4IiByeD0iNDYiIGZpbGw9InVybCgjYm9keUdyYWQpIiBzdHJva2U9IiNlM2Q4ZmIiIHN0cm9rZS13aWR0aD0iMS41Ii8+CiAgICAgIDxyZWN0IHg9IjExMCIgeT0iODgiIHdpZHRoPSI4MCIgaGVpZ2h0PSI1NCIgcng9IjI0IiBmaWxsPSIjMjQxYjNhIi8+CiAgICAgIDxyZWN0IGlkPSJ2aXNvckdsb3ciIHg9IjExMCIgeT0iODgiIHdpZHRoPSI4MCIgaGVpZ2h0PSI1NCIgcng9IjI0IiBmaWxsPSJub25lIiBzdHJva2U9IiNiNzliZmEiIHN0cm9rZS13aWR0aD0iMS41IiBvcGFjaXR5PSIuNiIvPgogICAgICA8ZWxsaXBzZSBpZD0iZXllTCIgY3g9IjEzNiIgY3k9IjExNSIgcng9IjgiIHJ5PSIxMCIgZmlsbD0iI2Y0ZWNmZiIvPjxlbGxpcHNlIGlkPSJleWVSIiBjeD0iMTY0IiBjeT0iMTE1IiByeD0iOCIgcnk9IjEwIiBmaWxsPSIjZjRlY2ZmIi8+CiAgICAgIDxwYXRoIGQ9Ik0gMTM2IDEzMiBRIDE1MCAxNDAgMTY0IDEzMiIgc3Ryb2tlPSIjYjc5YmZhIiBzdHJva2Utd2lkdGg9IjIuNSIgZmlsbD0ibm9uZSIgc3Ryb2tlLWxpbmVjYXA9InJvdW5kIi8+CiAgICAgIDxyZWN0IHg9IjEzOCIgeT0iMTYwIiB3aWR0aD0iMjQiIGhlaWdodD0iMTQiIHJ4PSI2IiBmaWxsPSIjZTNkOGZiIi8+CiAgICAgIDxyZWN0IHg9Ijg0IiB5PSIxNzAiIHdpZHRoPSIxMzIiIGhlaWdodD0iMTAwIiByeD0iMzgiIGZpbGw9InVybCgjYm9keUdyYWQpIiBzdHJva2U9IiNlM2Q4ZmIiIHN0cm9rZS13aWR0aD0iMS41Ii8+CiAgICAgIDxjaXJjbGUgY3g9IjE1MCIgY3k9IjIxNCIgcj0iMTMiIGZpbGw9InVybCgjZWFyR3JhZCkiLz48Y2lyY2xlIGN4PSIxNTAiIGN5PSIyMTQiIHI9IjUiIGZpbGw9IiNmZmYiLz4KICAgICAgPGc+PHJlY3QgeD0iNjIiIHk9IjE5NiIgd2lkdGg9IjIyIiBoZWlnaHQ9IjQ2IiByeD0iMTEiIGZpbGw9IiNkOWNkZmEiLz48Y2lyY2xlIGN4PSI3MyIgY3k9IjI0OCIgcj0iMTQiIGZpbGw9InVybCgjYm9keUdyYWQpIiBzdHJva2U9IiNlM2Q4ZmIiIHN0cm9rZS13aWR0aD0iMS41Ii8+PC9nPgogICAgICA8ZyBpZD0iYXJtUmlnaHQiPjxyZWN0IHg9IjIxNiIgeT0iMTk2IiB3aWR0aD0iMjIiIGhlaWdodD0iNDYiIHJ4PSIxMSIgZmlsbD0iI2Q5Y2RmYSIvPjxjaXJjbGUgY3g9IjIyNyIgY3k9IjI0OCIgcj0iMTQiIGZpbGw9InVybCgjYm9keUdyYWQpIiBzdHJva2U9IiNlM2Q4ZmIiIHN0cm9rZS13aWR0aD0iMS41Ii8+PC9nPgogICAgICA8ZWxsaXBzZSBjeD0iMTUwIiBjeT0iMjg0IiByeD0iNDYiIHJ5PSIxMCIgZmlsbD0iI2M0YjVmZCIgb3BhY2l0eT0iLjMiLz4KICAgIDwvc3ZnPgogIDwvZGl2PgogIDxkaXYgY2xhc3M9InNwbGFzaC10ZXh0Ij4KICAgIDxkaXYgY2xhc3M9InNwbGFzaC13ZWxjb21lIj5XZWxjb21lIHRvPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJzcGxhc2gtYnJhbmQiPjxzcGFuPk9yaW9uPC9zcGFuPiDigJQgU292ZXJlaWduIFdvcmtiZW5jaDwvZGl2PgogICAgPGRpdiBjbGFzcz0ic3BsYXNoLXF1b3RlIj5UaGlua2luZyBsb2NhbGx5LiBXb3JraW5nIGdsb2JhbGx5LjwvZGl2PgogICAgPGRpdiBjbGFzcz0ic3BsYXNoLWVudGVyIj48YnV0dG9uIG9uY2xpY2s9ImxlYXZlU3BsYXNoKGV2ZW50KSI+RW50ZXIgV29ya2JlbmNoPC9idXR0b24+PC9kaXY+CiAgPC9kaXY+CjwvZGl2PgoKPGRpdiBpZD0iYXBwIiBjbGFzcz0iaGlkZGVuIj4KICA8YXNpZGUgY2xhc3M9InNpZGViYXIiPgogICAgPGRpdiBjbGFzcz0iYnJhbmQiPjxkaXYgY2xhc3M9ImJyYW5kLWxlZnQiPjxzcGFuIGNsYXNzPSJicmFuZC1kb3QiPjwvc3Bhbj5PcmlvbiA8c3BhbiBzdHlsZT0iZm9udC13ZWlnaHQ6NTAwO2NvbG9yOnZhcigtLW11dGVkKTtmb250LXNpemU6MTFweDttYXJnaW4tbGVmdDozcHgiPsK3IFNvdmVyZWlnbiBXb3JrYmVuY2g8L3NwYW4+PC9kaXY+PGJ1dHRvbiBjbGFzcz0iaWNvbi1idG4iIG9uY2xpY2s9InRvZ2dsZVRoZW1lKCkiIGlkPSJ0aGVtZUJ0biI+8J+MmTwvYnV0dG9uPjwvZGl2PgogICAgPGRpdiBjbGFzcz0ic2VjdXJlIj5ORVRXT1JLOiBMT0NBTCAvIEFJUi1HQVBQRUQgTU9ERTwvZGl2PgogICAgPGRpdiBjbGFzcz0icHJvZmlsZSI+PGRpdiBjbGFzcz0ibmFtZSIgaWQ9InByb2ZpbGVOYW1lIj7igJQ8L2Rpdj48ZGl2IGNsYXNzPSJtZXRhIiBpZD0icHJvZmlsZU1ldGEiPuKAlDwvZGl2PjwvZGl2PgogICAgPGJ1dHRvbiBjbGFzcz0ibmV3LWNoYXQiIG9uY2xpY2s9Im5ld0NoYXQoKSI+77yLIE5ldyBDaGF0PC9idXR0b24+CiAgICA8ZGl2IGNsYXNzPSJzaWRlYmFyLW5hdiI+CiAgICAgIDxidXR0b24gY2xhc3M9InNpZGUtbmF2LWJ0biIgb25jbGljaz0ib3BlblByb2ZpbGUoKSI+8J+RpCBQcm9maWxlPC9idXR0b24+CiAgICAgIDxidXR0b24gY2xhc3M9InNpZGUtbmF2LWJ0biIgb25jbGljaz0ib3BlbkxpYnJhcnkoKSI+8J+TmiBMaWJyYXJ5PC9idXR0b24+CiAgICAgIDxidXR0b24gY2xhc3M9InNpZGUtbmF2LWJ0biIgb25jbGljaz0ib3BlblRlYW1DaGF0TmF2KCkiPvCfkaUgVGVhbSBDaGF0PC9idXR0b24+CiAgICAgIDxidXR0b24gY2xhc3M9InNpZGUtbmF2LWJ0biIgb25jbGljaz0ib3BlblNldHRpbmdzKCkiPuKamSBTZXR0aW5nczwvYnV0dG9uPgogICAgPC9kaXY+CgogICAgPGRpdiBjbGFzcz0ic2VjdGlvbi10aXRsZSI+Q2hhdHM8L2Rpdj4KICAgIDxkaXYgY2xhc3M9ImNoYXQtbGlzdCIgaWQ9ImNoYXRMaXN0Ij48L2Rpdj4KCiAgICA8ZGl2IGNsYXNzPSJzZWN0aW9uLXRpdGxlIj5UZWFtIENoYXRzPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJ0ZWFtLWxpc3QiIGlkPSJ0ZWFtTGlzdCI+PC9kaXY+CgogICAgPGRpdiBjbGFzcz0ic2lkZWJhci1ib3R0b20iPgogICAgICA8ZGl2IGNsYXNzPSJzdGF0dXMtcm93Ij48c3Bhbj5PbGxhbWE8L3NwYW4+PHNwYW4gaWQ9Im9sbGFtYVN0YXR1cyI+Y2hlY2tpbmfigKY8L3NwYW4+PC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9InN0YXR1cy1yb3ciPjxzcGFuPk1vZGVsPC9zcGFuPjxzcGFuIGlkPSJtb2RlbFN0YXR1cyI+ZHluYW1pYzwvc3Bhbj48L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0ic3RhdHVzLXJvdyI+PHNwYW4+Q3VycmVudCBjaGF0PC9zcGFuPjxzcGFuIGlkPSJjaGF0U3RhdHVzIj5uZXc8L3NwYW4+PC9kaXY+CiAgICAgIDxidXR0b24gY2xhc3M9ImxvZ291dCIgb25jbGljaz0ibG9nb3V0KCkiPlNpZ24gb3V0PC9idXR0b24+CiAgICAgIDxidXR0b24gY2xhc3M9ImxvZ291dCBoaWRkZW4iIGlkPSJhZG1pbkJ0biIgb25jbGljaz0ib3BlbkFkbWluKCkiPkFkbWluOiBSZWdpc3RyYXRpb24gJiBFbXBsb3llZXM8L2J1dHRvbj4KICAgICAgPGJ1dHRvbiBjbGFzcz0ibG9nb3V0IGhpZGRlbiIgaWQ9InRlYW1BZG1pbkJ0biIgb25jbGljaz0ib3BlblRlYW1NYW5hZ2VyKCkiPlRlYW0gTWFuYWdlbWVudDwvYnV0dG9uPgogICAgPC9kaXY+CiAgPC9hc2lkZT4KCiAgPG1haW4+CiAgICA8aGVhZGVyIGNsYXNzPSJoZWFkZXIiPgogICAgICA8ZGl2IGNsYXNzPSJoZWFkZXItdGl0bGUiPjxzcGFuIGlkPSJoZWFkZXJUaXRsZSI+T3Jpb24g4oCUIFNvdmVyZWlnbiBXb3JrYmVuY2g8L3NwYW4+PC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9ImhlYWRlci1yaWdodCI+PHNwYW4gY2xhc3M9Im1vZGUtcGlsbCI+MTAwJSBMb2NhbCBQcm9jZXNzaW5nPC9zcGFuPjxzcGFuIGlkPSJoZWFkZXJQcm9maWxlIj7igJQ8L3NwYW4+PC9kaXY+CiAgICA8L2hlYWRlcj4KCiAgICA8c2VjdGlvbiBjbGFzcz0id29ya3NwYWNlIiBpZD0id29ya3NwYWNlIj4KICAgICAgPGRpdiBjbGFzcz0iaGVybyIgaWQ9Imhlcm8iPgogICAgICAgIDxoMT5PcmlvbiDigJQgU292ZXJlaWduIFdvcmtiZW5jaDwvaDE+CiAgICAgICAgPGRpdiBzdHlsZT0ibWFyZ2luLXRvcDo2cHg7Y29sb3I6dmFyKC0tcHVycGxlKTtmb250LXdlaWdodDo3MDA7Zm9udC1zaXplOjEycHgiPlRoaW5raW5nIGxvY2FsbHkuIFdvcmtpbmcgZ2xvYmFsbHkuPC9kaXY+CiAgICAgICAgPHA+VXBsb2FkIGEgY29uZmlkZW50aWFsIGVudGVycHJpc2UgZG9jdW1lbnQgYW5kIGFzayBhIHF1ZXN0aW9uLiBUaGUgYmFja2VuZCByZXRyaWV2ZXMgcmVsZXZhbnQgZXZpZGVuY2UgYmVmb3JlIHNlbmRpbmcgY29udGV4dCB0byBhIGxvY2FsbHkgYXZhaWxhYmxlIG1vZGVsLjwvcD4KICAgICAgICA8ZGl2IGNsYXNzPSJoaW50LWdyaWQiPgogICAgICAgICAgPGRpdiBjbGFzcz0iaGludCI+PHN0cm9uZz5Eb2N1bWVudCBRJkE8L3N0cm9uZz5Bc2sgZ3JvdW5kZWQgcXVlc3Rpb25zIGFib3V0IFBERiwgRE9DWCwgUFBUWCwgVFhUIG9yIE1hcmtkb3duIGZpbGVzLjwvZGl2PgogICAgICAgICAgPGRpdiBjbGFzcz0iaGludCI+PHN0cm9uZz5EYXRhIEFuYWx5c2lzPC9zdHJvbmc+VXBsb2FkIENTVi9FeGNlbCBmaWxlcyBmb3IgZGV0ZXJtaW5pc3RpYyBsb2NhbCBhbmFseXNpcyBhbmQgdmVyaWZpZWQgY2hhcnRzLjwvZGl2PgogICAgICAgICAgPGRpdiBjbGFzcz0iaGludCI+PHN0cm9uZz5WaXN1YWwgUmVhc29uaW5nPC9zdHJvbmc+VXNlIGltYWdlcyBvciB0ZWNobmljYWwgZHJhd2luZ3Mgd2hlbiB2aXN1YWwgZXZpZGVuY2UgaXMgYWN0dWFsbHkgcmVxdWlyZWQuPC9kaXY+CiAgICAgICAgPC9kaXY+CiAgICAgIDwvZGl2PgogICAgICA8ZGl2IGlkPSJjaGF0Ij48L2Rpdj4KICAgIDwvc2VjdGlvbj4KCiAgICA8ZGl2IGNsYXNzPSJjb21wb3Nlci13cmFwIj4KICAgICAgPGRpdiBjbGFzcz0iYXR0YWNobWVudCIgaWQ9ImF0dGFjaG1lbnQiPjxzcGFuPvCfk44gPHN0cm9uZyBpZD0iYXR0YWNobWVudE5hbWUiPjwvc3Ryb25nPjwvc3Bhbj48YnV0dG9uIG9uY2xpY2s9ImNsZWFyRmlsZXMoKSI+w5c8L2J1dHRvbj48L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0iY29tcG9zZXIgZGlzYWJsZWQiIGlkPSJjb21wb3NlciI+CiAgICAgICAgPGxhYmVsIGNsYXNzPSJ1cGxvYWQiIHRpdGxlPSJVcGxvYWQgZmlsZXMiPvCfk448aW5wdXQgaWQ9ImZpbGVJbnB1dCIgdHlwZT0iZmlsZSIgbXVsdGlwbGUgYWNjZXB0PSIucGRmLC5kb2N4LC5wcHR4LC54bHN4LC54bHMsLmNzdiwubWQsLnR4dCwuanNvbiwubG9nLC5wbmcsLmpwZywuanBlZywud2VicCwuaHRtbCIgb25jaGFuZ2U9ImhhbmRsZUZpbGVTZWxlY3QoKSI+PC9sYWJlbD4KICAgICAgICA8aW5wdXQgaWQ9InF1ZXJ5IiB0eXBlPSJ0ZXh0IiBwbGFjZWhvbGRlcj0iVXBsb2FkIGEgZmlsZSwgdGhlbiBhc2sgYSBxdWVzdGlvbuKApiIgb25rZXlkb3duPSJpZihldmVudC5rZXk9PT0nRW50ZXInJiYhZXZlbnQuc2hpZnRLZXkpe2V2ZW50LnByZXZlbnREZWZhdWx0KCk7c2VuZE1lc3NhZ2UoKX0iPgogICAgICAgIDxidXR0b24gY2xhc3M9InNlbmQiIG9uY2xpY2s9InNlbmRNZXNzYWdlKCkiPuKepDwvYnV0dG9uPgogICAgICA8L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0iY29tcG9zZXItbm90ZSI+VGhlIGZpcnN0IHJlcXVlc3Qgb3IgdXBsb2FkZWQgZmlsZW5hbWUgaXMgdXNlZCB0byBjcmVhdGUgYW4gaW50ZWxsaWdlbnQgY2hhdCB0aXRsZSBhdXRvbWF0aWNhbGx5LjwvZGl2PgogICAgPC9kaXY+CiAgPC9tYWluPgo8L2Rpdj4KCjxkaXYgY2xhc3M9Im1vZGFsIGhpZGRlbiIgaWQ9InByb2ZpbGVNb2RhbCI+CiAgPGRpdiBjbGFzcz0ibW9kYWwtY2FyZCI+CiAgICA8ZGl2IGNsYXNzPSJtb2RhbC1oZWFkIj48aDM+UHJvZmlsZTwvaDM+PGJ1dHRvbiBjbGFzcz0iY2xvc2UiIG9uY2xpY2s9ImNsb3NlTW9kYWwoJ3Byb2ZpbGVNb2RhbCcpIj7DlzwvYnV0dG9uPjwvZGl2PgogICAgPGRpdiBjbGFzcz0ic2V0dGluZ3Mtcm93Ij48c3Bhbj5OYW1lPC9zcGFuPjxzcGFuIGNsYXNzPSJzZXR0aW5ncy12YWx1ZSIgaWQ9InByb2ZpbGVNb2RhbE5hbWUiPuKAlDwvc3Bhbj48L2Rpdj4KICAgIDxkaXYgY2xhc3M9InNldHRpbmdzLXJvdyI+PHNwYW4+RW1haWw8L3NwYW4+PHNwYW4gY2xhc3M9InNldHRpbmdzLXZhbHVlIiBpZD0icHJvZmlsZU1vZGFsRW1haWwiPuKAlDwvc3Bhbj48L2Rpdj4KICAgIDxkaXYgY2xhc3M9InNldHRpbmdzLXJvdyI+PHNwYW4+Q29tcGFueTwvc3Bhbj48c3BhbiBjbGFzcz0ic2V0dGluZ3MtdmFsdWUiIGlkPSJwcm9maWxlTW9kYWxDb21wYW55Ij7igJQ8L3NwYW4+PC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJzZXR0aW5ncy1yb3ciPjxzcGFuPlNlY3Rvcjwvc3Bhbj48c3BhbiBjbGFzcz0ic2V0dGluZ3MtdmFsdWUiIGlkPSJwcm9maWxlTW9kYWxTZWN0b3IiPuKAlDwvc3Bhbj48L2Rpdj4KICAgIDxkaXYgY2xhc3M9InNldHRpbmdzLXJvdyI+PHNwYW4+Um9sZTwvc3Bhbj48c3BhbiBjbGFzcz0ic2V0dGluZ3MtdmFsdWUiIGlkPSJwcm9maWxlTW9kYWxSb2xlIj7igJQ8L3NwYW4+PC9kaXY+CiAgPC9kaXY+CjwvZGl2PgoKPGRpdiBjbGFzcz0ibW9kYWwgaGlkZGVuIiBpZD0ibGlicmFyeU1vZGFsIj4KICA8ZGl2IGNsYXNzPSJtb2RhbC1jYXJkIj4KICAgIDxkaXYgY2xhc3M9Im1vZGFsLWhlYWQiPjxoMz5MaWJyYXJ5PC9oMz48YnV0dG9uIGNsYXNzPSJjbG9zZSIgb25jbGljaz0iY2xvc2VNb2RhbCgnbGlicmFyeU1vZGFsJykiPsOXPC9idXR0b24+PC9kaXY+CiAgICA8ZGl2IHN0eWxlPSJmb250LXNpemU6MTFweDtjb2xvcjp2YXIoLS1tdXRlZCk7bWFyZ2luOjEwcHggMCI+RmlsZXMgYXZhaWxhYmxlIHRvIHlvdXIgYWNjb3VudCBhbmQgdGVhbSBjaGF0cy48L2Rpdj4KICAgIDxkaXYgaWQ9ImxpYnJhcnlMaXN0IiBjbGFzcz0ibGlicmFyeS1saXN0Ij48L2Rpdj4KICA8L2Rpdj4KPC9kaXY+Cgo8ZGl2IGNsYXNzPSJtb2RhbCBoaWRkZW4iIGlkPSJzZXR0aW5nc01vZGFsIj4KICA8ZGl2IGNsYXNzPSJtb2RhbC1jYXJkIj4KICAgIDxkaXYgY2xhc3M9Im1vZGFsLWhlYWQiPjxoMz5TZXR0aW5nczwvaDM+PGJ1dHRvbiBjbGFzcz0iY2xvc2UiIG9uY2xpY2s9ImNsb3NlTW9kYWwoJ3NldHRpbmdzTW9kYWwnKSI+w5c8L2J1dHRvbj48L2Rpdj4KICAgIDxkaXYgY2xhc3M9InNldHRpbmdzLXJvdyI+PHNwYW4+UHJvY2Vzc2luZyBtb2RlPC9zcGFuPjxzcGFuIGNsYXNzPSJzZXR0aW5ncy12YWx1ZSI+TE9DQUwgLyBBSVItR0FQUEVEPC9zcGFuPjwvZGl2PgogICAgPGRpdiBjbGFzcz0ic2V0dGluZ3Mtcm93Ij48c3Bhbj5PbGxhbWE8L3NwYW4+PHNwYW4gY2xhc3M9InNldHRpbmdzLXZhbHVlIiBpZD0ic2V0dGluZ3NPbGxhbWEiPmNoZWNraW5n4oCmPC9zcGFuPjwvZGl2PgogICAgPGRpdiBjbGFzcz0ic2V0dGluZ3Mtcm93Ij48c3Bhbj5SZXNvbHZlZCBtb2RlbDwvc3Bhbj48c3BhbiBjbGFzcz0ic2V0dGluZ3MtdmFsdWUiIGlkPSJzZXR0aW5nc01vZGVsIj5keW5hbWljPC9zcGFuPjwvZGl2PgogICAgPGRpdiBjbGFzcz0ic2V0dGluZ3Mtcm93Ij48c3Bhbj5JbnN0YWxsZWQgbG9jYWwgbW9kZWxzPC9zcGFuPjxzcGFuIGNsYXNzPSJzZXR0aW5ncy12YWx1ZSIgaWQ9InNldHRpbmdzTW9kZWxzIj7igJQ8L3NwYW4+PC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJzZXR0aW5ncy1yb3ciPjxzcGFuPkNoYXQ8L3NwYW4+PHNwYW4gY2xhc3M9InNldHRpbmdzLXZhbHVlIj5QZXJzaXN0ZW50IGluIGxvY2FsIFNRTGl0ZTwvc3Bhbj48L2Rpdj4KICAgIDxkaXYgY2xhc3M9InNldHRpbmdzLXJvdyI+PHNwYW4+RXZpZGVuY2UgcG9saWN5PC9zcGFuPjxzcGFuIGNsYXNzPSJzZXR0aW5ncy12YWx1ZSI+U291cmNlLWdyb3VuZGVkPC9zcGFuPjwvZGl2PgogICAgPGRpdiBzdHlsZT0ibWFyZ2luLXRvcDoxMnB4O2ZvbnQtc2l6ZToxMHB4O2NvbG9yOnZhcigtLW11dGVkKSI+VGhlIGludGVyZmFjZSBkb2VzIG5vdCBzZW5kIGRvY3VtZW50cyBvciBwcm9tcHRzIHRvIGV4dGVybmFsIEFJIHByb3ZpZGVycy48L2Rpdj4KICA8L2Rpdj4KPC9kaXY+Cgo8ZGl2IGNsYXNzPSJtb2RhbCBoaWRkZW4iIGlkPSJhZG1pbk1vZGFsIj4KICA8ZGl2IGNsYXNzPSJtb2RhbC1jYXJkIj4KICAgIDxkaXYgY2xhc3M9Im1vZGFsLWhlYWQiPjxoMz5BZG1pbmlzdHJhdG9yIENlbnRlcjwvaDM+PGJ1dHRvbiBjbGFzcz0iY2xvc2UiIG9uY2xpY2s9ImNsb3NlQWRtaW4oKSI+w5c8L2J1dHRvbj48L2Rpdj4KICAgIDxkaXYgc3R5bGU9Im1hcmdpbi10b3A6MTJweDtmb250LXNpemU6MTJweDtjb2xvcjp2YXIoLS1tdXRlZCkiPlBlbmRpbmcgZW1wbG95ZWUgcmVnaXN0cmF0aW9uczwvZGl2PgogICAgPGRpdiBpZD0icGVuZGluZ1JlcXVlc3RzIiBzdHlsZT0ibWFyZ2luLXRvcDo4cHgiPjwvZGl2PgogICAgPGRpdiBzdHlsZT0ibWFyZ2luLXRvcDoxNHB4O2ZvbnQtc2l6ZToxMnB4O2NvbG9yOnZhcigtLW11dGVkKSI+Q3JlYXRlIGVtcGxveWVlIGRpcmVjdGx5PC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJhZG1pbi1ncmlkIiBzdHlsZT0ibWFyZ2luLXRvcDo4cHgiPgogICAgICA8ZGl2IGNsYXNzPSJmaWVsZCI+PGxhYmVsPk5hbWU8L2xhYmVsPjxpbnB1dCBpZD0ibmV3RW1wbG95ZWVOYW1lIj48L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0iZmllbGQiPjxsYWJlbD5FbWFpbDwvbGFiZWw+PGlucHV0IGlkPSJuZXdFbXBsb3llZUVtYWlsIiB0eXBlPSJlbWFpbCI+PC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9ImZpZWxkIj48bGFiZWw+UGFzc3dvcmQ8L2xhYmVsPjxpbnB1dCBpZD0ibmV3RW1wbG95ZWVQYXNzd29yZCIgdHlwZT0icGFzc3dvcmQiPjwvZGl2PgogICAgICA8ZGl2IGNsYXNzPSJmaWVsZCI+PGxhYmVsPlNlY3RvcjwvbGFiZWw+PGlucHV0IGlkPSJuZXdFbXBsb3llZVNlY3RvciI+PC9kaXY+CiAgICA8L2Rpdj4KICAgIDxkaXYgY2xhc3M9ImFkbWluLWFjdGlvbnMiPjxidXR0b24gY2xhc3M9Im5ldy1jaGF0IiBvbmNsaWNrPSJjbG9zZUFkbWluKCkiPkNsb3NlPC9idXR0b24+PGJ1dHRvbiBjbGFzcz0icHJpbWFyeSIgb25jbGljaz0iY3JlYXRlRW1wbG95ZWUoKSI+Q3JlYXRlIGFwcHJvdmVkIGVtcGxveWVlPC9idXR0b24+PC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJlcnJvciIgaWQ9ImFkbWluRXJyb3IiPjwvZGl2PgoKICAgIDxkaXYgY2xhc3M9InRlYW0tc2VjdGlvbiI+CiAgICAgIDxoND5BZGQgU09QIC8gS25vd2xlZGdlIERvY3VtZW50PC9oND4KICAgICAgPGRpdiBzdHlsZT0iZm9udC1zaXplOjEwcHg7Y29sb3I6dmFyKC0tbXV0ZWQpO21hcmdpbi1ib3R0b206OHB4Ij5Pbmx5IGFkbWluaXN0cmF0b3JzIGNhbiBhZGQgY29udHJvbGxlZCBTT1AgYW5kIHJlZmVyZW5jZSBkb2N1bWVudHMgdG8gdGhlIGxvY2FsIGtub3dsZWRnZSBiYXNlLjwvZGl2PgogICAgICA8ZGl2IGNsYXNzPSJmaWVsZCI+PGxhYmVsPlNlY3RvciAvIHNjb3BlPC9sYWJlbD48aW5wdXQgaWQ9InNvcFNlY3RvciIgcGxhY2Vob2xkZXI9ImUuZy4gRW5naW5lZXJpbmcsIEZpbmFuY2UsIElUIj48L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0iZmllbGQiPjxsYWJlbD5TT1AgZmlsZTwvbGFiZWw+PGlucHV0IGlkPSJzb3BGaWxlSW5wdXQiIHR5cGU9ImZpbGUiIGFjY2VwdD0iLnBkZiwuZG9jLC5kb2N4LC50eHQsLm1kIj48L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0iYWRtaW4tYWN0aW9ucyI+PGJ1dHRvbiBjbGFzcz0icHJpbWFyeSIgb25jbGljaz0idXBsb2FkU09QKCkiPlVwbG9hZCBTT1A8L2J1dHRvbj48L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0iZXJyb3IiIGlkPSJzb3BFcnJvciI+PC9kaXY+CiAgICA8L2Rpdj4KICA8L2Rpdj4KPC9kaXY+Cgo8ZGl2IGNsYXNzPSJtb2RhbCBoaWRkZW4iIGlkPSJ0ZWFtTW9kYWwiPgogIDxkaXYgY2xhc3M9Im1vZGFsLWNhcmQiPgogICAgPGRpdiBjbGFzcz0ibW9kYWwtaGVhZCI+PGgzPlRlYW0gQ2hhdDwvaDM+PGJ1dHRvbiBjbGFzcz0iY2xvc2UiIG9uY2xpY2s9ImNsb3NlVGVhbU1hbmFnZXIoKSI+w5c8L2J1dHRvbj48L2Rpdj4KCiAgICA8ZGl2IGNsYXNzPSJ0ZWFtLXNlY3Rpb24iIHN0eWxlPSJib3JkZXItdG9wOjA7bWFyZ2luLXRvcDowO3BhZGRpbmctdG9wOjhweCI+CiAgICAgIDxoND5DcmVhdGUgYSB0ZWFtIGNoYXQ8L2g0PgogICAgICA8ZGl2IGNsYXNzPSJmaWVsZCI+PGxhYmVsPlRlYW0gbmFtZTwvbGFiZWw+PGlucHV0IGlkPSJ0ZWFtTmFtZSIgcGxhY2Vob2xkZXI9ImUuZy4gRW5naW5lZXJpbmcgUmV2aWV3IFRlYW0iPjwvZGl2PgogICAgICA8ZGl2IHN0eWxlPSJmb250LXNpemU6MTFweDtjb2xvcjp2YXIoLS1tdXRlZCk7bWFyZ2luOjZweCAwIj5TZWxlY3QgYXBwcm92ZWQgZW1wbG95ZWVzLiBZb3Ugd2lsbCBhdXRvbWF0aWNhbGx5IGJlY29tZSB0aGUgdGVhbSBvd25lci48L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0ibWVtYmVyLWxpc3QiIGlkPSJhcHByb3ZlZEVtcGxveWVlTGlzdCI+PC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9InRlYW0tYWN0aW9ucyI+PGJ1dHRvbiBjbGFzcz0icHJpbWFyeSIgb25jbGljaz0iY3JlYXRlVGVhbUZyb21VSSgpIj7vvIsgQ3JlYXRlIFRlYW0gQ2hhdDwvYnV0dG9uPjwvZGl2PgogICAgPC9kaXY+CgogICAgPGRpdiBjbGFzcz0idGVhbS1zZWN0aW9uIj4KICAgICAgPGg0PkpvaW4gYW4gZXhpc3RpbmcgdGVhbTwvaDQ+CiAgICAgIDxkaXYgY2xhc3M9ImZpZWxkIj48bGFiZWw+SW52aXRlIGNvZGU8L2xhYmVsPjxpbnB1dCBpZD0iam9pblRlYW1Db2RlIiBwbGFjZWhvbGRlcj0iRW50ZXIgdGVhbSBpbnZpdGUgY29kZSI+PC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9InRlYW0tYWN0aW9ucyI+PGJ1dHRvbiBjbGFzcz0ibmV3LWNoYXQiIG9uY2xpY2s9ImpvaW5UZWFtRnJvbVVJKCkiPkpvaW4gVGVhbTwvYnV0dG9uPjwvZGl2PgogICAgPC9kaXY+CgogICAgPGRpdiBjbGFzcz0idGVhbS1zZWN0aW9uIj4KICAgICAgPGg0PkN1cnJlbnQgdGVhbTwvaDQ+CiAgICAgIDxkaXYgaWQ9ImN1cnJlbnRUZWFtSW5mbyIgY2xhc3M9ImVtcHR5LWxpc3QiPlNlbGVjdCBhIHRlYW0gZnJvbSB0aGUgbGVmdCBzaWRlYmFyLjwvZGl2PgogICAgICA8ZGl2IGlkPSJjdXJyZW50VGVhbU1lbWJlcnMiIGNsYXNzPSJtZW1iZXItbGlzdCIgc3R5bGU9Im1hcmdpbi10b3A6OHB4Ij48L2Rpdj4KICAgICAgPGRpdiBzdHlsZT0iZm9udC1zaXplOjExcHg7Y29sb3I6dmFyKC0tbXV0ZWQpO21hcmdpbjo4cHggMCA0cHgiPkFkZCBhcHByb3ZlZCBlbXBsb3llZXMgdG8gdGhlIHNlbGVjdGVkIHRlYW08L2Rpdj4KICAgICAgPGRpdiBpZD0iYWRkVGVhbU1lbWJlcnNMaXN0IiBjbGFzcz0ibWVtYmVyLWxpc3QiPjwvZGl2PgogICAgICA8ZGl2IGNsYXNzPSJ0ZWFtLWFjdGlvbnMiPjxidXR0b24gY2xhc3M9Im5ldy1jaGF0IiBvbmNsaWNrPSJhZGRTZWxlY3RlZFRlYW1NZW1iZXJzKCkiPkFkZCBTZWxlY3RlZCBNZW1iZXJzPC9idXR0b24+PGJ1dHRvbiBjbGFzcz0ibmV3LWNoYXQiIG9uY2xpY2s9Im9wZW5TZWxlY3RlZFRlYW1DaGF0KCkiPk9wZW4gVGVhbSBDaGF0PC9idXR0b24+PC9kaXY+CiAgICA8L2Rpdj4KICAgIDxkaXYgY2xhc3M9ImFkbWluLWFjdGlvbnMiPjxidXR0b24gY2xhc3M9Im5ldy1jaGF0IiBvbmNsaWNrPSJjbG9zZVRlYW1NYW5hZ2VyKCkiPkNsb3NlPC9idXR0b24+PC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJlcnJvciIgaWQ9InRlYW1FcnJvciI+PC9kaXY+CiAgPC9kaXY+CjwvZGl2PgoKPHNjcmlwdD4KY29uc3QgJCA9IGlkID0+IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKGlkKTsKbGV0IGN1cnJlbnRVc2VyID0gbnVsbDsKbGV0IGN1cnJlbnRDaGF0SWQgPSBudWxsOwpsZXQgY3VycmVudFRlYW1JZCA9IG51bGw7CmxldCBzZWxlY3RlZEZpbGVzID0gW107CmxldCBzcGxhc2hTZWVuID0gZmFsc2U7CgooZnVuY3Rpb24gYnVpbGRTcGxhc2hTdGFycygpe2NvbnN0IGJveD0kKCdzcGxhc2hTdGFycycpO2lmKCFib3gpcmV0dXJuO2xldCBoPScnO2ZvcihsZXQgaT0wO2k8NDA7aSsrKXtoKz1gPHNwYW4gc3R5bGU9ImxlZnQ6JHtNYXRoLnJhbmRvbSgpKjEwMH0lO3RvcDoke01hdGgucmFuZG9tKCkqMTAwfSU7YW5pbWF0aW9uLWRlbGF5OiR7KE1hdGgucmFuZG9tKCkqMy40KS50b0ZpeGVkKDIpfXMiPjwvc3Bhbj5gO31ib3guaW5uZXJIVE1MPWg7fSkoKTsKCmFzeW5jIGZ1bmN0aW9uIGFwaSh1cmwsIG9wdGlvbnM9e30pIHsKICBjb25zdCByZXMgPSBhd2FpdCBmZXRjaCh1cmwsIG9wdGlvbnMpOwogIGxldCBkYXRhID0gbnVsbDsKICB0cnkgeyBkYXRhID0gYXdhaXQgcmVzLmpzb24oKTsgfSBjYXRjaCAoXykge30KICBpZiAocmVzLnN0YXR1cyA9PT0gNDAxKSB7IHNob3dMb2dpbigpOyB0aHJvdyBuZXcgRXJyb3IoJ0F1dGhlbnRpY2F0aW9uIHJlcXVpcmVkLicpOyB9CiAgaWYgKCFyZXMub2sgfHwgKGRhdGEgJiYgZGF0YS5zdGF0dXMgPT09ICdlcnJvcicpKSB0aHJvdyBuZXcgRXJyb3IoKGRhdGEgJiYgZGF0YS5tZXNzYWdlKSB8fCBgUmVxdWVzdCBmYWlsZWQgKCR7cmVzLnN0YXR1c30pYCk7CiAgcmV0dXJuIGRhdGE7Cn0KCmZ1bmN0aW9uIGVzYyh2YWx1ZSkgewogIHJldHVybiBTdHJpbmcodmFsdWUgPz8gJycpLnJlcGxhY2UoL1smPD4nIl0vZywgYyA9PiAoeycmJzonJmFtcDsnLCc8JzonJmx0OycsJz4nOicmZ3Q7JywiJyI6JyYjMzk7JywnIic6JyZxdW90Oyd9W2NdKSk7Cn0KZnVuY3Rpb24gc2Nyb2xsVG9Cb3R0b20oKXsgY29uc3Qgdz0kKCd3b3Jrc3BhY2UnKTsgdy5zY3JvbGxUb3A9dy5zY3JvbGxIZWlnaHQ7IH0KYXN5bmMgZnVuY3Rpb24gZG93bmxvYWRBcnRpZmFjdCh1cmwsIGZpbGVuYW1lPSdvcmlvbi1vdXRwdXQnKXsKICBpZighdXJsKXJldHVybjsKICB0cnl7CiAgICBjb25zdCByZXM9YXdhaXQgZmV0Y2godXJsKTsKICAgIGlmKCFyZXMub2spdGhyb3cgbmV3IEVycm9yKGBEb3dubG9hZCBmYWlsZWQgKCR7cmVzLnN0YXR1c30pYCk7CiAgICBjb25zdCBibG9iPWF3YWl0IHJlcy5ibG9iKCk7CiAgICBjb25zdCBvYmplY3RVcmw9VVJMLmNyZWF0ZU9iamVjdFVSTChibG9iKTsKICAgIGNvbnN0IGE9ZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYScpO2EuaHJlZj1vYmplY3RVcmw7YS5kb3dubG9hZD1maWxlbmFtZTthLnN0eWxlLmRpc3BsYXk9J25vbmUnOwogICAgZG9jdW1lbnQuYm9keS5hcHBlbmRDaGlsZChhKTthLmNsaWNrKCk7YS5yZW1vdmUoKTtzZXRUaW1lb3V0KCgpPT5VUkwucmV2b2tlT2JqZWN0VVJMKG9iamVjdFVybCksMTUwMCk7CiAgfWNhdGNoKGVycil7YWxlcnQoZXJyLm1lc3NhZ2V8fCdVbmFibGUgdG8gZG93bmxvYWQgdGhlIGZpbGUuJyk7fQp9CmZ1bmN0aW9uIGRvd25sb2FkVGV4dEZpbGUoZmlsZW5hbWUsdGV4dCl7CiAgY29uc3QgYmxvYj1uZXcgQmxvYihbU3RyaW5nKHRleHR8fCcnKV0se3R5cGU6J3RleHQvcGxhaW47Y2hhcnNldD11dGYtOCd9KTsKICBjb25zdCB1cmw9VVJMLmNyZWF0ZU9iamVjdFVSTChibG9iKTtjb25zdCBhPWRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2EnKTthLmhyZWY9dXJsO2EuZG93bmxvYWQ9ZmlsZW5hbWV8fCdvcmlvbi1yZXNwb25zZS50eHQnOwogIGRvY3VtZW50LmJvZHkuYXBwZW5kQ2hpbGQoYSk7YS5jbGljaygpO2EucmVtb3ZlKCk7c2V0VGltZW91dCgoKT0+VVJMLnJldm9rZU9iamVjdFVSTCh1cmwpLDEwMDApOwp9CmFzeW5jIGZ1bmN0aW9uIGRvd25sb2FkQW5zd2VyRG9jeChjb250ZW50LGJ1dHRvbil7CiAgaWYoYnV0dG9uKWJ1dHRvbi5jbGFzc0xpc3QuYWRkKCdkb3dubG9hZC1idXN5Jyk7CiAgdHJ5ewogICAgY29uc3QgZGF0YT1hd2FpdCBhcGkoJy9hcGkvZ2VuZXJhdGUtYW5zd2VyLWRvY3gnLHttZXRob2Q6J1BPU1QnLGhlYWRlcnM6eydDb250ZW50LVR5cGUnOidhcHBsaWNhdGlvbi9qc29uJ30sYm9keTpKU09OLnN0cmluZ2lmeSh7Y29udGVudDpTdHJpbmcoY29udGVudHx8JycpfSl9KTsKICAgIGF3YWl0IGRvd25sb2FkQXJ0aWZhY3QoZGF0YS5kb3dubG9hZF91cmx8fGRhdGEuZmlsZV91cmwsZGF0YS5maWxlbmFtZXx8J29yaW9uLWFuc3dlci5kb2N4Jyk7CiAgfWNhdGNoKGVycil7YWxlcnQoZXJyLm1lc3NhZ2V8fCdVbmFibGUgdG8gY3JlYXRlIERPQ1guJyk7fQogIGZpbmFsbHl7aWYoYnV0dG9uKWJ1dHRvbi5jbGFzc0xpc3QucmVtb3ZlKCdkb3dubG9hZC1idXN5Jyk7fQp9CmZ1bmN0aW9uIHJlc2V0V29ya3NwYWNlU2Nyb2xsKCl7ICQoJ3dvcmtzcGFjZScpLnNjcm9sbFRvcD0wOyB9CmZ1bmN0aW9uIGxlYXZlU3BsYXNoKGV2ZW50KXsgaWYoZXZlbnQpIGV2ZW50LnN0b3BQcm9wYWdhdGlvbigpOyAkKCdzcGxhc2gnKS5jbGFzc0xpc3QuYWRkKCdoaWRlJyk7IHNwbGFzaFNlZW49dHJ1ZTsgc2V0VGltZW91dCgoKT0+eyAkKCdzcGxhc2gnKS5jbGFzc0xpc3QuYWRkKCdoaWRkZW4nKTsgJCgnbG9naW4nKS5jbGFzc0xpc3QucmVtb3ZlKCdoaWRkZW4nKTsgJCgnbG9naW5FbWFpbCcpLmZvY3VzKCk7IH0sODAwKTsgfQpmdW5jdGlvbiBlbnRlcldvcmtiZW5jaCgpeyBsZWF2ZVNwbGFzaCgpOyB9CnNldFRpbWVvdXQoKCk9PntpZighJCgnc3BsYXNoJykuY2xhc3NMaXN0LmNvbnRhaW5zKCdoaWRlJykpIGxlYXZlU3BsYXNoKCk7fSw2NTAwKTsKCmZ1bmN0aW9uIHVwZGF0ZVRoZW1lQnV0dG9ucygpewogIGNvbnN0IGljb249ZG9jdW1lbnQuYm9keS5jbGFzc0xpc3QuY29udGFpbnMoJ2xpZ2h0Jyk/J+KYgO+4jyc6J/CfjJknOwogIFsndGhlbWVCdG4nLCdsb2dpblRoZW1lQnRuJywnc3BsYXNoVGhlbWVCdG4nXS5mb3JFYWNoKGlkPT57aWYoJChpZCkpJChpZCkudGV4dENvbnRlbnQ9aWNvbjt9KTsKfQpmdW5jdGlvbiB0b2dnbGVUaGVtZSgpewogIGNvbnN0IGlzTGlnaHQ9IWRvY3VtZW50LmJvZHkuY2xhc3NMaXN0LmNvbnRhaW5zKCdsaWdodCcpOwogIGRvY3VtZW50LmJvZHkuY2xhc3NMaXN0LnRvZ2dsZSgnbGlnaHQnLGlzTGlnaHQpOwogIGxvY2FsU3RvcmFnZS5zZXRJdGVtKCdzb3ZlcmVpZ25fdGhlbWUnLGlzTGlnaHQ/J2xpZ2h0JzonZGFyaycpOwogIHVwZGF0ZVRoZW1lQnV0dG9ucygpOwp9CmZ1bmN0aW9uIGFwcGx5VGhlbWUoKXsKICBjb25zdCBzYXZlZD1sb2NhbFN0b3JhZ2UuZ2V0SXRlbSgnc292ZXJlaWduX3RoZW1lJyl8fCdsaWdodCc7CiAgZG9jdW1lbnQuYm9keS5jbGFzc0xpc3QudG9nZ2xlKCdsaWdodCcsc2F2ZWQhPT0nZGFyaycpOwogIHVwZGF0ZVRoZW1lQnV0dG9ucygpOwp9CgpmdW5jdGlvbiBzaG93QXV0aE1vZGUobW9kZSl7CiAgY29uc3QgbG9naW5Nb2RlPW1vZGU9PT0nbG9naW4nOwogICQoJ2xvZ2luRm9ybScpLmNsYXNzTGlzdC50b2dnbGUoJ2hpZGRlbicsIWxvZ2luTW9kZSk7CiAgJCgncmVnaXN0ZXJGb3JtJykuY2xhc3NMaXN0LnRvZ2dsZSgnaGlkZGVuJyxsb2dpbk1vZGUpOwogICQoJ2xvZ2luVGFiJykuY2xhc3NMaXN0LnRvZ2dsZSgnYWN0aXZlJyxsb2dpbk1vZGUpOwogICQoJ2VtcGxveWVlVGFiJykuY2xhc3NMaXN0LnRvZ2dsZSgnYWN0aXZlJyxtb2RlPT09J2VtcGxveWVlJyk7CiAgJCgnYWRtaW5UYWInKS5jbGFzc0xpc3QudG9nZ2xlKCdhY3RpdmUnLG1vZGU9PT0nYWRtaW4nKTsKICAkKCdyZWdTZWN0b3JXcmFwJykuY2xhc3NMaXN0LnRvZ2dsZSgnaGlkZGVuJyxtb2RlPT09J2FkbWluJyk7CiAgJCgncmVnaXN0ZXJTdWJtaXQnKS50ZXh0Q29udGVudD1tb2RlPT09J2FkbWluJz8nUmVnaXN0ZXIgYWRtaW5pc3RyYXRvcic6J1JlZ2lzdGVyIGVtcGxveWVlJzsKICAkKCdyZWdpc3RlckZvcm0nKS5kYXRhc2V0Lm1vZGU9bW9kZTsKICAkKCdsb2dpbkVycm9yJykudGV4dENvbnRlbnQ9Jyc7JCgncmVnaXN0ZXJFcnJvcicpLnRleHRDb250ZW50PScnOwp9Cgphc3luYyBmdW5jdGlvbiBsb2dpbihldmVudCl7CiAgZXZlbnQucHJldmVudERlZmF1bHQoKTskKCdsb2dpbkVycm9yJykudGV4dENvbnRlbnQ9Jyc7CiAgdHJ5ewogICAgY29uc3QgZGF0YT1hd2FpdCBhcGkoJy9hcGkvbG9naW4nLHttZXRob2Q6J1BPU1QnLGhlYWRlcnM6eydDb250ZW50LVR5cGUnOidhcHBsaWNhdGlvbi9qc29uJ30sYm9keTpKU09OLnN0cmluZ2lmeSh7ZW1haWw6JCgnbG9naW5FbWFpbCcpLnZhbHVlLnRyaW0oKSxwYXNzd29yZDokKCdsb2dpblBhc3N3b3JkJykudmFsdWV9KX0pOwogICAgY3VycmVudFVzZXI9ZGF0YS51c2VyOyQoJ2xvZ2luJykuY2xhc3NMaXN0LmFkZCgnaGlkZGVuJyk7JCgnc3BsYXNoJykuY2xhc3NMaXN0LmFkZCgnaGlkZGVuJyk7JCgnYXBwJykuY2xhc3NMaXN0LnJlbW92ZSgnaGlkZGVuJyk7YXdhaXQgaW5pdGlhbGl6ZVdvcmtiZW5jaCgpOwogIH1jYXRjaChlcnIpeyQoJ2xvZ2luRXJyb3InKS50ZXh0Q29udGVudD1lcnIubWVzc2FnZTt9Cn0KCmFzeW5jIGZ1bmN0aW9uIHJlZ2lzdGVyQWNjb3VudChldmVudCl7CiAgZXZlbnQucHJldmVudERlZmF1bHQoKTskKCdyZWdpc3RlckVycm9yJykudGV4dENvbnRlbnQ9Jyc7CiAgY29uc3QgbW9kZT0kKCdyZWdpc3RlckZvcm0nKS5kYXRhc2V0Lm1vZGV8fCdlbXBsb3llZSc7CiAgY29uc3QgcGF5bG9hZD17bmFtZTokKCdyZWdOYW1lJykudmFsdWUudHJpbSgpLGVtYWlsOiQoJ3JlZ0VtYWlsJykudmFsdWUudHJpbSgpLHBhc3N3b3JkOiQoJ3JlZ1Bhc3N3b3JkJykudmFsdWUsY29tcGFueTokKCdyZWdDb21wYW55JykudmFsdWUudHJpbSgpfTsKICBpZihtb2RlPT09J2VtcGxveWVlJylwYXlsb2FkLnNlY3Rvcj0kKCdyZWdTZWN0b3InKS52YWx1ZS50cmltKCk7CiAgdHJ5ewogICAgY29uc3QgZGF0YT1hd2FpdCBhcGkobW9kZT09PSdhZG1pbic/Jy9hcGkvcmVnaXN0ZXIvYWRtaW4nOicvYXBpL3JlZ2lzdGVyL2VtcGxveWVlJyx7bWV0aG9kOidQT1NUJyxoZWFkZXJzOnsnQ29udGVudC1UeXBlJzonYXBwbGljYXRpb24vanNvbid9LGJvZHk6SlNPTi5zdHJpbmdpZnkocGF5bG9hZCl9KTsKICAgICQoJ3JlZ2lzdGVyRm9ybScpLnJlc2V0KCk7CiAgICBpZihtb2RlPT09J2VtcGxveWVlJyl7YWxlcnQoJ1JlZ2lzdHJhdGlvbiBzdWJtaXR0ZWQuIEFuIGFkbWluaXN0cmF0b3IgbXVzdCBhcHByb3ZlIHlvdXIgYWNjb3VudCBiZWZvcmUgeW91IGNhbiBzaWduIGluLicpO3Nob3dBdXRoTW9kZSgnbG9naW4nKTt9CiAgICBlbHNle2FsZXJ0KCdBZG1pbmlzdHJhdG9yIHJlZ2lzdGVyZWQuIFlvdSBjYW4gc2lnbiBpbiBub3cuJyk7c2hvd0F1dGhNb2RlKCdsb2dpbicpO30KICB9Y2F0Y2goZXJyKXskKCdyZWdpc3RlckVycm9yJykudGV4dENvbnRlbnQ9ZXJyLm1lc3NhZ2U7fQp9CgpmdW5jdGlvbiBzaG93TG9naW4oKXsgY3VycmVudFVzZXI9bnVsbDtjdXJyZW50Q2hhdElkPW51bGw7Y3VycmVudFRlYW1JZD1udWxsOyQoJ2FwcCcpLmNsYXNzTGlzdC5hZGQoJ2hpZGRlbicpOyQoJ3NwbGFzaCcpLmNsYXNzTGlzdC5hZGQoJ2hpZGRlbicpOyQoJ2xvZ2luJykuY2xhc3NMaXN0LnJlbW92ZSgnaGlkZGVuJyk7IH0KYXN5bmMgZnVuY3Rpb24gbG9nb3V0KCl7IHRyeXthd2FpdCBmZXRjaCgnL2FwaS9sb2dvdXQnLHttZXRob2Q6J1BPU1QnfSk7fWNhdGNoKF8pe30gc2hvd0xvZ2luKCk7IH0KCmFzeW5jIGZ1bmN0aW9uIGluaXRpYWxpemVXb3JrYmVuY2goKXsKICAkKCdwcm9maWxlTmFtZScpLnRleHRDb250ZW50PWN1cnJlbnRVc2VyLm5hbWV8fGN1cnJlbnRVc2VyLmVtYWlsfHwnVXNlcic7CiAgJCgncHJvZmlsZU1ldGEnKS5pbm5lckhUTUw9YCR7ZXNjKGN1cnJlbnRVc2VyLmVtYWlsKX08YnI+JHtlc2MoY3VycmVudFVzZXIuY29tcGFueSl9JHtjdXJyZW50VXNlci5zZWN0b3I/JyDCtyAnK2VzYyhjdXJyZW50VXNlci5zZWN0b3IpOicnfTxicj4ke2VzYyhjdXJyZW50VXNlci5yb2xlKX0ke2N1cnJlbnRVc2VyLmlzX2FkbWluPycgwrcgQWRtaW5pc3RyYXRvcic6Jyd9YDsKICAkKCdoZWFkZXJQcm9maWxlJykudGV4dENvbnRlbnQ9YCR7Y3VycmVudFVzZXIuY29tcGFueX0ke2N1cnJlbnRVc2VyLnNlY3Rvcj8nIMK3ICcrY3VycmVudFVzZXIuc2VjdG9yOicnfWA7CiAgJCgnY29tcG9zZXInKS5jbGFzc0xpc3QucmVtb3ZlKCdkaXNhYmxlZCcpOwogICQoJ2FkbWluQnRuJykuY2xhc3NMaXN0LnRvZ2dsZSgnaGlkZGVuJywhY3VycmVudFVzZXIuaXNfYWRtaW4pOwogICQoJ3RlYW1BZG1pbkJ0bicpLmNsYXNzTGlzdC5yZW1vdmUoJ2hpZGRlbicpOwogIGlmKGN1cnJlbnRVc2VyLmlzX2FkbWluKSAkKCdhZG1pbkJ0bicpLnRleHRDb250ZW50PSdBZG1pbjogUmVnaXN0cmF0aW9uICYgRW1wbG95ZWVzJzsKICBhd2FpdCByZWZyZXNoQ2hhdHMoKTsgYXdhaXQgcmVmcmVzaFRlYW1zKCk7IGF3YWl0IHJlZnJlc2hIZWFsdGgoKTsKfQoKYXN5bmMgZnVuY3Rpb24gcmVmcmVzaEhlYWx0aCgpewogIHRyeXsKICAgIGNvbnN0IGRhdGE9YXdhaXQgYXBpKCcvYXBpL2hlYWx0aCcpOwogICAgJCgnb2xsYW1hU3RhdHVzJykudGV4dENvbnRlbnQ9ZGF0YS5vbGxhbWFfcmVhY2hhYmxlPydvbmxpbmUnOidvZmZsaW5lJzsKICAgICQoJ29sbGFtYVN0YXR1cycpLmNsYXNzTmFtZT1kYXRhLm9sbGFtYV9yZWFjaGFibGU/J3N0YXR1cy1vayc6Jyc7CiAgICAkKCdtb2RlbFN0YXR1cycpLnRleHRDb250ZW50PWRhdGEucmVzb2x2ZWRfbW9kZWx8fCdkeW5hbWljJzsKICB9Y2F0Y2goZXJyKXskKCdvbGxhbWFTdGF0dXMnKS50ZXh0Q29udGVudD0ndW5hdmFpbGFibGUnOyQoJ29sbGFtYVN0YXR1cycpLmNsYXNzTmFtZT0nZGFuZ2VyLW5vdGUnO30KfQoKYXN5bmMgZnVuY3Rpb24gcmVmcmVzaENoYXRzKCl7CiAgY29uc3QgZGF0YT1hd2FpdCBhcGkoJy9hcGkvY2hhdHMnKTsgY29uc3QgbGlzdD0kKCdjaGF0TGlzdCcpOyBsaXN0LmlubmVySFRNTD0nJzsKICBjb25zdCBjaGF0cz0oZGF0YS5jaGF0c3x8W10pLmZpbHRlcihjPT4hYy50ZWFtX2lkKTsKICBpZighY2hhdHMubGVuZ3RoKXtsaXN0LmlubmVySFRNTD0nPGRpdiBjbGFzcz0iZW1wdHktbGlzdCI+Tm8gY2hhdHMgeWV0LjwvZGl2Pic7cmV0dXJuO30KICBjaGF0cy5mb3JFYWNoKGM9PnsKICAgIGNvbnN0IGJ0bj1kb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTtidG4uY2xhc3NOYW1lPSdjaGF0LWl0ZW0nKyhOdW1iZXIoYy5pZCk9PT1OdW1iZXIoY3VycmVudENoYXRJZCk/JyBhY3RpdmUnOicnKTsKICAgIGJ0bi5pbm5lckhUTUw9YDxzcGFuIGNsYXNzPSJkb3QiPjwvc3Bhbj48c3BhbiBjbGFzcz0idGl0bGUiPiR7ZXNjKGMudGl0bGV8fCdOZXcgQ2hhdCcpfTwvc3Bhbj5gOwogICAgYnRuLm9uY2xpY2s9KCk9Pm9wZW5DaGF0KGMuaWQpO2xpc3QuYXBwZW5kQ2hpbGQoYnRuKTsKICB9KTsKfQoKYXN5bmMgZnVuY3Rpb24gcmVmcmVzaFRlYW1zKCl7CiAgdHJ5ewogICAgY29uc3QgZGF0YT1hd2FpdCBhcGkoJy9hcGkvdGVhbXMnKTsgY29uc3QgbGlzdD0kKCd0ZWFtTGlzdCcpO2xpc3QuaW5uZXJIVE1MPScnOwogICAgaWYoIShkYXRhLnRlYW1zfHxbXSkubGVuZ3RoKXtsaXN0LmlubmVySFRNTD0nPGRpdiBjbGFzcz0iZW1wdHktbGlzdCI+Tm8gdGVhbSBjaGF0cy48L2Rpdj4nO3JldHVybjt9CiAgICBkYXRhLnRlYW1zLmZvckVhY2godD0+ewogICAgICBjb25zdCBidG49ZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7YnRuLmNsYXNzTmFtZT0nY2hhdC1pdGVtJysoTnVtYmVyKGN1cnJlbnRUZWFtSWQpPT09TnVtYmVyKHQuaWQpPycgYWN0aXZlJzonJyk7CiAgICAgIGJ0bi5pbm5lckhUTUw9YDxzcGFuIGNsYXNzPSJkb3QiPjwvc3Bhbj48c3BhbiBjbGFzcz0idGl0bGUiPiR7ZXNjKHQubmFtZSl9PC9zcGFuPjxzcGFuIGNsYXNzPSJ0ZWFtLXRhZyI+VEVBTTwvc3Bhbj5gOwogICAgICBidG4ub25jbGljaz0oKT0+b3BlblRlYW0odCk7bGlzdC5hcHBlbmRDaGlsZChidG4pOwogICAgfSk7CiAgfWNhdGNoKGVycil7JCgndGVhbUxpc3QnKS5pbm5lckhUTUw9JzxkaXYgY2xhc3M9ImVtcHR5LWxpc3QiPlVuYWJsZSB0byBsb2FkIHRlYW1zLjwvZGl2Pic7fQp9Cgphc3luYyBmdW5jdGlvbiBuZXdDaGF0KCl7CiAgdHJ5ewogICAgY29uc3QgZGF0YT1hd2FpdCBhcGkoJy9hcGkvY2hhdHMvbmV3Jyx7bWV0aG9kOidQT1NUJyxoZWFkZXJzOnsnQ29udGVudC1UeXBlJzonYXBwbGljYXRpb24vanNvbid9LGJvZHk6SlNPTi5zdHJpbmdpZnkoe3RpdGxlOidOZXcgQ2hhdCd9KX0pOwogICAgY3VycmVudENoYXRJZD1kYXRhLmNoYXRfaWQ7Y3VycmVudFRlYW1JZD1udWxsOyQoJ2NoYXRTdGF0dXMnKS50ZXh0Q29udGVudD1jdXJyZW50Q2hhdElkOyQoJ2hlYWRlclRpdGxlJykudGV4dENvbnRlbnQ9J09yaW9uIOKAlCBTb3ZlcmVpZ24gV29ya2JlbmNoJzsKICAgICQoJ2NoYXQnKS5pbm5lckhUTUw9Jyc7JCgnaGVybycpLmNsYXNzTGlzdC5yZW1vdmUoJ2hpZGRlbicpOwogICAgY2xlYXJGaWxlcygpOyQoJ3F1ZXJ5JykudmFsdWU9Jyc7cmVzZXRXb3Jrc3BhY2VTY3JvbGwoKTsKICAgIGF3YWl0IHJlZnJlc2hDaGF0cygpO2F3YWl0IHJlZnJlc2hUZWFtcygpO3Jlc2V0V29ya3NwYWNlU2Nyb2xsKCk7JCgncXVlcnknKS5mb2N1cygpOwogIH1jYXRjaChlcnIpe3JlbmRlck1lc3NhZ2UoJ2Fzc2lzdGFudCcsJ1VuYWJsZSB0byBjcmVhdGUgYSBuZXcgY2hhdDogJytlcnIubWVzc2FnZSk7fQp9Cgphc3luYyBmdW5jdGlvbiBvcGVuQ2hhdChpZCl7CiAgdHJ5ewogICAgY29uc3QgZGF0YT1hd2FpdCBhcGkoYC9hcGkvY2hhdHMvJHtpZH1gKTtjdXJyZW50Q2hhdElkPWlkO2N1cnJlbnRUZWFtSWQ9KGRhdGEuY2hhdCYmZGF0YS5jaGF0LnRlYW1faWQpfHxudWxsOwogICAgJCgnY2hhdFN0YXR1cycpLnRleHRDb250ZW50PWlkOyQoJ2hlYWRlclRpdGxlJykudGV4dENvbnRlbnQ9KGRhdGEuY2hhdCYmZGF0YS5jaGF0LnRpdGxlKXx8J09yaW9uIOKAlCBTb3ZlcmVpZ24gV29ya2JlbmNoJzsKICAgICQoJ2NoYXQnKS5pbm5lckhUTUw9Jyc7JCgnaGVybycpLmNsYXNzTGlzdC5hZGQoJ2hpZGRlbicpOwogICAgKGRhdGEubWVzc2FnZXN8fFtdKS5mb3JFYWNoKG09PnJlbmRlck1lc3NhZ2UobS5zZW5kZXJfdHlwZT09PSd1c2VyJz8ndXNlcic6J2Fzc2lzdGFudCcsbS5jb250ZW50LG0uZmlsZXN8fFtdLGZhbHNlLG0pKTsKICAgIGF3YWl0IGFjdGl2YXRlQ2hhdFN0YXRlKGlkKTthd2FpdCByZWZyZXNoQ2hhdHMoKTthd2FpdCByZWZyZXNoVGVhbXMoKTtzY3JvbGxUb0JvdHRvbSgpOwogIH1jYXRjaChlcnIpe3JlbmRlck1lc3NhZ2UoJ2Fzc2lzdGFudCcsJ1VuYWJsZSB0byBvcGVuIGNoYXQ6ICcrZXJyLm1lc3NhZ2UpO30KfQoKYXN5bmMgZnVuY3Rpb24gYWN0aXZhdGVDaGF0U3RhdGUoaWQpewogIHRyeXthd2FpdCBhcGkoYC9hcGkvY2hhdHMvJHtpZH0vYWN0aXZhdGVgLHttZXRob2Q6J1BPU1QnfSk7fWNhdGNoKF8pe30KfQphc3luYyBmdW5jdGlvbiBvcGVuVGVhbUNoYXROYXYoKXsKICB0cnl7Y29uc3QgZGF0YT1hd2FpdCBhcGkoJy9hcGkvdGVhbXMnKTtjb25zdCB0ZWFtcz1kYXRhLnRlYW1zfHxbXTtpZighdGVhbXMubGVuZ3RoKXthbGVydCgnWW91IGFyZSBub3QgYSBtZW1iZXIgb2YgYW55IHRlYW0geWV0LicpO3JldHVybjt9aWYodGVhbXMubGVuZ3RoPT09MSl7YXdhaXQgb3BlblRlYW0odGVhbXNbMF0pO3JldHVybjt9Y29uc3QgbmFtZXM9dGVhbXMubWFwKCh0LGkpPT5gJHtpKzF9LiAke3QubmFtZX1gKS5qb2luKCdcbicpO2NvbnN0IGNob2ljZT1wcm9tcHQoJ0Nob29zZSBhIHRlYW0gYnkgbnVtYmVyOlxuJytuYW1lcyk7Y29uc3QgaWR4PU51bWJlcihjaG9pY2UpLTE7aWYoTnVtYmVyLmlzSW50ZWdlcihpZHgpJiZ0ZWFtc1tpZHhdKWF3YWl0IG9wZW5UZWFtKHRlYW1zW2lkeF0pO30KICBjYXRjaChlcnIpe3JlbmRlck1lc3NhZ2UoJ2Fzc2lzdGFudCcsJ1VuYWJsZSB0byBsb2FkIHRlYW0gY2hhdHM6ICcrZXJyLm1lc3NhZ2UpO30KfQoKYXN5bmMgZnVuY3Rpb24gb3BlblRlYW0odGVhbSl7CiAgdHJ5ewogICAgY29uc3QgdGVhbUlkPU51bWJlcih0ZWFtLmlkKTsKICAgIGN1cnJlbnRUZWFtSWQ9dGVhbUlkOwogICAgY29uc3QgY2hhdHNEYXRhPWF3YWl0IGFwaSgnL2FwaS9jaGF0cycpOwogICAgY29uc3QgZXhpc3Rpbmc9KGNoYXRzRGF0YS5jaGF0c3x8W10pLmZpbmQoYz0+TnVtYmVyKGMudGVhbV9pZCk9PT10ZWFtSWQpOwogICAgbGV0IGNoYXRJZD1leGlzdGluZz9leGlzdGluZy5pZDpudWxsOwogICAgaWYoIWNoYXRJZCl7CiAgICAgIGNvbnN0IGRhdGE9YXdhaXQgYXBpKCcvYXBpL2NoYXRzL25ldycse21ldGhvZDonUE9TVCcsaGVhZGVyczp7J0NvbnRlbnQtVHlwZSc6J2FwcGxpY2F0aW9uL2pzb24nfSxib2R5OkpTT04uc3RyaW5naWZ5KHt0aXRsZTp0ZWFtLm5hbWV8fCdUZWFtIENoYXQnLHRlYW1faWQ6dGVhbUlkfSl9KTsKICAgICAgY2hhdElkPWRhdGEuY2hhdF9pZDsKICAgIH0KICAgIGN1cnJlbnRDaGF0SWQ9Y2hhdElkOwogICAgJCgnY2hhdFN0YXR1cycpLnRleHRDb250ZW50PWB0ZWFtOiR7dGVhbUlkfWA7CiAgICAkKCdoZWFkZXJUaXRsZScpLnRleHRDb250ZW50PXRlYW0ubmFtZXx8J1RlYW0gQ2hhdCc7CiAgICBhd2FpdCBvcGVuQ2hhdChjaGF0SWQpOwogIH1jYXRjaChlcnIpe3JlbmRlck1lc3NhZ2UoJ2Fzc2lzdGFudCcsJ1VuYWJsZSB0byBvcGVuIHRlYW0gY2hhdDogJytlcnIubWVzc2FnZSk7fQp9CgpmdW5jdGlvbiBmaWxlSWNvbihuYW1lKXsKICBjb25zdCBleHQ9KFN0cmluZyhuYW1lfHwnJykuc3BsaXQoJy4nKS5wb3AoKXx8JycpLnRvTG93ZXJDYXNlKCk7CiAgcmV0dXJuICh7cGRmOifwn5OVJyxkb2N4Oifwn5OYJyxwcHR4Oifwn5OZJyx4bHN4Oifwn5OXJyx4bHM6J/Cfk5cnLGNzdjon8J+TiicscG5nOifwn5a877iPJyxqcGc6J/CflrzvuI8nLGpwZWc6J/CflrzvuI8nLHdlYnA6J/CflrzvuI8nLHR4dDon8J+ThCcsbWQ6J/Cfk50nLGpzb246J/CflKcnLGxvZzon8J+TnCd9KVtleHRdfHwn8J+Tjic7Cn0KZnVuY3Rpb24gbm9ybWFsaXplRmlsZXMoZmlsZXMpewogIHJldHVybiAoZmlsZXN8fFtdKS5tYXAoZj0+dHlwZW9mIGY9PT0nc3RyaW5nJz97bmFtZTpmfTpmKS5maWx0ZXIoZj0+ZiYmZi5uYW1lKTsKfQpmdW5jdGlvbiByZW5kZXJGaWxlQ2FyZHMoZmlsZXMpewogIGNvbnN0IGFycj1ub3JtYWxpemVGaWxlcyhmaWxlcyk7IGlmKCFhcnIubGVuZ3RoKXJldHVybiAnJzsKICByZXR1cm4gYDxkaXYgY2xhc3M9ImZpbGUtY2FyZHMiPiR7YXJyLm1hcChmPT57CiAgICBjb25zdCBuYW1lPWVzYyhmLm5hbWUpOyBjb25zdCBocmVmPWYuZG93bmxvYWRfdXJsfHxmLnVybHx8Jyc7IGNvbnN0IHRhcmdldD1ocmVmP2BocmVmPSIke2VzYyhocmVmKX0iIHRhcmdldD0iX2JsYW5rImA6Jyc7CiAgICBjb25zdCBsb2NhbD1mLmxvY2FsX3VybD9gPGltZyBzcmM9IiR7ZXNjKGYubG9jYWxfdXJsKX0iIGFsdD0iIiBzdHlsZT0id2lkdGg6MzBweDtoZWlnaHQ6MzBweDtvYmplY3QtZml0OmNvdmVyO2JvcmRlci1yYWRpdXM6N3B4Ij5gOmA8ZGl2IGNsYXNzPSJmaWxlLWljb24iPiR7ZmlsZUljb24oZi5uYW1lKX08L2Rpdj5gOwogICAgY29uc3QgaW5uZXI9YCR7bG9jYWx9PGRpdiBzdHlsZT0ibWluLXdpZHRoOjAiPjxkaXYgY2xhc3M9ImZpbGUtY2FyZC1uYW1lIj4ke25hbWV9PC9kaXY+PGRpdiBjbGFzcz0iZmlsZS1jYXJkLW1ldGEiPiR7aHJlZj8nT3BlbiAvIGRvd25sb2FkJzonQXR0YWNoZWQgZmlsZSd9PC9kaXY+PC9kaXY+YDsKICAgIHJldHVybiBocmVmP2A8YSBjbGFzcz0iZmlsZS1jYXJkIiAke3RhcmdldH0+JHtpbm5lcn08L2E+YDpgPGRpdiBjbGFzcz0iZmlsZS1jYXJkIj4ke2lubmVyfTwvZGl2PmA7CiAgfSkuam9pbignJyl9PC9kaXY+YDsKfQpmdW5jdGlvbiByZW5kZXJNZXNzYWdlKHNlbmRlcixjb250ZW50LGZpbGVzPVtdLGF1dG9TY3JvbGw9dHJ1ZSxtZXRhPW51bGwpewogICQoJ2hlcm8nKS5jbGFzc0xpc3QuYWRkKCdoaWRkZW4nKTsKICBjb25zdCByb3c9ZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7cm93LmNsYXNzTmFtZT0nbWVzc2FnZSAnK3NlbmRlcjsKICBjb25zdCBidWJibGU9ZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7YnViYmxlLmNsYXNzTmFtZT0nYnViYmxlJzsKICBjb25zdCBzZW5kZXJOYW1lPW1ldGE/LnNlbmRlcl9uYW1lIHx8IChzZW5kZXI9PT0nYXNzaXN0YW50Jz8nT3Jpb24nOmN1cnJlbnRVc2VyPy5uYW1lfHwnWW91Jyk7CiAgY29uc3Qgc2VuZGVyRW1haWw9bWV0YT8uc2VuZGVyX2VtYWlsP2A8c3BhbiBjbGFzcz0ic2VuZGVyLWVtYWlsIj4ke2VzYyhtZXRhLnNlbmRlcl9lbWFpbCl9PC9zcGFuPmA6Jyc7CiAgYnViYmxlLmlubmVySFRNTD1gPGRpdiBjbGFzcz0ic2VuZGVyLWxhYmVsIj4ke2VzYyhzZW5kZXJOYW1lKX0ke3NlbmRlckVtYWlsfTwvZGl2PiR7cmVuZGVyRmlsZUNhcmRzKGZpbGVzKX08ZGl2IGNsYXNzPSIke3NlbmRlcj09PSd1c2VyJz8ndXNlci1jb250ZW50JzonYXNzaXN0YW50LWJvZHknfSI+JHtlc2MoY29udGVudCl9PC9kaXY+YDsKICBpZihzZW5kZXI9PT0nYXNzaXN0YW50Jyl7CiAgICBjb25zdCB0b29scz1kb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTt0b29scy5jbGFzc05hbWU9J2Fuc3dlci10b29scyc7CiAgICBjb25zdCB0eHQ9ZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7dHh0LnRleHRDb250ZW50PSfirIcgRG93bmxvYWQgVFhUJzt0eHQub25jbGljaz0oKT0+ZG93bmxvYWRUZXh0RmlsZSgnT3Jpb25fQW5zd2VyLnR4dCcsU3RyaW5nKGNvbnRlbnR8fCcnKSk7dG9vbHMuYXBwZW5kQ2hpbGQodHh0KTsKICAgIGNvbnN0IGRvY3hCdG49ZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnYnV0dG9uJyk7ZG9jeEJ0bi5jbGFzc05hbWU9J2Rvd25sb2FkLWRvY3gnO2RvY3hCdG4udGV4dENvbnRlbnQ9J/Cfk4QgRG93bmxvYWQgRE9DWCc7ZG9jeEJ0bi5vbmNsaWNrPSgpPT5kb3dubG9hZEFuc3dlckRvY3goU3RyaW5nKGNvbnRlbnR8fCcnKSxkb2N4QnRuKTt0b29scy5hcHBlbmRDaGlsZChkb2N4QnRuKTsKICAgIGNvbnN0IG1kPWRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2J1dHRvbicpO21kLnRleHRDb250ZW50PSfirIcgRG93bmxvYWQgTWFya2Rvd24nO21kLm9uY2xpY2s9KCk9PmRvd25sb2FkVGV4dEZpbGUoJ09yaW9uX0Fuc3dlci5tZCcsU3RyaW5nKGNvbnRlbnR8fCcnKSk7dG9vbHMuYXBwZW5kQ2hpbGQobWQpOwogICAgY29uc3QgY29weT1kb2N1bWVudC5jcmVhdGVFbGVtZW50KCdidXR0b24nKTtjb3B5LnRleHRDb250ZW50PSdDb3B5IGFuc3dlcic7Y29weS5vbmNsaWNrPSgpPT5uYXZpZ2F0b3IuY2xpcGJvYXJkPy53cml0ZVRleHQoU3RyaW5nKGNvbnRlbnR8fCcnKSk7dG9vbHMuYXBwZW5kQ2hpbGQoY29weSk7CiAgICBidWJibGUuYXBwZW5kQ2hpbGQodG9vbHMpOwogIH0KICByb3cuYXBwZW5kQ2hpbGQoYnViYmxlKTskKCdjaGF0JykuYXBwZW5kQ2hpbGQocm93KTtpZihhdXRvU2Nyb2xsKXNjcm9sbFRvQm90dG9tKCk7cmV0dXJuIHJvdzsKfQoKZnVuY3Rpb24gaGFuZGxlRmlsZVNlbGVjdCgpewogIHNlbGVjdGVkRmlsZXM9QXJyYXkuZnJvbSgkKCdmaWxlSW5wdXQnKS5maWxlc3x8W10pOwogIGNvbnN0IGJhZGdlPSQoJ2F0dGFjaG1lbnQnKTsKICBpZihzZWxlY3RlZEZpbGVzLmxlbmd0aCl7CiAgICAkKCdhdHRhY2htZW50TmFtZScpLnRleHRDb250ZW50PXNlbGVjdGVkRmlsZXMubGVuZ3RoPT09MT9zZWxlY3RlZEZpbGVzWzBdLm5hbWU6YCR7c2VsZWN0ZWRGaWxlcy5sZW5ndGh9IGZpbGVzIHNlbGVjdGVkYDsKICAgIGJhZGdlLnN0eWxlLmRpc3BsYXk9J2ZsZXgnOwogIH0gZWxzZSBiYWRnZS5zdHlsZS5kaXNwbGF5PSdub25lJzsKfQpmdW5jdGlvbiBjbGVhckZpbGVzKCl7c2VsZWN0ZWRGaWxlcz1bXTskKCdmaWxlSW5wdXQnKS52YWx1ZT0nJzskKCdhdHRhY2htZW50Jykuc3R5bGUuZGlzcGxheT0nbm9uZSc7JCgnYXR0YWNobWVudE5hbWUnKS50ZXh0Q29udGVudD0nJzt9Cgphc3luYyBmdW5jdGlvbiBzZW5kTWVzc2FnZSgpewogIGNvbnN0IHR5cGVkUXVlcnk9JCgncXVlcnknKS52YWx1ZS50cmltKCk7CiAgaWYoIXR5cGVkUXVlcnkgJiYgIXNlbGVjdGVkRmlsZXMubGVuZ3RoKXskKCdxdWVyeScpLmZvY3VzKCk7cmV0dXJuO30KICBpZighY3VycmVudENoYXRJZCl7CiAgICB0cnl7CiAgICAgIGNvbnN0IGRhdGE9YXdhaXQgYXBpKCcvYXBpL2NoYXRzL25ldycse21ldGhvZDonUE9TVCcsaGVhZGVyczp7J0NvbnRlbnQtVHlwZSc6J2FwcGxpY2F0aW9uL2pzb24nfSxib2R5OkpTT04uc3RyaW5naWZ5KHt0aXRsZTonTmV3IENoYXQnLHRlYW1faWQ6Y3VycmVudFRlYW1JZHx8bnVsbH0pfSk7CiAgICAgIGN1cnJlbnRDaGF0SWQ9ZGF0YS5jaGF0X2lkOwogICAgICAkKCdjaGF0U3RhdHVzJykudGV4dENvbnRlbnQ9Y3VycmVudFRlYW1JZD9gdGVhbToke2N1cnJlbnRUZWFtSWR9YDpjdXJyZW50Q2hhdElkOwogICAgICAkKCdoZWFkZXJUaXRsZScpLnRleHRDb250ZW50PWN1cnJlbnRUZWFtSWQ/J1RlYW0gQ2hhdCc6J09yaW9uIOKAlCBTb3ZlcmVpZ24gV29ya2JlbmNoJzsKICAgICAgJCgnY2hhdCcpLmlubmVySFRNTD0nJzskKCdoZXJvJykuY2xhc3NMaXN0LmFkZCgnaGlkZGVuJyk7YXdhaXQgcmVmcmVzaENoYXRzKCk7YXdhaXQgcmVmcmVzaFRlYW1zKCk7CiAgICB9Y2F0Y2goZXJyKXtyZW5kZXJNZXNzYWdlKCdhc3Npc3RhbnQnLCdVbmFibGUgdG8gY3JlYXRlIGEgbmV3IGNoYXQ6ICcrZXJyLm1lc3NhZ2UpO3JldHVybjt9CiAgfQoKICBjb25zdCBkaXNwbGF5UXVlcnk9dHlwZWRRdWVyeXx8J0FuYWx5emUgdGhlIHVwbG9hZGVkIGZpbGVzJzsKICBjb25zdCBwcmV2aWV3RmlsZXM9c2VsZWN0ZWRGaWxlcy5tYXAoZj0+KHtuYW1lOmYubmFtZSxsb2NhbF91cmw6VVJMLmNyZWF0ZU9iamVjdFVSTChmKX0pKTsKICByZW5kZXJNZXNzYWdlKCd1c2VyJyxkaXNwbGF5UXVlcnkscHJldmlld0ZpbGVzLHRydWUse3NlbmRlcl9uYW1lOmN1cnJlbnRVc2VyPy5uYW1lfHwnWW91JyxzZW5kZXJfZW1haWw6Y3VycmVudFVzZXI/LmVtYWlsfHwnJ30pOwogICQoJ3F1ZXJ5JykudmFsdWU9Jyc7CiAgY29uc3QgbG9hZGluZz1yZW5kZXJNZXNzYWdlKCdhc3Npc3RhbnQnLCdQcm9jZXNzaW5nIGxvY2FsbHnigKYnKTsKCiAgdHJ5ewogICAgaWYoIXNlbGVjdGVkRmlsZXMubGVuZ3RoKXsKICAgICAgY29uc3QgZGF0YT1hd2FpdCBhcGkoJy9hcGkvY2hhdCcse21ldGhvZDonUE9TVCcsaGVhZGVyczp7J0NvbnRlbnQtVHlwZSc6J2FwcGxpY2F0aW9uL2pzb24nfSxib2R5OkpTT04uc3RyaW5naWZ5KHtjaGF0X2lkOmN1cnJlbnRDaGF0SWQscXVlcnk6dHlwZWRRdWVyeX0pfSk7CiAgICAgIGxvYWRpbmcucmVtb3ZlKCk7cmVuZGVyQXNzaXN0YW50UmVzdWx0KGRhdGEsbnVsbCxudWxsLG51bGwsW10pOwogICAgICBwZXJzaXN0QXJ0aWZhY3RzKGN1cnJlbnRDaGF0SWQse2Fuc3dlcjpkYXRhLnJlcG9ydHx8JycsY3JlYXRlZF9hdDpEYXRlLm5vdygpfSk7CiAgICAgIGF3YWl0IHJlZnJlc2hDaGF0cygpO2F3YWl0IHJlZnJlc2hUZWFtcygpO2F3YWl0IHJlZnJlc2hIZWFsdGgoKTsKICAgICAgcmV0dXJuOwogICAgfQoKICAgIGNvbnN0IGZvcm09bmV3IEZvcm1EYXRhKCk7c2VsZWN0ZWRGaWxlcy5mb3JFYWNoKGY9PmZvcm0uYXBwZW5kKCdmaWxlJyxmKSk7Zm9ybS5hcHBlbmQoJ3F1ZXJ5JyxkaXNwbGF5UXVlcnkpO2Zvcm0uYXBwZW5kKCd1c2VyX3F1ZXJ5Jyx0eXBlZFF1ZXJ5KTtmb3JtLmFwcGVuZCgnY2hhdF9pZCcsY3VycmVudENoYXRJZCk7CiAgICBjb25zdCBzZW50TmFtZXM9c2VsZWN0ZWRGaWxlcy5tYXAoZj0+Zi5uYW1lKTtjbGVhckZpbGVzKCk7CiAgICBjb25zdCBkYXRhPWF3YWl0IGFwaSgnL2FwaS9hbmFseXplLWZpbGUnLHttZXRob2Q6J1BPU1QnLGJvZHk6Zm9ybX0pOwogICAgbG9hZGluZy5yZW1vdmUoKTsKICAgIGxldCBjaGFydD1udWxsLGRpYWdyYW09bnVsbCxkb2M9bnVsbDsKICAgIGlmKGRhdGEuY2hhcnRfcmVxdWlyZWQpewogICAgICB0cnl7Y2hhcnQ9YXdhaXQgYXBpKCcvYXBpL2dlbmVyYXRlLWNoYXJ0Jyx7bWV0aG9kOidQT1NUJyxoZWFkZXJzOnsnQ29udGVudC1UeXBlJzonYXBwbGljYXRpb24vanNvbid9LGJvZHk6SlNPTi5zdHJpbmdpZnkoe2NoYXRfaWQ6Y3VycmVudENoYXRJZCxxdWVyeTp0eXBlZFF1ZXJ5fHwnQ3JlYXRlIGEgY2hhcnQgZnJvbSB0aGUgdXBsb2FkZWQgZGF0YScsY2hhcnRfdHlwZTpkYXRhLmNoYXJ0X3R5cGV8fG51bGx9KX0pO30KICAgICAgY2F0Y2goZSl7Y2hhcnQ9e3N0YXR1czonZXJyb3InLG1lc3NhZ2U6ZS5tZXNzYWdlfTt9CiAgICB9CiAgICBpZihkYXRhLmRpYWdyYW1fcmVxdWlyZWQpewogICAgICB0cnl7ZGlhZ3JhbT1hd2FpdCBhcGkoJy9hcGkvZ2VuZXJhdGUtZGlhZ3JhbScse21ldGhvZDonUE9TVCcsaGVhZGVyczp7J0NvbnRlbnQtVHlwZSc6J2FwcGxpY2F0aW9uL2pzb24nfSxib2R5OkpTT04uc3RyaW5naWZ5KHtjaGF0X2lkOmN1cnJlbnRDaGF0SWQscXVlcnk6dHlwZWRRdWVyeXx8J0NyZWF0ZSBhIGRpYWdyYW0gZnJvbSB0aGUgdXBsb2FkZWQgcHJvY2Vzcyd9KX0pO30KICAgICAgY2F0Y2goZSl7ZGlhZ3JhbT17c3RhdHVzOidlcnJvcicsbWVzc2FnZTplLm1lc3NhZ2V9O30KICAgIH0KICAgIGlmKHR5cGVkUXVlcnkmJi9kb2N1bWVudHxyZXBvcnR8YXBwcm92YWwgbm90ZXxtZW1vfGRlbGl2ZXJhYmxlfHdvcmR8ZG9jeHxwcmVzZW50YXRpb258cHB0eHxleGNlbHxjcmVhdGUuKnJlcG9ydHxnZW5lcmF0ZS4qcmVwb3J0L2kudGVzdCh0eXBlZFF1ZXJ5KSl7CiAgICAgIHRyeXtkb2M9YXdhaXQgYXBpKCcvYXBpL2dlbmVyYXRlLWRvYycse21ldGhvZDonUE9TVCcsaGVhZGVyczp7J0NvbnRlbnQtVHlwZSc6J2FwcGxpY2F0aW9uL2pzb24nfSxib2R5OkpTT04uc3RyaW5naWZ5KHtjaGF0X2lkOmN1cnJlbnRDaGF0SWQscXVlcnk6dHlwZWRRdWVyeSxmb3JtYXQ6J2RvY3gnfSl9KTt9CiAgICAgIGNhdGNoKGUpe2RvYz17c3RhdHVzOidlcnJvcicsbWVzc2FnZTplLm1lc3NhZ2V9O30KICAgIH0KICAgIHJlbmRlckFzc2lzdGFudFJlc3VsdChkYXRhLGNoYXJ0LGRpYWdyYW0sZG9jLHNlbnROYW1lcyk7CiAgICBwZXJzaXN0QXJ0aWZhY3RzKGN1cnJlbnRDaGF0SWQse2NoYXJ0LGRpYWdyYW0sZG9jLGFuc3dlcjpkYXRhLnJlcG9ydHx8JycsZmlsZXM6c2VudE5hbWVzLGNyZWF0ZWRfYXQ6RGF0ZS5ub3coKX0pOwogICAgYXdhaXQgcmVmcmVzaENoYXRzKCk7YXdhaXQgcmVmcmVzaFRlYW1zKCk7YXdhaXQgcmVmcmVzaEhlYWx0aCgpOwogIH1jYXRjaChlcnIpe2xvYWRpbmcucmVtb3ZlKCk7cmVuZGVyTWVzc2FnZSgnYXNzaXN0YW50Jywn4pqg77iPICcrZXJyLm1lc3NhZ2UpO30KfQoKZnVuY3Rpb24gYXJ0aWZhY3RTdG9yZUtleShjaGF0SWQpe3JldHVybiBgc292ZXJlaWduX2FydGlmYWN0c18ke2NoYXRJZH1gO30KZnVuY3Rpb24gcGVyc2lzdEFydGlmYWN0cyhjaGF0SWQscmVjb3JkKXsKICBpZighY2hhdElkKXJldHVybjsKICB0cnl7CiAgICBjb25zdCBrZXk9YXJ0aWZhY3RTdG9yZUtleShjaGF0SWQpO2NvbnN0IGV4aXN0aW5nPUpTT04ucGFyc2UobG9jYWxTdG9yYWdlLmdldEl0ZW0oa2V5KXx8J1tdJyk7CiAgICBleGlzdGluZy5wdXNoKHJlY29yZCk7bG9jYWxTdG9yYWdlLnNldEl0ZW0oa2V5LEpTT04uc3RyaW5naWZ5KGV4aXN0aW5nLnNsaWNlKC01MCkpKTsKICB9Y2F0Y2goXyl7IH0KfQpmdW5jdGlvbiByZXN0b3JlQXJ0aWZhY3RzKGNoYXRJZCl7CiAgaWYoIWNoYXRJZClyZXR1cm47CiAgdHJ5ewogICAgY29uc3QgcmVjb3Jkcz1KU09OLnBhcnNlKGxvY2FsU3RvcmFnZS5nZXRJdGVtKGFydGlmYWN0U3RvcmVLZXkoY2hhdElkKSl8fCdbXScpOwogICAgcmVjb3Jkcy5mb3JFYWNoKHI9PnsKICAgICAgY29uc3QgaGFzPXIuY2hhcnQ/LmRvd25sb2FkX3VybHx8ci5jaGFydD8uaW1hZ2VfdXJsfHxyLmRpYWdyYW0/LmRvd25sb2FkX3VybHx8ci5kaWFncmFtPy5pbWFnZV91cmx8fHIuZG9jPy5maWxlX3VybDsKICAgICAgaWYoaGFzKXJlbmRlckFzc2lzdGFudFJlc3VsdCh7cmVwb3J0OicnLG1vZGVsX3VzZWQ6J2xvY2FsJyx0YXNrX3R5cGU6J2FydGlmYWN0J30sci5jaGFydCxyLmRpYWdyYW0sci5kb2MsW10pOwogICAgfSk7CiAgfWNhdGNoKF8peyB9Cn0KCmZ1bmN0aW9uIHJlbmRlckFzc2lzdGFudFJlc3VsdChkYXRhLGNoYXJ0LGRpYWdyYW0sZG9jLG5hbWVzKXsKICBjb25zdCByb3c9ZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7cm93LmNsYXNzTmFtZT0nbWVzc2FnZSBhc3Npc3RhbnQnOwogIGNvbnN0IGJ1YmJsZT1kb2N1bWVudC5jcmVhdGVFbGVtZW50KCdkaXYnKTtidWJibGUuY2xhc3NOYW1lPSdidWJibGUnOwogIGxldCBodG1sPSc8ZGl2IGNsYXNzPSJzZW5kZXItbGFiZWwiPk9yaW9uPC9kaXY+JzsKICBpZihkYXRhLnJlcG9ydClodG1sKz1gPGRpdiBjbGFzcz0iYXNzaXN0YW50LWJvZHkiPiR7ZXNjKGRhdGEucmVwb3J0KX08L2Rpdj5gOwogIGlmKG5hbWVzPy5sZW5ndGgpaHRtbCs9cmVuZGVyRmlsZUNhcmRzKG5hbWVzLm1hcChuYW1lPT4oe25hbWV9KSkpOwogIGlmKGRhdGEudmVyaWZpZWQpaHRtbCs9YDxkaXYgY2xhc3M9ImZpbGVzLWxpbmUiPuKckyBBbnN3ZXIgZ3JvdW5kZWQgaW4gdXBsb2FkZWQgZXZpZGVuY2Uke2RhdGEuY2hhcnRfdHlwZT9gIMK3IGNoYXJ0IHR5cGU6ICR7ZXNjKGRhdGEuY2hhcnRfdHlwZSl9YDonJ308L2Rpdj5gOwogIGNvbnN0IGFjdGlvbnM9W107CiAgaWYoY2hhcnQ/LmRvd25sb2FkX3VybClhY3Rpb25zLnB1c2goYDxidXR0b24gb25jbGljaz0nZG93bmxvYWRBcnRpZmFjdCgke0pTT04uc3RyaW5naWZ5KGNoYXJ0LmRvd25sb2FkX3VybCl9LCR7SlNPTi5zdHJpbmdpZnkoY2hhcnQuZmlsZW5hbWV8fCdvcmlvbi1jaGFydC5wbmcnKX0pJz7wn5OKIERvd25sb2FkIFBORzwvYnV0dG9uPmApOwogIGlmKGNoYXJ0Py5pbWFnZV91cmwpaHRtbCs9YDxkaXY+PGltZyBjbGFzcz0iYXJ0aWZhY3QiIHNyYz0iJHtlc2MoY2hhcnQuaW1hZ2VfdXJsKX0iIGFsdD0iR2VuZXJhdGVkIGNoYXJ0IiBvbmNsaWNrPSJ3aW5kb3cub3Blbigke0pTT04uc3RyaW5naWZ5KGNoYXJ0LmltYWdlX3VybCl9LCdfYmxhbmsnKSI+PC9kaXY+YDsKICBpZihkaWFncmFtPy5kb3dubG9hZF91cmwpYWN0aW9ucy5wdXNoKGA8YnV0dG9uIG9uY2xpY2s9J2Rvd25sb2FkQXJ0aWZhY3QoJHtKU09OLnN0cmluZ2lmeShkaWFncmFtLmRvd25sb2FkX3VybCl9LCR7SlNPTi5zdHJpbmdpZnkoZGlhZ3JhbS5maWxlbmFtZXx8J29yaW9uLWRpYWdyYW0ucG5nJyl9KSc+8J+Xuu+4jyBEb3dubG9hZCBQTkc8L2J1dHRvbj5gKTsKICBpZihkaWFncmFtPy5pbWFnZV91cmwpaHRtbCs9YDxkaXY+PGltZyBjbGFzcz0iYXJ0aWZhY3QiIHNyYz0iJHtlc2MoZGlhZ3JhbS5pbWFnZV91cmwpfSIgYWx0PSJHZW5lcmF0ZWQgZGlhZ3JhbSIgb25jbGljaz0id2luZG93Lm9wZW4oJHtKU09OLnN0cmluZ2lmeShkaWFncmFtLmltYWdlX3VybCl9LCdfYmxhbmsnKSI+PC9kaXY+YDsKICBpZihkb2M/LmZpbGVfdXJsKWFjdGlvbnMucHVzaChgPGJ1dHRvbiBvbmNsaWNrPSdkb3dubG9hZEFydGlmYWN0KCR7SlNPTi5zdHJpbmdpZnkoZG9jLmZpbGVfdXJsKX0sJHtKU09OLnN0cmluZ2lmeShkb2MuZmlsZW5hbWV8fCdvcmlvbi1yZXBvcnQuZG9jeCcpfSknPvCfk4QgRG93bmxvYWQgRE9DWCByZXBvcnQ8L2J1dHRvbj5gKTsKICBpZihjaGFydD8uc3RhdHVzPT09J2Vycm9yJylhY3Rpb25zLnB1c2goYDxzcGFuIGNsYXNzPSJkYW5nZXItbm90ZSI+Q2hhcnQ6ICR7ZXNjKGNoYXJ0Lm1lc3NhZ2UpfTwvc3Bhbj5gKTsKICBpZihkaWFncmFtPy5zdGF0dXM9PT0nZXJyb3InKWFjdGlvbnMucHVzaChgPHNwYW4gY2xhc3M9ImRhbmdlci1ub3RlIj5EaWFncmFtOiAke2VzYyhkaWFncmFtLm1lc3NhZ2UpfTwvc3Bhbj5gKTsKICBpZihkb2M/LnN0YXR1cz09PSdlcnJvcicpYWN0aW9ucy5wdXNoKGA8c3BhbiBjbGFzcz0iZGFuZ2VyLW5vdGUiPkRvY3VtZW50OiAke2VzYyhkb2MubWVzc2FnZSl9PC9zcGFuPmApOwogIC8vIEV2ZXJ5IE9yaW9uIGFuc3dlciBnZXRzIGxvY2FsIGV4cG9ydCBjb250cm9scywgZXZlbiB3aGVuIG5vIGNoYXJ0L2RvY3VtZW50IHdhcyByZXF1ZXN0ZWQuCiAgaWYoZGF0YS5yZXBvcnQpewogICAgYWN0aW9ucy5wdXNoKGA8YnV0dG9uIG9uY2xpY2s9J2Rvd25sb2FkVGV4dEZpbGUoXCJPcmlvbl9BbnN3ZXIudHh0XCIsJHtKU09OLnN0cmluZ2lmeShkYXRhLnJlcG9ydCl9KSc+4qyHIERvd25sb2FkIFRYVDwvYnV0dG9uPmApOwogICAgYWN0aW9ucy5wdXNoKGA8YnV0dG9uIG9uY2xpY2s9J2Rvd25sb2FkQW5zd2VyRG9jeCgke0pTT04uc3RyaW5naWZ5KGRhdGEucmVwb3J0KX0sdGhpcyknPvCfk4QgRG93bmxvYWQgRE9DWDwvYnV0dG9uPmApOwogICAgYWN0aW9ucy5wdXNoKGA8YnV0dG9uIG9uY2xpY2s9J2Rvd25sb2FkVGV4dEZpbGUoXCJPcmlvbl9BbnN3ZXIubWRcIiwke0pTT04uc3RyaW5naWZ5KGRhdGEucmVwb3J0KX0pJz7irIcgRG93bmxvYWQgTWFya2Rvd248L2J1dHRvbj5gKTsKICAgIGFjdGlvbnMucHVzaChgPGJ1dHRvbiBvbmNsaWNrPSduYXZpZ2F0b3IuY2xpcGJvYXJkPy53cml0ZVRleHQoJHtKU09OLnN0cmluZ2lmeShkYXRhLnJlcG9ydCl9KSc+Q29weSBhbnN3ZXI8L2J1dHRvbj5gKTsKICB9CiAgaWYoYWN0aW9ucy5sZW5ndGgpaHRtbCs9YDxkaXYgY2xhc3M9ImFydGlmYWN0LWFjdGlvbnMiPiR7YWN0aW9ucy5qb2luKCcnKX08L2Rpdj5gOwogIGNvbnN0IGtiPShkYXRhLmtiX3NvdXJjZXNfdXNlZHx8W10pLmpvaW4oJywgJyl8fCdub25lIG1hdGNoZWQnOwogIGh0bWwrPWA8ZGl2IGNsYXNzPSJtZXRhIj48c3Bhbj48Yj5Nb2RlbDwvYj4gJHtlc2MoZGF0YS5tb2RlbF91c2VkfHwnbG9jYWwnKX08L3NwYW4+PHNwYW4+PGI+VGFzazwvYj4gJHtlc2MoZGF0YS50YXNrX3R5cGV8fCdnZW5lcmFsJyl9PC9zcGFuPjxzcGFuPjxiPlNlY3RvcjwvYj4gJHtlc2MoZGF0YS5zZWN0b3J8fGN1cnJlbnRVc2VyPy5zZWN0b3J8fCduL2EnKX08L3NwYW4+PHNwYW4+PGI+S0I8L2I+ICR7ZXNjKGtiKX08L3NwYW4+PC9kaXY+YDsKICBidWJibGUuaW5uZXJIVE1MPWh0bWw7cm93LmFwcGVuZENoaWxkKGJ1YmJsZSk7JCgnY2hhdCcpLmFwcGVuZENoaWxkKHJvdyk7c2Nyb2xsVG9Cb3R0b20oKTsKfQoKZnVuY3Rpb24gY2xvc2VNb2RhbChpZCl7JChpZCkuY2xhc3NMaXN0LmFkZCgnaGlkZGVuJyk7fQpmdW5jdGlvbiBvcGVuUHJvZmlsZSgpewogIGlmKCFjdXJyZW50VXNlcilyZXR1cm47CiAgJCgncHJvZmlsZU1vZGFsTmFtZScpLnRleHRDb250ZW50PWN1cnJlbnRVc2VyLm5hbWV8fCfigJQnOyQoJ3Byb2ZpbGVNb2RhbEVtYWlsJykudGV4dENvbnRlbnQ9Y3VycmVudFVzZXIuZW1haWx8fCfigJQnOyQoJ3Byb2ZpbGVNb2RhbENvbXBhbnknKS50ZXh0Q29udGVudD1jdXJyZW50VXNlci5jb21wYW55fHwn4oCUJzskKCdwcm9maWxlTW9kYWxTZWN0b3InKS50ZXh0Q29udGVudD1jdXJyZW50VXNlci5zZWN0b3J8fCfigJQnOyQoJ3Byb2ZpbGVNb2RhbFJvbGUnKS50ZXh0Q29udGVudD1jdXJyZW50VXNlci5yb2xlfHwnRW1wbG95ZWUnOyQoJ3Byb2ZpbGVNb2RhbCcpLmNsYXNzTGlzdC5yZW1vdmUoJ2hpZGRlbicpOwp9CmFzeW5jIGZ1bmN0aW9uIG9wZW5MaWJyYXJ5KCl7CiAgJCgnbGlicmFyeU1vZGFsJykuY2xhc3NMaXN0LnJlbW92ZSgnaGlkZGVuJyk7Y29uc3QgYm94PSQoJ2xpYnJhcnlMaXN0Jyk7Ym94LmlubmVySFRNTD0nPGRpdiBjbGFzcz0iZW1wdHktbGlzdCI+TG9hZGluZyBsaWJyYXJ54oCmPC9kaXY+JzsKICB0cnl7Y29uc3QgZGF0YT1hd2FpdCBhcGkoJy9hcGkvZmlsZXMnKTtjb25zdCBmaWxlcz1kYXRhLmZpbGVzfHxbXTtib3guaW5uZXJIVE1MPScnO2lmKCFmaWxlcy5sZW5ndGgpe2JveC5pbm5lckhUTUw9JzxkaXYgY2xhc3M9ImVtcHR5LWxpc3QiPk5vIHVwbG9hZGVkIGZpbGVzIHlldC48L2Rpdj4nO3JldHVybjt9ZmlsZXMuZm9yRWFjaChmPT57Y29uc3Qgcm93PWRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpO3Jvdy5jbGFzc05hbWU9J2xpYnJhcnktaXRlbSc7cm93LmlubmVySFRNTD1gPGRpdj48ZGl2IGNsYXNzPSJuYW1lIj4ke2VzYyhmLm9yaWdpbmFsX25hbWUpfTwvZGl2PjxkaXYgY2xhc3M9Im1ldGEiPiR7ZXNjKGYuc2VjdG9yfHwnR2VuZXJhbCcpfSDCtyBjaGF0ICR7ZXNjKGYuY2hhdF9pZHx8J3VuYXNzaWduZWQnKX08L2Rpdj48L2Rpdj48YSBjbGFzcz0ic21hbGwtYnRuIiBocmVmPSIvYXBpL2ZpbGVzLyR7TnVtYmVyKGYuaWQpfS9kb3dubG9hZCIgdGFyZ2V0PSJfYmxhbmsiPk9wZW48L2E+YDtib3guYXBwZW5kQ2hpbGQocm93KTt9KTt9CiAgY2F0Y2goZXJyKXtib3guaW5uZXJIVE1MPWA8ZGl2IGNsYXNzPSJkYW5nZXItbm90ZSI+JHtlc2MoZXJyLm1lc3NhZ2UpfTwvZGl2PmA7fQp9CmFzeW5jIGZ1bmN0aW9uIG9wZW5TZXR0aW5ncygpewogICQoJ3NldHRpbmdzTW9kYWwnKS5jbGFzc0xpc3QucmVtb3ZlKCdoaWRkZW4nKTsKICB0cnl7Y29uc3QgZGF0YT1hd2FpdCBhcGkoJy9hcGkvaGVhbHRoJyk7JCgnc2V0dGluZ3NPbGxhbWEnKS50ZXh0Q29udGVudD1kYXRhLm9sbGFtYV9yZWFjaGFibGU/J29ubGluZSc6J29mZmxpbmUnOyQoJ3NldHRpbmdzTW9kZWwnKS50ZXh0Q29udGVudD1kYXRhLnJlc29sdmVkX21vZGVsfHwnZHluYW1pYyc7JCgnc2V0dGluZ3NNb2RlbHMnKS50ZXh0Q29udGVudD0oZGF0YS5hdmFpbGFibGVfbW9kZWxzfHxbXSkuam9pbignLCAnKXx8J25vbmUgZGV0ZWN0ZWQnO30KICBjYXRjaChlcnIpeyQoJ3NldHRpbmdzT2xsYW1hJykudGV4dENvbnRlbnQ9J3VuYXZhaWxhYmxlJzskKCdzZXR0aW5nc01vZGVsJykudGV4dENvbnRlbnQ9J+KAlCc7JCgnc2V0dGluZ3NNb2RlbHMnKS50ZXh0Q29udGVudD0n4oCUJzt9Cn0KCmFzeW5jIGZ1bmN0aW9uIG9wZW5BZG1pbigpewogICQoJ2FkbWluTW9kYWwnKS5jbGFzc0xpc3QucmVtb3ZlKCdoaWRkZW4nKTskKCdhZG1pbkVycm9yJykudGV4dENvbnRlbnQ9Jyc7YXdhaXQgcmVmcmVzaFBlbmRpbmdSZXF1ZXN0cygpOwp9CmZ1bmN0aW9uIGNsb3NlQWRtaW4oKXskKCdhZG1pbk1vZGFsJykuY2xhc3NMaXN0LmFkZCgnaGlkZGVuJyk7fQphc3luYyBmdW5jdGlvbiByZWZyZXNoUGVuZGluZ1JlcXVlc3RzKCl7CiAgY29uc3QgYm94PSQoJ3BlbmRpbmdSZXF1ZXN0cycpO2JveC5pbm5lckhUTUw9JzxkaXYgY2xhc3M9ImVtcHR5LWxpc3QiPkxvYWRpbmfigKY8L2Rpdj4nOwogIHRyeXsKICAgIGNvbnN0IGRhdGE9YXdhaXQgYXBpKCcvYXBpL2FkbWluL3JlZ2lzdHJhdGlvbi1yZXF1ZXN0cycpO2JveC5pbm5lckhUTUw9Jyc7CiAgICBpZighKGRhdGEucmVxdWVzdHN8fFtdKS5sZW5ndGgpe2JveC5pbm5lckhUTUw9JzxkaXYgY2xhc3M9ImVtcHR5LWxpc3QiPk5vIHBlbmRpbmcgZW1wbG95ZWUgcmVnaXN0cmF0aW9ucy48L2Rpdj4nO3JldHVybjt9CiAgICBkYXRhLnJlcXVlc3RzLmZvckVhY2gocj0+ewogICAgICBjb25zdCBjYXJkPWRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpO2NhcmQuY2xhc3NOYW1lPSdyZXF1ZXN0LWNhcmQnOwogICAgICBjYXJkLmlubmVySFRNTD1gPHN0cm9uZz4ke2VzYyhyLm5hbWV8fCdVbm5hbWVkIGVtcGxveWVlJyl9PC9zdHJvbmc+PGRpdj4ke2VzYyhyLmVtYWlsfHwnJyl9IMK3ICR7ZXNjKHIuc2VjdG9yfHwnJyl9PC9kaXY+PGRpdiBjbGFzcz0icmVxdWVzdC1hY3Rpb25zIj48YnV0dG9uIGNsYXNzPSJzbWFsbC1idG4iIG9uY2xpY2s9ImRlY2lkZVJlZ2lzdHJhdGlvbigke051bWJlcihyLmlkKX0sJ3JlamVjdGVkJykiPlJlamVjdDwvYnV0dG9uPjxidXR0b24gY2xhc3M9InByaW1hcnkiIHN0eWxlPSJwYWRkaW5nOjdweCIgb25jbGljaz0iZGVjaWRlUmVnaXN0cmF0aW9uKCR7TnVtYmVyKHIuaWQpfSwnYXBwcm92ZWQnKSI+QXBwcm92ZTwvYnV0dG9uPjwvZGl2PmA7Ym94LmFwcGVuZENoaWxkKGNhcmQpOwogICAgfSk7CiAgfWNhdGNoKGVycil7Ym94LmlubmVySFRNTD1gPGRpdiBjbGFzcz0iZGFuZ2VyLW5vdGUiPiR7ZXNjKGVyci5tZXNzYWdlKX08L2Rpdj5gO30KfQphc3luYyBmdW5jdGlvbiBkZWNpZGVSZWdpc3RyYXRpb24oaWQsZGVjaXNpb24pewogIHRyeXthd2FpdCBhcGkoYC9hcGkvYWRtaW4vcmVnaXN0cmF0aW9uLXJlcXVlc3RzLyR7aWR9YCx7bWV0aG9kOidQT1NUJyxoZWFkZXJzOnsnQ29udGVudC1UeXBlJzonYXBwbGljYXRpb24vanNvbid9LGJvZHk6SlNPTi5zdHJpbmdpZnkoe2RlY2lzaW9ufSl9KTthd2FpdCByZWZyZXNoUGVuZGluZ1JlcXVlc3RzKCk7YXdhaXQgcmVmcmVzaEhlYWx0aCgpO30KICBjYXRjaChlcnIpeyQoJ2FkbWluRXJyb3InKS50ZXh0Q29udGVudD1lcnIubWVzc2FnZTt9Cn0KYXN5bmMgZnVuY3Rpb24gY3JlYXRlRW1wbG95ZWUoKXsKICAkKCdhZG1pbkVycm9yJykudGV4dENvbnRlbnQ9Jyc7CiAgdHJ5ewogICAgY29uc3QgZGF0YT1hd2FpdCBhcGkoJy9hcGkvZW1wbG95ZWVzL2NyZWF0ZScse21ldGhvZDonUE9TVCcsaGVhZGVyczp7J0NvbnRlbnQtVHlwZSc6J2FwcGxpY2F0aW9uL2pzb24nfSxib2R5OkpTT04uc3RyaW5naWZ5KHtuYW1lOiQoJ25ld0VtcGxveWVlTmFtZScpLnZhbHVlLnRyaW0oKSxlbWFpbDokKCduZXdFbXBsb3llZUVtYWlsJykudmFsdWUudHJpbSgpLHBhc3N3b3JkOiQoJ25ld0VtcGxveWVlUGFzc3dvcmQnKS52YWx1ZSxzZWN0b3I6JCgnbmV3RW1wbG95ZWVTZWN0b3InKS52YWx1ZS50cmltKCkscm9sZTonRW1wbG95ZWUnfSl9KTsKICAgIGFsZXJ0KGBFbXBsb3llZSBjcmVhdGVkLiBVc2VyIElEOiAke2RhdGEudXNlcl9pZH1gKTskKCduZXdFbXBsb3llZU5hbWUnKS52YWx1ZT0nJzskKCduZXdFbXBsb3llZUVtYWlsJykudmFsdWU9Jyc7JCgnbmV3RW1wbG95ZWVQYXNzd29yZCcpLnZhbHVlPScnOyQoJ25ld0VtcGxveWVlU2VjdG9yJykudmFsdWU9Jyc7YXdhaXQgcmVmcmVzaFBlbmRpbmdSZXF1ZXN0cygpOwogIH1jYXRjaChlcnIpeyQoJ2FkbWluRXJyb3InKS50ZXh0Q29udGVudD1lcnIubWVzc2FnZTt9Cn0KCmFzeW5jIGZ1bmN0aW9uIHVwbG9hZFNPUCgpewogICQoJ3NvcEVycm9yJykudGV4dENvbnRlbnQ9Jyc7CiAgY29uc3QgZmlsZUlucHV0PSQoJ3NvcEZpbGVJbnB1dCcpOwogIGNvbnN0IHNlY3Rvcj0kKCdzb3BTZWN0b3InKS52YWx1ZS50cmltKCk7CiAgY29uc3QgZmlsZT1maWxlSW5wdXQuZmlsZXMmJmZpbGVJbnB1dC5maWxlc1swXTsKICBpZighZmlsZSl7JCgnc29wRXJyb3InKS50ZXh0Q29udGVudD0nU2VsZWN0IGFuIFNPUCBmaWxlIGZpcnN0Lic7cmV0dXJuO30KICB0cnl7CiAgICBjb25zdCBmb3JtPW5ldyBGb3JtRGF0YSgpOwogICAgZm9ybS5hcHBlbmQoJ2ZpbGUnLGZpbGUpOwogICAgaWYoc2VjdG9yKWZvcm0uYXBwZW5kKCdzZWN0b3InLHNlY3Rvcik7CiAgICBjb25zdCBkYXRhPWF3YWl0IGFwaSgnL2FwaS9zb3AtdXBsb2FkJyx7bWV0aG9kOidQT1NUJyxib2R5OmZvcm19KTsKICAgIGFsZXJ0KGRhdGEubWVzc2FnZXx8J1NPUCB1cGxvYWRlZCBzdWNjZXNzZnVsbHkuJyk7CiAgICBmaWxlSW5wdXQudmFsdWU9Jyc7JCgnc29wU2VjdG9yJykudmFsdWU9Jyc7CiAgfWNhdGNoKGVycil7JCgnc29wRXJyb3InKS50ZXh0Q29udGVudD1lcnIubWVzc2FnZTt9Cn0KCmFzeW5jIGZ1bmN0aW9uIG9wZW5UZWFtTWFuYWdlcigpewogICQoJ3RlYW1Nb2RhbCcpLmNsYXNzTGlzdC5yZW1vdmUoJ2hpZGRlbicpOyQoJ3RlYW1FcnJvcicpLnRleHRDb250ZW50PScnOwogIGF3YWl0IGxvYWRBcHByb3ZlZEVtcGxveWVlcygpOwogIGlmKGN1cnJlbnRUZWFtSWQpIGF3YWl0IGxvYWRDdXJyZW50VGVhbU1hbmFnZXIoY3VycmVudFRlYW1JZCk7CiAgZWxzZSAkKCdjdXJyZW50VGVhbUluZm8nKS50ZXh0Q29udGVudD0nU2VsZWN0IGEgdGVhbSBmcm9tIHRoZSBsZWZ0IHNpZGViYXIuJzsKfQpmdW5jdGlvbiBjbG9zZVRlYW1NYW5hZ2VyKCl7JCgndGVhbU1vZGFsJykuY2xhc3NMaXN0LmFkZCgnaGlkZGVuJyk7fQphc3luYyBmdW5jdGlvbiBsb2FkQXBwcm92ZWRFbXBsb3llZXMoKXsKICBjb25zdCBib3g9JCgnYXBwcm92ZWRFbXBsb3llZUxpc3QnKTtib3guaW5uZXJIVE1MPSc8ZGl2IGNsYXNzPSJlbXB0eS1saXN0Ij5Mb2FkaW5nIGFwcHJvdmVkIGVtcGxveWVlc+KApjwvZGl2Pic7CiAgdHJ5ewogICAgY29uc3QgZGF0YT1hd2FpdCBhcGkoJy9hcGkvdXNlcnMnKTtib3guaW5uZXJIVE1MPScnOwogICAgY29uc3QgdXNlcnM9KGRhdGEudXNlcnN8fFtdKS5maWx0ZXIodT0+TnVtYmVyKHUuaWQpIT09TnVtYmVyKGN1cnJlbnRVc2VyPy5pZCkpOwogICAgaWYoIXVzZXJzLmxlbmd0aCl7Ym94LmlubmVySFRNTD0nPGRpdiBjbGFzcz0iZW1wdHktbGlzdCI+Tm8gb3RoZXIgYXBwcm92ZWQgZW1wbG95ZWVzIGFyZSBhdmFpbGFibGUgeWV0LjwvZGl2Pic7cmV0dXJuO30KICAgIHVzZXJzLmZvckVhY2godT0+ewogICAgICBjb25zdCByb3c9ZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnbGFiZWwnKTtyb3cuY2xhc3NOYW1lPSdtZW1iZXItY2hlY2snOwogICAgICByb3cuaW5uZXJIVE1MPWA8aW5wdXQgdHlwZT0iY2hlY2tib3giIHZhbHVlPSIke051bWJlcih1LmlkKX0iPjxzcGFuPiR7ZXNjKHUubmFtZXx8dS5lbWFpbCl9PGJyPjxzcGFuIGNsYXNzPSJtdXRlZCI+JHtlc2ModS5lbWFpbCl9IMK3ICR7ZXNjKHUuc2VjdG9yfHwnJyl9PC9zcGFuPjwvc3Bhbj5gOwogICAgICBib3guYXBwZW5kQ2hpbGQocm93KTsKICAgIH0pOwogIH1jYXRjaChlcnIpe2JveC5pbm5lckhUTUw9YDxkaXYgY2xhc3M9ImRhbmdlci1ub3RlIj4ke2VzYyhlcnIubWVzc2FnZSl9PC9kaXY+YDt9Cn0KYXN5bmMgZnVuY3Rpb24gY3JlYXRlVGVhbUZyb21VSSgpewogICQoJ3RlYW1FcnJvcicpLnRleHRDb250ZW50PScnO2NvbnN0IG5hbWU9JCgndGVhbU5hbWUnKS52YWx1ZS50cmltKCk7CiAgY29uc3QgbWVtYmVyX2lkcz1BcnJheS5mcm9tKGRvY3VtZW50LnF1ZXJ5U2VsZWN0b3JBbGwoJyNhcHByb3ZlZEVtcGxveWVlTGlzdCBpbnB1dFt0eXBlPWNoZWNrYm94XTpjaGVja2VkJykpLm1hcCh4PT5OdW1iZXIoeC52YWx1ZSkpOwogIGlmKCFuYW1lKXskKCd0ZWFtRXJyb3InKS50ZXh0Q29udGVudD0nVGVhbSBuYW1lIGlzIHJlcXVpcmVkLic7cmV0dXJuO30KICB0cnl7CiAgICBjb25zdCBkYXRhPWF3YWl0IGFwaSgnL2FwaS90ZWFtcy9jcmVhdGUnLHttZXRob2Q6J1BPU1QnLGhlYWRlcnM6eydDb250ZW50LVR5cGUnOidhcHBsaWNhdGlvbi9qc29uJ30sYm9keTpKU09OLnN0cmluZ2lmeSh7bmFtZSxtZW1iZXJfaWRzfSl9KTsKICAgICQoJ3RlYW1OYW1lJykudmFsdWU9Jyc7JCgnam9pblRlYW1Db2RlJykudmFsdWU9Jyc7YXdhaXQgcmVmcmVzaFRlYW1zKCk7CiAgICBjb25zdCB0ZWFtPShhd2FpdCBhcGkoJy9hcGkvdGVhbXMnKSkudGVhbXMuZmluZCh0PT5OdW1iZXIodC5pZCk9PT1OdW1iZXIoZGF0YS50ZWFtX2lkKSk7CiAgICBpZih0ZWFtKXthd2FpdCBvcGVuVGVhbSh0ZWFtKTthd2FpdCBsb2FkQ3VycmVudFRlYW1NYW5hZ2VyKHRlYW0uaWQpO30KICAgIGFsZXJ0KGBUZWFtIGNyZWF0ZWQuIEludml0ZSBjb2RlOiAke2RhdGEuaW52aXRlX2NvZGV9YCk7CiAgfWNhdGNoKGVycil7JCgndGVhbUVycm9yJykudGV4dENvbnRlbnQ9ZXJyLm1lc3NhZ2U7fQp9CmFzeW5jIGZ1bmN0aW9uIGpvaW5UZWFtRnJvbVVJKCl7CiAgJCgndGVhbUVycm9yJykudGV4dENvbnRlbnQ9Jyc7Y29uc3QgY29kZT0kKCdqb2luVGVhbUNvZGUnKS52YWx1ZS50cmltKCk7CiAgaWYoIWNvZGUpeyQoJ3RlYW1FcnJvcicpLnRleHRDb250ZW50PSdJbnZpdGUgY29kZSBpcyByZXF1aXJlZC4nO3JldHVybjt9CiAgdHJ5ewogICAgY29uc3QgZGF0YT1hd2FpdCBhcGkoJy9hcGkvdGVhbXMvam9pbicse21ldGhvZDonUE9TVCcsaGVhZGVyczp7J0NvbnRlbnQtVHlwZSc6J2FwcGxpY2F0aW9uL2pzb24nfSxib2R5OkpTT04uc3RyaW5naWZ5KHtpbnZpdGVfY29kZTpjb2RlfSl9KTsKICAgICQoJ2pvaW5UZWFtQ29kZScpLnZhbHVlPScnO2F3YWl0IHJlZnJlc2hUZWFtcygpOwogICAgaWYoZGF0YS50ZWFtKSB7YXdhaXQgb3BlblRlYW0oZGF0YS50ZWFtKTthd2FpdCBsb2FkQ3VycmVudFRlYW1NYW5hZ2VyKGRhdGEudGVhbS5pZCk7fQogIH1jYXRjaChlcnIpeyQoJ3RlYW1FcnJvcicpLnRleHRDb250ZW50PWVyci5tZXNzYWdlO30KfQphc3luYyBmdW5jdGlvbiBsb2FkQ3VycmVudFRlYW1NYW5hZ2VyKHRlYW1JZCl7CiAgY3VycmVudFRlYW1JZD1OdW1iZXIodGVhbUlkKTtjb25zdCB0ZWFtcz0oYXdhaXQgYXBpKCcvYXBpL3RlYW1zJykpLnRlYW1zfHxbXTtjb25zdCB0ZWFtPXRlYW1zLmZpbmQodD0+TnVtYmVyKHQuaWQpPT09TnVtYmVyKHRlYW1JZCkpOwogIGlmKCF0ZWFtKXskKCdjdXJyZW50VGVhbUluZm8nKS50ZXh0Q29udGVudD0nVGVhbSBpcyBubyBsb25nZXIgYXZhaWxhYmxlLic7cmV0dXJuO30KICAkKCdjdXJyZW50VGVhbUluZm8nKS5pbm5lckhUTUw9YDxkaXY+PHN0cm9uZz4ke2VzYyh0ZWFtLm5hbWUpfTwvc3Ryb25nPjwvZGl2PjxkaXYgY2xhc3M9InRlYW0tY29kZSI+SW52aXRlIGNvZGU6ICR7ZXNjKHRlYW0uaW52aXRlX2NvZGUpfTwvZGl2PmA7CiAgdHJ5ewogICAgY29uc3QgbWVtYmVyc0RhdGE9YXdhaXQgYXBpKGAvYXBpL3RlYW1zLyR7TnVtYmVyKHRlYW1JZCl9L21lbWJlcnNgKTtjb25zdCBtZW1iZXJzPW1lbWJlcnNEYXRhLm1lbWJlcnN8fFtdOwogICAgJCgnY3VycmVudFRlYW1NZW1iZXJzJykuaW5uZXJIVE1MPW1lbWJlcnMubGVuZ3RoP21lbWJlcnMubWFwKG09PmA8ZGl2IGNsYXNzPSJtZW1iZXItcm93Ij48c3Bhbj4ke2VzYyhtLm5hbWV8fG0uZW1haWwpfTxicj48c3BhbiBjbGFzcz0ibXV0ZWQiPiR7ZXNjKG0uZW1haWx8fCcnKX0gwrcgJHtlc2MobS50ZWFtX3JvbGV8fCdtZW1iZXInKX08L3NwYW4+PC9zcGFuPjwvZGl2PmApLmpvaW4oJycpOic8ZGl2IGNsYXNzPSJlbXB0eS1saXN0Ij5ObyBtZW1iZXJzLjwvZGl2Pic7CiAgICBjb25zdCBhbGw9KGF3YWl0IGFwaSgnL2FwaS91c2VycycpKS51c2Vyc3x8W107Y29uc3QgZXhpc3Rpbmc9bmV3IFNldChtZW1iZXJzLm1hcChtPT5OdW1iZXIobS5pZCkpKTsKICAgIGNvbnN0IGFkZEJveD0kKCdhZGRUZWFtTWVtYmVyc0xpc3QnKTthZGRCb3guaW5uZXJIVE1MPScnOwogICAgYWxsLmZpbHRlcih1PT4hZXhpc3RpbmcuaGFzKE51bWJlcih1LmlkKSkpLmZvckVhY2godT0+ewogICAgICBjb25zdCByb3c9ZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnbGFiZWwnKTtyb3cuY2xhc3NOYW1lPSdtZW1iZXItY2hlY2snO3Jvdy5pbm5lckhUTUw9YDxpbnB1dCB0eXBlPSJjaGVja2JveCIgdmFsdWU9IiR7TnVtYmVyKHUuaWQpfSI+PHNwYW4+JHtlc2ModS5uYW1lfHx1LmVtYWlsKX08YnI+PHNwYW4gY2xhc3M9Im11dGVkIj4ke2VzYyh1LmVtYWlsfHwnJyl9IMK3ICR7ZXNjKHUuc2VjdG9yfHwnJyl9PC9zcGFuPjwvc3Bhbj5gO2FkZEJveC5hcHBlbmRDaGlsZChyb3cpOwogICAgfSk7CiAgICBpZighYWRkQm94LmNoaWxkcmVuLmxlbmd0aClhZGRCb3guaW5uZXJIVE1MPSc8ZGl2IGNsYXNzPSJlbXB0eS1saXN0Ij5BbGwgYXBwcm92ZWQgZW1wbG95ZWVzIGFyZSBhbHJlYWR5IG1lbWJlcnMuPC9kaXY+JzsKICB9Y2F0Y2goZXJyKXskKCdjdXJyZW50VGVhbU1lbWJlcnMnKS5pbm5lckhUTUw9YDxkaXYgY2xhc3M9ImRhbmdlci1ub3RlIj4ke2VzYyhlcnIubWVzc2FnZSl9PC9kaXY+YDt9Cn0KYXN5bmMgZnVuY3Rpb24gYWRkU2VsZWN0ZWRUZWFtTWVtYmVycygpewogICQoJ3RlYW1FcnJvcicpLnRleHRDb250ZW50PScnO2lmKCFjdXJyZW50VGVhbUlkKXskKCd0ZWFtRXJyb3InKS50ZXh0Q29udGVudD0nU2VsZWN0IGEgdGVhbSBmaXJzdC4nO3JldHVybjt9CiAgY29uc3QgaWRzPUFycmF5LmZyb20oZG9jdW1lbnQucXVlcnlTZWxlY3RvckFsbCgnI2FkZFRlYW1NZW1iZXJzTGlzdCBpbnB1dFt0eXBlPWNoZWNrYm94XTpjaGVja2VkJykpLm1hcCh4PT5OdW1iZXIoeC52YWx1ZSkpOwogIGlmKCFpZHMubGVuZ3RoKXskKCd0ZWFtRXJyb3InKS50ZXh0Q29udGVudD0nU2VsZWN0IGF0IGxlYXN0IG9uZSBlbXBsb3llZS4nO3JldHVybjt9CiAgdHJ5ewogICAgZm9yKGNvbnN0IHVzZXJfaWQgb2YgaWRzKXthd2FpdCBhcGkoYC9hcGkvdGVhbXMvJHtOdW1iZXIoY3VycmVudFRlYW1JZCl9L21lbWJlcnMvYWRkYCx7bWV0aG9kOidQT1NUJyxoZWFkZXJzOnsnQ29udGVudC1UeXBlJzonYXBwbGljYXRpb24vanNvbid9LGJvZHk6SlNPTi5zdHJpbmdpZnkoe3VzZXJfaWR9KX0pO30KICAgIGF3YWl0IHJlZnJlc2hUZWFtcygpO2F3YWl0IGxvYWRDdXJyZW50VGVhbU1hbmFnZXIoY3VycmVudFRlYW1JZCk7YXdhaXQgb3BlblRlYW0oe2lkOmN1cnJlbnRUZWFtSWR9KTsKICB9Y2F0Y2goZXJyKXskKCd0ZWFtRXJyb3InKS50ZXh0Q29udGVudD1lcnIubWVzc2FnZTt9Cn0KYXN5bmMgZnVuY3Rpb24gb3BlblNlbGVjdGVkVGVhbUNoYXQoKXsKICBpZighY3VycmVudFRlYW1JZCl7JCgndGVhbUVycm9yJykudGV4dENvbnRlbnQ9J1NlbGVjdCBhIHRlYW0gZmlyc3QuJztyZXR1cm47fQogIGNvbnN0IHRlYW1zPShhd2FpdCBhcGkoJy9hcGkvdGVhbXMnKSkudGVhbXN8fFtdO2NvbnN0IHRlYW09dGVhbXMuZmluZCh0PT5OdW1iZXIodC5pZCk9PT1OdW1iZXIoY3VycmVudFRlYW1JZCkpOwogIGlmKHRlYW0pe2Nsb3NlVGVhbU1hbmFnZXIoKTthd2FpdCBvcGVuVGVhbSh0ZWFtKTt9ZWxzZSAkKCd0ZWFtRXJyb3InKS50ZXh0Q29udGVudD0nVGVhbSBub3QgZm91bmQuJzsKfQp3aW5kb3cuYWRkRXZlbnRMaXN0ZW5lcignbG9hZCcsYXN5bmMoKT0+ewogIGFwcGx5VGhlbWUoKTsKICB0cnl7CiAgICBjb25zdCBkYXRhPWF3YWl0IGZldGNoKCcvYXBpL21lJykudGhlbihyPT5yLmpzb24oKSk7CiAgICBpZihkYXRhLmF1dGhlbnRpY2F0ZWQpe2N1cnJlbnRVc2VyPWRhdGEudXNlcjskKCdsb2dpbicpLmNsYXNzTGlzdC5hZGQoJ2hpZGRlbicpOyQoJ3NwbGFzaCcpLmNsYXNzTGlzdC5hZGQoJ2hpZGRlbicpOyQoJ2FwcCcpLmNsYXNzTGlzdC5yZW1vdmUoJ2hpZGRlbicpO2F3YWl0IGluaXRpYWxpemVXb3JrYmVuY2goKTt9CiAgfWNhdGNoKF8pe30KfSk7Cjwvc2NyaXB0Pgo8L2JvZHk+CjwvaHRtbD4='

@app.route("/")
def home():
    # Always serve the exact original Orion UI. Restore it if the project-root
    # copy was accidentally deleted. No Jinja rendering is used.
    index_path = BASE_DIR / "index.html"
    if not index_path.exists():
        try:
            index_path.write_bytes(base64.b64decode(_ORIGINAL_INDEX_B64))
        except Exception:
            pass
    if index_path.exists():
        return send_file(index_path)
    return "Orion UI file could not be restored.", 500


@app.post("/api/register/admin")
def register_admin():
    data=request.get_json(silent=True) or {}
    try: uid=_store_create_admin(data.get("name"),data.get("email"),data.get("password"),data.get("company"))
    except ValueError as exc: return jsonify({"status":"error","message":str(exc)}),400
    except sqlite3.IntegrityError: return jsonify({"status":"error","message":"An account with this email already exists."}),409
    except Exception: logger.exception("Administrator registration failed"); return jsonify({"status":"error","message":"Unable to register administrator."}),500
    return jsonify({"status":"success","message":"Administrator registered. You can sign in now.","user_id":uid})

@app.post("/api/register/employee")
def register_employee():
    data=request.get_json(silent=True) or {}
    try:
        uid=_store_create_employee(data.get("name"),data.get("email"),data.get("password"),data.get("company"),data.get("sector"),data.get("role") or "Employee","pending")
        rid=create_registration_request(uid)
    except ValueError as exc: return jsonify({"status":"error","message":str(exc)}),400
    except sqlite3.IntegrityError: return jsonify({"status":"error","message":"An account with this email already exists."}),409
    except Exception: logger.exception("Employee registration failed"); return jsonify({"status":"error","message":"Unable to register employee."}),500
    return jsonify({"status":"success","message":"Registration submitted. An administrator must approve your account before sign-in.","user_id":uid,"request_id":rid,"pending":True})

@app.post("/api/setup")
def setup_first_admin():
    if _count_users() > 0:
        return jsonify({"status": "error", "message": "Initial setup is already complete."}), 409
    data = request.get_json(silent=True) or {}
    username, password = str(data.get("username", "")).strip(), str(data.get("password", ""))
    company = str(data.get("company", "")).strip()
    sector = str(data.get("sector", "")).strip()
    if len(username) < 3 or len(password) < 8:
        return jsonify({"status": "error", "message": "Username must be 3+ characters and password 8+ characters."}), 400
    user = _create_user_compat(username, password, company, sector, role="Administrator", is_admin=True)
    session["user_id"] = user["id"]
    return jsonify({"status": "success", "user": _public_user(user)})


@app.post("/api/login")
def login():
    data = request.get_json(silent=True) or {}
    user = authenticate(str(data.get("email") or data.get("username") or ""), str(data.get("password", "")))
    if not user:
        return jsonify({"status": "error", "message": "Invalid credentials."}), 401
    session.clear()
    session["user_id"] = user["id"]
    return jsonify({"status": "success", "user": _public_user(user)})


@app.post("/api/logout")
def logout():
    session.clear()
    return jsonify({"status": "success"})


def _public_user(user: dict[str, Any]) -> dict[str, Any]:
    return {"id":user.get("id"),"name":user.get("name") or str(user.get("email") or "").split("@",1)[0],"email":user.get("email") or "","employee_id":user.get("employee_id") or "","company":user.get("company") or "","sector":user.get("sector") or "","role":user.get("role") or "Employee","is_admin":bool(user.get("is_admin")),"status":user.get("status") or "approved"}



@app.get("/api/me")
def me():
    user = current_user()
    if not user:
        return jsonify({"authenticated": False})
    return jsonify({"authenticated": True, "user": _public_user(user)})


@app.get("/api/status")
@app.get("/api/health")
def health():
    status = model_status()
    status["ok"] = status["local_only"] and status["ollama_reachable"]
    status["external_calls_allowed"] = False
    return jsonify(status)


@app.post("/api/chats/new")
@login_required
def new_chat():
    user = current_user()
    data = request.get_json(silent=True) or {}
    team_id = data.get("team_id")
    # Employees may create their normal private chats, but only an administrator
    # may create a chat attached to a team. This preserves the existing chat UI.
    if team_id:
        if not bool(user.get("is_admin")):
            return jsonify({"status": "error", "message": "Only an administrator can create a team chat."}), 403
        if not is_team_member(user["id"], team_id):
            return jsonify({"status": "error", "message": "You are not a member of that team."}), 403
    chat = create_chat(user["id"], str(data.get("title") or "New chat"), team_id)
    return jsonify({"status": "success", "chat": chat, "chat_id": chat.get("id") if isinstance(chat, dict) else chat})


@app.get("/api/chats")
@login_required
def chats():
    return jsonify({"status": "success", "chats": list_chats_for_user(current_user()["id"])})


@app.get("/api/library")
@login_required
def library():
    """Return all files the authenticated user is allowed to access.

    The current local_store schema is chat-scoped, so the library is built
    from the user's authorized chats instead of introducing a second store.
    """
    user = current_user()
    files_by_id = {}
    for chat in list_chats_for_user(user["id"]):
        chat_id = chat.get("id")
        if not chat_id:
            continue
        for record in list_files_for_chat(user["id"], chat_id):
            item = dict(record)
            item["chat_title"] = chat.get("title") or "New chat"
            files_by_id[str(item.get("id"))] = item
    files = list(files_by_id.values())
    files.sort(key=lambda item: float(item.get("created_at") or 0), reverse=True)
    return jsonify({"status": "success", "files": files})


@app.get("/api/admin/registration-requests")
@admin_required
def admin_registration_requests():
    user=current_user()
    return jsonify({"status":"success","requests":get_pending_requests(user["company"])})

# Keep the existing UI unchanged.  The frontend may use any of the
# approval paths below depending on which original Orion screen is active.
# All aliases execute the same backend approval logic.
@app.post("/api/admin/registration-requests/<int:request_id>")
@app.post("/api/admin/registration-request/<int:request_id>")
@app.post("/api/admin/approve-registration/<int:request_id>")
@app.post("/api/admin/registration/<int:request_id>/approve")
@app.post("/api/admin/registration-requests/<int:request_id>/decision")
@admin_required
def admin_registration_decision(request_id):
    user = current_user()
    data = request.get_json(silent=True) or {}

    # Accept both names so older/newer backend calls remain compatible.
    decision = str(
        data.get("decision")
        or data.get("action")
        or data.get("status")
        or ""
    ).strip().lower()

    # Some clients send approved/rejected instead of approve/reject.
    if decision == "approved":
        decision = "approve"
    elif decision in {"rejected", "deny", "denied"}:
        decision = "reject"

    try:
        result = approve_registration(request_id, user["id"], decision)
    except ValueError as exc:
        return jsonify({"status": "error", "message": str(exc)}), 400
    except Exception as exc:
        logger.exception("Registration decision failed for request %s", request_id)
        return jsonify({
            "status": "error",
            "message": f"Unable to process registration request: {exc}",
        }), 500

    if not result:
        return jsonify({
            "status": "error",
            "message": "Registration request not found, already processed, or organization mismatch.",
        }), 404

    return jsonify({
        "status": "success",
        "message": (
            "Employee registration approved."
            if decision == "approve"
            else "Employee registration rejected."
        ),
        "request_id": request_id,
        "user": _public_user(result),
    })


@app.get("/api/notifications")
@login_required
def notifications():
    """Return persisted admin/registration notifications plus existing chat activity."""
    user = current_user()
    events = []
    now = time.time()
    # Registration notifications are persisted locally so an administrator can
    # see the request even after refreshing/restarting Orion.
    try:
        conn=_store_conn()
        rows=conn.execute(
            "SELECT id,kind,title,message,related_user_id,created_at,read_at FROM notifications WHERE recipient_id=? ORDER BY created_at DESC LIMIT 50",
            (int(user["id"]),),
        ).fetchall()
        conn.close()
        for row in rows:
            events.append({
                "id": row["id"], "kind": row["kind"], "title": row["title"],
                "message": row["message"], "created_at": float(row["created_at"] or 0),
                "read": row["read_at"] is not None,
                "related_user_id": row["related_user_id"],
            })
    except Exception:
        logger.exception("Unable to load persisted notifications")
    for chat in list_chats_for_user(user["id"]):
        chat_id = chat.get("id")
        if not chat_id:
            continue
        title = chat.get("title") or "New chat"
        for record in list_files_for_chat(user["id"], chat_id):
            created = float(record.get("created_at") or 0)
            events.append({
                "title": "File uploaded",
                "message": f"{record.get('original_name') or 'File'} is available in {title}.",
                "created_at": created,
                "read": (now - created) > 300,
            })
        messages = list_messages(chat_id)
        if messages:
            last = messages[-1]
            created = float(last.get("created_at") or 0)
            if last.get("sender_type") == "assistant":
                events.append({
                    "title": "Chat updated",
                    "message": f"{title} received a new response.",
                    "created_at": created,
                    "read": (now - created) > 300,
                })
    events.sort(key=lambda item: item["created_at"], reverse=True)
    return jsonify({"status": "success", "notifications": events[:25]})


@app.get("/api/chats/<chat_id>")
@login_required
def chat_detail(chat_id: str):
    user = current_user()
    chat = _authorized_chat(user["id"], chat_id)
    if not chat:
        return jsonify({"status": "error", "message": "Chat not found or unauthorized."}), 404
    return jsonify({"status": "success", "chat": chat, "messages": list_messages(chat_id),
                    "files": list_files_for_chat(user["id"], chat_id)})


@app.post("/api/chats/<chat_id>/activate")
@login_required
def activate_chat(chat_id: str):
    user = current_user()
    if not user_can_access_chat(user["id"], chat_id):
        return jsonify({"status": "error", "message": "Chat not found or unauthorized."}), 404
    session["active_chat_id"] = chat_id
    return jsonify({"status": "success", "chat_id": chat_id})


@app.get("/api/files")
@login_required
def files_list():
    return jsonify({"status": "success", "files": _store_list_files(current_user()["id"]) or []})


@app.get("/api/files/<file_id>/download")
@login_required
def download_file(file_id: str):
    user = current_user()
    if not user_can_access_file(user["id"], file_id):
        return jsonify({"status": "error", "message": "File not found or unauthorized."}), 404
    record = get_file(file_id)
    if not record:
        return jsonify({"status": "error", "message": "File not found."}), 404
    path = Path(record["stored_path"]).resolve()
    allowed_roots = [UPLOAD_ROOT.resolve(), GENERATED_ROOT.resolve()]
    if not any(str(path).startswith(str(root) + os.sep) or path == root for root in allowed_roots):
        return jsonify({"status": "error", "message": "Invalid stored file path."}), 403
    if not path.is_file():
        return jsonify({"status": "error", "message": "Stored file is missing."}), 404
    return send_file(
        path,
        as_attachment=request.args.get("inline") != "1",
        download_name=record["original_name"],
    )


# ---------------------------------------------------------------------------
# Teams and admin
# ---------------------------------------------------------------------------
@app.get("/api/teams")
@login_required
def teams():
    return jsonify({"status": "success", "teams": list_teams(current_user()["id"])})


@app.post("/api/teams/create")
@admin_required
def teams_create():
    user = current_user()
    data = request.get_json(silent=True) or {}
    name = str(data.get("name", "")).strip()
    if not name:
        return jsonify({"status": "error", "message": "Team name is required."}), 400

    raw_member_ids = data.get("member_ids") or data.get("members") or []
    if not isinstance(raw_member_ids, list):
        raw_member_ids = [raw_member_ids]
    member_ids = []
    for raw_id in raw_member_ids:
        try:
            mid = int(raw_id)
        except (TypeError, ValueError):
            continue
        if mid != int(user["id"]) and mid not in member_ids:
            member_ids.append(mid)

    # Only approved employees from the administrator's company can be added.
    valid_member_ids = []
    for mid in member_ids:
        try:
            member = get_user(mid)
        except Exception:
            member = None
        if member and str(member.get("company")) == str(user.get("company")) and not bool(member.get("is_admin")) and str(member.get("status") or "").lower() == "approved":
            valid_member_ids.append(mid)

    team_id, invite_code = create_team(name, user["company"], user["id"])
    for mid in valid_member_ids:
        try:
            add_team_member(team_id, user["id"], mid)
        except Exception:
            logger.exception("Unable to add employee %s to team %s", mid, team_id)

    team = get_team(team_id) or {"id": team_id, "name": name, "company": user["company"], "owner_id": user["id"], "invite_code": invite_code}
    return jsonify({"status": "success", "team": team, "team_id": team_id, "invite_code": invite_code,
                    "member_ids": valid_member_ids})


@app.post("/api/teams/join")
@login_required
def teams_join():
    user = current_user()
    data = request.get_json(silent=True) or {}
    code = str(data.get("invite_code") or data.get("code") or "").strip()
    if not code:
        return jsonify({"status": "error", "message": "Invite code is required."}), 400
    team = join_team_by_code(code, user["id"])
    if not team:
        return jsonify({"status": "error", "message": "Invalid invite code or organization mismatch."}), 403
    return jsonify({"status": "success", "team": team})


@app.get("/api/teams/<team_id>/members")
@login_required
def team_members_list_api(team_id: str):
    user = current_user()
    if not is_team_member(user["id"], team_id):
        return jsonify({"status": "error", "message": "You are not a member of this team."}), 403
    try:
        members = team_members(int(team_id), int(user["id"])) or []
    except Exception:
        members = []
    cleaned = []
    for member in members:
        item = dict(member)
        item["name"] = str(item.get("name") or item.get("email") or item.get("employee_id") or "Employee").strip()
        item["email"] = str(item.get("email") or "")
        item["employee_id"] = str(item.get("employee_id") or "")
        cleaned.append(item)
    return jsonify({"status": "success", "members": cleaned})


@app.post("/api/teams/<team_id>/members")
@app.post("/api/teams/<team_id>/members/add")
@admin_required
def team_member_add(team_id: str):
    user = current_user()
    if not is_team_member(user["id"], team_id):
        return jsonify({"status": "error", "message": "You are not a member of this team."}), 403
    data = request.get_json(silent=True) or {}
    member_id = str(data.get("user_id") or "").strip()
    if not member_id.isdigit():
        return jsonify({"status": "error", "message": "A valid employee ID is required."}), 400
    member = get_user(int(member_id))
    if not member or str(member.get("company")) != str(user.get("company")):
        return jsonify({"status": "error", "message": "Employee is not in your organization."}), 403
    if bool(member.get("is_admin")) or str(member.get("status") or "").lower() != "approved":
        return jsonify({"status": "error", "message": "Only approved employees can be added to a team."}), 403
    if not add_team_member(team_id, user["id"], int(member_id)):
        return jsonify({"status": "error", "message": "Membership request denied."}), 403
    return jsonify({"status": "success", "member": _public_user(member)})


@app.get("/api/users")
@login_required
def users():
    user = current_user()
    # The existing UI uses this endpoint while the administrator creates a
    # team. Return real approved employee names/emails from the users table.
    try:
        approved = list_users(user["company"], approved_only=True, employees_only=True)
    except TypeError:
        approved = [u for u in (list_users(user["company"]) or []) if not bool(u.get("is_admin")) and str(u.get("status") or "approved").lower() == "approved"]
    cleaned = []
    for row in approved or []:
        item = dict(row)
        item["name"] = str(item.get("name") or item.get("email") or item.get("employee_id") or "Employee").strip()
        item["email"] = str(item.get("email") or "")
        item["employee_id"] = str(item.get("employee_id") or "")
        item["status"] = str(item.get("status") or "approved")
        cleaned.append(item)
    return jsonify({"status": "success", "users": cleaned})


@app.post("/api/users")
@admin_required
def create_user_api():
    actor = current_user()
    data = request.get_json(silent=True) or {}
    try:
        user = _create_user_compat(str(data.get("username") or data.get("email") or ""), str(data.get("password", "")), actor["company"], str(data.get("sector", "")), str(data.get("role", "User")), bool(data.get("is_admin", False)))
    except Exception as exc:
        return jsonify({"status": "error", "message": "Unable to create user."}), 400
    return jsonify({"status": "success", "user": _public_user(user)})


@app.post("/api/sop-upload")
@admin_required
def sop_upload():
    user = current_user()
    uploaded = request.files.get("file")
    if not uploaded or not uploaded.filename:
        return jsonify({"status": "error", "message": "SOP file is required."}), 400
    if not allowed_file(uploaded.filename):
        return jsonify({"status": "error", "message": "Unsupported SOP file type."}), 400
    # SOPs inherit authenticated company/sector. The frontend cannot override it.
    target = KB_ROOT / user["company"] / (secure_filename(uploaded.filename) or "sop")
    target.parent.mkdir(parents=True, exist_ok=True)
    uploaded.save(target)
    return jsonify({
        "status": "success",
        "filename": uploaded.filename,
        "company": user["company"],
        "sector": user["sector"],
    })


# ---------------------------------------------------------------------------
# Text-only chat endpoint
# ---------------------------------------------------------------------------
@app.post("/api/chat")
@login_required
def chat_endpoint():
    user=current_user(); data=request.get_json(silent=True) or {}; query=str(data.get("query") or "").strip()
    if not query:return jsonify({"status":"error","message":"Query is required."}),400
    raw=data.get("chat_id") or session.get("active_chat_id")
    if raw:
        try: chat_id=int(raw)
        except (TypeError,ValueError):return jsonify({"status":"error","message":"Invalid chat ID."}),400
        chat=_authorized_chat(user["id"],str(chat_id))
        if not chat:return jsonify({"status":"error","message":"Chat not found or unauthorized."}),403
    else:
        chat=create_chat(user["id"],query[:70] or "New Chat"); chat_id=chat["id"]; session["active_chat_id"]=chat_id
    files=list_files_for_chat(user["id"],chat_id); parts=[]; images=[]
    for rec in files:
        path=rec.get("stored_path")
        if not path or not Path(path).is_file():continue
        try:
            text,isv,img=extract_document_text(path,rec.get("original_name") or Path(path).name,query=query); parts.append(f"[Source: {rec.get('original_name')}]\n{text}")
            if isv and img:images.append(img)
        except Exception as exc:parts.append(f"[Source unavailable: {rec.get('original_name')}] {exc}")
    context=retrieve_relevant_passages("\n\n".join(parts),query) if parts else "No file is attached. Answer the general question normally using the local model."
    add_message(chat_id,user["id"],"user",query,[])
    task=classify_task(query); selected_model,model_err=resolve_model_for_task("vision" if images else task,vision=bool(images))
    if not selected_model:
        report=f"I could not produce a local answer. {model_err}"; add_message(chat_id,user["id"],"assistant",report,{"model_error":model_err}); return jsonify({"status":"error","message":report,"chat_id":chat_id}),503
    prompt=f"You are Orion, a precise local enterprise assistant. Answer the user's question. When files are present, use only the supplied local context and do not invent facts.\n\nQuestion: {query}\n\nContext:\n{context[:MAX_CONTEXT_CHARS]}"
    report,err,used_model=query_ollama(prompt,images=images or None,vision=bool(images),task_type=task,model=selected_model,timeout=OLLAMA_TIMEOUT)
    report=report or f"I could not produce a reliable local answer. {err or 'Unknown local model error.'}"
    add_message(chat_id,user["id"],"assistant",report,{"model_used":used_model,"task_type":task}); touch_chat(chat_id,title=query[:80])
    return jsonify({"status":"success" if used_model else "error","chat_id":chat_id,"report":report,"model_used":used_model,"task_type":task,"file_ids":[f.get("id") for f in files]})


# ---------------------------------------------------------------------------
# Analyze uploaded sources and persist both sides of every turn
# ---------------------------------------------------------------------------
@app.post("/api/analyze-file")
@app.post("/api/analyze-pdf")
@app.post("/api/analyze")
@app.post("/api/analyze-documents")
@login_required
def analyze_file():
    user = current_user()
    data = request.get_json(silent=True) or {}
    chat_id = (request.form.get("chat_id") if request.form else None) or data.get("chat_id") or session.get("active_chat_id")
    chat = _authorized_chat(user["id"], str(chat_id) if chat_id else None)
    if not chat:
        chat = create_chat(user["id"], "New chat")
        chat_id = chat["id"]
        session["active_chat_id"] = chat_id

    query = ((request.form.get("query") if request.form else None) or data.get("query") or "").strip()
    # Client-provided sector/company are deliberately ignored for authorization
    # and routing. Trusted identity comes from the authenticated user record.
    # Accept both names so older and newer frontends remain compatible.
    uploads = []
    if request.files:
        uploads = request.files.getlist("files") or request.files.getlist("file")
    if not query and not uploads:
        return jsonify({"status": "error", "message": "A question or uploaded file is required."}), 400
    if not query:
        query = "Summarize the supplied sources"

    source_records = []
    combined_parts = []
    vision_images = []
    for upload in uploads:
        original = upload.filename or ""
        if not allowed_file(original):
            return jsonify({"status": "error", "message": f"Unsupported file type: {original}"}), 400
        target = _safe_upload_path(user["id"], chat_id, original)
        upload.save(target)
        size = target.stat().st_size
        if size > MAX_UPLOAD_BYTES:
            target.unlink(missing_ok=True)
            return jsonify({"status": "error", "message": f"File exceeds the configured {MAX_UPLOAD_BYTES // 1024 // 1024} MB limit."}), 413
        file_id = save_file(user["id"],chat_id,chat.get("team_id"),original,str(target),user.get("sector") or "General")
        rec = get_file(file_id)
        try:
            text, is_vision, image_b64 = extract_document_text(str(target), original, query=query)
        except Exception as exc:
            logger.exception("File parsing failed")
            return jsonify({"status": "error", "message": f"File parsing failed: {exc}"}), 400
        source_label = f"[Source: {original}]"
        combined_parts.append(f"{source_label}\n{text}")
        source_records.append({"id": file_id, "filename": original, "path": str(target), "vision": is_vision})
        if is_vision and image_b64:
            vision_images.append(image_b64)

    combined = "\n\n".join(combined_parts)
    chart_needed = detect_chart_need(query)
    diagram_needed = detect_diagram_need(query)
    task = classify_task(query)

    # Deterministic data path. Never return before persisting the user turn.
    user_message = add_message(chat_id, user["id"], "user", query, {
        "sources": [x["filename"] for x in source_records],
        "file_ids": [x["id"] for x in source_records],
        "company": user["company"],
    })

    data_records = [r for r in source_records if Path(r["path"]).suffix.lower() in {".csv", ".tsv", ".xlsx", ".xls", ".ods"}]
    direct_answer = None
    direct_spec = None
    if data_records and not vision_images:
        direct_answer, direct_spec = answer_from_dataframe(Path(data_records[0]["path"]), query)
        if direct_answer:
            metadata = {
                "model_used": "deterministic-pandas",
                "task_type": "data",
                "sources": [r["filename"] for r in data_records],
                "spec": direct_spec,
            }
            add_message(chat_id, user["id"], "assistant", direct_answer, metadata)
            touch_chat(chat_id, title=query[:80])
            return jsonify({
                "status": "success", "chat_id": chat_id,
                "active_file": ", ".join(r["filename"] for r in source_records),
                "report": direct_answer, "chart_required": chart_needed,
                "diagram_required": diagram_needed, "task_type": "data",
                "model_used": "deterministic-pandas", "file_ids": [r["id"] for r in source_records],
                "kb_sources_used": [],
            })

    # Source-aware local retrieval. Broad requests such as "analyse the file"
    # use distributed whole-document evidence rather than treating the phrase as
    # a literal keyword query. Targeted questions use relevance retrieval.
    kb_snippets = []
    if not vision_images:
        try:
            kb_snippets = search_kb(query, top_k=3, min_score=0.15, company=user["company"])
        except Exception:
            kb_snippets = []

    whole_document = bool(source_records) and is_whole_document_analysis_request(query)
    if whole_document and not vision_images:
        context = build_whole_document_context(combined, query)
    else:
        context = retrieve_relevant_passages(combined, query, top_k=MAX_RETRIEVAL_CHUNKS)
    if kb_snippets and not whole_document:
        context += "\n\n" + "\n---\n".join(
            f"[Local KB source: {src}]\n{chunk}" for chunk, src, _ in kb_snippets
        )
    context = context[:(12000 if whole_document else MAX_CONTEXT_CHARS)]

    selected_model, model_err = resolve_model_for_task(
        "vision" if vision_images else task,
        vision=bool(vision_images),
    )
    if not selected_model:
        report = (
            "I could not produce a model-based answer because the required local "
            f"capability is unavailable. {model_err}"
        )
        add_message(chat_id, user["id"], "assistant", report, {"model_error": model_err})
        return jsonify({"status": "error", "message": report, "chat_id": chat_id}), 503

    if vision_images:
        prompt = (
            "Answer only from the visible uploaded images and their local transcript. "
            "Do not invent labels, numbers or connections. Clearly separate "
            "'Explicitly shown:' from 'Inference:'. Mark unreadable content. "
            "Preserve exact visible text where possible.\n\n"
            f"Question: {query}\nLocal transcript:\n{context}"
        )
        report, err, used_model = query_ollama(
            prompt, images=vision_images, vision=True, task_type="vision",
            model=selected_model, timeout=OLLAMA_TIMEOUT,
        )
    else:
        if whole_document:
            prompt = (
                "You are Orion, an evidence-only local enterprise document analyst. "
                "The user asked for a WHOLE-DOCUMENT ANALYSIS. Do not search for the "
                "literal words in the user's request. Instead, synthesize the uploaded "
                "document as a report. The evidence below is deliberately sampled "
                "across the document and includes query-relevant passages. Cover the "
                "document's purpose, major sections/topics, key facts and figures, "
                "financial or operational performance when present, governance/people "
                "when present, important dates/events/resolutions, risks/issues, and "
                "a concise set of key takeaways. Preserve page/source markers when "
                "available. Do not invent facts. Clearly label an inference as an "
                "inference. If a requested category is not supported by the evidence, "
                "say that it was not established rather than guessing.\n\n"
                f"Document: {', '.join(r['filename'] for r in source_records)}\n"
                f"User request: {query}\n\nDistributed local evidence:\n{context}"
            )
        else:
            prompt = (
                "You are Orion, an evidence-only local enterprise document analyst. "
                "Answer the exact question using ONLY the supplied retrieved excerpts "
                "from the uploaded source and optional local KB excerpts. The uploaded "
                "document may be hundreds of pages; the excerpts are the locally "
                "retrieved evidence, not the whole document. Never invent a number, "
                "entity, date, page, slide, sheet, citation or relationship. Preserve "
                "page/source markers. If the excerpts do not contain enough evidence, "
                "say so clearly instead of guessing. Separate observations from "
                "inferences.\n\n"
                f"Question: {query}\n\nSources:\n{context}"
            )
        report, err, used_model = query_ollama(
            prompt, task_type=task, model=selected_model, timeout=OLLAMA_TIMEOUT,
        )

        # Retry with less context only when the local model actually fails.
        if not report and err:
            if whole_document:
                compact_context = build_whole_document_context(combined, query)[:6000]
            else:
                compact_context = retrieve_relevant_passages(
                    combined, query, top_k=max(3, MAX_RETRIEVAL_CHUNKS // 2)
                )[:4500]
            retry_prompt = (
                ("Produce a concise whole-document analysis using ONLY this distributed "
                 "local evidence. Cover purpose, major findings, key facts, important "
                 "dates/events and takeaways. Do not guess.\n\n" if whole_document else
                 "Answer using ONLY these retrieved local document excerpts. Be concise "
                 "and evidence-based. Do not guess.\n\n")
                + f"Question: {query}\n\nEvidence:\n{compact_context}"
            )
            report, err, used_model = query_ollama(
                retry_prompt, task_type=task, model=selected_model,
                timeout=min(OLLAMA_RETRY_TIMEOUT, OLLAMA_TIMEOUT),
            )

    if not report:
        if whole_document and source_records and not vision_images:
            report = build_local_document_overview(combined, ", ".join(r["filename"] for r in source_records))
            report += "\n\n[Local model note: the document was processed locally; the final generative synthesis was unavailable."
            if err:
                report += f" {err}]"
            else:
                report += "]"
        else:
            report = (
                "The file was processed locally, but the local model did not finish "
                f"the final reasoning step. {err or 'Unknown local model error.'}"
            )

    metadata = {
        "model_used": used_model,
        "task_type": task,
        "sources": [r["filename"] for r in source_records],
        "kb_sources": [s for _, s, _ in kb_snippets],
    }
    add_message(chat_id, user["id"], "assistant", report, metadata)
    touch_chat(chat_id, title=query[:80])

    return jsonify({
        "status": "success" if used_model else "error",
        "chat_id": chat_id,
        "active_file": ", ".join(r["filename"] for r in source_records),
        "file_freshly_uploaded_this_turn": True,
        "report": report,
        "chart_required": chart_needed,
        "diagram_required": diagram_needed,
        "task_type": task,
        "model_used": used_model,
        "file_ids": [r["id"] for r in source_records],
        "sector": user["sector"],
        "kb_sources_used": [s for _, s, _ in kb_snippets],
    })


# ---------------------------------------------------------------------------
# Artifact helpers
# ---------------------------------------------------------------------------
def _artifact_path(user_id: str, chat_id: str, suffix: str) -> Path:
    folder = GENERATED_ROOT / str(user_id) / str(chat_id)
    folder.mkdir(parents=True, exist_ok=True)
    return folder / f"{uuid.uuid4().hex}{suffix}"


def _register_artifact(user: dict, chat_id: str, path: Path, filename: str, mime: str) -> str:
    fid = save_file(user["id"],chat_id,(get_chat(chat_id) or {}).get("team_id"),filename,str(path),user.get("sector") or "General")
    return fid


@app.post("/api/generate-chart")
@app.post("/api/chart")
@app.post("/api/generate-graph")
@app.post("/api/graph")
@login_required
def generate_chart():
    user = current_user()
    data = request.get_json(silent=True) or {}
    chat_id = data.get("chat_id") or session.get("active_chat_id")
    chat = _authorized_chat(user["id"], chat_id)
    if not chat:
        return jsonify({"status": "error", "message": "Chat not found or unauthorized."}), 404
    files = list_files_for_chat(user["id"], chat_id)
    tabular = [f for f in files if Path(f["stored_path"]).suffix.lower() in {".csv", ".tsv", ".xlsx", ".xls", ".ods"}]
    if not tabular:
        return jsonify({"status": "error", "message": "No tabular source is attached to this chat."}), 400
    path = Path(tabular[0]["stored_path"])
    try:
        df = _load_dataframe(path, str(data.get("query", "")))
        spec = _resolve_chart_spec(df, str(data.get("query", "")))
        if not spec or spec.get("chart_type") == "none":
            return jsonify({"status": "error", "message": "The request could not be mapped to a valid chart using the actual columns."}), 400
        out = _artifact_path(user["id"], chat_id, ".png")
        render_chart(df, spec, out)
        fid = _register_artifact(user, chat_id, out, "chart.png", "image/png")
    except Exception as exc:
        logger.exception("Chart generation failed")
        return jsonify({"status": "error", "message": f"Chart generation failed: {exc}"}), 400

    url = f"/api/files/{fid}/download"
    add_message(chat_id, user["id"], "assistant", "Chart generated from verified uploaded data.", {
        "artifact_file_id": fid, "artifact_type": "chart", "spec": spec,
    })
    return jsonify({"status": "success", "download_url": url, "image_url": f"{url}?inline=1", "filename": "orion-chart.png", "spec": spec})


@app.post("/api/generate-diagram")
@app.post("/api/diagram")
@app.post("/api/generate-workflow")
@app.post("/api/generate-architecture")
@login_required
def generate_diagram():
    user = current_user()
    data = request.get_json(silent=True) or {}
    chat_id = data.get("chat_id") or session.get("active_chat_id")
    chat = _authorized_chat(user["id"], chat_id)
    if not chat:
        return jsonify({"status": "error", "message": "Chat not found or unauthorized."}), 404
    query = str(data.get("query", "Create a diagram"))
    sources = list_files_for_chat(user["id"], chat_id)
    context_parts = []
    for f in sources:
        try:
            text, _, _ = extract_document_text(f["stored_path"], f["original_name"], query=query)
            context_parts.append(f"[Source: {f['original_name']}]\n{text[:5000]}")
        except Exception:
            continue
    model, err = resolve_model_for_task("diagram")
    if not model:
        return jsonify({"status": "error", "message": err}), 503
    raw, err, used_model = query_ollama(
        build_diagram_prompt(query, "\n\n".join(context_parts)),
        task_type="diagram", model=model, timeout=OLLAMA_TIMEOUT,
    )
    graph = _parse_json_object(raw)
    if not graph or not validate_graph(graph):
        return jsonify({"status": "error", "message": "The local model did not return a valid diagram structure."}), 502
    try:
        out = _artifact_path(user["id"], chat_id, ".png")
        render_diagram(graph, out)
        fid = _register_artifact(user, chat_id, out, "diagram.png", "image/png")
    except Exception as exc:
        return jsonify({"status": "error", "message": f"Diagram rendering failed: {exc}"}), 500
    return jsonify({"status": "success", "download_url": f"/api/files/{fid}/download",
                    "model_used": used_model, "graph": graph})


@app.post("/api/generate-doc")
@app.post("/api/generate-deliverable")
@login_required
def generate_deliverable():
    user = current_user()
    data = request.get_json(silent=True) or {}
    chat_id = data.get("chat_id") or session.get("active_chat_id")
    chat = _authorized_chat(user["id"], chat_id)
    if not chat:
        return jsonify({"status": "error", "message": "Chat not found or unauthorized."}), 404
    file_type = str(data.get("type", "docx")).lower()
    title = str(data.get("title", "Sovereign AI Deliverable"))
    content = str(data.get("content", ""))
    notice = "CONFIDENTIAL — processed within the organization's controlled infrastructure."

    if file_type == "docx":
        if docx is None:
            return jsonify({"status": "error", "message": "python-docx is not installed."}), 500
        d = docx.Document()
        d.add_heading(title, level=1)
        d.add_paragraph(notice)
        d.add_paragraph(content)
        path = _artifact_path(user["id"], chat_id, ".docx")
        d.save(path)
        filename, mime = "Deliverable.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    elif file_type == "pptx":
        if Presentation is None:
            return jsonify({"status": "error", "message": "python-pptx is not installed."}), 500
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = title
        slide.placeholders[1].text = f"{notice}\n\n{content[:3000]}"
        path = _artifact_path(user["id"], chat_id, ".pptx")
        prs.save(path)
        filename, mime = "Deliverable.pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    elif file_type == "xlsx":
        if openpyxl is None:
            return jsonify({"status": "error", "message": "openpyxl is not installed."}), 500
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Deliverable"
        ws.append(["Title", title])
        ws.append(["Confidentiality", notice])
        ws.append(["Content", content[:30000]])
        path = _artifact_path(user["id"], chat_id, ".xlsx")
        wb.save(path)
        filename, mime = "Deliverable.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    elif file_type in {"txt", "md"}:
        path = _artifact_path(user["id"], chat_id, "." + file_type)
        path.write_text(f"# {title}\n\n{notice}\n\n{content}", encoding="utf-8")
        filename, mime = f"Deliverable.{file_type}", "text/plain"
    else:
        return jsonify({"status": "error", "message": "Supported deliverables: DOCX, PPTX, XLSX, TXT, MD."}), 400

    fid = _register_artifact(user, chat_id, path, filename, mime)
    add_message(chat_id, user["id"], "assistant", f"Generated {filename}.", {
        "artifact_file_id": fid, "artifact_type": file_type,
    })
    return jsonify({"status": "success", "download_url": f"/api/files/{fid}/download", "filename": filename})


@app.errorhandler(413)
def too_large(_):
    return jsonify({"status": "error", "message": "Uploaded file is larger than the configured limit."}), 413


@app.errorhandler(404)
def not_found(_):
    return jsonify({"status": "error", "message": "Endpoint not found."}), 404


@app.errorhandler(500)
def internal_error(_):
    logger.exception("Unhandled server error")
    return jsonify({"status": "error", "message": "Internal server error. Check the local server log."}), 500


if __name__ == "__main__":
    logger.info("Starting Sovereign local workbench")
    logger.info("Ollama endpoint: %s", OLLAMA_BASE)
    logger.info("Discovered models: %s", [m["name"] for m in get_model_registry()])
    app.run(host=os.getenv("SOVEREIGN_HOST", "127.0.0.1"),
            port=int(os.getenv("SOVEREIGN_PORT", "8000")),
            debug=os.getenv("SOVEREIGN_DEBUG", "0") == "1",
            threaded=True, use_reloader=False)
