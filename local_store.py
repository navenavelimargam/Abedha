"""
Offline persistence for the Sovereign AI prototype.

All data is stored in a local SQLite database. No network service is used.

This is intentionally simple for the prototype; production should replace
password login with the organization's SSO/AD/LDAP and use encrypted storage.
"""

import hashlib
import os
import secrets
import sqlite3
import time


DB_PATH = os.getenv("SOVEREIGN_DB", "sovereign_local.db")


# ============================================================
# DATABASE CONNECTION
# ============================================================

def _conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


# ============================================================
# PASSWORD HASHING
# ============================================================

def _hash_password(password, salt=None):
    salt = salt or secrets.token_hex(16)

    digest = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode(),
        salt.encode(),
        120_000,
    ).hex()

    return salt, digest


# ============================================================
# INITIALIZE DATABASE
# ============================================================

def init_db():
    conn = _conn()

    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_id TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            salt TEXT NOT NULL,
            company TEXT NOT NULL,
            sector TEXT NOT NULL,
            role TEXT NOT NULL,
            is_admin INTEGER NOT NULL DEFAULT 0,
            created_at REAL NOT NULL
        );

        CREATE TABLE IF NOT EXISTS chats (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            owner_id INTEGER NOT NULL,
            team_id INTEGER,
            title TEXT NOT NULL,
            created_at REAL NOT NULL,
            updated_at REAL NOT NULL,
            FOREIGN KEY(owner_id)
                REFERENCES users(id)
                ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS messages (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            chat_id INTEGER NOT NULL,
            sender_id INTEGER,
            sender_type TEXT NOT NULL,
            content TEXT NOT NULL,
            files_json TEXT,
            created_at REAL NOT NULL,
            FOREIGN KEY(chat_id)
                REFERENCES chats(id)
                ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS teams (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company TEXT NOT NULL,
            name TEXT NOT NULL,
            owner_id INTEGER NOT NULL,
            invite_code TEXT UNIQUE NOT NULL,
            created_at REAL NOT NULL,
            FOREIGN KEY(owner_id)
                REFERENCES users(id)
                ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS team_members (
            team_id INTEGER NOT NULL,
            user_id INTEGER NOT NULL,
            role TEXT NOT NULL DEFAULT 'member',
            PRIMARY KEY(team_id, user_id),
            FOREIGN KEY(team_id)
                REFERENCES teams(id)
                ON DELETE CASCADE,
            FOREIGN KEY(user_id)
                REFERENCES users(id)
                ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS files (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            owner_id INTEGER NOT NULL,
            chat_id INTEGER,
            team_id INTEGER,
            original_name TEXT NOT NULL,
            stored_path TEXT NOT NULL,
            sector TEXT NOT NULL,
            created_at REAL NOT NULL,
            FOREIGN KEY(owner_id)
                REFERENCES users(id)
                ON DELETE CASCADE,
            FOREIGN KEY(chat_id)
                REFERENCES chats(id)
                ON DELETE SET NULL,
            FOREIGN KEY(team_id)
                REFERENCES teams(id)
                ON DELETE SET NULL
        );
        """
    )

    # ========================================================
    # DEMO ACCOUNTS
    # ========================================================
    # Change these before real deployment.

    seeds = [
        (
            "ADMIN001",
            "admin123",
            "TCS",
            "General",
            "Administrator",
            1,
        ),
        (
            "EMP001",
            "emp123",
            "TCS",
            "Engineering",
            "Engineer",
            0,
        ),
        (
            "EMP002",
            "emp123",
            "TCS",
            "Finance",
            "Analyst",
            0,
        ),
        (
            "EMP003",
            "emp123",
            "TCS",
            "IT",
            "Developer",
            0,
        ),
        (
            "EMP004",
            "emp123",
            "TCS",
            "Operations",
            "Operations Executive",
            0,
        ),
    ]

    for (
        employee_id,
        password,
        company,
        sector,
        role,
        is_admin,
    ) in seeds:

        existing = conn.execute(
            "SELECT 1 FROM users WHERE employee_id=?",
            (employee_id,),
        ).fetchone()

        if existing:
            continue

        salt, digest = _hash_password(password)

        conn.execute(
            """
            INSERT INTO users(
                employee_id,
                password_hash,
                salt,
                company,
                sector,
                role,
                is_admin,
                created_at
            )
            VALUES(?,?,?,?,?,?,?,?)
            """,
            (
                employee_id,
                digest,
                salt,
                company,
                sector,
                role,
                is_admin,
                time.time(),
            ),
        )

    conn.commit()
    conn.close()


# ============================================================
# AUTHENTICATION
# ============================================================

def authenticate(employee_id, password):
    conn = _conn()

    row = conn.execute(
        "SELECT * FROM users WHERE employee_id=?",
        (employee_id.strip(),),
    ).fetchone()

    conn.close()

    if not row:
        return None

    _, digest = _hash_password(
        password,
        row["salt"],
    )

    if secrets.compare_digest(
        digest,
        row["password_hash"],
    ):
        return dict(row)

    return None


# ============================================================
# USERS
# ============================================================

def get_user(user_id):
    conn = _conn()

    row = conn.execute(
        "SELECT * FROM users WHERE id=?",
        (user_id,),
    ).fetchone()

    conn.close()

    return dict(row) if row else None


def list_users(company, sector=None):
    conn = _conn()

    if sector and sector not in ("General", "All"):

        rows = conn.execute(
            """
            SELECT
                id,
                employee_id,
                company,
                sector,
                role,
                is_admin
            FROM users
            WHERE company=?
              AND sector=?
            ORDER BY employee_id
            """,
            (company, sector),
        ).fetchall()

    else:

        rows = conn.execute(
            """
            SELECT
                id,
                employee_id,
                company,
                sector,
                role,
                is_admin
            FROM users
            WHERE company=?
            ORDER BY sector, employee_id
            """,
            (company,),
        ).fetchall()

    conn.close()

    return [dict(row) for row in rows]


# ============================================================
# CHATS
# ============================================================

def create_chat(owner_id, title="New Chat", team_id=None):
    now = time.time()

    conn = _conn()

    cur = conn.execute(
        """
        INSERT INTO chats(
            owner_id,
            team_id,
            title,
            created_at,
            updated_at
        )
        VALUES(?,?,?,?,?)
        """,
        (
            owner_id,
            team_id,
            title,
            now,
            now,
        ),
    )

    chat_id = cur.lastrowid

    conn.commit()
    conn.close()

    return chat_id


def list_chats(owner_id):
    conn = _conn()

    rows = conn.execute(
        """
        SELECT DISTINCT
            c.id,
            c.team_id,
            c.title,
            c.created_at,
            c.updated_at
        FROM chats c
        LEFT JOIN team_members tm
            ON tm.team_id=c.team_id
        WHERE c.owner_id=?
           OR tm.user_id=?
        ORDER BY c.updated_at DESC
        """,
        (
            owner_id,
            owner_id,
        ),
    ).fetchall()

    conn.close()

    return [dict(row) for row in rows]


def get_or_create_team_chat(team_id, owner_id, title):
    conn = _conn()

    allowed = conn.execute(
        """
        SELECT 1
        FROM team_members
        WHERE team_id=?
          AND user_id=?
        """,
        (
            team_id,
            owner_id,
        ),
    ).fetchone()

    if not allowed:
        conn.close()
        return None

    row = conn.execute(
        """
        SELECT id
        FROM chats
        WHERE team_id=?
        ORDER BY id
        LIMIT 1
        """,
        (team_id,),
    ).fetchone()

    if row:
        chat_id = row["id"]
        conn.close()
        return chat_id

    now = time.time()

    cur = conn.execute(
        """
        INSERT INTO chats(
            owner_id,
            team_id,
            title,
            created_at,
            updated_at
        )
        VALUES(?,?,?,?,?)
        """,
        (
            owner_id,
            team_id,
            title,
            now,
            now,
        ),
    )

    chat_id = cur.lastrowid

    conn.commit()
    conn.close()

    return chat_id


# ============================================================
# MESSAGES
# ============================================================

def add_message(chat_id, sender_id, sender_type, content, files_json='[]'):
    if not isinstance(files_json, str):
        import json
        files_json = json.dumps(files_json, ensure_ascii=False)

    now = time.time()
    conn = _conn()
    conn.execute(
        'INSERT INTO messages(chat_id,sender_id,sender_type,content,files_json,created_at) VALUES(?,?,?,?,?,?)',
        (chat_id, sender_id, sender_type, content, files_json or '[]', now)
    )
    conn.execute(
        'UPDATE chats SET updated_at=? WHERE id=?',
        (now, chat_id)
    )
    conn.commit()
    conn.close()

def get_messages(chat_id, owner_id):
    conn = _conn()

    allowed = conn.execute(
        """
        SELECT 1
        FROM chats c
        WHERE c.id=?
          AND (
                c.owner_id=?
                OR EXISTS(
                    SELECT 1
                    FROM team_members tm
                    WHERE tm.team_id=c.team_id
                      AND tm.user_id=?
                )
          )
        """,
        (
            chat_id,
            owner_id,
            owner_id,
        ),
    ).fetchone()

    if not allowed:
        conn.close()
        return None

    rows = conn.execute(
        """
        SELECT
            id,
            sender_id,
            sender_type,
            content,
            files_json,
            created_at
        FROM messages
        WHERE chat_id=?
        ORDER BY id
        """,
        (chat_id,),
    ).fetchall()

    conn.close()

    return [dict(row) for row in rows]


# ============================================================
# TEAMS
# ============================================================

def create_team(company, name, owner_id):
    code = (
        secrets.token_urlsafe(8)
        .replace("-", "")
        .replace("_", "")[:10]
        .upper()
    )

    conn = _conn()

    cur = conn.execute(
        """
        INSERT INTO teams(
            company,
            name,
            owner_id,
            invite_code,
            created_at
        )
        VALUES(?,?,?,?,?)
        """,
        (
            company,
            name,
            owner_id,
            code,
            time.time(),
        ),
    )

    team_id = cur.lastrowid

    conn.execute(
        """
        INSERT INTO team_members(
            team_id,
            user_id,
            role
        )
        VALUES(?,?,?)
        """,
        (
            team_id,
            owner_id,
            "owner",
        ),
    )

    conn.commit()
    conn.close()

    return team_id, code


def add_team_member(team_id, user_id):
    conn = _conn()

    conn.execute(
        """
        INSERT OR IGNORE INTO team_members(
            team_id,
            user_id
        )
        VALUES(?,?)
        """,
        (
            team_id,
            user_id,
        ),
    )

    conn.commit()
    conn.close()


def join_team_by_code(invite_code, user_id):
    conn = _conn()

    team = conn.execute(
        """
        SELECT *
        FROM teams
        WHERE invite_code=?
        """,
        (invite_code.strip().upper(),),
    ).fetchone()

    if not team:
        conn.close()
        return None

    user = conn.execute(
        """
        SELECT company
        FROM users
        WHERE id=?
        """,
        (user_id,),
    ).fetchone()

    if not user or user["company"] != team["company"]:
        conn.close()
        return None

    conn.execute(
        """
        INSERT OR IGNORE INTO team_members(
            team_id,
            user_id
        )
        VALUES(?,?)
        """,
        (
            team["id"],
            user_id,
        ),
    )

    conn.commit()
    conn.close()

    return dict(team)


def list_teams(user_id):
    conn = _conn()

    rows = conn.execute(
        """
        SELECT
            t.id,
            t.name,
            t.company,
            t.invite_code,
            t.owner_id,
            t.created_at
        FROM teams t
        JOIN team_members tm
            ON tm.team_id=t.id
        WHERE tm.user_id=?
        ORDER BY t.created_at DESC
        """,
        (user_id,),
    ).fetchall()

    conn.close()

    return [dict(row) for row in rows]


def team_members(team_id, requester_id):
    conn = _conn()

    allowed = conn.execute(
        """
        SELECT 1
        FROM team_members
        WHERE team_id=?
          AND user_id=?
        """,
        (
            team_id,
            requester_id,
        ),
    ).fetchone()

    if not allowed:
        conn.close()
        return None

    rows = conn.execute(
        """
        SELECT
            u.id,
            u.employee_id,
            u.company,
            u.sector,
            u.role,
            tm.role AS team_role
        FROM users u
        JOIN team_members tm
            ON tm.user_id=u.id
        WHERE tm.team_id=?
        ORDER BY u.employee_id
        """,
        (team_id,),
    ).fetchall()

    conn.close()

    return [dict(row) for row in rows]


# ============================================================
# FILES
# ============================================================

def save_file(
    owner_id,
    chat_id,
    team_id,
    original_name,
    stored_path,
    sector,
):
    conn = _conn()

    cur = conn.execute(
        """
        INSERT INTO files(
            owner_id,
            chat_id,
            team_id,
            original_name,
            stored_path,
            sector,
            created_at
        )
        VALUES(?,?,?,?,?,?,?)
        """,
        (
            owner_id,
            chat_id,
            team_id,
            original_name,
            stored_path,
            sector,
            time.time(),
        ),
    )

    conn.commit()

    file_id = cur.lastrowid

    conn.close()

    return file_id


# ============================================================
# COMPATIBILITY FUNCTION FOR app.py
# ============================================================

def add_file(
    owner_id,
    chat_id,
    team_id,
    original_name,
    stored_path,
    sector,
):
    """
    Compatibility wrapper.

    The database implementation is save_file(), while newer
    app.py code may call add_file().
    """

    return save_file(
        owner_id=owner_id,
        chat_id=chat_id,
        team_id=team_id,
        original_name=original_name,
        stored_path=stored_path,
        sector=sector,
    )


def list_files(owner_id):
    conn = _conn()

    rows = conn.execute(
        """
        SELECT
            id,
            chat_id,
            team_id,
            original_name,
            stored_path,
            sector,
            created_at
        FROM files
        WHERE owner_id=?
        ORDER BY created_at DESC
        """,
        (owner_id,),
    ).fetchall()

    conn.close()

    return [dict(row) for row in rows]


def get_chat_files(chat_id, owner_id):
    conn = _conn()

    allowed = conn.execute(
        """
        SELECT 1
        FROM chats c
        WHERE c.id=?
          AND (
                c.owner_id=?
                OR EXISTS(
                    SELECT 1
                    FROM team_members tm
                    WHERE tm.team_id=c.team_id
                      AND tm.user_id=?
                )
          )
        """,
        (
            chat_id,
            owner_id,
            owner_id,
        ),
    ).fetchone()

    if not allowed:
        conn.close()
        return None

    rows = conn.execute(
        """
        SELECT
            id,
            original_name,
            stored_path,
            sector,
            created_at
        FROM files
        WHERE chat_id=?
           OR team_id=(
                SELECT team_id
                FROM chats
                WHERE id=?
           )
        ORDER BY id
        """,
        (
            chat_id,
            chat_id,
        ),
    ).fetchall()

    conn.close()

    return [dict(row) for row in rows]