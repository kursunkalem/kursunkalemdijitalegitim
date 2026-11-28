import base64
import html
import io
import json
import os
import re
import time
import urllib.parse
import uuid
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import iyzipay
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from google import genai
from google.genai import types

import hashlib
import hmac
import secrets
import smtplib
import sqlite3
import ssl
import string
import threading
from email.message import EmailMessage
from decimal import Decimal
from types import SimpleNamespace
from typing import Iterable, Sequence
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_BREAK


# ==============================================================================
# VERİ KATMANI: kullanıcılar, sınıflar, ödevler, sonuçlar, ödemeler (SQLite)
# ==============================================================================
DB_PATH = Path(os.environ.get("KK_DB_PATH", "kursun_kalem.db"))
PG_URL: Optional[str] = os.environ.get("KK_DATABASE_URL") or None
ROLES = ("ogrenci", "veli", "ogretmen")
ROLE_LABELS = {"ogrenci": "Öğrenci", "veli": "Veli", "ogretmen": "Öğretmen"}

_lock = threading.Lock()

_SCHEMA_SQLITE = """
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    email TEXT NOT NULL UNIQUE,
    name TEXT NOT NULL,
    role TEXT NOT NULL,
    pw_hash TEXT NOT NULL,
    grade TEXT,
    link_code TEXT UNIQUE,
    parent_of INTEGER REFERENCES users(id),
    daily_goal INTEGER NOT NULL DEFAULT 20,
    created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS user_state (
    user_id INTEGER PRIMARY KEY REFERENCES users(id) ON DELETE CASCADE,
    payload TEXT NOT NULL,
    updated_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS activity_days (
    user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    day TEXT NOT NULL,
    solved INTEGER NOT NULL DEFAULT 0,
    correct INTEGER NOT NULL DEFAULT 0,
    wrong INTEGER NOT NULL DEFAULT 0,
    blank INTEGER NOT NULL DEFAULT 0,
    PRIMARY KEY (user_id, day)
);
CREATE TABLE IF NOT EXISTS exam_results (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    created_at TEXT NOT NULL,
    source TEXT NOT NULL,
    correct INTEGER NOT NULL,
    wrong INTEGER NOT NULL,
    blank INTEGER NOT NULL,
    accuracy REAL NOT NULL,
    assignment_id INTEGER
);
CREATE TABLE IF NOT EXISTS badges (
    user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    code TEXT NOT NULL,
    earned_at TEXT NOT NULL,
    PRIMARY KEY (user_id, code)
);
CREATE TABLE IF NOT EXISTS classes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    code TEXT NOT NULL UNIQUE,
    teacher_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    grade TEXT,
    created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS class_members (
    class_id INTEGER NOT NULL REFERENCES classes(id) ON DELETE CASCADE,
    user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    joined_at TEXT NOT NULL,
    PRIMARY KEY (class_id, user_id)
);
CREATE TABLE IF NOT EXISTS assignments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    class_id INTEGER NOT NULL REFERENCES classes(id) ON DELETE CASCADE,
    title TEXT NOT NULL,
    grade TEXT,
    branch TEXT,
    outcome TEXT,
    content TEXT NOT NULL,
    due_date TEXT,
    created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS submissions (
    assignment_id INTEGER NOT NULL REFERENCES assignments(id) ON DELETE CASCADE,
    user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    correct INTEGER NOT NULL,
    wrong INTEGER NOT NULL,
    blank INTEGER NOT NULL,
    net REAL NOT NULL,
    accuracy REAL NOT NULL,
    submitted_at TEXT NOT NULL,
    PRIMARY KEY (assignment_id, user_id)
);
CREATE TABLE IF NOT EXISTS payments (
    token TEXT PRIMARY KEY,
    user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    amount INTEGER NOT NULL,
    billing_cycle TEXT NOT NULL,
    plan_json TEXT NOT NULL,
    status TEXT NOT NULL,
    created_at TEXT NOT NULL,
    verified_at TEXT
);
CREATE TABLE IF NOT EXISTS bank_questions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    teacher_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    grade TEXT,
    branch TEXT,
    outcome TEXT,
    difficulty TEXT,
    stem TEXT NOT NULL,
    options_json TEXT NOT NULL,
    correct TEXT,
    explanation TEXT,
    status TEXT NOT NULL DEFAULT 'taslak',
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_bank_teacher ON bank_questions (teacher_id, status);
CREATE TABLE IF NOT EXISTS password_resets (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    code_hash TEXT NOT NULL,
    created_at TEXT NOT NULL,
    expires_at TEXT NOT NULL,
    attempts INTEGER NOT NULL DEFAULT 0,
    used_at TEXT
);
CREATE TABLE IF NOT EXISTS teacher_profiles (
    user_id INTEGER PRIMARY KEY REFERENCES users(id) ON DELETE CASCADE,
    org_name TEXT,
    signature TEXT,
    updated_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS report_log (
    user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    week TEXT NOT NULL,
    sent_at TEXT NOT NULL,
    PRIMARY KEY (user_id, week)
);
"""


def _now() -> str:
    return datetime.now().isoformat(timespec="seconds")


_TABLES = ["users", "user_state", "activity_days", "exam_results", "badges", "classes",
           "class_members", "assignments", "submissions", "payments", "report_log",
           "bank_questions", "teacher_profiles", "password_resets"]

# Var olan veritabanlarına sonradan eklenen sütunlar: (tablo, sütun, tür)
_MIGRATIONS = [("assignments", "video_url", "TEXT")]

# Postgres şeması SQLite şemasından türetilir (tek kaynak, iki lehçe)
_SCHEMA_PG = (
    _SCHEMA_SQLITE.replace("INTEGER PRIMARY KEY AUTOINCREMENT", "BIGINT GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY")
    .replace(" INTEGER", " BIGINT")
    .replace(" REAL", " DOUBLE PRECISION")
)
# Supabase'in herkese açık REST API'si bu tablolara erişemesin: RLS açık, politika yok.
# (Uygulama veritabanı sahibi olarak bağlandığı için kendisi etkilenmez.)
_SCHEMA_PG += "\n".join(f"ALTER TABLE {t} ENABLE ROW LEVEL SECURITY;" for t in _TABLES)

_pg_conn: Any = None


def configure(url: Optional[str]) -> None:
    """Postgres/Supabase bağlantı adresini ayarlar. Boşsa SQLite kullanılır."""
    global PG_URL, _pg_conn
    url = (url or "").strip() or None
    if url != PG_URL:
        PG_URL = url
        _pg_conn = None


def backend() -> str:
    return "postgres" if PG_URL else "sqlite"


def _pg() -> Any:
    """Tek, paylaşılan Postgres bağlantısı (kopmuşsa yeniden kurulur)."""
    global _pg_conn
    import psycopg2  # sadece Postgres kullanılırken gerekir
    import psycopg2.extras

    if _pg_conn is None or getattr(_pg_conn, "closed", 1):
        _pg_conn = psycopg2.connect(PG_URL, connect_timeout=10, cursor_factory=psycopg2.extras.RealDictCursor)
        _pg_conn.autocommit = True
    return _pg_conn


def _sqlite() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def _clean(row: Any) -> Dict[str, Any]:
    out = dict(row)
    for k, v in out.items():
        if isinstance(v, Decimal):
            out[k] = float(v)
    return out


def _run(sql: str, params: tuple, fetch: str) -> Any:
    """fetch: 'rows' | 'id' | 'count'. SQL'ler '?' yer tutucusuyla yazılır."""
    with _lock:
        if PG_URL:
            pg_sql = sql.replace("?", "%s")
            if fetch == "id":
                pg_sql += " RETURNING id"
            for attempt in range(2):
                try:
                    with _pg().cursor() as cur:
                        cur.execute(pg_sql, params)
                        if fetch == "rows":
                            return [_clean(r) for r in cur.fetchall()]
                        if fetch == "id":
                            return int(cur.fetchone()["id"])
                        return cur.rowcount
                except Exception as exc:
                    # Bağlantı koptuysa bir kez yeniden dene
                    if attempt == 0 and exc.__class__.__name__ in ("OperationalError", "InterfaceError"):
                        global _pg_conn
                        _pg_conn = None
                        continue
                    raise
        conn = _sqlite()
        try:
            cur = conn.execute(sql, params)
            if fetch == "rows":
                return [_clean(r) for r in cur.fetchall()]
            conn.commit()
            return cur.lastrowid if fetch == "id" else cur.rowcount
        finally:
            conn.close()


def init_db() -> None:
    with _lock:
        if PG_URL:
            with _pg().cursor() as cur:
                cur.execute(_SCHEMA_PG)
                for table, col, typ in _MIGRATIONS:
                    cur.execute(f"ALTER TABLE {table} ADD COLUMN IF NOT EXISTS {col} {typ}")
            return
        conn = _sqlite()
        try:
            conn.executescript(_SCHEMA_SQLITE)
            for table, col, typ in _MIGRATIONS:
                cols = {r[1] for r in conn.execute(f"PRAGMA table_info({table})").fetchall()}
                if col not in cols:
                    conn.execute(f"ALTER TABLE {table} ADD COLUMN {col} {typ}")
            conn.commit()
        finally:
            conn.close()


def _q(sql: str, params: tuple = (), one: bool = False) -> Any:
    rows = _run(sql, params, "rows")
    if one:
        return rows[0] if rows else None
    return rows


def _x(sql: str, params: tuple = ()) -> int:
    """Yazma sorgusu; etkilenen satır sayısını döner."""
    return _run(sql, params, "count")


def _insert_id(sql: str, params: tuple = ()) -> int:
    """INSERT; yeni satırın id'sini döner."""
    return _run(sql, params, "id")


# ------------------------------------------------------------------------------
# Şifreleme ve giriş
# ------------------------------------------------------------------------------
_PBKDF2_ITER = 200_000


def hash_password(password: str) -> str:
    salt = secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), bytes.fromhex(salt), _PBKDF2_ITER).hex()
    return f"pbkdf2${_PBKDF2_ITER}${salt}${digest}"


def verify_password(password: str, stored: str) -> bool:
    try:
        _algo, iters, salt, digest = stored.split("$")
        calc = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), bytes.fromhex(salt), int(iters)).hex()
        return hmac.compare_digest(calc, digest)
    except Exception:
        return False


def _new_code(length: int = 6) -> str:
    alphabet = string.ascii_uppercase.replace("O", "").replace("I", "") + "23456789"
    return "".join(secrets.choice(alphabet) for _ in range(length))


def create_user(email: str, name: str, password: str, role: str, grade: Optional[str] = None) -> Dict[str, Any]:
    email = email.strip().lower()
    if role not in ROLES:
        raise ValueError("Geçersiz rol.")
    if "@" not in email or "." not in email.split("@")[-1]:
        raise ValueError("Geçerli bir e-posta adresi girin.")
    if len(password) < 8:
        raise ValueError("Şifre en az 8 karakter olmalı.")
    if not name.strip():
        raise ValueError("Ad soyad boş olamaz.")
    if get_user_by_email(email):
        raise ValueError("Bu e-posta ile kayıtlı bir hesap zaten var.")
    link_code = _new_code() if role == "ogrenci" else None
    uid = _insert_id(
        "INSERT INTO users (email, name, role, pw_hash, grade, link_code, created_at) VALUES (?,?,?,?,?,?,?)",
        (email, name.strip(), role, hash_password(password), grade, link_code, _now()),
    )
    return get_user(uid)


def authenticate(email: str, password: str) -> Optional[Dict[str, Any]]:
    row = _q("SELECT * FROM users WHERE email = ?", (email.strip().lower(),), one=True)
    if row and verify_password(password, row["pw_hash"]):
        return _public(row)
    return None


def _set_password(user_id: int, new_password: str) -> None:
    if len(new_password) < 8:
        raise ValueError("Şifre en az 8 karakter olmalı.")
    _x("UPDATE users SET pw_hash = ? WHERE id = ?", (hash_password(new_password), user_id))


def change_password(user_id: int, old_password: str, new_password: str) -> None:
    row = _q("SELECT pw_hash FROM users WHERE id = ?", (user_id,), one=True)
    if not row or not verify_password(old_password, row["pw_hash"]):
        raise ValueError("Mevcut şifre hatalı.")
    _set_password(user_id, new_password)


def _code_hash(code: str) -> str:
    return hashlib.sha256(code.strip().upper().encode("utf-8")).hexdigest()


RESET_TTL_MIN = 30
RESET_MAX_ATTEMPTS = 5
GENERATED_EMAIL_DOMAIN = "@ogrenci.kursunkalem"


def create_reset_code(email: str) -> Optional[Tuple[Dict[str, Any], str]]:
    """Kayıtlı e-posta için 8 haneli kod üretir; kayıt yoksa None (ekranda bunu belli etme)."""
    user = get_user_by_email(email)
    if not user or user["email"].endswith(GENERATED_EMAIL_DOMAIN):
        return None
    last = _q("SELECT created_at FROM password_resets WHERE user_id = ? ORDER BY id DESC LIMIT 1",
              (user["id"],), one=True)
    if last and last["created_at"] > (datetime.now() - timedelta(seconds=60)).isoformat(timespec="seconds"):
        raise ValueError("Yeni kod istemek için bir dakika bekleyin.")
    _x("UPDATE password_resets SET used_at = ? WHERE user_id = ? AND used_at IS NULL", (_now(), user["id"]))
    code = _new_code(8)
    _insert_id(
        "INSERT INTO password_resets (user_id, code_hash, created_at, expires_at) VALUES (?,?,?,?)",
        (user["id"], _code_hash(code), _now(),
         (datetime.now() + timedelta(minutes=RESET_TTL_MIN)).isoformat(timespec="seconds")),
    )
    return user, code


def reset_password_with_code(email: str, code: str, new_password: str) -> None:
    generic = "Kod hatalı ya da süresi dolmuş. Yeni kod isteyin."
    user = get_user_by_email(email)
    if not user:
        raise ValueError(generic)
    row = _q("SELECT * FROM password_resets WHERE user_id = ? AND used_at IS NULL ORDER BY id DESC LIMIT 1",
             (user["id"],), one=True)
    if not row or row["expires_at"] < _now() or int(row["attempts"]) >= RESET_MAX_ATTEMPTS:
        raise ValueError(generic)
    if not hmac.compare_digest(row["code_hash"], _code_hash(code)):
        _x("UPDATE password_resets SET attempts = attempts + 1 WHERE id = ?", (row["id"],))
        raise ValueError(generic)
    _set_password(user["id"], new_password)
    _x("UPDATE password_resets SET used_at = ? WHERE id = ?", (_now(), row["id"]))


def teacher_reset_student_password(teacher_id: int, student_id: int) -> str:
    """Öğretmen, kendi sınıfındaki ve e-postası olmayan (toplu açılmış) öğrencinin şifresini sıfırlar."""
    row = _q(
        "SELECT u.id, u.email FROM users u JOIN class_members m ON m.user_id = u.id "
        "JOIN classes c ON c.id = m.class_id WHERE u.id = ? AND u.role = 'ogrenci' AND c.teacher_id = ?",
        (student_id, teacher_id), one=True,
    )
    if not row:
        raise ValueError("Bu öğrenci sınıflarınızda değil.")
    if not row["email"].endswith(GENERATED_EMAIL_DOMAIN):
        raise ValueError("Bu öğrencinin kendi e-postası var; girişte 'Şifremi unuttum' ile sıfırlamalı.")
    password = "".join(secrets.choice("abcdefghjkmnpqrstuvwxyz23456789") for _ in range(8))
    _set_password(student_id, password)
    return password


def _public(row: Optional[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    if not row:
        return None
    row = dict(row)
    row.pop("pw_hash", None)
    return row


def get_user(user_id: int) -> Optional[Dict[str, Any]]:
    return _public(_q("SELECT * FROM users WHERE id = ?", (user_id,), one=True))


def get_user_by_email(email: str) -> Optional[Dict[str, Any]]:
    return _public(_q("SELECT * FROM users WHERE email = ?", (email.strip().lower(),), one=True))


def link_parent(parent_id: int, link_code: str) -> Dict[str, Any]:
    student = _q("SELECT * FROM users WHERE link_code = ? AND role = 'ogrenci'", (link_code.strip().upper(),), one=True)
    if not student:
        raise ValueError("Bu koda ait bir öğrenci bulunamadı.")
    _x("UPDATE users SET parent_of = ? WHERE id = ?", (student["id"], parent_id))
    return _public(student)


def set_daily_goal(user_id: int, goal: int) -> None:
    _x("UPDATE users SET daily_goal = ? WHERE id = ?", (max(5, min(200, int(goal))), user_id))


def update_grade(user_id: int, grade: str) -> None:
    _x("UPDATE users SET grade = ? WHERE id = ?", (grade, user_id))


def list_parents_with_students() -> List[Dict[str, Any]]:
    return _q(
        "SELECT p.id AS parent_id, p.name AS parent_name, p.email AS parent_email, "
        "s.id AS student_id, s.name AS student_name, s.grade AS student_grade "
        "FROM users p JOIN users s ON s.id = p.parent_of WHERE p.role = 'veli'"
    )


# ------------------------------------------------------------------------------
# Kişiye özel uygulama durumu (eski ortak JSON dosyasının yerini alır)
# ------------------------------------------------------------------------------
def load_state(user_id: int) -> Dict[str, Any]:
    row = _q("SELECT payload FROM user_state WHERE user_id = ?", (user_id,), one=True)
    if not row:
        return {}
    try:
        data = json.loads(row["payload"])
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_state(user_id: int, payload: Dict[str, Any]) -> None:
    _x(
        "INSERT INTO user_state (user_id, payload, updated_at) VALUES (?,?,?) "
        "ON CONFLICT(user_id) DO UPDATE SET payload = excluded.payload, updated_at = excluded.updated_at",
        (user_id, json.dumps(payload, ensure_ascii=False), _now()),
    )


def delete_state(user_id: int) -> None:
    _x("DELETE FROM user_state WHERE user_id = ?", (user_id,))


# ------------------------------------------------------------------------------
# Aktivite, deneme sonuçları, rozetler
# ------------------------------------------------------------------------------
def record_practice(user_id: int, correct: int, wrong: int, blank: int, day: Optional[date] = None) -> None:
    d = (day or date.today()).isoformat()
    solved = correct + wrong + blank
    _x(
        "INSERT INTO activity_days (user_id, day, solved, correct, wrong, blank) VALUES (?,?,?,?,?,?) "
        "ON CONFLICT(user_id, day) DO UPDATE SET solved = activity_days.solved + excluded.solved, "
        "correct = activity_days.correct + excluded.correct, wrong = activity_days.wrong + excluded.wrong, "
        "blank = activity_days.blank + excluded.blank",
        (user_id, d, solved, correct, wrong, blank),
    )


def activity_days(user_id: int, since: Optional[date] = None) -> List[Dict[str, Any]]:
    if since:
        return _q("SELECT * FROM activity_days WHERE user_id = ? AND day >= ? ORDER BY day", (user_id, since.isoformat()))
    return _q("SELECT * FROM activity_days WHERE user_id = ? ORDER BY day", (user_id,))


def add_exam_result(user_id: int, source: str, correct: int, wrong: int, blank: int,
                    assignment_id: Optional[int] = None) -> None:
    graded = correct + wrong
    acc = round(100.0 * correct / graded, 1) if graded else 0.0
    _x(
        "INSERT INTO exam_results (user_id, created_at, source, correct, wrong, blank, accuracy, assignment_id) "
        "VALUES (?,?,?,?,?,?,?,?)",
        (user_id, _now(), source, correct, wrong, blank, acc, assignment_id),
    )


def recent_exam_results(user_id: int, limit: int = 10) -> List[Dict[str, Any]]:
    return _q("SELECT * FROM exam_results WHERE user_id = ? ORDER BY id DESC LIMIT ?", (user_id, limit))


def exam_count(user_id: int) -> int:
    row = _q("SELECT COUNT(*) AS n FROM exam_results WHERE user_id = ?", (user_id,), one=True)
    return int(row["n"]) if row else 0


def award_badge(user_id: int, code: str) -> bool:
    """Rozeti verir. Yeni kazanıldıysa True döner."""
    n = _x(
        "INSERT INTO badges (user_id, code, earned_at) VALUES (?,?,?) ON CONFLICT DO NOTHING",
        (user_id, code, _now()),
    )
    return n > 0


def user_badges(user_id: int) -> Dict[str, str]:
    return {r["code"]: r["earned_at"] for r in _q("SELECT code, earned_at FROM badges WHERE user_id = ?", (user_id,))}


def weekly_stats(user_id: int, today: Optional[date] = None) -> Dict[str, Any]:
    today = today or date.today()
    since = today - timedelta(days=6)
    rows = activity_days(user_id, since)
    solved = sum(r["solved"] for r in rows)
    correct = sum(r["correct"] for r in rows)
    wrong = sum(r["wrong"] for r in rows)
    blank = sum(r["blank"] for r in rows)
    graded = correct + wrong
    return {
        "weekly_solved": solved,
        "correct": correct,
        "wrong": wrong,
        "blank": blank,
        "accuracy": round(100.0 * correct / graded, 1) if graded else 0.0,
        "active_days": len([r for r in rows if r["solved"] > 0]),
        "by_day": {r["day"]: r["solved"] for r in rows},
    }


# ------------------------------------------------------------------------------
# Sınıflar, ödevler, teslimler
# ------------------------------------------------------------------------------
def create_class(teacher_id: int, name: str, grade: Optional[str]) -> Dict[str, Any]:
    if not name.strip():
        raise ValueError("Sınıf adı boş olamaz.")
    for _ in range(10):
        code = _new_code()
        if not _q("SELECT 1 FROM classes WHERE code = ?", (code,), one=True):
            break
    cid = _insert_id(
        "INSERT INTO classes (name, code, teacher_id, grade, created_at) VALUES (?,?,?,?,?)",
        (name.strip(), code, teacher_id, grade, _now()),
    )
    return _q("SELECT * FROM classes WHERE id = ?", (cid,), one=True)


def teacher_classes(teacher_id: int) -> List[Dict[str, Any]]:
    return _q(
        "SELECT c.*, (SELECT COUNT(*) FROM class_members m WHERE m.class_id = c.id) AS member_count "
        "FROM classes c WHERE teacher_id = ? ORDER BY c.id DESC",
        (teacher_id,),
    )


def join_class(user_id: int, code: str) -> Dict[str, Any]:
    cls = _q("SELECT * FROM classes WHERE code = ?", (code.strip().upper(),), one=True)
    if not cls:
        raise ValueError("Bu koda ait bir sınıf bulunamadı.")
    _x(
        "INSERT INTO class_members (class_id, user_id, joined_at) VALUES (?,?,?) ON CONFLICT DO NOTHING",
        (cls["id"], user_id, _now()),
    )
    return cls


def student_classes(user_id: int) -> List[Dict[str, Any]]:
    return _q(
        "SELECT c.*, u.name AS teacher_name FROM classes c JOIN class_members m ON m.class_id = c.id "
        "JOIN users u ON u.id = c.teacher_id WHERE m.user_id = ? ORDER BY c.name",
        (user_id,),
    )


def class_members(class_id: int) -> List[Dict[str, Any]]:
    return _q(
        "SELECT u.id, u.name, u.email, u.grade FROM users u JOIN class_members m ON m.user_id = u.id "
        "WHERE m.class_id = ? ORDER BY u.name",
        (class_id,),
    )


def clean_video_url(url: str) -> str:
    """Yalnızca http(s) bağlantılarını kabul eder; boşsa boş döner."""
    url = (url or "").strip()
    if not url:
        return ""
    if not re.match(r"^https?://[^\s<>\"']+$", url, flags=re.IGNORECASE):
        raise ValueError("Video bağlantısı https:// ile başlayan geçerli bir adres olmalı.")
    return url


def create_assignment(class_id: int, title: str, content: str, grade: str = "", branch: str = "",
                      outcome: str = "", due_date: Optional[str] = None, video_url: str = "") -> int:
    if not content.strip():
        raise ValueError("Ödev içeriği boş olamaz.")
    return _insert_id(
        "INSERT INTO assignments (class_id, title, grade, branch, outcome, content, due_date, created_at, video_url) "
        "VALUES (?,?,?,?,?,?,?,?,?)",
        (class_id, title.strip() or "Ödev", grade, branch, outcome, content, due_date, _now(),
         clean_video_url(video_url) or None),
    )


def class_assignments(class_id: int) -> List[Dict[str, Any]]:
    return _q(
        "SELECT a.*, (SELECT COUNT(*) FROM submissions s WHERE s.assignment_id = a.id) AS submission_count "
        "FROM assignments a WHERE class_id = ? ORDER BY a.id DESC",
        (class_id,),
    )


def student_assignments(user_id: int) -> List[Dict[str, Any]]:
    return _q(
        "SELECT a.*, c.name AS class_name, s.accuracy AS my_accuracy, s.net AS my_net, s.submitted_at "
        "FROM assignments a JOIN class_members m ON m.class_id = a.class_id "
        "JOIN classes c ON c.id = a.class_id "
        "LEFT JOIN submissions s ON s.assignment_id = a.id AND s.user_id = m.user_id "
        "WHERE m.user_id = ? ORDER BY (s.submitted_at IS NOT NULL), a.id DESC",
        (user_id,),
    )


def get_assignment(assignment_id: int) -> Optional[Dict[str, Any]]:
    return _q("SELECT * FROM assignments WHERE id = ?", (assignment_id,), one=True)


def submit_assignment(assignment_id: int, user_id: int, correct: int, wrong: int, blank: int) -> None:
    graded = correct + wrong
    acc = round(100.0 * correct / graded, 1) if graded else 0.0
    net = round(max(0.0, correct - wrong / 3.0), 2)
    _x(
        "INSERT INTO submissions (assignment_id, user_id, correct, wrong, blank, net, accuracy, submitted_at) "
        "VALUES (?,?,?,?,?,?,?,?) ON CONFLICT(assignment_id, user_id) DO UPDATE SET "
        "correct = excluded.correct, wrong = excluded.wrong, blank = excluded.blank, net = excluded.net, "
        "accuracy = excluded.accuracy, submitted_at = excluded.submitted_at",
        (assignment_id, user_id, correct, wrong, blank, net, acc, _now()),
    )


def assignment_results(assignment_id: int) -> List[Dict[str, Any]]:
    """Sınıftaki her öğrenci için teslim durumu (teslim etmeyenler de listelenir)."""
    return _q(
        "SELECT u.id AS user_id, u.name, s.correct, s.wrong, s.blank, s.net, s.accuracy, s.submitted_at "
        "FROM assignments a JOIN class_members m ON m.class_id = a.class_id JOIN users u ON u.id = m.user_id "
        "LEFT JOIN submissions s ON s.assignment_id = a.id AND s.user_id = u.id "
        "WHERE a.id = ? ORDER BY u.name",
        (assignment_id,),
    )


def remove_member(teacher_id: int, class_id: int, user_id: int) -> int:
    """Öğrenciyi sınıftan çıkarır (yalnızca sınıfın öğretmeni)."""
    return _x(
        "DELETE FROM class_members WHERE class_id = ? AND user_id = ? "
        "AND class_id IN (SELECT id FROM classes WHERE teacher_id = ?)",
        (class_id, user_id, teacher_id),
    )


def _slug(text: str) -> str:
    tr = str.maketrans("çğıöşüÇĞİÖŞÜ", "cgiosuCGIOSU")
    out = re.sub(r"[^a-z0-9]+", ".", text.translate(tr).lower()).strip(".")
    return out[:24] or "ogrenci"


def bulk_create_students(teacher_id: int, class_id: int, rows: List[Dict[str, str]],
                         grade: Optional[str] = None) -> List[Dict[str, str]]:
    """rows: [{"name": ..., "email": opsiyonel}]. Hesap açar, sınıfa ekler, geçici şifreleri döner.
    E-postası olmayan öğrenciye giriş için okunabilir bir kullanıcı adresi üretilir."""
    cls = _q("SELECT * FROM classes WHERE id = ? AND teacher_id = ?", (class_id, teacher_id), one=True)
    if not cls:
        raise ValueError("Sınıf bulunamadı.")
    results: List[Dict[str, str]] = []
    pw_alphabet = "abcdefghjkmnpqrstuvwxyz23456789"
    for row in rows:
        name = (row.get("name") or "").strip()
        if not name:
            continue
        email = (row.get("email") or "").strip().lower()
        existing = get_user_by_email(email) if email else None
        if existing:
            if existing["role"] != "ogrenci":
                results.append({"name": name, "email": email, "password": "", "status": "Atlandı: öğrenci hesabı değil"})
                continue
            join_class(existing["id"], cls["code"])
            results.append({"name": name, "email": email, "password": "", "status": "Var olan hesap sınıfa eklendi"})
            continue
        if not email:
            for _ in range(20):
                email = f"{_slug(name)}.{secrets.randbelow(9000) + 1000}{GENERATED_EMAIL_DOMAIN}"
                if not get_user_by_email(email):
                    break
        password = "".join(secrets.choice(pw_alphabet) for _ in range(8))
        try:
            user = create_user(email, name, password, "ogrenci", grade or cls.get("grade"))
        except ValueError as exc:
            results.append({"name": name, "email": email, "password": "", "status": f"Hata: {exc}"})
            continue
        join_class(user["id"], cls["code"])
        results.append({"name": name, "email": email, "password": password, "status": "Oluşturuldu"})
    return results


def class_parents(class_id: int) -> List[Dict[str, Any]]:
    """Sınıftaki öğrencilere bağlı veli hesapları."""
    return _q(
        "SELECT p.id AS parent_id, p.name AS parent_name, p.email AS parent_email, "
        "s.id AS student_id, s.name AS student_name "
        "FROM class_members m JOIN users s ON s.id = m.user_id JOIN users p ON p.parent_of = s.id "
        "WHERE m.class_id = ? AND p.role = 'veli' ORDER BY s.name",
        (class_id,),
    )


def get_teacher_profile(user_id: int) -> Dict[str, Any]:
    row = _q("SELECT * FROM teacher_profiles WHERE user_id = ?", (user_id,), one=True)
    return row or {"user_id": user_id, "org_name": "", "signature": ""}


def save_teacher_profile(user_id: int, org_name: str, signature: str) -> None:
    _x(
        "INSERT INTO teacher_profiles (user_id, org_name, signature, updated_at) VALUES (?,?,?,?) "
        "ON CONFLICT(user_id) DO UPDATE SET org_name = excluded.org_name, signature = excluded.signature, "
        "updated_at = excluded.updated_at",
        (user_id, org_name.strip(), signature.strip(), _now()),
    )


def student_teachers(student_id: int) -> List[Dict[str, Any]]:
    """Öğrencinin öğretmenleri (veli raporunda imza için)."""
    return _q(
        "SELECT DISTINCT u.id, u.name, c.name AS class_name, tp.org_name, tp.signature "
        "FROM class_members m JOIN classes c ON c.id = m.class_id JOIN users u ON u.id = c.teacher_id "
        "LEFT JOIN teacher_profiles tp ON tp.user_id = u.id WHERE m.user_id = ? ORDER BY c.name",
        (student_id,),
    )


# ------------------------------------------------------------------------------
# Öğretmen soru bankası (taslak → onaylı)
# ------------------------------------------------------------------------------
BANK_STATUSES = ("taslak", "onayli")


def _bank_row(row: Optional[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    if not row:
        return None
    try:
        row["options"] = json.loads(row.get("options_json") or "{}")
    except ValueError:
        row["options"] = {}
    return row


def add_bank_questions(teacher_id: int, questions: List[Dict[str, Any]], grade: str = "", branch: str = "",
                       outcome: str = "", difficulty: str = "", status: str = "taslak") -> List[int]:
    if status not in BANK_STATUSES:
        raise ValueError("Geçersiz durum.")
    ids: List[int] = []
    now = _now()
    for q in questions:
        stem = (q.get("stem") or "").strip()
        options = {k: str(v).strip() for k, v in (q.get("options") or {}).items() if str(v).strip()}
        if not stem or len(options) < 2:
            continue
        correct = (q.get("correct") or "").strip().upper() or None
        ids.append(_insert_id(
            "INSERT INTO bank_questions (teacher_id, grade, branch, outcome, difficulty, stem, options_json, "
            "correct, explanation, status, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
            (teacher_id, grade, branch, q.get("outcome") or outcome, difficulty, stem,
             json.dumps(options, ensure_ascii=False), correct, (q.get("explanation") or "").strip(),
             status, now, now),
        ))
    return ids


def list_bank_questions(teacher_id: int, status: Optional[str] = None, grade: Optional[str] = None,
                        branch: Optional[str] = None, search: str = "") -> List[Dict[str, Any]]:
    sql = "SELECT * FROM bank_questions WHERE teacher_id = ?"
    params: List[Any] = [teacher_id]
    if status:
        sql += " AND status = ?"
        params.append(status)
    if grade:
        sql += " AND grade = ?"
        params.append(grade)
    if branch:
        sql += " AND branch = ?"
        params.append(branch)
    if search.strip():
        sql += " AND (LOWER(stem) LIKE ? OR LOWER(outcome) LIKE ?)"
        like = f"%{search.strip().lower()}%"
        params += [like, like]
    sql += " ORDER BY id DESC"
    return [_bank_row(r) for r in _q(sql, tuple(params))]


def get_bank_questions(teacher_id: int, ids: List[int]) -> List[Dict[str, Any]]:
    if not ids:
        return []
    marks = ",".join("?" for _ in ids)
    rows = _q(f"SELECT * FROM bank_questions WHERE teacher_id = ? AND id IN ({marks})", (teacher_id, *ids))
    by_id = {r["id"]: _bank_row(r) for r in rows}
    return [by_id[i] for i in ids if i in by_id]


def update_bank_question(teacher_id: int, qid: int, stem: str, options: Dict[str, str], correct: Optional[str],
                         explanation: str = "", outcome: Optional[str] = None) -> int:
    options = {k: v.strip() for k, v in options.items() if v and v.strip()}
    if not stem.strip():
        raise ValueError("Soru kökü boş olamaz.")
    if len(options) < 2:
        raise ValueError("En az iki şık gerekli.")
    if correct and correct not in options:
        raise ValueError("Doğru cevap şıklardan biri olmalı.")
    sql = "UPDATE bank_questions SET stem = ?, options_json = ?, correct = ?, explanation = ?, updated_at = ?"
    params: List[Any] = [stem.strip(), json.dumps(options, ensure_ascii=False), correct or None, explanation.strip(), _now()]
    if outcome is not None:
        sql += ", outcome = ?"
        params.append(outcome.strip())
    sql += " WHERE id = ? AND teacher_id = ?"
    params += [qid, teacher_id]
    return _x(sql, tuple(params))


def set_bank_status(teacher_id: int, qids: List[int], status: str) -> int:
    if status not in BANK_STATUSES:
        raise ValueError("Geçersiz durum.")
    if not qids:
        return 0
    if status == "onayli":
        # Cevabı belirsiz soru onaylanamaz
        marks = ",".join("?" for _ in qids)
        bad = _q(f"SELECT id FROM bank_questions WHERE teacher_id = ? AND id IN ({marks}) "
                 "AND (correct IS NULL OR correct = '')", (teacher_id, *qids))
        if bad:
            raise ValueError(f"{len(bad)} sorunun doğru cevabı seçilmemiş; önce cevabı işaretleyin.")
    marks = ",".join("?" for _ in qids)
    return _x(f"UPDATE bank_questions SET status = ?, updated_at = ? WHERE teacher_id = ? AND id IN ({marks})",
              (status, _now(), teacher_id, *qids))


def delete_bank_questions(teacher_id: int, qids: List[int]) -> int:
    if not qids:
        return 0
    marks = ",".join("?" for _ in qids)
    return _x(f"DELETE FROM bank_questions WHERE teacher_id = ? AND id IN ({marks})", (teacher_id, *qids))


def bank_counts(teacher_id: int) -> Dict[str, int]:
    rows = _q("SELECT status, COUNT(*) AS n FROM bank_questions WHERE teacher_id = ? GROUP BY status", (teacher_id,))
    out = {s: 0 for s in BANK_STATUSES}
    for r in rows:
        out[r["status"]] = int(r["n"])
    return out


def bank_filters(teacher_id: int) -> Dict[str, List[str]]:
    grades = [r["grade"] for r in _q("SELECT DISTINCT grade FROM bank_questions WHERE teacher_id = ? AND grade <> '' ORDER BY grade", (teacher_id,))]
    branches = [r["branch"] for r in _q("SELECT DISTINCT branch FROM bank_questions WHERE teacher_id = ? AND branch <> '' ORDER BY branch", (teacher_id,))]
    return {"grades": grades, "branches": branches}


def questions_to_text(questions: List[Dict[str, Any]]) -> str:
    """Bankadaki soruları portalın çözebildiği metin biçimine çevirir (Soru N / A) ... / Cevap Anahtarı)."""
    blocks: List[str] = []
    key: List[str] = []
    for i, q in enumerate(questions, start=1):
        opts = "\n".join(f"{k}) {v}" for k, v in q["options"].items())
        blocks.append(f"Soru {i}\n{q['stem']}\n{opts}")
        if q.get("correct"):
            key.append(f"{i}-{q['correct']}")
    text = "\n\n".join(blocks)
    if key:
        text += "\n\nCevap Anahtarı\n" + "\n".join(key)
    return text


def teacher_has_student(teacher_id: int, student_id: int) -> bool:
    return bool(_q(
        "SELECT 1 AS ok FROM class_members m JOIN classes c ON c.id = m.class_id "
        "WHERE m.user_id = ? AND c.teacher_id = ? LIMIT 1", (student_id, teacher_id), one=True))


def student_xray(student_id: int, teacher_id: Optional[int] = None) -> Dict[str, Any]:
    """Gerçek ödev/deneme kayıtlarından öğrenci röntgeni.
    teacher_id verilirse yalnızca o öğretmenin sınıflarındaki ödevler sayılır."""
    t_filter = " AND c.teacher_id = ?" if teacher_id else ""
    t_param: tuple = (teacher_id,) if teacher_id else ()
    mine = _q(
        "SELECT a.id, a.class_id, a.title, COALESCE(NULLIF(a.outcome, ''), a.title) AS outcome, "
        "COALESCE(a.branch, '') AS branch, a.due_date, a.created_at, s.accuracy, s.net, s.submitted_at "
        "FROM assignments a JOIN classes c ON c.id = a.class_id "
        "JOIN class_members m ON m.class_id = a.class_id AND m.user_id = ? "
        "LEFT JOIN submissions s ON s.assignment_id = a.id AND s.user_id = ? "
        "WHERE 1 = 1" + t_filter + " ORDER BY a.id",
        (student_id, student_id, *t_param),
    )
    done = [r for r in mine if r.get("submitted_at")]
    # Sınıf ortalaması: aynı ödevlerde tüm sınıfın doğruluğu
    class_avg: Dict[int, float] = {}
    ids = [r["id"] for r in done]
    if ids:
        marks = ",".join("?" for _ in ids)
        for row in _q(f"SELECT assignment_id, AVG(accuracy) AS acc FROM submissions WHERE assignment_id IN ({marks}) "
                      "GROUP BY assignment_id", tuple(ids)):
            class_avg[int(row["assignment_id"])] = float(row["acc"])
    groups: Dict[Tuple[str, str], Dict[str, Any]] = {}
    for r in done:
        g = groups.setdefault((r["outcome"], r["branch"]), {"outcome": r["outcome"], "branch": r["branch"],
                                                             "n": 0, "acc": 0.0, "class_acc": 0.0, "net": 0.0})
        g["n"] += 1
        g["acc"] += float(r["accuracy"])
        g["net"] += float(r["net"])
        g["class_acc"] += class_avg.get(int(r["id"]), float(r["accuracy"]))
    outcomes = []
    for g in groups.values():
        n = g["n"]
        outcomes.append({"outcome": g["outcome"], "branch": g["branch"], "n": n,
                         "acc": round(g["acc"] / n, 1), "class_acc": round(g["class_acc"] / n, 1),
                         "net": round(g["net"] / n, 2)})
    outcomes.sort(key=lambda x: (x["acc"], -x["n"]))
    today = datetime.now().date().isoformat()
    missing = [r for r in mine if not r.get("submitted_at")]
    overdue = [r for r in missing if r.get("due_date") and str(r["due_date"])[:10] < today]
    exams = list(reversed(_q("SELECT created_at, source, accuracy, correct, wrong, blank FROM exam_results "
                             "WHERE user_id = ? ORDER BY id DESC LIMIT 20", (student_id,))))
    avg_acc = round(sum(float(r["accuracy"]) for r in done) / len(done), 1) if done else None
    avg_class = round(sum(class_avg.get(int(r["id"]), float(r["accuracy"])) for r in done) / len(done), 1) if done else None
    return {"outcomes": outcomes, "assigned": len(mine), "done": len(done), "missing": missing,
            "overdue": overdue, "exams": exams, "avg_acc": avg_acc, "avg_class": avg_class}


def class_overview(class_id: int) -> Dict[str, Any]:
    """Sınıf karşılaştırma paneli için kazanım ve öğrenci bazlı özet."""
    by_outcome = _q(
        "SELECT COALESCE(NULLIF(a.outcome, ''), a.title) AS outcome, a.branch, "
        "ROUND(CAST(AVG(s.accuracy) AS NUMERIC), 1) AS avg_accuracy, COUNT(s.user_id) AS submissions "
        "FROM assignments a JOIN submissions s ON s.assignment_id = a.id "
        "WHERE a.class_id = ? GROUP BY 1, 2 ORDER BY avg_accuracy ASC NULLS LAST",
        (class_id,),
    )
    by_student = _q(
        "SELECT u.name, COUNT(s.assignment_id) AS done, ROUND(CAST(AVG(s.accuracy) AS NUMERIC), 1) AS avg_accuracy, "
        "ROUND(CAST(AVG(s.net) AS NUMERIC), 2) AS avg_net "
        "FROM class_members m JOIN users u ON u.id = m.user_id "
        "LEFT JOIN submissions s ON s.user_id = u.id AND s.assignment_id IN "
        "(SELECT id FROM assignments WHERE class_id = ?) "
        "WHERE m.class_id = ? GROUP BY u.id, u.name ORDER BY avg_accuracy DESC NULLS LAST, u.name",
        (class_id, class_id),
    )
    total_assignments = _q("SELECT COUNT(*) AS n FROM assignments WHERE class_id = ?", (class_id,), one=True)["n"]
    return {"by_outcome": by_outcome, "by_student": by_student, "total_assignments": total_assignments}


# ------------------------------------------------------------------------------
# Ödemeler
# ------------------------------------------------------------------------------
def create_payment(token: str, user_id: int, amount: int, billing_cycle: str, plan: Dict[str, Any]) -> None:
    _x(
        "INSERT INTO payments (token, user_id, amount, billing_cycle, plan_json, status, created_at) "
        "VALUES (?,?,?,?,?,?,?) ON CONFLICT(token) DO UPDATE SET user_id = excluded.user_id, "
        "amount = excluded.amount, billing_cycle = excluded.billing_cycle, plan_json = excluded.plan_json, "
        "status = excluded.status, created_at = excluded.created_at",
        (token, user_id, int(amount), billing_cycle, json.dumps(plan, ensure_ascii=False), "pending", _now()),
    )


def get_payment(token: str) -> Optional[Dict[str, Any]]:
    row = _q("SELECT * FROM payments WHERE token = ?", (token,), one=True)
    if row:
        row["plan"] = json.loads(row.pop("plan_json"))
    return row


def latest_pending_payment(user_id: int) -> Optional[Dict[str, Any]]:
    row = _q(
        "SELECT * FROM payments WHERE user_id = ? AND status = 'pending' ORDER BY created_at DESC LIMIT 1",
        (user_id,), one=True,
    )
    if row:
        row["plan"] = json.loads(row.pop("plan_json"))
    return row


def mark_payment(token: str, status: str) -> None:
    _x("UPDATE payments SET status = ?, verified_at = ? WHERE token = ?", (status, _now(), token))


# ------------------------------------------------------------------------------
# Haftalık rapor kaydı
# ------------------------------------------------------------------------------
def week_key(d: Optional[date] = None) -> str:
    y, w, _ = (d or date.today()).isocalendar()
    return f"{y}-W{w:02d}"


def report_already_sent(user_id: int, week: str) -> bool:
    return _q("SELECT 1 FROM report_log WHERE user_id = ? AND week = ?", (user_id, week), one=True) is not None


def log_report_sent(user_id: int, week: str) -> None:
    _x("INSERT INTO report_log (user_id, week, sent_at) VALUES (?,?,?) ON CONFLICT DO NOTHING", (user_id, week, _now()))


db = SimpleNamespace(
    BANK_STATUSES=BANK_STATUSES,
    DB_PATH=DB_PATH,
    GENERATED_EMAIL_DOMAIN=GENERATED_EMAIL_DOMAIN,
    PG_URL=PG_URL,
    RESET_MAX_ATTEMPTS=RESET_MAX_ATTEMPTS,
    RESET_TTL_MIN=RESET_TTL_MIN,
    ROLES=ROLES,
    ROLE_LABELS=ROLE_LABELS,
    activity_days=activity_days,
    add_bank_questions=add_bank_questions,
    add_exam_result=add_exam_result,
    assignment_results=assignment_results,
    authenticate=authenticate,
    award_badge=award_badge,
    backend=backend,
    bank_counts=bank_counts,
    bank_filters=bank_filters,
    bulk_create_students=bulk_create_students,
    change_password=change_password,
    class_assignments=class_assignments,
    class_members=class_members,
    class_overview=class_overview,
    class_parents=class_parents,
    clean_video_url=clean_video_url,
    configure=configure,
    create_assignment=create_assignment,
    create_class=create_class,
    create_payment=create_payment,
    create_reset_code=create_reset_code,
    create_user=create_user,
    delete_bank_questions=delete_bank_questions,
    delete_state=delete_state,
    exam_count=exam_count,
    get_assignment=get_assignment,
    get_bank_questions=get_bank_questions,
    get_payment=get_payment,
    get_teacher_profile=get_teacher_profile,
    get_user=get_user,
    get_user_by_email=get_user_by_email,
    hash_password=hash_password,
    init_db=init_db,
    join_class=join_class,
    latest_pending_payment=latest_pending_payment,
    link_parent=link_parent,
    list_bank_questions=list_bank_questions,
    list_parents_with_students=list_parents_with_students,
    load_state=load_state,
    log_report_sent=log_report_sent,
    mark_payment=mark_payment,
    questions_to_text=questions_to_text,
    recent_exam_results=recent_exam_results,
    record_practice=record_practice,
    remove_member=remove_member,
    report_already_sent=report_already_sent,
    reset_password_with_code=reset_password_with_code,
    save_state=save_state,
    save_teacher_profile=save_teacher_profile,
    set_bank_status=set_bank_status,
    set_daily_goal=set_daily_goal,
    student_assignments=student_assignments,
    student_classes=student_classes,
    student_teachers=student_teachers,
    student_xray=student_xray,
    submit_assignment=submit_assignment,
    teacher_classes=teacher_classes,
    teacher_has_student=teacher_has_student,
    teacher_reset_student_password=teacher_reset_student_password,
    update_bank_question=update_bank_question,
    update_grade=update_grade,
    user_badges=user_badges,
    verify_password=verify_password,
    week_key=week_key,
    weekly_stats=weekly_stats,
)


# ==============================================================================
# OYUNLAŞTIRMA: seri, rozetler, aralıklı tekrar, zorluk önerisi
# ==============================================================================
DIFFICULTY_LEVELS = ["Kolay", "Orta", "Orta-Zor", "Zor (LGS/ÖSYM)"]

# ------------------------------------------------------------------------------
# Seri
# ------------------------------------------------------------------------------
def streak(active_days: Iterable[str], today: Optional[date] = None) -> int:
    """Üst üste çalışılan gün sayısı. Bugün henüz çalışılmadıysa dünden geriye sayar
    (böylece seri gün içinde 'kopmuş' görünmez)."""
    today = today or date.today()
    days = {date.fromisoformat(d) for d in active_days}
    cursor = today if today in days else today - timedelta(days=1)
    count = 0
    while cursor in days:
        count += 1
        cursor -= timedelta(days=1)
    return count


# ------------------------------------------------------------------------------
# Rozetler
# ------------------------------------------------------------------------------
# kod: (ad, nasıl kazanılır, renk)
BADGES: Dict[str, Tuple[str, str, str]] = {
    "ilk_deneme": ("İlk Adım", "İlk mini denemeni bitir", "#1D4ED8"),
    "seri_3": ("Isınıyoruz", "3 gün üst üste çalış", "#F97316"),
    "seri_7": ("Haftanın Yıldızı", "7 gün üst üste çalış", "#DC2626"),
    "seri_30": ("Durdurulamaz", "30 gün üst üste çalış", "#7C3AED"),
    "soru_100": ("Yüzlük", "Toplam 100 soru çöz", "#059669"),
    "soru_500": ("Soru Canavarı", "Toplam 500 soru çöz", "#0D9488"),
    "isabet_90": ("Keskin Nişancı", "En az 5 soruluk denemede %90 doğruluk", "#DB2777"),
    "tekrar_10": ("Hata Avcısı", "Hata defterinden 10 soruyu tekrarla tamamla", "#CA8A04"),
    "foto_cozum": ("Fotoğrafçı", "Fotoğrafla ilk sorunu çözdür", "#2563EB"),
    "odev_5": ("Ödev Kahramanı", "5 ödev teslim et", "#9333EA"),
}


def earned_badges(ctx: Dict[str, Any]) -> List[str]:
    """ctx: streak, total_solved, exams, last_exam_size, last_exam_accuracy,
    reviews_completed, photo_solves, assignments_done"""
    out: List[str] = []
    s = int(ctx.get("streak", 0))
    total = int(ctx.get("total_solved", 0))
    if int(ctx.get("exams", 0)) >= 1:
        out.append("ilk_deneme")
    for need, code in ((3, "seri_3"), (7, "seri_7"), (30, "seri_30")):
        if s >= need:
            out.append(code)
    if total >= 100:
        out.append("soru_100")
    if total >= 500:
        out.append("soru_500")
    if int(ctx.get("last_exam_size", 0)) >= 5 and float(ctx.get("last_exam_accuracy", 0)) >= 90:
        out.append("isabet_90")
    if int(ctx.get("reviews_completed", 0)) >= 10:
        out.append("tekrar_10")
    if int(ctx.get("photo_solves", 0)) >= 1:
        out.append("foto_cozum")
    if int(ctx.get("assignments_done", 0)) >= 5:
        out.append("odev_5")
    return out


# ------------------------------------------------------------------------------
# Aralıklı tekrar (1. gün -> 3. gün -> 7. gün)
# ------------------------------------------------------------------------------
REVIEW_GAPS = [1, 3, 7]


def _parse_tr_date(s: str) -> Optional[date]:
    try:
        return datetime.strptime(s, "%d.%m.%Y").date()
    except Exception:
        return None


def next_review_date(item: Dict[str, Any]) -> Optional[date]:
    stage = int(item.get("stage", 0))
    if stage >= len(REVIEW_GAPS):
        return None
    key = f"review_{REVIEW_GAPS[stage]}d"
    return _parse_tr_date(str(item.get(key, "")))


def due_reviews(notebook: List[Dict[str, Any]], today: Optional[date] = None) -> List[Dict[str, Any]]:
    today = today or date.today()
    out = []
    for item in notebook:
        if item.get("done"):
            continue
        d = next_review_date(item)
        if d is not None and d <= today:
            out.append(item)
    return out


def advance_review(item: Dict[str, Any], answered_correctly: bool, today: Optional[date] = None) -> None:
    """Doğru: bir sonraki aşamaya geçer, 3 aşama bitince 'done'.
    Yanlış: takvim bugünden itibaren yeniden başlar."""
    today = today or date.today()
    if answered_correctly:
        item["stage"] = int(item.get("stage", 0)) + 1
        if item["stage"] >= len(REVIEW_GAPS):
            item["done"] = True
    else:
        item["stage"] = 0
        item["done"] = False
        for g in REVIEW_GAPS:
            item[f"review_{g}d"] = (today + timedelta(days=g)).strftime("%d.%m.%Y")


_LEADING_NUM = re.compile(r"^\s*(?:#{1,6}\s*)?(?:\*\*)?\s*(?:Soru\s*)?\d+\s*[\.\)\:\-]?\s*(?:\*\*)?", re.IGNORECASE)


def build_review_exam(items: List[Dict[str, Any]]) -> Tuple[str, Dict[int, str]]:
    """Hata defteri kayıtlarından yeniden numaralandırılmış bir mini deneme metni üretir.
    Dönen sözlük: yeni soru numarası -> hata defteri kayıt id'si."""
    parts: List[str] = []
    key_lines: List[str] = []
    mapping: Dict[int, str] = {}
    for i, item in enumerate(items, start=1):
        body = _LEADING_NUM.sub("", str(item.get("question_text", "")).strip(), count=1).strip()
        parts.append(f"Soru {i}\n{body}")
        key_lines.append(f"{i}-{item.get('correct_ans', '')}")
        mapping[i] = str(item.get("id"))
    text = "\n\n".join(parts) + "\n\nCevap Anahtarı\n" + "\n".join(key_lines)
    return text, mapping


# ------------------------------------------------------------------------------
# Uyarlanabilir zorluk
# ------------------------------------------------------------------------------
def suggest_difficulty(recent_accuracies: List[float], current: str = "Orta-Zor") -> Tuple[str, str]:
    """Son (en fazla) 3 denemenin ortalamasına göre bir üst/alt seviyeyi önerir."""
    if current not in DIFFICULTY_LEVELS:
        current = "Orta-Zor"
    vals = [float(a) for a in recent_accuracies[:3]]
    if not vals:
        return current, "Henüz deneme sonucun yok; varsayılan seviye kullanılıyor."
    avg = sum(vals) / len(vals)
    idx = DIFFICULTY_LEVELS.index(current)
    if avg >= 85 and idx < len(DIFFICULTY_LEVELS) - 1:
        return DIFFICULTY_LEVELS[idx + 1], f"Son denemelerinde ortalama %{avg:.0f} doğruluk: bir üst seviye önerilir."
    if avg < 55 and idx > 0:
        return DIFFICULTY_LEVELS[idx - 1], f"Son denemelerinde ortalama %{avg:.0f} doğruluk: bir alt seviyeyle pekiştir."
    return current, f"Son denemelerinde ortalama %{avg:.0f} doğruluk: seviye sana uygun."


gamify = SimpleNamespace(
    BADGES=BADGES,
    DIFFICULTY_LEVELS=DIFFICULTY_LEVELS,
    REVIEW_GAPS=REVIEW_GAPS,
    advance_review=advance_review,
    build_review_exam=build_review_exam,
    due_reviews=due_reviews,
    earned_badges=earned_badges,
    next_review_date=next_review_date,
    streak=streak,
    suggest_difficulty=suggest_difficulty,
)


# ==============================================================================
# RAPORLAR: veli raporu, e-posta, optik formlu deneme kitapçığı
# ==============================================================================
INK = RGBColor(15, 23, 42)
MUTED = RGBColor(100, 116, 139)


# ------------------------------------------------------------------------------
# E-posta
# ------------------------------------------------------------------------------
def smtp_ready(cfg: Optional[Dict[str, Any]]) -> bool:
    return bool(cfg and cfg.get("host") and cfg.get("user") and cfg.get("password"))


def send_email(cfg: Dict[str, Any], to: str, subject: str, body: str,
               attachments: Sequence[Tuple[str, bytes, str]] = ()) -> None:
    """cfg: host, port (587), user, password, sender (isteğe bağlı).
    attachments: (dosya adı, içerik, 'ana/alt' mime türü)"""
    msg = EmailMessage()
    msg["From"] = cfg.get("sender") or cfg["user"]
    msg["To"] = to
    msg["Subject"] = subject
    msg.set_content(body)
    for filename, data, mime in attachments:
        maintype, subtype = mime.split("/", 1)
        msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=filename)
    port = int(cfg.get("port", 587))
    ctx = ssl.create_default_context()
    if port == 465:
        with smtplib.SMTP_SSL(cfg["host"], port, context=ctx, timeout=20) as s:
            s.login(cfg["user"], cfg["password"])
            s.send_message(msg)
    else:
        with smtplib.SMTP(cfg["host"], port, timeout=20) as s:
            s.starttls(context=ctx)
            s.login(cfg["user"], cfg["password"])
            s.send_message(msg)


def weekly_report_text(brand: str, parent_name: str, student_name: str, stats: Dict[str, Any],
                       streak_days: int, new_badges: List[str], weak_topic: str = "",
                       teacher_line: str = "") -> str:
    lines = [
        f"Sayın {parent_name},",
        "",
        f"{student_name} için bu haftanın özeti:",
        "",
        f"• Çözülen soru: {stats.get('weekly_solved', 0)}",
        f"• Doğru / Yanlış / Boş: {stats.get('correct', 0)} / {stats.get('wrong', 0)} / {stats.get('blank', 0)}",
        f"• Doğruluk oranı: %{stats.get('accuracy', 0)}",
        f"• Çalıştığı gün sayısı: {stats.get('active_days', 0)} / 7",
        f"• Güncel çalışma serisi: {streak_days} gün",
    ]
    if new_badges:
        lines.append(f"• Kazandığı rozetler: {', '.join(new_badges)}")
    if weak_topic:
        lines += ["", f"Pekiştirilmesi önerilen konu: {weak_topic}"]
    if int(stats.get("weekly_solved", 0)) == 0:
        lines += ["", "Bu hafta platformda çalışma kaydı görünmüyor. Kısa da olsa günlük bir rutin önerilir."]
    lines += ["", "Ayrıntılar için Veli Takip Paneli'ne göz atabilirsiniz.", ""]
    if teacher_line:
        lines.append(f"Öğretmen: {teacher_line}")
    lines.append(brand)
    return "\n".join(lines)


# ------------------------------------------------------------------------------
# Basılabilir deneme kitapçığı + optik form
# ------------------------------------------------------------------------------
def _margins(doc: Document, inch: float = 0.7) -> None:
    for s in doc.sections:
        s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = Inches(inch)


def _p(doc: Document, text: str, size: float = 11, bold: bool = False, color: RGBColor = INK,
       align: Optional[int] = None, space_after: float = 4) -> None:
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.bold = bold
    r.font.color.rgb = color


def build_exam_booklet_docx(questions: List[Dict[str, Any]], title: str, meta: str,
                            logo_png: Optional[bytes] = None, include_key: bool = True) -> bytes:
    """questions: app'teki parse_quiz_questions çıktısı (num, stem, options, correct)."""
    doc = Document()
    _margins(doc)
    if logo_png:
        try:
            lp = doc.add_paragraph()
            lp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            lp.add_run().add_picture(io.BytesIO(logo_png), width=Inches(0.9))
        except Exception:
            pass
    _p(doc, title.upper(), 15, True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _p(doc, meta, 10, color=MUTED, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)
    _p(doc, "Ad Soyad: ______________________   Sınıf/No: __________   Tarih: ____________", 10.5, space_after=14)

    letters = sorted({k for q in questions for k in q.get("options", {})}) or ["A", "B", "C", "D"]

    for i, q in enumerate(questions, start=1):
        _p(doc, f"{i}. {q.get('stem', '').strip()}", 11, bold=False, space_after=3)
        for letter, text in q.get("options", {}).items():
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0.3)
            p.paragraph_format.space_after = Pt(1)
            r = p.add_run(f"{letter}) {text}")
            r.font.size = Pt(10.5)
        doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # Optik form
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    _p(doc, "CEVAP KÂĞIDI (OPTİK FORM)", 14, True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _p(doc, "Doğru seçeneğin dairesini tamamen doldurun. Silinti bırakmayın.", 9.5, color=MUTED,
       align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)
    _p(doc, "Ad Soyad: ______________________   Sınıf/No: __________", 10.5, space_after=10)

    n = len(questions)
    half = (n + 1) // 2
    cols_per_block = 1 + len(letters)
    table = doc.add_table(rows=half + 1, cols=cols_per_block * 2 + 1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    header = ["No"] + letters
    for block in range(2):
        for j, h in enumerate(header):
            cell = table.cell(0, block * (cols_per_block + 1) + j)
            cell.text = h
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell.paragraphs[0].runs[0].bold = True
    for i in range(n):
        block, row = (0, i + 1) if i < half else (1, i - half + 1)
        base = block * (cols_per_block + 1)
        table.cell(row, base).text = str(i + 1)
        for j in range(len(letters)):
            c = table.cell(row, base + 1 + j)
            c.text = "◯"
            c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    # ortadaki boşluk sütunu
    for r in range(half + 1):
        table.cell(r, cols_per_block).text = ""

    if include_key and any(q.get("correct") for q in questions):
        doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
        _p(doc, "CEVAP ANAHTARI (öğretmen nüshası)", 13, True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)
        key = "   ".join(f"{i}-{q.get('correct') or '?'}" for i, q in enumerate(questions, start=1))
        _p(doc, key, 11)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


reports = SimpleNamespace(
    INK=INK,
    MUTED=MUTED,
    build_exam_booklet_docx=build_exam_booklet_docx,
    send_email=send_email,
    smtp_ready=smtp_ready,
    weekly_report_text=weekly_report_text,
)


# ==============================================================================
# SAYFA YAPILANDIRMASI VE KURUMSAL SABİTLER (2027 KURGUSU)
# ==============================================================================
# Streamlit'in kendi bileşenlerinin (sekme, kaydırıcı, onay kutusu) rengi: kırmızı yerine marka mavisi.
# Kalıcı ve kesin yol .streamlit/config.toml'dur; bu satır o dosya yoksa da çalışmayı dener.
try:
    from streamlit import config as _st_config

    if not os.path.exists(".streamlit/config.toml"):
        _st_config.set_option("theme.primaryColor", "#1D4ED8")
except Exception:
    pass

st.set_page_config(
    page_title="KURŞUN KALEM DİJİTAL OFİS • 2027",
    page_icon="✏️",
    layout="wide",
    initial_sidebar_state="expanded",
)

BRAND_NAME = "KURŞUN KALEM DİJİTAL OFİS"
# İletişim: kendi numaranı yazana kadar boş bırak; boşsa sitede telefon gösterilmez.
SUPPORT_PHONE = ""  # örn. "0532 000 00 00"
SUPPORT_EMAIL = "destek@kursunkalemdijitalofis.com.tr"


def _support_phone_html() -> str:
    if not SUPPORT_PHONE:
        return ""
    digits = "".join(ch for ch in SUPPORT_PHONE if ch.isdigit() or ch == "+")
    return f'<a href="tel:{digits}">{html.escape(SUPPORT_PHONE)}</a>'


def _support_line_html() -> str:
    parts = [_support_phone_html(), f'<a href="mailto:{SUPPORT_EMAIL}">{SUPPORT_EMAIL}</a>' if SUPPORT_EMAIL else ""]
    parts = [p for p in parts if p]
    return ("Destek: " + " · ".join(parts)) if parts else ""
BRAND_SHORT = "Kurşun Kalem"
BRAND_TAGLINE = "2027 Hedefli Soru & Redaksiyon Fabrikası • Dijimentor • Veli & B2B Dijital Ofis"
DEFAULT_MODEL = "gemini-2.5-flash"
FALLBACK_MODEL = "gemini-3.5-flash-lite"
STATE_FILE_PATH = Path(".kursun_kalem_ofis_2028_state.json")
STATE_BACKUP_PATH = Path(".kursun_kalem_ofis_2028_state.bak.json")
EXTERNAL_CSS_PATH = Path("assets/style.css")

GRADE_LEVELS_2027: List[str] = [
    "5. Sınıf",
    "6. Sınıf",
    "7. Sınıf",
    "8. Sınıf (LGS)",
    "9. Sınıf",
    "10. Sınıf",
    "11. Sınıf (YKS)",
    "12. Sınıf (YKS)",
]

# Logo: mavi rozet üzerine onay işareti çizen sarı kurşun kalem
EMBEDDED_LOGO_SVG = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64" width="512" height="512">
  <rect x="3" y="3" width="58" height="58" rx="16" fill="#1D4ED8" stroke="#0F172A" stroke-width="3"/>
  <path d="M14 34 L23 43 L37 27" fill="none" stroke="#FDE047" stroke-width="6" stroke-linecap="round" stroke-linejoin="round"/>
  <g transform="translate(37.5 26.5) rotate(45) scale(0.64)">
    <polygon points="0,0 -6,-11 6,-11" fill="#FDE7B0" stroke="#0F172A" stroke-width="2.6" stroke-linejoin="round"/>
    <polygon points="0,0 -2.6,-4.8 2.6,-4.8" fill="#0F172A"/>
    <rect x="-6" y="-33" width="12" height="22" fill="#FACC15" stroke="#0F172A" stroke-width="2.6"/>
    <line x1="0" y1="-33" x2="0" y2="-11" stroke="#EAB308" stroke-width="2"/>
    <rect x="-6" y="-38" width="12" height="5" fill="#CBD5E1" stroke="#0F172A" stroke-width="2.6"/>
    <rect x="-6" y="-45" width="12" height="7" rx="3" fill="#F472B6" stroke="#0F172A" stroke-width="2.6"/>
  </g>
</svg>"""

SCHOOL_BACKPACK_SVG = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 260 300" width="260" height="300">
  <defs>
    <linearGradient id="bagBody" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#172c56"/>
      <stop offset="100%" stop-color="#0f172a"/>
    </linearGradient>
    <linearGradient id="bagPocket" x1="0%" y1="0%" x2="0%" y2="100%">
      <stop offset="0%" stop-color="#334155"/>
      <stop offset="100%" stop-color="#1e293b"/>
    </linearGradient>
  </defs>
  <path d="M 98 56 Q 130 20 162 56" fill="none" stroke="#0f172a" stroke-width="10" stroke-linecap="round"/>
  <path d="M 88 64 C 58 95, 54 165, 66 245" fill="none" stroke="#334155" stroke-width="15" stroke-linecap="round"/>
  <path d="M 172 64 C 202 95, 206 165, 194 245" fill="none" stroke="#334155" stroke-width="15" stroke-linecap="round"/>
  <rect x="64" y="56" width="132" height="205" rx="32" fill="url(#bagBody)" stroke="#0f172a" stroke-width="3"/>
  <path d="M 68 92 Q 130 114 192 92 L 192 78 Q 130 56 68 78 Z" fill="#d4af37" opacity="0.28"/>
  <rect x="82" y="138" width="96" height="96" rx="20" fill="url(#bagPocket)" stroke="#d4af37" stroke-width="2"/>
  <circle cx="130" cy="186" r="22" fill="#ffffff" stroke="#0f172a" stroke-width="2"/>
  <text x="130" y="184" text-anchor="middle" font-family="Inter, Arial, sans-serif" font-size="8.5" font-weight="800" fill="#0f172a">OFİS</text>
  <text x="130" y="196" text-anchor="middle" font-family="Inter, Arial, sans-serif" font-size="7.5" font-weight="700" fill="#475569">ÇANTASI</text>
  <rect x="72" y="242" width="116" height="18" rx="9" fill="#090d16"/>
</svg>"""

HERO_SHOWCASE_SVG = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 440 330" width="440" height="330">
  <defs>
    <linearGradient id="bgCardGrad" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#EEF2FF"/>
      <stop offset="100%" stop-color="#E0E7FF"/>
    </linearGradient>
    <linearGradient id="navyBanner" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#1E2952"/>
      <stop offset="100%" stop-color="#2563EB"/>
    </linearGradient>
    <linearGradient id="goldBadgeGrad" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#FACC15"/>
      <stop offset="100%" stop-color="#D97706"/>
    </linearGradient>
    <linearGradient id="pencilWoodGrad" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#FDE047"/>
      <stop offset="100%" stop-color="#EAB308"/>
    </linearGradient>
  </defs>

  <rect x="22" y="18" width="396" height="292" rx="24" fill="url(#bgCardGrad)" stroke="#C7D2FE" stroke-width="1.5"/>

  <rect x="44" y="42" width="214" height="244" rx="18" fill="#FFFFFF" stroke="#E2E8F0" stroke-width="1.5"/>
  <rect x="44" y="42" width="214" height="58" rx="18" fill="url(#navyBanner)"/>
  <rect x="44" y="76" width="214" height="24" fill="url(#navyBanner)"/>
  <text x="62" y="68" font-family="Inter, Arial, sans-serif" font-size="10.5" font-weight="800" fill="#FACC15" letter-spacing="1.2">KURŞUN KALEM • 2027</text>
  <text x="62" y="87" font-family="Inter, Arial, sans-serif" font-size="13" font-weight="800" fill="#FFFFFF">Ölçme &amp; Redaksiyon Ofisi</text>

  <rect x="60" y="114" width="182" height="42" rx="10" fill="#F8FAFC" stroke="#E2E8F0"/>
  <circle cx="79" cy="135" r="8" fill="#2563EB"/>
  <text x="94" y="132" font-family="Inter, Arial, sans-serif" font-size="10.5" font-weight="700" fill="#1E2952">Maarif Kazanım Uyumu</text>
  <text x="94" y="146" font-family="Inter, Arial, sans-serif" font-size="9.5" font-weight="600" fill="#64748B">LGS &amp; YKS 2027 Müfredatı</text>

  <rect x="60" y="164" width="182" height="42" rx="10" fill="#F8FAFC" stroke="#E2E8F0"/>
  <circle cx="79" cy="185" r="8" fill="#4F46E5"/>
  <text x="94" y="182" font-family="Inter, Arial, sans-serif" font-size="10.5" font-weight="700" fill="#1E2952">Çeldirici &amp; Kök Denetimi</text>
  <text x="94" y="196" font-family="Inter, Arial, sans-serif" font-size="9.5" font-weight="600" fill="#64748B">Çift Katmanlı Editör Kontrolü</text>

  <rect x="60" y="214" width="182" height="52" rx="10" fill="#EFF6FF" stroke="#BFDBFE"/>
  <text x="74" y="233" font-family="Inter, Arial, sans-serif" font-size="9.5" font-weight="700" fill="#1D4ED8">HEDEF NET &amp; İSABET İVMESİ</text>
  <rect x="74" y="243" width="152" height="8" rx="4" fill="#DBEAFE"/>
  <rect x="74" y="243" width="124" height="8" rx="4" fill="#2563EB"/>

  <rect x="276" y="42" width="120" height="114" rx="18" fill="#FFFFFF" stroke="#E2E8F0" stroke-width="1.5"/>
  <circle cx="336" cy="82" r="22" fill="url(#goldBadgeGrad)"/>
  <polygon points="336,68 340,77 350,78 342,85 345,95 336,90 327,95 330,85 322,78 332,77" fill="#1E2952"/>
  <text x="336" y="121" text-anchor="middle" font-family="Inter, Arial, sans-serif" font-size="11" font-weight="800" fill="#1E2952">2027 HEDEF</text>
  <text x="336" y="136" text-anchor="middle" font-family="Inter, Arial, sans-serif" font-size="9.5" font-weight="600" fill="#64748B">Dijimentor &amp; Röntgen</text>

  <rect x="276" y="170" width="120" height="116" rx="18" fill="#1E2952"/>
  <g transform="translate(336, 214) rotate(-32)">
    <rect x="-8" y="-34" width="16" height="52" rx="2" fill="url(#pencilWoodGrad)"/>
    <rect x="-8" y="-38" width="16" height="6" rx="1.5" fill="#F43F5E"/>
    <polygon points="-8,18 8,18 0,34" fill="#FDE68A"/>
    <polygon points="-2.5,28 2.5,28 0,34" fill="#0F172A"/>
  </g>
  <text x="336" y="266" text-anchor="middle" font-family="Inter, Arial, sans-serif" font-size="10" font-weight="800" fill="#FACC15" letter-spacing="0.8">KURŞUN KALEM</text>
  <text x="336" y="279" text-anchor="middle" font-family="Inter, Arial, sans-serif" font-size="8.5" font-weight="600" fill="#E2E8F0">Dijital Ofis Vitrini</text>
</svg>"""


def embed_html(markup: str, height: int) -> None:
    """HTML/script gömme: yeni Streamlit'te st.iframe, eskisinde components.html."""
    frame = getattr(st, "iframe", None)
    if frame is not None:
        frame(markup, height=max(1, int(height)))
    else:
        import streamlit.components.v1 as components

        components.html(markup, height=height, scrolling=True)


def tr_upper(text: str) -> str:
    """Türkçe büyük harf: i -> İ, ı -> I (CSS'in text-transform'u bunu bilmez)."""
    return str(text).replace("i", "İ").replace("ı", "I").upper()


def svg_to_data_uri(svg_str: str) -> str:
    encoded = base64.b64encode(svg_str.encode("utf-8")).decode("ascii")
    return f"data:image/svg+xml;base64,{encoded}"


def get_embedded_logo_png_bytes() -> Optional[bytes]:
    try:
        import cairosvg  # type: ignore

        png = cairosvg.svg2png(bytestring=EMBEDDED_LOGO_SVG.encode("utf-8"), scale=2.0)
        if png:
            return png
    except Exception:
        pass

    try:
        from PIL import Image, ImageDraw  # type: ignore

        size = 320
        img = Image.new("RGBA", (size, size), (255, 255, 255, 0))
        draw = ImageDraw.Draw(img)
        draw.rounded_rectangle([15, 15, size - 15, size - 15], radius=80, fill=(29, 78, 216, 255),
                               outline=(15, 23, 42, 255), width=15)
        draw.line([(70, 170), (115, 215), (185, 135)], fill=(253, 224, 71, 255), width=30, joint="curve")
        out = io.BytesIO()
        img.save(out, format="PNG")
        return out.getvalue()
    except Exception:
        return None


# Resmi tarihler: MEB 2027 Sınav Uygulama Takvimi (LGS) ve ÖSYM 2027 takvimi (YKS).
# Başlama saatleri henüz açıklanmadı; sayaç için tahmini saat kullanılıyor.
LGS_TARGET = datetime(2027, 6, 13, 9, 30, 0)
YKS_TARGET = datetime(2027, 6, 19, 10, 15, 0)
# Açılış sayfasındaki rakamlar: yalnızca gerçek ve doğrulanabilir değerler yaz.
# Boş bırakılırsa rakam yerine özellik maddeleri gösterilir.
LANDING_STATS: List[Tuple[str, str]] = []  # örn. [("1.250", "kayıtlı öğrenci")]

PAGES = [
    "Ana Sayfa",
    "Soru & Redaksiyon Fabrikası",
    "Öğrenci Portalı & Çözüm",
    "Dijimentor",
    "Veli Takip Paneli",
    "Öğrenci Röntgeni",
    "B2B Lisans & Kurum",
    "Paketler & Abonelik Vitrini",
    "Sınıflarım",
    "Profilim & Rozetler",
    "Fotoğrafla Çöz",
    "Soru Bankam",
]

# Hangi rol hangi sayfaları görür
ROLE_PAGES: Dict[str, List[str]] = {
    "ogrenci": [
        "Ana Sayfa", "Soru & Redaksiyon Fabrikası", "Öğrenci Portalı & Çözüm", "Fotoğrafla Çöz", "Dijimentor",
        "Öğrenci Röntgeni", "Paketler & Abonelik Vitrini", "Profilim & Rozetler",
    ],
    "veli": ["Ana Sayfa", "Veli Takip Paneli", "Öğrenci Röntgeni", "Paketler & Abonelik Vitrini", "Profilim & Rozetler"],
    "ogretmen": [
        "Ana Sayfa", "Soru & Redaksiyon Fabrikası", "Soru Bankam", "Sınıflarım", "Öğrenci Röntgeni",
        "B2B Lisans & Kurum", "Paketler & Abonelik Vitrini", "Profilim & Rozetler",
    ],
}

# Kayıt olmadan denemek isteyenler için ücretsiz başlangıç kotası (iş kararı: değiştirebilirsin)
FREE_PLAN: Dict[str, Any] = {
    "name": "Ücretsiz Başlangıç",
    "segment": "Deneme",
    "badge": "Ücretsiz",
    "monthly_price": 0,
    "monthly_quota": 50,
    "coach_quota": 3,
    "total_count": 0,
}

PAGE_DESCRIPTIONS: Dict[str, Tuple[str, str]] = {
    "Ana Sayfa": (
        "Genel Bakış, 2027 Kademe Vitrini & Kişisel Akış",
        "Kademe seçimi, eksik kazanım öncelik sırası, 2027 LGS ve YKS geri sayımı ile dijital ofis özeti.",
    ),
    "Soru & Redaksiyon Fabrikası": (
        "Maarif Modeli Soru & Kalite Stüdyosu",
        "5–12. sınıf müfredatı, kademeli zorluk serisi, çeldirici denetimi, soru havuzu ve DOCX çıktı.",
    ),
    "Öğrenci Portalı & Çözüm": (
        "İnteraktif Mini Deneme, Hata Defteri & Çözüm Rehberi",
        "Süre takipli mini deneme çözün, netinizi görün, yanlışları Hata Defteri'ne alın ve adımlı çözüm inceleyin.",
    ),
    "Dijimentor": (
        "Akıllı Sınav, Faz Planı & Performans Mentorluğu",
        "2027 sınav takvimine göre aşamalı kamp ve haftalık aksiyon planı üretin.",
    ),
    "Veli Takip Paneli": (
        "Veli Karne Özeti & Bilgilendirme Merkezi",
        "Haftalık değişim kartları, kazanım ivmesi ve hazır veli bilgilendirme mesajları.",
    ),
    "Öğrenci Röntgeni": (
        "Öğrenci Röntgeni & Akademik Check-Up",
        "Ödev ve deneme sonuçlarından kazanım bazlı başarı, sınıf karşılaştırması ve telafi önerisi.",
    ),
    "B2B Lisans & Kurum": (
        "Kurumsal Lisans, Zümre & Deneme Planlayıcı",
        "Koltuk/kota planlaması, zümre dağılımı, kurum içi deneme takvimi ve DOCX teklif çıktısı.",
    ),
    "Fotoğrafla Çöz": (
        "Fotoğrafla Çöz",
        "Kitaptaki ya da defterdeki sorunun fotoğrafını yükle; Dijimentor önce ipucu versin ya da adım adım çözsün.",
    ),
    "Soru Bankam": (
        "Soru Bankam: Taslak → Onay → Ödev",
        "Yapay zekânın ürettiği sorular önce taslak olarak gelir. Siz düzeltip onaylarsınız; öğrenciye yalnızca onaylı soru gider.",
    ),
    "Sınıflarım": (
        "Sınıflarım, Ödevler & Sınıf Analizi",
        "Sınıf oluşturun, öğrencileri kodla davet edin, test atayın ve kazanım bazlı sınıf karşılaştırmasını görün.",
    ),
    "Profilim & Rozetler": (
        "Profilim, Seri & Rozetler",
        "Günlük hedefini ayarla, serini ve rozetlerini gör, veli bağlantı kodunu paylaş.",
    ),
    "Paketler & Abonelik Vitrini": (
        "Abonelik Paketleri & Ofis Çantası",
        "Öğrenci, 2027 YKS, Veli ve Kurumsal paketleri çantanıza ekleyin, kotaları yönetin ve iyzico planını inceleyin.",
    ),
}

HERO_PILL_CHIPS: List[Dict[str, str]] = [
    {"label": "YKS 2027", "grade": "12. Sınıf (YKS)", "exam": "2027 TYT - AYT", "page": "Dijimentor"},
    {"label": "LGS 2027", "grade": "8. Sınıf (LGS)", "exam": "2027 LGS", "page": "Soru & Redaksiyon Fabrikası"},
    {"label": "11. Sınıf", "grade": "11. Sınıf (YKS)", "exam": "2027 TYT - AYT", "page": "Soru & Redaksiyon Fabrikası"},
    {"label": "10. Sınıf", "grade": "10. Sınıf", "exam": "2027 TYT - AYT", "page": "Soru & Redaksiyon Fabrikası"},
    {"label": "9. Sınıf", "grade": "9. Sınıf", "exam": "2027 TYT - AYT", "page": "Soru & Redaksiyon Fabrikası"},
    {"label": "7. Sınıf", "grade": "7. Sınıf", "exam": "5-7. Sınıf Akademik Gelişim", "page": "Soru & Redaksiyon Fabrikası"},
    {"label": "6. Sınıf", "grade": "6. Sınıf", "exam": "5-7. Sınıf Akademik Gelişim", "page": "Soru & Redaksiyon Fabrikası"},
    {"label": "5. Sınıf", "grade": "5. Sınıf", "exam": "5-7. Sınıf Akademik Gelişim", "page": "Soru & Redaksiyon Fabrikası"},
    {"label": "Veli & Kurum", "grade": "8. Sınıf (LGS)", "exam": "2027 LGS", "page": "Veli Takip Paneli"},
    {"label": "Zümre & B2B", "grade": "8. Sınıf (LGS)", "exam": "2027 LGS", "page": "B2B Lisans & Kurum"},
]

KADEME_VITRIN_KARTLARI: List[Dict[str, str]] = [
    {
        "tag": "5–7. SINIF • 2027",
        "cta": "5–7. Sınıf • Aç",
        "title": "Ortaokul & Maarif Temeli",
        "desc": "Kazanım odaklı sorular, okuma hızı ve düzenli çalışma alışkanlığı.",
        "accent": "#2563EB",
        "grade": "7. Sınıf",
        "exam": "5-7. Sınıf Akademik Gelişim",
        "target_page": "Soru & Redaksiyon Fabrikası",
    },
    {
        "tag": "2027 LGS",
        "cta": "2027 LGS • Aç",
        "title": "8. Sınıf LGS Hazırlık",
        "desc": "Sözel-sayısal muhakeme, çeldirici analizi ve net yükselten deneme akışı.",
        "accent": "#4F46E5",
        "grade": "8. Sınıf (LGS)",
        "exam": "2027 LGS",
        "target_page": "Soru & Redaksiyon Fabrikası",
    },
    {
        "tag": "2027 YKS",
        "cta": "2027 YKS • Aç",
        "title": "TYT & AYT Lise Kampı",
        "desc": "Süre yönetimi, paragraf-analitik denge ve derece hedefleyen mentorluk.",
        "accent": "#7C3AED",
        "grade": "12. Sınıf (YKS)",
        "exam": "2027 TYT - AYT",
        "target_page": "Dijimentor",
    },
    {
        "tag": "VELİ & KURUM • 2027",
        "cta": "Veli & Kurum • Aç",
        "title": "Veli Takip & B2B Ofis",
        "desc": "Haftalık çalışma-soru özeti, veli mesajları ve zümre kota yönetimi.",
        "accent": "#0284C7",
        "grade": "8. Sınıf (LGS)",
        "exam": "2027 LGS",
        "target_page": "Veli Takip Paneli",
    },
]

# Ana sayfadaki sınıf seçici: sekme -> o sekmedeki seçenekler (HERO_PILL_CHIPS etiketleri)
GRADE_GROUPS: Dict[str, List[str]] = {
    "Ortaokul": ["5. Sınıf", "6. Sınıf", "7. Sınıf", "LGS 2027"],
    "Lise": ["9. Sınıf", "10. Sınıf", "11. Sınıf", "YKS 2027"],
    "Veli & Kurum": ["Veli & Kurum", "Zümre & B2B"],
}
CHIP_BY_LABEL: Dict[str, Dict[str, str]] = {c["label"]: c for c in HERO_PILL_CHIPS}

# Sol menü: (başlık, [(sayfa adı, menüde görünen kısa ad, ikon), ...])
NAV_GROUPS: List[Tuple[Optional[str], List[Tuple[str, str, str]]]] = [
    (None, [("Ana Sayfa", "Ana Sayfa", ":material/home:")]),
    ("Modüller", [
        ("Soru & Redaksiyon Fabrikası", "Soru & Redaksiyon", ":material/edit_note:"),
        ("Soru Bankam", "Soru Bankam", ":material/inventory_2:"),
        ("Öğrenci Portalı & Çözüm", "Öğrenci Portalı", ":material/quiz:"),
        ("Fotoğrafla Çöz", "Fotoğrafla Çöz", ":material/photo_camera:"),
        ("Dijimentor", "Dijimentor", ":material/event_note:"),
        ("Öğrenci Röntgeni", "Öğrenci Röntgeni", ":material/monitor_heart:"),
    ]),
    ("Okul & Veli", [
        ("Sınıflarım", "Sınıflarım", ":material/groups:"),
        ("Veli Takip Paneli", "Veli Takip Paneli", ":material/family_restroom:"),
        ("B2B Lisans & Kurum", "B2B Lisans & Kurum", ":material/apartment:"),
    ]),
    ("Hesap", [
        ("Profilim & Rozetler", "Profilim & Rozetler", ":material/military_tech:"),
        ("Paketler & Abonelik Vitrini", "Paketler", ":material/shopping_bag:"),
    ]),
]

# Ana sayfadaki modül kartları: (kısaltma, başlık, açıklama, hedef sayfa, arka plan, yazı rengi)
MODULE_CARDS: List[Tuple[str, str, str, str, str, str]] = [
    ("SR", "Soru & Redaksiyon", "Kazanım etiketli soru üretimi; yayın öncesi öğretmen kontrolü için düzenlenebilir.", "Soru & Redaksiyon Fabrikası", "#EFF6FF", "#1D4ED8"),
    ("DM", "Dijimentor", "Seviyene göre her hafta yenilenen kişisel çalışma planı.", "Dijimentor", "#F0FDF4", "#15803D"),
    ("MD", "Mini Deneme", "Süre takipli denemeler ve otomatik hata defteri.", "Öğrenci Portalı & Çözüm", "#FEF3C7", "#B45309"),
    ("FC", "Fotoğrafla Çöz", "Sorunun fotoğrafını çek, ipucu al ya da adım adım çözümünü gör.", "Fotoğrafla Çöz", "#ECFEFF", "#0E7490"),
    ("ÖR", "Öğrenci Röntgeni", "Konu konu güçlü ve zayıf yönlerin net haritası.", "Öğrenci Röntgeni", "#F5F3FF", "#6D28D9"),
    ("VT", "Veli Takip", "Haftalık özet, ilerleme ve veli mesajları tek yerde.", "Veli Takip Paneli", "#FFF1F2", "#BE123C"),
]

# Modül kartlarındaki çizgi ikonlar (kısaltma -> SVG içi)
_SVG = '<svg aria-hidden="true" width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">{}</svg>'
MODULE_ICONS: Dict[str, str] = {
    "SR": _SVG.format('<path d="M17 3l4 4L8 20H4v-4L17 3z"/><path d="M14 6l4 4"/>'),
    "DM": _SVG.format('<rect x="3" y="4" width="18" height="17" rx="3"/><path d="M3 9h18M8 2v4M16 2v4M8 14l2 2 4-4"/>'),
    "MD": _SVG.format('<circle cx="12" cy="13" r="8"/><path d="M12 9v4l2.5 2.5M9 2h6"/>'),
    "ÖR": _SVG.format('<path d="M3 12h4l3-8 4 16 3-8h4"/>'),
    "VT": _SVG.format('<circle cx="9" cy="8" r="3"/><circle cx="17" cy="10" r="2.5"/><path d="M3 20c0-3 3-5 6-5s6 2 6 5M15 20c0-2 1.5-3.5 4-3.5"/>'),
    "FC": _SVG.format('<path d="M4 8h3l2-3h6l2 3h3a1 1 0 0 1 1 1v10a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1V9a1 1 0 0 1 1-1z"/><circle cx="12" cy="13.5" r="3.5"/>'),
    "SE": _SVG.format('<path d="M12 3c1 3.5 5 5.5 5 10a5 5 0 0 1-10 0c0-2.5 1.5-4 2.5-5 .3 1.8 1.2 2.8 2.5 3.2C11.5 8.5 11 6 12 3z"/>'),
    "SO": _SVG.format('<rect x="3" y="4" width="18" height="12" rx="2"/><path d="M7 20h10M12 16v4M7 9h6M7 12h4"/>'),
}

# Eğlenceli tema açık/kapalı (False yaparsan sade görünüme döner)
FUN_THEME = True

# ==============================================================================
# 5–12. SINIF TAM MÜFREDAT MOTORU (MAARİF MODELİ, LGS & YKS)
# ==============================================================================
_MIDDLE_COMMON: Dict[str, Dict[str, List[str]]] = {
    "Türkçe": {
        "Okuma ve Anlama": [
            "Bağlamdan yararlanarak sözcük ve söz öbeklerinin anlamını tahmin etme",
            "Metindeki ana düşünce, yardımcı düşünceler ve başlık ilişkisini kurma",
            "Sözel mantık, tablo/grafik yorumlama ve muhakeme gerektiren metinleri çözümleme",
            "Metinler arası karşılaştırma yapma ve çıkarımda bulunma",
        ],
        "Dil Bilgisi ve Yazım": [
            "Sözcük türleri, fiilimsiler ve cümle ögeleri işlevini belirleme",
            "Yazım kuralları ve noktalama işaretlerinin anlamı belirlemedeki işlevi",
            "Cümlede anlam ilişkileri ve anlatım akıcılığını denetleme",
        ],
    },
    "Matematik": {
        "Sayılar ve İşlemler": [
            "Doğal sayılar, tam sayılar, kesirler ve ondalık gösterimlerle çok adımlı işlemler",
            "Çarpanlar ve katlar, EBOB-EKOK problem modellemeleri",
            "Üslü ve kareköklü ifadelerde akıl yürütme, tahmin ve cebirsel ifadeler",
        ],
        "Geometri ve Veri Analizi": [
            "Doğrular, açılar, çokgenler, üçgenler ve geometrik cisimler",
            "Veri toplama, grafik yorumlama, olasılık ve merkezi eğilim ölçüleri",
        ],
    },
    "Fen Bilimleri": {
        "Fiziksel Süreçler ve Kuvvet": [
            "Kuvvet, hareket, basınç, iş-enerji ve basit makineler",
            "Mevsimlerin oluşumu, ışık, ses ve elektrik enerjisi uygulamaları",
        ],
        "Madde, Doğa ve Canlılar": [
            "Maddenin yapısı, karışımlar, kimyasal tepkimeler ve asit-baz kavramları",
            "Hücre, sistemler, DNA-genetik kod, kalıtım ve ekoloji/sürdürülebilirlik",
        ],
    },
    "Sosyal Bilgiler / İnkılap Tarihi": {
        "Tarih ve Kültür": [
            "Birey, toplum, kültür ve miras; Millî Mücadele ve Atatürk İlkeleri",
            "Cumhuriyet kazanımları, çağdaş gelişmeler ve tarihsel kanıt okuma",
        ],
        "İnsan, Yer ve Çevre": [
            "Harita okuryazarlığı, nüfus, ekonomik faaliyetler ve çevre bilinci",
            "Haklar, sorumluluklar, demokratik süreçler ve küresel bağlantılar",
        ],
    },
    "İngilizce": {
        "Reading & Comprehension": [
            "Contextual vocabulary, inference and main idea in short texts",
            "Everyday communication, invitations, preferences and plans",
        ],
        "Language & Reasoning": [
            "Tenses, modals, comparatives and cohesive connectors in context",
            "Table/chart interpretation and situational dialogue completion",
        ],
    },
}

_HIGH_COMMON: Dict[str, Dict[str, List[str]]] = {
    "Türk Dili ve Edebiyatı": {
        "Anlama ve Yorumlama": [
            "Nitelikli metinlerde yapı, bakış açısı, anlatım biçimleri ve anlam çözümlemesi",
            "Dönem/zihniyet ilişkisi, edebi sanatlar ve metinler arası etkileşim",
        ],
        "Dil Bilgisi ve Yazım": [
            "Sözcükte ve cümlede anlam, paragraf kurgusu ve çeldirici analizi",
            "Ses bilgisi, yazım kuralları, noktalama işaretleri ve anlatım bozukluğu",
        ],
    },
    "Matematik": {
        "Temel Kavramlar ve Cebir": [
            "Temel kavramlar, bölünebilme, rasyonel/üslü/köklü ifadeler ve mutlak değer",
            "Denklemler, eşitsizlikler, oran-orantı, problemler ve fonksiyonlar",
        ],
        "Analitik, Geometri ve İleri Konular": [
            "Üçgenler, çokgenler, çember-daire, katı cisimler ve analitik geometri",
            "Polinomlar, logaritma, diziler, trigonometri, limit, türev ve integral",
        ],
    },
    "Fizik": {
        "Mekanik ve Hareket": [
            "Fizik bilimine giriş, hareket, Newton hareket yasaları, iş-güç-enerji",
            "İtme-momentum, tork-denge, basit makineler, kaldırma kuvveti ve basınç",
        ],
        "Elektrik, Manyetizma ve Dalgalar": [
            "Elektrostatik, elektrik akımı, manyetizma, elektromanyetik indüksiyon",
            "Dalgalar, optik, çembersel hareket, harmonik hareket ve modern fizik",
        ],
    },
    "Kimya": {
        "Madde ve Tepkimeler": [
            "Kimya bilimi, atom ve periyodik sistem, kimyasal türler arası etkileşimler",
            "Maddenin halleri, mol kavramı, kimyasal hesaplamalar ve tepkime hızı/denge",
        ],
        "Çözeltiler ve Enerji": [
            "Karışımlar, asitler-bazlar-tuzlar, kimya ve enerji, elektrokimya ve organik kimya",
        ],
    },
    "Biyoloji": {
        "Canlıların Temelleri ve Sistemler": [
            "Yaşam bilimi biyoloji, hücre zarından madde geçişi, canlıların sınıflandırılması",
            "Hücre bölünmeleri, kalıtım, ekosistem ekolojisi ve insan fizyolojisi sistemleri",
        ],
        "Genden Proteine ve Enerji": [
            "Nükleik asitler, protein sentezi, fotosentez ve hücresel solunum",
        ],
    },
    "Tarih": {
        "Tarihsel Süreç ve İnkılap": [
            "Tarih ve zaman, insanlığın ilk dönemleri, İslam tarihi ve Türk-İslam devletleri",
            "Osmanlı kurumları ve diplomasi, Kurtuluş Savaşı ve Atatürk İlke ve İnkılapları",
        ],
    },
    "Coğrafya": {
        "Doğal ve Beşeri Sistemler": [
            "Doğa ve insan, coğrafi koordinat ve iklim bilgisi, iç-dış kuvvetler",
            "Nüfus, yerleşme, ekonomik faaliyetler, bölgeler ve küresel çevre sorunları",
        ],
    },
}

FULL_CURRICULUM_2027: Dict[str, Dict[str, Dict[str, List[str]]]] = {
    "5. Sınıf": _MIDDLE_COMMON,
    "6. Sınıf": _MIDDLE_COMMON,
    "7. Sınıf": _MIDDLE_COMMON,
    "8. Sınıf (LGS)": {
        "Türkçe": {
            "Okuma ve Muhakeme": [
                "T.8.1.9. Sözel mantık ve muhakeme gerektiren metinleri analiz etme",
                "Metindeki ana düşünce, çıkarım ve başlık/görsel ilişkisini kurma",
                "Metinler arası karşılaştırma yapma ve yargı/kanıt türlerini ayırt etme",
            ],
            "Dil Bilgisi ve Yazım": [
                "Fiilimsiler (isim-fiil, sıfat-fiil, zarf-fiil) ve cümle ögeleri",
                "Yazım kuralları ve noktalama işaretlerinin anlamı belirlemedeki işlevi",
            ],
        },
        "Matematik": {
            "Sayılar ve Cebir": [
                "Çarpanlar ve katlar, EBOB-EKOK problem modellemeleri",
                "Üslü ve kareköklü ifadelerde akıl yürütme, tahmin ve işlemler",
                "Cebirsel ifadeler, özdeşlikler ve birinci dereceden denklemler",
            ],
            "Geometri, Olasılık ve Veri": [
                "Üçgenler, kenarortay/açıortay/yükseklik, eşlik-benzerlik ve dönüşüm geometrisi",
                "Olasılık ve istatistiksel veri analizi / grafik dönüştürme",
            ],
        },
        "Fen Bilimleri": {
            "2027 LGS Fen Kazanımları": [
                "Mevsimlerin oluşumu ve iklim/hava hareketleri",
                "DNA, genetik kod, mutasyon, modifikasyon, adaptasyon ve biyoteknoloji",
                "Basınç (katı, sıvı, gaz) ve madde-endüstri uygulamaları",
                "Basit makineler, enerji dönüşümleri, elektrik yükleri ve elektrik enerjisi",
            ],
        },
        "T.C. İnkılap Tarihi ve Atatürkçülük": {
            "Millî Mücadele ve Çağdaş Türkiye": [
                "Bir Kahraman Doğuyor ve Millî Uyanış: Bağımsızlık Yolunda Atılan Adımlar",
                "Millî Bir Destan: Ya İstiklal Ya Ölüm ve Atatürkçülük/Çağdaşlaşan Türkiye",
                "Demokratikleşme çabaları ve Atatürk'ten Sonraki Dış Politika",
            ],
        },
        "İngilizce": {
            "2027 LGS İngilizce Temaları": [
                "Friendship, Teen Life, In the Kitchen, On the Phone, The Internet",
                "Tourism, Chores, Science, Natural Forces & Reading Comprehension",
            ],
        },
    },
    "9. Sınıf": _HIGH_COMMON,
    "10. Sınıf": _HIGH_COMMON,
    "11. Sınıf (YKS)": _HIGH_COMMON,
    "12. Sınıf (YKS)": _HIGH_COMMON,
}


def get_branches_for_grade_2027(grade: str) -> List[str]:
    g_map = FULL_CURRICULUM_2027.get(grade)
    if g_map:
        return list(g_map.keys())
    return list(_MIDDLE_COMMON.keys())


def get_areas_for_grade_branch_2027(grade: str, branch: str) -> List[str]:
    g_map = FULL_CURRICULUM_2027.get(grade, {})
    if branch in g_map:
        return list(g_map[branch].keys())
    for other_g in FULL_CURRICULUM_2027.values():
        if branch in other_g:
            return list(other_g[branch].keys())
    return ["Genel Kazanımlar"]


def get_outcomes_for_2027(grade: str, branch: str, area: str) -> List[str]:
    g_map = FULL_CURRICULUM_2027.get(grade, {})
    if branch in g_map and area in g_map[branch]:
        return list(g_map[branch][area])
    for other_g in FULL_CURRICULUM_2027.values():
        if branch in other_g and area in other_g[branch]:
            return list(other_g[branch][area])
    return ["Kazanım odaklı beceri ve muhakeme sorusu üretimi"]



ONCELIKLI_EKSIK_SIRASI: List[Dict[str, str]] = [
    {
        "priority": "1. ÖNCELİK • %46 İSABET",
        "pct": 46,
        "topic": "T.8.1.9. Sözel mantık ve muhakeme gerektiren metinleri analiz etme",
        "branch": "Türkçe",
        "grade": "8. Sınıf (LGS)",
        "action": "Tablo/şema kurma ve telafi sorusu",
    },
    {
        "priority": "2. ÖNCELİK • %52 İSABET",
        "pct": 52,
        "topic": "Çarpanlar ve katlar, EBOB-EKOK problem modellemeleri",
        "branch": "Matematik",
        "grade": "8. Sınıf (LGS)",
        "action": "Yeni nesil kurgu & çeldirici analizi",
    },
    {
        "priority": "3. ÖNCELİK • DİKKAT",
        "topic": "Yazım kuralları ve noktalama işaretlerinin anlamı belirlemedeki işlevi",
        "branch": "Türkçe",
        "grade": "8. Sınıf (LGS)",
        "action": "Olumsuz soru kökü hız antrenmanı",
    },
]

DIJIMENTOR_PRESETS: Dict[str, Dict[str, Any]] = {
    "2027 LGS Derece & Net İvmesi": {
        "exam": "2027 LGS",
        "hours": 20,
        "goal": "Fen liseleri ve nitelikli Anadolu liseleri bandı (470+ puan ve 82+ net hedefi)",
        "focus": "Deneme Net Artırma & Zaman Yönetimi",
        "weak": "Sözel mantık ve muhakeme metinleri, cebirsel ifadeler ve çarpanlar-katlar yeni nesil problemleri, süre baskısı altında çeldirici ayıklama",
        "status": "Orta-üst seviye; Türkçe ve Matematikte son 15 soruda süre baskısı yaşıyor ve deneme analizini haftalık sistematik yapamıyor",
        "branch": "Türkçe",
        "grade": "8. Sınıf (LGS)",
        "topic": "T.8.1.9. Sözel mantık ve muhakeme gerektiren metinleri analiz etme",
    },
    "2027 YKS TYT-AYT Disiplin Rutini": {
        "exam": "2027 TYT - AYT",
        "hours": 30,
        "goal": "TYT 98+ net ve AYT alanında ilk 15 bin derece hedefi",
        "focus": "Haftalık Dengeli Çalışma Programı",
        "weak": "TYT Türkçe paragraf süresi ve anlam-yorum çeldiricileri, problem kurma hızı, AYT konu tekrar gecikmeleri ve deneme tempo dalgalanması",
        "status": "Düzenli çalışma alışkanlığı var ancak TYT süresini verimli bölme ve hafta sonu deneme analizlerini puan/net aksiyonuna çevirme konusunda gelişime açık",
        "branch": "Türk Dili ve Edebiyatı",
        "grade": "12. Sınıf (YKS)",
        "topic": "Nitelikli metinlerde yapı, bakış açısı, anlatım biçimleri ve anlam çözümlemesi",
    },
    "Dikkat & Çeldirici Tuzağı Temizliği": {
        "exam": "2027 LGS",
        "hours": 16,
        "goal": "Dikkat ve okuma kaynaklı net kayıplarını %70 azaltmak ve istikrarlı net artışı sağlamak",
        "focus": "Dikkat ve Çeldirici Tuzağı Azaltma",
        "weak": "Soru kökündeki 'ulaşılamaz / çıkarılamaz / değildir' ifadelerini kaçırma, öncül atlama ve çeldirici şıklara erken yönelme",
        "status": "Konu bilgisi yeterli olmasına rağmen hızlı okuma ve aceleci işaretleme nedeniyle bildiği sorularda hata yapıyor",
        "branch": "Türkçe",
        "grade": "8. Sınıf (LGS)",
        "topic": "Yazım kuralları ve noktalama işaretlerinin anlamı belirlemedeki işlevi",
    },
}

SUBSCRIPTION_TIERS: List[Dict[str, Any]] = [
    {
        "id": "ogrenci_temel",
        "segment": "Öğrenci",
        "name": "2027 LGS Başlangıç Dijital",
        "monthly_price": 1990,
        "monthly_quota": 400,
        "coach_quota": 30,
        "badge": "Temel",
        "desc": "5–8. sınıf öğrencileri için kazanım tekrarları, çeldirici farkındalığı ve adımlı çözüm desteği.",
        "features": [
            "Aylık 400 Maarif Modeli & LGS soru üretimi",
            "Öğrenci Portalı mini deneme ve ipucu/çözüm asistanı",
            "2027 LGS canlı geri sayım ve haftalık hedef takibi",
            "Çeldirici analizli standart DOCX test çıktısı",
            "Hata defteri ve kazanım özeti",
            "Mobil & masaüstü uyumlu kesintisiz çalışma arayüzü",
        ],
    },
    {
        "id": "ogrenci_pro",
        "segment": "Öğrenci",
        "name": "2027 LGS Hedef Pro",
        "monthly_price": 3290,
        "monthly_quota": 1500,
        "coach_quota": 90,
        "badge": "En Çok Tercih Edilen",
        "desc": "Yeni nesil muhakeme soruları, otomatik redaksiyon, Dijimentor planı ve veli raporu.",
        "features": [
            "Aylık 1.500 yapay zekâ soru ve redaksiyon hakkı",
            "Otomatik redaktör ile çeldirici ve soru kökü denetimi",
            "Dijimentor haftalık akıllı çalışma ve tempo programı",
            "Veli Takip Paneli canlı bildirim ve özet erişimi",
            "Soru köküne uygun başlıklandırılmış DOCX test çıktısı",
            "Yanlış soruları otomatik Hata Defteri'ne aktarma",
        ],
    },
    {
        "id": "ogrenci_elit",
        "segment": "Öğrenci",
        "name": "2027 LGS Tam Donanım Elit",
        "monthly_price": 4890,
        "monthly_quota": 3500,
        "coach_quota": 200,
        "badge": "Kapsamlı",
        "desc": "Yüksek soru üretim kotası, yoğun kalite denetimi, çeldirici raporu ve özel mentorluk senaryoları.",
        "features": [
            "Aylık 3.500 soru üretimi + düzenlenebilir taslaklar",
            "Dijimentor günlük deneme analizi & çeldirici tuzağı raporu",
            "Veli Takip Paneli + haftalık veli bilgilendirme merkezi",
            "Özel kurum/öğrenci başlıklı düzenlenebilir DOCX testler",
            "Zayıf kazanımlardan anında telafi testi üretimi",
            "Süre yönetimi ve dikkat kaybı azaltma protokolü",
        ],
    },
    {
        "id": "ogrenci_kamp",
        "segment": "Öğrenci",
        "name": "2027 LGS Final & Deneme Kampı",
        "monthly_price": 5690,
        "monthly_quota": 5000,
        "coach_quota": 260,
        "badge": "Kamp",
        "desc": "Sınava hazırlık maratonunda yüksek hacimli deneme üretimi, hız antrenmanı ve net ivme planı.",
        "features": [
            "Aylık 5.000 yeni nesil deneme ve kazanım sorusu",
            "Yoğun süre yönetimi ve son tur hızlandırma planları",
            "Dijimentor haftalık net artırma ve kritik kazanım takibi",
            "Ayrıntılı doğru/yanlış/net tablosu ve hata analizi",
            "Baskıya hazır DOCX deneme ve çalışma fasikülleri",
            "Veliye haftalık kamp ilerleme raporu paylaşımı",
        ],
    },
    {
        "id": "yks_sayisal",
        "segment": "YKS",
        "name": "2027 YKS MF (Sayısal) Akademi",
        "monthly_price": 3690,
        "monthly_quota": 1800,
        "coach_quota": 110,
        "badge": "TYT + AYT",
        "desc": "TYT ve AYT Matematik, Fizik, Kimya, Biyoloji ve Türkçe odaklı ileri hazırlık paketi.",
        "features": [
            "Aylık 1.800 TYT & AYT kazanım ve analitik soru üretimi",
            "Dijimentor net artırma, süre bölümü ve tekrar döngüsü",
            "Adımlı çözüm asistanı ile mantık ve işlem adımı analizi",
            "Konu bazlı zorluk kademeli test ve deneme üretimi",
            "Düzenlenebilir DOCX soru ve çözüm çıktısı",
            "Haftalık çalışma kapasitesi (%45/%30/%25) optimizasyonu",
        ],
    },
    {
        "id": "yks_ea_sozel",
        "segment": "YKS",
        "name": "2027 YKS EA & Sözel Atölye",
        "monthly_price": 3490,
        "monthly_quota": 1800,
        "coach_quota": 110,
        "badge": "TYT + AYT",
        "desc": "Türk Dili ve Edebiyatı, Tarih, Coğrafya, Matematik ve paragraf yorum-analiz paketi.",
        "features": [
            "Aylık 1.800 paragraf, edebiyat ve sosyal bilimler sorusu",
            "Çeldirici şık ayıklama ve metin analizi redaksiyonu",
            "Dijimentor okuma hızı ve deneme stratejisi mentorluğu",
            "Mini deneme modülü ile anlık net ve doğruluk hesabı",
            "Düzenlenebilir DOCX çalışma kağıdı ve test çıktısı",
            "Eksik kazanım odaklı haftalık tekrar planı",
        ],
    },
    {
        "id": "yks_vip",
        "segment": "YKS",
        "name": "2027 YKS Derece VIP",
        "monthly_price": 5490,
        "monthly_quota": 4500,
        "coach_quota": 280,
        "badge": "VIP",
        "desc": "Derece hedefleyen öğrenciler için yüksek soru kotası, derin analiz ve özel mentorluk.",
        "features": [
            "Aylık 4.500 soru üretimi ve kalite redaksiyonu",
            "Dijimentor hedef derece simülasyonu ve haftalık rapor",
            "Zor (LGS/ÖSYM) seviye muhakeme ve çeldirici setleri",
            "Yanlış sorulardan kişiye özel pekiştirme akışı",
            "Kurumsal standartta DOCX test ve plan çıktıları",
            "Öncelikli kapsam ve yoğun deneme hazırlık desteği",
        ],
    },
    {
        "id": "yks_tyt_sprint",
        "segment": "YKS",
        "name": "2027 TYT Hız & Paragraf Sprinti",
        "monthly_price": 3190,
        "monthly_quota": 1400,
        "coach_quota": 90,
        "badge": "TYT Odaklı",
        "desc": "TYT Türkçe ve temel matematikte okuma hızı, dikkat ve süre baskısını aşmaya yönelik paket.",
        "features": [
            "Aylık 1.400 TYT paragraf ve akıl yürütme sorusu",
            "Dikkat ve çeldirici tuzağı azaltma mentorluk planı",
            "İnteraktif mini deneme ile hız ve net takibi",
            "Zorluk seviyeli (Kolay–Zor) kademeli soru setleri",
            "DOCX formatında pratik test ve değerlendirme çıktısı",
            "Haftalık tempo ve dikkat KPI kontrol listesi",
        ],
    },
    {
        "id": "veli_aile",
        "segment": "Veli",
        "name": "Aile Takip & Rehberlik",
        "monthly_price": 2190,
        "monthly_quota": 600,
        "coach_quota": 60,
        "badge": "Veli Odaklı",
        "desc": "Öğrencinin çalışma düzenini, deneme ivmesini ve eksik kazanımlarını şeffaf izleyin.",
        "features": [
            "Veli Takip Paneli haftalık performans ve net özeti",
            "4 kategorili Veli Bilgilendirme & Otomatik Mesaj Merkezi",
            "WhatsApp, SMS ve E-posta için hazır veli mesaj şablonları",
            "Dijimentor çalışma disiplini ve motivasyon planı",
            "Aylık 600 soru üretimi ve 1 öğrenci profili dahil",
            "Veli mesajlarını ve raporları DOCX olarak indirme",
        ],
    },
    {
        "id": "veli_coklu",
        "segment": "Veli",
        "name": "Çoklu Kardeş Aile Paketi",
        "monthly_price": 3690,
        "monthly_quota": 2200,
        "coach_quota": 130,
        "badge": "2 Öğrenci",
        "desc": "İki kardeş için ayrı sınıf/sınav seviyelerinde takip, soru üretimi ve Dijimentor desteği.",
        "features": [
            "2 ayrı öğrenci profili (5–12. sınıf / 2027 LGS & YKS)",
            "Aylık 2.200 ortak soru üretim ve redaksiyon kotası",
            "Veli Takip Paneli karşılaştırmalı gelişim görünümü",
            "Otomatik veli bilgilendirme mesajları ve duyuru akışı",
            "İki kademe için ayrı DOCX test ve plan çıktısı",
            "Aile bütçesinde avantajlı çoklu paket kotası",
        ],
    },
    {
        "id": "veli_rehber_plus",
        "segment": "Veli",
        "name": "Aile Rehberlik & Deneme Plus",
        "monthly_price": 4390,
        "monthly_quota": 3000,
        "coach_quota": 160,
        "badge": "Plus",
        "desc": "Sınav yılındaki öğrencinin deneme, kazanım ve çalışma disiplini takibini üst düzeye taşıyan paket.",
        "features": [
            "Aylık 3.000 soru üretimi + otomatik redaksiyon",
            "Dijimentor haftalık çalışma ve sınav kaygısı desteği",
            "Veli Takip Paneli ders ve kazanım bazlı grafikler",
            "Düzenli veli bilgilendirme şablonları ve DOCX dışa aktarım",
            "Eksik kazanım telafi takibi ve özet bildirimleri",
            "Yüksek mentorluk kotası ile kesintisiz planlama",
        ],
    },
    {
        "id": "b2b_ogretmen",
        "segment": "Kurum / B2B",
        "name": "Öğretmen & Zümre Stüdyosu",
        "monthly_price": 7900,
        "monthly_quota": 6000,
        "coach_quota": 220,
        "badge": "Zümre",
        "desc": "Branş öğretmenleri ve zümre başkanları için kazanım testi, redaksiyon ve zümre üretim aracı.",
        "features": [
            "Aylık 6.000 soru üretimi + otomatik redaksiyon denetimi",
            "Türkiye Yüzyılı Maarif Modeli kazanım ve tema uyumu",
            "Kurum başlıklı düzenlenebilir DOCX sınav ve test çıktısı",
            "5 öğretmen/editör koltuğu ve zümre paylaşım akışı",
            "Çeldirici kalitesi ve dil/yazım doğruluğu kontrolü",
            "Hızlı kazanım tarama ve ölçme-değerlendirme setleri",
        ],
    },
    {
        "id": "b2b_kurs",
        "segment": "Kurum / B2B",
        "name": "Etüt Merkezi & Kurs Paketi",
        "monthly_price": 19900,
        "monthly_quota": 18000,
        "coach_quota": 600,
        "badge": "Kurumsal",
        "desc": "Kurs merkezleri ve butik okullar için ölçeklenebilir soru fabrikası, portal ve zümre yönetimi.",
        "features": [
            "Aylık 18.000 soru üretimi ve toplu test hazırlama",
            "Öğrenci Portalı, Dijimentor ve Veli Paneli kurumsal erişimi",
            "Zümre soru kotası dağıtımı ve kurum içi raporlama",
            "Kurum logolu/başlıklı DOCX sınav ve özet çıktıları",
            "25 öğretmen + 250 öğrenci hesabı kapasitesi",
            "Veli bilgilendirme şablonları ile kurumsal iletişim",
        ],
    },
    {
        "id": "b2b_kampus",
        "segment": "Kurum / B2B",
        "name": "Okul & Kampüs Lisansı",
        "monthly_price": 39900,
        "monthly_quota": 50000,
        "coach_quota": 1500,
        "badge": "Kampüs",
        "desc": "Tüm kademelerde standart ölçme-değerlendirme, zümre havuzu ve akıllı mentorluk altyapısı.",
        "features": [
            "Aylık 50.000 soru üretimi + öncelikli redaksiyon",
            "Tüm modüllerde kurumsal yetkilendirme ve takip",
            "Branş zümreleri arası otomatik kota ve deneme planı",
            "İdareci ve zümre performans/kapasite raporları",
            "Sınırsız DOCX test/teklif çıktısı ve kurumsal şablon",
            "Kampüs geneli standart sınav ve kazanım uyumu",
        ],
    },
    {
        "id": "b2b_ozel",
        "segment": "Kurum / B2B",
        "name": "Kurumsal Zincir & Yayınevi",
        "monthly_price": 74900,
        "monthly_quota": 140000,
        "coach_quota": 3500,
        "badge": "Enterprise",
        "desc": "Çok şubeli eğitim kurumları ve içerik ekipleri için yüksek kapasite ve özel editörlük.",
        "features": [
            "Aylık 140.000 yüksek hacimli soru üretim kotası",
            "Çok şubeli kampüs ve zümre kapasite yönetimi",
            "Sıkı redaksiyon, çeldirici ve kazanım kalite standardı",
            "Kurumsal teklif, özet ve sınav DOCX otomasyonu",
            "Yüksek hacimli Dijimentor mentorluk ve analiz kotası",
            "Kurumsal onboarding ve özel yapılandırma desteği",
        ],
    },
    {
        "id": "b2b_yayinevi_pro",
        "segment": "Kurum / B2B",
        "name": "Akademik İçerik & Soru Bankası Pro",
        "monthly_price": 98900,
        "monthly_quota": 220000,
        "coach_quota": 5000,
        "badge": "Ultra B2B",
        "desc": "Yayınevleri ve büyük eğitim grupları için tam kapasiteli soru fabrikası ve ölçme stüdyosu.",
        "features": [
            "Aylık 220.000 soru üretim ve redaksiyon kapasitesi",
            "Tüm sınıf kademeleri ve sınav türlerinde çoklu üretim",
            "Gelişmiş zümre dağılımı ve kurumsal raporlama",
            "5.000 Dijimentor plan kotası ile geniş mentorluk ağı",
            "Standartlaştırılmış DOCX kitapçık ve rapor çıktıları",
            "Kurumsal ölçekte kesintisiz içerik üretim altyapısı",
        ],
    },
]

PARENT_MESSAGE_TOPICS: Dict[str, Dict[str, Any]] = {
    "1. Akademik İlerleme ve Başarı Takibi": {
        "subtopics": [
            "Haftalık Soru ve Net Özeti",
            "Türkiye Yüzyılı Maarif Modeli Kazanım Analizi",
            "Deneme Sınavı ve 2027 LGS Değerlendirmesi",
        ],
        "templates": {
            "Haftalık Soru ve Net Özeti": (
                "Sayın {parent_name},\n\n"
                "Kurşun Kalem Dijital Ofis platformunda öğrencimiz {student_name} ({grade}) "
                "bu hafta toplam {weekly_solved} soru çözmüştür. "
                "{correct_count} doğru, {wrong_count} yanlış ve {blank_count} boş ile "
                "%{accuracy} başarı oranı elde edilmiştir.\n\n"
                "Güçlü Kazanım: {strong_topic}\n"
                "Pekiştirilmesi Önerilen Alan: {weak_topic}\n\n"
                "Kurşun Kalem Dijital Ofis"
            ),
            "Türkiye Yüzyılı Maarif Modeli Kazanım Analizi": (
                "Sayın {parent_name},\n\n"
                "Öğrencimiz {student_name}'in Türkiye Yüzyılı Maarif Modeli beceri temelli kazanım tablosu güncellenmiştir. "
                "'{strong_topic}' alanında güçlü bir ilerleme kaydedilirken, '{weak_topic}' kazanımında "
                "ek pekiştirme testleri tanımlanmıştır.\n\n"
                "Kurşun Kalem Dijital Ofis"
            ),
            "Deneme Sınavı ve 2027 LGS Değerlendirmesi": (
                "Sayın {parent_name},\n\n"
                "Öğrencimiz {student_name}'in son deneme performansında doğruluk oranı %{accuracy} seviyesindedir. "
                "Bir sonraki deneme oturumumuz {next_exam_date} tarihinde '{next_exam_name}' adıyla yapılacaktır.\n\n"
                "Kurşun Kalem Dijital Ofis"
            ),
        },
    },
    "2. Çalışma Disiplini ve Dijimentor Takibi": {
        "subtopics": [
            "Dijimentor Çalışma Planı Uyumu",
            "Aktif Kullanım ve Düzenlilik",
            "Eksik Telafi Uyarıları",
        ],
        "templates": {
            "Dijimentor Çalışma Planı Uyumu": (
                "Sayın {parent_name},\n\n"
                "Öğrencimiz {student_name}, haftalık Dijimentor çalışma planının %{mentor_completion} "
                "kısmını tamamlamıştır. Düzenli çalışma temposu net artışını desteklemektedir.\n\n"
                "Kurşun Kalem Dijital Ofis"
            ),
            "Aktif Kullanım ve Düzenlilik": (
                "Sayın {parent_name},\n\n"
                "{student_name} bu hafta platformda düzenli soru ve mini test çözümü gerçekleştirmiştir. "
                "Günlük kısa tekrar rutinlerinin sürdürülmesi önerilmektedir.\n\n"
                "Kurşun Kalem Dijital Ofis"
            ),
            "Eksik Telafi Uyarıları": (
                "Sayın {parent_name},\n\n"
                "Öğrencimiz {student_name}'in '{weak_topic}' kazanımındaki telafi görevlerinin tamamlanması "
                "konu bütünlüğü açısından önem taşımaktadır.\n\n"
                "Kurşun Kalem Dijital Ofis"
            ),
        },
    },
    "3. Sınav, Ödev ve Takvim Duyuruları": {
        "subtopics": [
            "Yaklaşan Sınav ve Deneme Takvimi",
            "Ödev ve Etkinlik Bilgilendirmesi",
            "Özel Dönem Bilgilendirmeleri",
        ],
        "templates": {
            "Yaklaşan Sınav ve Deneme Takvimi": (
                "Sayın {parent_name},\n\n"
                "{next_exam_date} tarihinde '{next_exam_name}' sınav oturumu yapılacaktır. "
                "Öğrencimiz {student_name}'in sınav saatinden 15 dakika önce hazır bulunmasını rica ederiz.\n\n"
                "Kurşun Kalem Dijital Ofis"
            ),
            "Ödev ve Etkinlik Bilgilendirmesi": (
                "Sayın {parent_name},\n\n"
                "Öğrencimiz {student_name} için '{weak_topic}' ve '{strong_topic}' odaklı "
                "haftalık çalışma görevi tanımlanmıştır.\n\n"
                "Kurşun Kalem Dijital Ofis"
            ),
            "Özel Dönem Bilgilendirmeleri": (
                "Sayın {parent_name},\n\n"
                "Kurşun Kalem Dijital Ofis 2027 kamp ve kazanım tekrar programları aktif edilmiştir. "
                "Detaylara panel üzerinden ulaşabilirsiniz.\n\n"
                "Kurşun Kalem Dijital Ofis"
            ),
        },
    },
    "4. Abonelik, Paket ve Kurumsal İşlemler": {
        "subtopics": [
            "Ofis Çantası & Paket Durumu",
            "Ödeme ve Fatura Bilgilendirmesi",
            "Destek ve Veli-Öğretmen İletişimi",
        ],
        "templates": {
            "Ofis Çantası & Paket Durumu": (
                "Sayın {parent_name},\n\n"
                "Öğrencimiz {student_name} hesabında {rem_questions} soru üretim kotası ve "
                "{rem_mentor} Dijimentor plan hakkı tanımlıdır.\n\n"
                "Kurşun Kalem Dijital Ofis"
            ),
            "Ödeme ve Fatura Bilgilendirmesi": (
                "Sayın {parent_name},\n\n"
                "Abonelik ve ödeme bilgilendirmeleri {parent_email} adresine iletilmektedir. "
                "Her türlü soru için Veli Takip Paneli üzerinden bize ulaşabilirsiniz.\n\n"
                "Kurşun Kalem Dijital Ofis"
            ),
            "Destek ve Veli-Öğretmen İletişimi": (
                "Sayın {parent_name},\n\n"
                "Öğrencimiz {student_name}'in gelişim süreciyle ilgili rehberlik ve öğretmen görüşmesi "
                "taleplerinizi Veli Takip Paneli üzerinden iletebilirsiniz.\n\n"
                "Kurşun Kalem Dijital Ofis"
            ),
        },
    },
}

LOGO_DATA_URI = svg_to_data_uri(EMBEDDED_LOGO_SVG)
BACKPACK_DATA_URI = svg_to_data_uri(SCHOOL_BACKPACK_SVG)
HERO_SHOWCASE_URI = svg_to_data_uri(HERO_SHOWCASE_SVG)
LOGO_PNG_BYTES = get_embedded_logo_png_bytes()

# ==============================================================================
# KUNDUZ İLHAMLI KURUMSAL TASARIM SİSTEMİ + GİRİŞ ANİMASYONLARI
# ==============================================================================
REFINED_OFFICE_CSS = """
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Plus+Jakarta+Sans:wght@600;700;800&family=Baloo+2:wght@600;700;800&display=swap');

:root {
    --kp-navy: #1E2952;
    --kp-navy-2: #2E3F8F;
    --kp-indigo: #5A52E8;
    --kp-purple-bar: #6E47E5;
    --kp-blue-accent: #2563EB;
    --kp-navy-soft: #EEF2FF;
    --kp-gold: #D97706;
    --kp-gold-bright: #FACC15;
    --kp-bg: #FFFFFF;
    --kp-surface: #FFFFFF;
    --kp-surface-alt: #F8FAFC;
    --kp-border: #E2E8F0;
    --kp-text: #1E293B;
    --kp-muted: #64748B;
}

@keyframes kpSoftEnter {
    from { opacity: 0.72; transform: translateY(5px); }
    to { opacity: 1; transform: translateY(0); }
}

@keyframes kpHeroEntrance {
    0% { opacity: 0; transform: translateY(10px) scale(0.99); }
    100% { opacity: 1; transform: translateY(0) scale(1); }
}

@keyframes kpStaggerCard {
    0% { opacity: 0; transform: translateY(8px); }
    100% { opacity: 1; transform: translateY(0); }
}

@keyframes kpTopProgressOnce {
    0% { width: 0%; opacity: 0.9; }
    70% { width: 85%; opacity: 0.9; }
    100% { width: 100%; opacity: 0; }
}

@keyframes kpSkeletonShimmer {
    0% { background-position: -200% 0; }
    100% { background-position: 200% 0; }
}

.kp-top-loader {
    position: fixed;
    top: 0;
    left: 0;
    height: 3px;
    z-index: 9998;
    pointer-events: none;
    background: linear-gradient(90deg, var(--kp-purple-bar) 0%, var(--kp-blue-accent) 60%, var(--kp-gold-bright) 100%);
    animation: kpTopProgressOnce 0.45s ease-out forwards;
}

.kp-skeleton-card {
    background: #FFFFFF;
    border: 1px solid var(--kp-border);
    border-radius: 14px;
    padding: 1rem 1.15rem;
    margin-bottom: 0.85rem;
}
.kp-skel-line {
    height: 12px;
    border-radius: 8px;
    margin-bottom: 0.55rem;
    background: linear-gradient(90deg, #F1F5F9 25%, #E2E8F0 50%, #F1F5F9 75%);
    background-size: 200% 100%;
    animation: kpSkeletonShimmer 1.35s ease-in-out infinite;
}

html, body, [class*="css"] {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    color: var(--kp-text);
}
h1, h2, h3, .kp-kd-hero-title, .kp-page-strip-title, .kp-pkg-title, .kp-mod-title {
    font-family: 'Plus Jakarta Sans', 'Inter', sans-serif;
}
.stApp { background: #FFFFFF; }
[data-testid="stHeader"] {
    background: rgba(255, 255, 255, 0.96);
    backdrop-filter: blur(8px);
    border-bottom: 1px solid var(--kp-border);
}
[data-testid="stDecoration"] { display: none; }
.block-container {
    /* Streamlit'in üstteki sabit şeridi (~60px) içeriği örtmesin */
    padding-top: 4.5rem;
    padding-bottom: 4.5rem;
    padding-left: clamp(1rem, 3vw, 3rem);
    padding-right: clamp(1rem, 3vw, 3rem);
    /* Geniş ekranlarda kenarlar boş kalmasın */
    max-width: min(1600px, 96vw);
    animation: kpSoftEnter 0.24s cubic-bezier(0.22, 1, 0.36, 1) both;
}

section[data-testid="stSidebar"] {
    background: #FFFFFF;
    border-right: 1px solid var(--kp-border);
}
section[data-testid="stSidebar"] [data-testid="stVerticalBlock"] {
    gap: 0.34rem;
}
.kp-sidebar-brand {
    display: flex;
    align-items: center;
    gap: 0.72rem;
    padding: 0.15rem 0.2rem 0.85rem 0.2rem;
    border-bottom: 1px solid var(--kp-border);
    margin-bottom: 0.65rem;
}
.kp-sidebar-brand img {
    width: 42px;
    height: 42px;
    object-fit: contain;
}
.kp-sidebar-brand-title {
    font-size: 0.98rem;
    font-weight: 800;
    color: #0F172A;
    line-height: 1.2;
}
.kp-sidebar-brand-sub {
    font-size: 0.78rem;
    color: var(--kp-muted);
    margin-top: 2px;
}
/* ---------- Sol menü: çerçevesiz, ikonlu linkler ---------- */
[data-testid="stSidebar"] .stButton > button {
    width: 100% !important;
    display: flex !important;
    justify-content: flex-start !important;
    align-items: center !important;
    gap: 0.55rem;
    border-radius: 8px !important;
    padding: 0.5rem 0.75rem !important;
    min-height: 2.5rem;
    font-size: 0.9rem !important;
    font-weight: 500 !important;
    border: 0 !important;
    background: transparent !important;
    color: #334155 !important;
    box-shadow: none !important;
    transition: background-color 0.14s ease, color 0.14s ease !important;
}
[data-testid="stSidebar"] .stButton > button > div { justify-content: flex-start !important; }
[data-testid="stSidebar"] .stButton > button:hover {
    background: #F1F5F9 !important;
    color: #0F172A !important;
}
[data-testid="stSidebar"] .stButton > button[kind="primary"] {
    background: #EFF6FF !important;
    color: #1D4ED8 !important;
    font-weight: 600 !important;
}
.kp-nav-title {
    font-size: 0.7rem;
    font-weight: 700;
    letter-spacing: 0.06em;
    line-height: 1.2;
    color: #64748B;
    padding: 1rem 0.75rem 0.35rem 0.75rem;
}
/* Streamlit metin bloklarına negatif alt boşluk veriyor; başlıklar butonun altına kaymasın */
[data-testid="stSidebar"] [data-testid="stMarkdown"]:has(.kp-nav-title),
[data-testid="stSidebar"] .stMarkdown:has(.kp-nav-title),
[data-testid="stSidebar"] [data-testid="stElementContainer"]:has(.kp-nav-title),
[data-testid="stSidebar"] .element-container:has(.kp-nav-title) {
    margin-bottom: 0 !important;
    height: auto !important;
    overflow: visible !important;
}
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"]:has(.kp-nav-title) { margin-bottom: 0 !important; }
.kp-sidebar-meta {
    background: #F8FAFC;
    border: 1px solid #E2E8F0;
    border-radius: 12px;
    padding: 0.85rem 0.9rem;
    margin-top: 1rem;
    font-size: 0.8rem;
    color: #334155;
    line-height: 1.45;
}
.kp-meta-head { display: flex; justify-content: space-between; align-items: center; font-weight: 700; color: #0F172A; }
.kp-meta-sub { color: #475569; margin: 0.15rem 0 0.6rem 0; }
.kp-meter-label { display: flex; justify-content: space-between; font-size: 0.76rem; color: #475569; margin: 0.45rem 0 0.2rem 0; }
.kp-meter { height: 6px; background: #E2E8F0; border-radius: 3px; overflow: hidden; }
.kp-meter > span { display: block; height: 100%; background: #1D4ED8; border-radius: 3px; }
.kp-status { display: flex; align-items: center; gap: 0.4rem; font-size: 0.76rem; color: #475569; margin-top: 0.7rem; padding-top: 0.6rem; border-top: 1px solid #E2E8F0; }
.kp-status-dot { width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; }

/* ---------- Üst şerit ---------- */
.kp-topbar {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 0.75rem;
    flex-wrap: wrap;
    padding: 0.2rem 0 0.9rem 0;
    margin-bottom: 1.5rem;
    border-bottom: 1px solid #E2E8F0;
    font-size: 0.85rem;
    color: #475569;
}
.kp-topbar-brand { font-weight: 700; color: #0F172A; }
.kp-topbar a { color: #1D4ED8; text-decoration: none; }

/* ---------- Hero ---------- */
.kp-eyebrow {
    display: inline-block;
    font-size: 0.82rem;
    font-weight: 600;
    color: #1D4ED8;
    background: #EFF6FF;
    padding: 0.3rem 0.75rem;
    border-radius: 999px;
    margin-bottom: 1rem;
}
.kp-hero-title {
    font-family: 'Plus Jakarta Sans', 'Inter', sans-serif;
    font-size: clamp(2rem, 3.6vw, 3.1rem);
    font-weight: 800;
    line-height: 1.1;
    letter-spacing: -0.02em;
    color: #0F172A;
    margin: 0 0 1rem 0;
    padding: 0;
}
.kp-hero-title em { font-style: normal; color: #1D4ED8; }
.kp-hero-lead { font-size: 1.1rem; line-height: 1.6; color: #475569; margin: 0 0 1.4rem 0; max-width: 42rem; }
.kp-stats { display: flex; gap: 2.2rem; flex-wrap: wrap; margin-top: 1.4rem; }
.kp-stat-num { font-family: 'Plus Jakarta Sans', sans-serif; font-size: 1.7rem; font-weight: 800; color: #0F172A; line-height: 1.2; }
.kp-stat-lbl { font-size: 0.82rem; color: #64748B; }
.st-key-hero_cta_primary button, .st-key-hero_cta_secondary button {
    min-height: 3rem;
    font-size: 0.95rem !important;
    border-radius: 10px !important;
}

.kp-preview {
    background: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 16px;
    padding: 1.3rem;
    box-shadow: 0 20px 50px rgba(15, 23, 42, 0.08);
}
.kp-preview-head { display: flex; justify-content: space-between; align-items: flex-start; gap: 0.5rem; }
.kp-preview-head small { display: block; font-size: 0.78rem; color: #64748B; }
.kp-preview-head strong { font-size: 1.05rem; color: #0F172A; }
.kp-pill { font-size: 0.75rem; font-weight: 600; color: #1D4ED8; background: #EFF6FF; padding: 0.2rem 0.6rem; border-radius: 999px; white-space: nowrap; }
.kp-preview-stats { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 0.6rem; margin: 1.1rem 0; }
.kp-preview-stat { background: #F8FAFC; border-radius: 10px; padding: 0.7rem; }
.kp-preview-stat b { display: block; font-size: 1.25rem; color: #0F172A; }
.kp-preview-stat span { font-size: 0.75rem; color: #64748B; }
.kp-preview-row { display: flex; justify-content: space-between; gap: 0.75rem; font-size: 0.86rem; color: #334155; padding: 0.6rem 0; border-top: 1px solid #E2E8F0; }
.kp-preview-row span:last-child { color: #0F172A; font-weight: 600; text-align: right; }

/* ---------- Bölüm başlıkları ---------- */
.kp-section-title { font-family: 'Plus Jakarta Sans', sans-serif; font-size: 1.45rem; font-weight: 700; color: #0F172A; margin: 0; letter-spacing: -0.01em; }
.kp-section-sub { font-size: 0.95rem; color: #475569; margin: 0.3rem 0 0.9rem 0; }
.kp-gap { height: 3rem; }

/* ---------- Kartlı kapsayıcılar (st.container(border=True)) ---------- */
[data-testid="stVerticalBlockBorderWrapper"] { border-radius: 14px !important; }

/* Sınıf seçici: hap şeklinde butonlar */
[class*="st-key-grade_chip_"] button { border-radius: 999px !important; min-height: 2.75rem; font-weight: 500 !important; font-size: 0.9rem !important; }

/* Modül kartları */
.kp-mod-icon { width: 40px; height: 40px; border-radius: 10px; display: flex; align-items: center; justify-content: center; font-weight: 800; font-size: 0.85rem; margin-bottom: 0.7rem; }
.kp-mod-title { font-family: 'Plus Jakarta Sans', sans-serif; font-weight: 700; font-size: 1rem; color: #0F172A; }
.kp-mod-desc { font-size: 0.86rem; line-height: 1.5; color: #475569; margin-top: 0.3rem; min-height: 3.9rem; }
[class*="st-key-mod_open_"] button, .st-key-prio_all button, [class*="st-key-landing_try_"] button, .st-key-landing_kvkk button {
    border: 0 !important;
    background: transparent !important;
    color: #1D4ED8 !important;
    padding: 0 !important;
    min-height: 0 !important;
    font-weight: 600 !important;
    font-size: 0.88rem !important;
    justify-content: flex-start !important;
}
[class*="st-key-mod_open_"] button:hover, .st-key-prio_all button:hover, [class*="st-key-landing_try_"] button:hover { color: #1E3A8A !important; text-decoration: underline; }

/* Öncelik satırları */
.kp-prio-num { width: 32px; height: 32px; border-radius: 50%; background: #F1F5F9; font-weight: 700; font-size: 0.9rem; display: flex; align-items: center; justify-content: center; color: #0F172A; }
.kp-prio-title { font-weight: 600; font-size: 0.95rem; color: #0F172A; line-height: 1.4; }
.kp-prio-sub { font-size: 0.8rem; color: #64748B; margin-top: 0.15rem; }
.kp-prio-divider { border-top: 1px solid #E2E8F0; margin: 0.35rem 0; }
/* Kullanıcı kartı, seri, bugün kartları */
.kp-user-card { display: flex; align-items: flex-start; gap: 0.6rem; padding: 0.6rem 0.7rem; margin: 0 0 0.6rem 0; border: 1px solid #E2E8F0; border-radius: 12px; background: #FFFFFF; }
.kp-avatar { width: 34px; height: 34px; border-radius: 50%; background: #1D4ED8; color: #FFFFFF; font-weight: 800; display: flex; align-items: center; justify-content: center; flex-shrink: 0; }
.kp-user-name { font-weight: 700; font-size: 0.88rem; color: #0F172A; line-height: 1.25; overflow-wrap: anywhere; }
.kp-user-role { font-size: 0.75rem; color: #64748B; }
.kp-streak-pill { display: inline-block; margin-top: 0.3rem; font-size: 0.72rem; font-weight: 800; color: #9A3412; background: #FFEDD5; border-radius: 999px; padding: 0.15rem 0.55rem; white-space: nowrap; }
.kp-today-label { font-size: 0.78rem; font-weight: 700; color: #64748B; letter-spacing: 0.05em; }
.kp-today-big { font-family: 'Plus Jakarta Sans', sans-serif; font-size: 2.2rem; font-weight: 800; color: #0F172A; line-height: 1.15; margin-top: 0.2rem; }
.kp-today-big span { font-size: 0.95rem; font-weight: 600; color: #64748B; }
.kp-today-sub { font-size: 0.86rem; color: #475569; min-height: 2.6rem; margin-bottom: 0.6rem; }
.kp-banner { border-radius: 12px; padding: 0.75rem 1rem; margin-bottom: 0.9rem; font-size: 0.92rem; color: #0F172A; }
.kp-banner-hw { background: #EFF6FF; border: 1px solid #BFDBFE; }
.kp-banner-review { background: #FEF9C3; border: 1px solid #FDE047; }
.kp-code { font-family: ui-monospace, monospace; font-weight: 800; font-size: 1.05rem; letter-spacing: 0.12em; background: #FFFFFF; border: 1px dashed #1D4ED8; border-radius: 8px; padding: 0.1rem 0.5rem; }
.kp-step { padding: 0.5rem 0; }
.kp-streak-hero { display: flex; align-items: center; gap: 1.2rem; background: #FFF7ED; border: 1px solid #FED7AA; border-radius: 16px; padding: 1.2rem 1.4rem; }
.kp-streak-num { font-family: 'Plus Jakarta Sans', sans-serif; font-size: 3.4rem; font-weight: 800; color: #EA580C; line-height: 1; }

.kp-empty-chart { display: flex; align-items: center; justify-content: center; text-align: center; border: 2px dashed #CBD5E1; border-radius: 14px; color: #64748B; font-size: 0.92rem; line-height: 1.6; background: rgba(255,255,255,0.6); }
.kp-empty-chart b { color: #0F172A; }

/* Rozetler */
.kp-badge-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(118px, 1fr)); gap: 0.7rem; }
.kp-badge { text-align: center; padding: 0.8rem 0.5rem; border-radius: 14px; border: 1px solid #E2E8F0; background: #FFFFFF; }
.kp-badge-off { opacity: 0.75; }
.kp-badge-off .kp-badge-name { color: #64748B; }
.kp-badge-name { font-weight: 700; font-size: 0.85rem; color: #0F172A; margin-top: 0.3rem; }
.kp-badge-sub { font-size: 0.72rem; color: #64748B; margin-top: 0.15rem; line-height: 1.3; }

.kp-tag-warn { font-size: 0.75rem; font-weight: 600; color: #92400E; background: #FEF3C7; padding: 0.2rem 0.6rem; border-radius: 999px; }
.kp-page-strip {
    display: flex;
    align-items: baseline;
    justify-content: space-between;
    gap: 1rem;
    padding: 0.05rem 0 0.7rem 0;
    margin-bottom: 1rem;
    border-bottom: 1px solid var(--kp-border);
    animation: kpHeroEntrance 0.32s cubic-bezier(0.22, 1, 0.36, 1) both;
}
.kp-page-strip-title {
    font-size: 1.28rem;
    font-weight: 800;
    color: var(--kp-navy);
    margin: 0;
}
.kp-page-strip-desc {
    font-size: 0.83rem;
    color: var(--kp-muted);
    margin: 0.2rem 0 0 0;
}

.kp-kademe-card {
    background: #FFFFFF;
    border: 1px solid var(--kp-border);
    border-radius: 16px;
    padding: 1rem 1.1rem;
    height: 146px;
    display: flex;
    flex-direction: column;
    justify-content: flex-start;
    margin-bottom: 0.55rem;
    box-shadow: 0 3px 12px rgba(37, 99, 235, 0.04);
    transition: transform 0.16s ease, box-shadow 0.16s ease;
    animation: kpStaggerCard 0.44s cubic-bezier(0.22, 1, 0.36, 1) both;
}
.kp-kademe-card:hover {
    transform: translateY(-2px);
    box-shadow: 0 8px 18px rgba(37, 99, 235, 0.09);
}
.kp-kademe-tag {
    display: inline-block;
    width: fit-content;
    font-size: 0.7rem;
    font-weight: 800;
    color: #FFFFFF;
    padding: 0.22rem 0.65rem;
    border-radius: 999px;
}
.kp-kademe-title {
    font-size: 1rem;
    font-weight: 800;
    color: var(--kp-navy);
    margin: 0.55rem 0 0.25rem 0;
}
.kp-kademe-desc {
    font-size: 0.8rem;
    color: var(--kp-muted);
    line-height: 1.42;
    margin: 0;
}

.kp-priority-card {
    background: #FFFFFF;
    border: 1px solid var(--kp-border);
    border-left: 4px solid #2563EB;
    border-radius: 14px;
    padding: 0.95rem 1.05rem;
    height: 132px;
    margin-bottom: 0.55rem;
    animation: kpStaggerCard 0.46s cubic-bezier(0.22, 1, 0.36, 1) both;
}
.kp-priority-badge {
    font-size: 0.7rem;
    font-weight: 800;
    color: #1D4ED8;
    background: #EFF6FF;
    padding: 0.18rem 0.58rem;
    border-radius: 999px;
    display: inline-block;
}
.kp-priority-title {
    font-size: 0.94rem;
    font-weight: 800;
    color: var(--kp-navy);
    margin: 0.45rem 0 0.2rem 0;
}
.kp-priority-sub {
    font-size: 0.79rem;
    color: var(--kp-muted);
    margin: 0;
}

.kp-value-strip {
    background: linear-gradient(90deg, #1E2952 0%, #2563EB 65%, #5A52E8 100%);
    color: #FFFFFF;
    border-radius: 14px;
    padding: 0.85rem 1.2rem;
    margin-bottom: 1.05rem;
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 1rem;
    flex-wrap: wrap;
}
.kp-value-strip-item {
    display: flex;
    align-items: center;
    gap: 0.48rem;
    font-size: 0.82rem;
    font-weight: 600;
}
.kp-value-dot {
    width: 8px;
    height: 8px;
    border-radius: 50%;
    background: #FACC15;
    display: inline-block;
}

div[data-testid="stMetric"] {
    background: #FFFFFF;
    border: 1px solid var(--kp-border);
    border-radius: 14px;
    padding: 0.92rem 1.05rem;
    box-shadow: 0 3px 10px rgba(15, 23, 42, 0.04);
    animation: kpStaggerCard 0.42s cubic-bezier(0.22, 1, 0.36, 1) both;
}
div[data-testid="stMetric"] [data-testid="stMetricValue"] {
    font-size: 1.52rem;
    font-weight: 800;
    color: var(--kp-navy);
}

.kp-card, .kp-bag-shell, .kp-pkg-card, .kp-pay-box, .kp-kpi-card, .kp-metric-box, .kp-studio-banner {
    background: #FFFFFF;
    border: 1px solid var(--kp-border);
    border-radius: 16px;
    box-shadow: 0 4px 14px rgba(15, 23, 42, 0.04);
    animation: kpStaggerCard 0.42s cubic-bezier(0.22, 1, 0.36, 1) both;
}
.kp-card, .kp-bag-shell, .kp-pkg-card, .kp-pay-box, .kp-studio-banner {
    padding: 1.15rem 1.25rem;
}
.kp-card { margin-bottom: 0.65rem; }

.kp-studio-banner {
    margin-bottom: 1.05rem;
    border-left: 4px solid #2563EB;
}
.kp-studio-title {
    font-size: 1.14rem;
    font-weight: 800;
    color: var(--kp-navy);
    margin: 0;
}
.kp-studio-sub {
    font-size: 0.84rem;
    color: var(--kp-muted);
    margin: 0.22rem 0 0 0;
}
.kp-kpi-card { padding: 0.92rem 1.02rem; }
.kp-kpi-label {
    font-size: 0.7rem;
    font-weight: 700;
    color: #64748B;
}
.kp-kpi-value {
    font-size: 1.25rem;
    font-weight: 800;
    color: var(--kp-navy);
    margin-top: 0.2rem;
}
.kp-kpi-sub {
    font-size: 0.76rem;
    color: var(--kp-muted);
    margin-top: 0.15rem;
}

.kp-metric-box {
    padding: 1.05rem 0.95rem;
    text-align: center;
    border-bottom: 4px solid #2563EB;
}
.kp-metric-num {
    font-size: 1.62rem;
    font-weight: 800;
    color: var(--kp-navy);
    line-height: 1.15;
    font-variant-numeric: tabular-nums;
}
.kp-metric-lbl {
    font-size: 0.71rem;
    color: #64748B;
    margin-top: 0.3rem;
    letter-spacing: 0.08em;
    font-weight: 700;
}

.kp-bag-shell {
    margin-bottom: 1.2rem;
    border-top: 4px solid #2563EB;
}
.kp-bag-visual {
    background: linear-gradient(160deg, #1E2952 0%, #2563EB 100%);
    border-radius: 14px;
    padding: 1rem 0.85rem;
    display: flex;
    flex-direction: column;
    align-items: center;
    color: #FFFFFF;
}
.kp-bag-visual img {
    width: 102px;
    height: 118px;
    object-fit: contain;
}
.kp-bag-count-pill {
    margin-top: 0.5rem;
    background: #FACC15;
    color: #0F172A;
    font-size: 0.74rem;
    font-weight: 800;
    padding: 0.2rem 0.7rem;
    border-radius: 999px;
}
.kp-bag-item-chip {
    background: #F8FAFC;
    border: 1px solid #E2E8F0;
    border-left: 3.5px solid #2563EB;
    border-radius: 10px;
    padding: 0.62rem 0.85rem;
    margin-bottom: 0.45rem;
    display: flex;
    justify-content: space-between;
    align-items: center;
}

.kp-pkg-card { margin-bottom: 0.65rem; }
.kp-pkg-featured {
    border: 2px solid #2563EB !important;
    background: linear-gradient(180deg, #EFF6FF 0%, #FFFFFF 35%) !important;
}
.kp-pkg-top {
    display: flex;
    justify-content: space-between;
    align-items: center;
    gap: 0.5rem;
}
.kp-pkg-seg {
    font-size: 0.71rem;
    font-weight: 700;
    color: #1D4ED8;
    background: #EFF6FF;
    padding: 0.2rem 0.58rem;
    border-radius: 6px;
}
.kp-pkg-badge {
    font-size: 0.71rem;
    font-weight: 800;
    color: #92400E;
    background: #FEF3C7;
    border: 1px solid #F59E0B;
    padding: 0.2rem 0.58rem;
    border-radius: 6px;
}
.kp-pkg-title {
    font-size: 1.04rem;
    font-weight: 800;
    color: var(--kp-navy);
    margin: 0.55rem 0 0.2rem 0;
}
.kp-pkg-price {
    font-size: 1.38rem;
    font-weight: 800;
    color: var(--kp-text);
    margin-bottom: 0.4rem;
}
.kp-pkg-price span {
    font-size: 0.8rem;
    font-weight: 500;
    color: #64748B;
}
.kp-pkg-desc {
    font-size: 0.82rem;
    color: var(--kp-muted);
    line-height: 1.42;
    margin-bottom: 0.7rem;
    min-height: 3.1rem;
}
.kp-pkg-feats {
    list-style: none;
    padding: 0.65rem 0 0 0;
    margin: 0;
    border-top: 1px solid var(--kp-border);
}
.kp-pkg-feats li {
    font-size: 0.79rem;
    color: #334155;
    padding: 0.16rem 0;
}

.kp-dist-bar {
    display: flex;
    height: 9px;
    border-radius: 999px;
    overflow: hidden;
    background: #E2E8F0;
    margin: 0.5rem 0 0.45rem 0;
}
.kp-dist-legend {
    display: flex;
    flex-wrap: wrap;
    gap: 1rem;
    font-size: 0.78rem;
    color: var(--kp-muted);
}

.stButton > button,
div[data-testid="stFormSubmitButton"] > button {
    border-radius: 8px;
    font-weight: 600;
    font-size: 0.9rem;
    border: 1px solid #CBD5E1;
    background-color: #FFFFFF;
    color: #1E293B;
}
.stButton > button:hover,
div[data-testid="stFormSubmitButton"] > button:hover {
    border-color: #2563EB;
    background-color: #EFF6FF;
    color: #1D4ED8;
}
.stButton > button[kind="primary"],
div[data-testid="stFormSubmitButton"] > button[kind="primary"] {
    background: #2563EB !important;
    border-color: #2563EB !important;
    color: #FFFFFF !important;
}
.stButton > button[kind="primary"]:hover,
div[data-testid="stFormSubmitButton"] > button[kind="primary"]:hover {
    background: #1D4ED8 !important;
    border-color: #1D4ED8 !important;
}

/* ---------- Streamlit'in varsayılan kırmızısını marka mavisine çevir ---------- */
button[kind="segmented_controlActive"], [data-testid="stBaseButton-segmented_controlActive"] {
    background: #1D4ED8 !important; border-color: #1D4ED8 !important; color: #FFFFFF !important;
}
button[kind="segmented_control"]:hover, [data-testid="stBaseButton-segmented_control"]:hover {
    border-color: #1D4ED8 !important; color: #1D4ED8 !important;
}
[data-baseweb="tab-highlight"] { background-color: #1D4ED8 !important; }
button[data-baseweb="tab"][aria-selected="true"], button[data-baseweb="tab"][aria-selected="true"] p { color: #1D4ED8 !important; }
[data-testid="stSlider"] [role="slider"] { background-color: #1D4ED8 !important; border-color: #1D4ED8 !important; }
[data-testid="stSlider"] [data-testid="stThumbValue"] { color: #1D4ED8 !important; }
[data-testid="stCheckbox"] label span[aria-checked], [data-baseweb="checkbox"] [data-checked="true"] { background-color: #1D4ED8 !important; border-color: #1D4ED8 !important; }
a { color: #1D4ED8; }

/* ---------- Sol menü: kartlar arası nefes payı ---------- */
.kp-sidebar-meta { margin-bottom: 1.1rem; }
[data-testid="stSidebar"] [data-testid="stExpander"] { margin-top: 0.4rem; }
.kp-user-card { margin-bottom: 0.9rem; }

/* ---------- Deneme sonrası kutlama ---------- */
.kp-celebrate { display: flex; align-items: center; gap: 0.8rem; background: #DCFCE7; border: 1px solid #86EFAC; border-radius: 14px; padding: 0.8rem 1rem; margin: 0.6rem 0 1rem; font-size: 0.95rem; color: #14532D; }
.kp-celebrate b { font-size: 1.1rem; }

/* ---------- Telefon ---------- */
@media (max-width: 640px) {
    .block-container { padding-left: 1rem; padding-right: 1rem; }
    .kp-hero-title { font-size: 2rem !important; }
    .kp-hero-lead { font-size: 1rem; }
    .kp-stats { gap: 1.2rem; }
    .kp-stat-num { font-size: 1.5rem !important; }
    .kp-preview { transform: none !important; }
    .kp-sticker { right: 6px; top: -14px; font-size: 0.75rem; }
    .kp-topbar { font-size: 0.78rem; }
    .kp-today-big { font-size: 1.8rem; }
    .kp-badge-grid { grid-template-columns: repeat(auto-fill, minmax(96px, 1fr)); }
    .kp-gap { height: 1.8rem; }
    [data-testid="stHorizontalBlock"] { flex-wrap: wrap; gap: 0.75rem; }
    [data-testid="stHorizontalBlock"] > [data-testid="stColumn"] { min-width: calc(50% - 0.5rem) !important; flex: 1 1 calc(50% - 0.5rem) !important; }
}

@media (prefers-reduced-motion: reduce) {
    .block-container,
    .kp-top-loader,
    .kp-top-compact-bar,
    .kp-kd-kicker,
    .kp-kd-hero-title,
    .kp-kd-hero-sub,
    .kp-kd-social-row,
    .kp-feature-strip-card,
    .kp-showcase-frame,
    .kp-page-strip,
    .kp-kademe-card,
    .kp-priority-card,
    div[data-testid="stMetric"],
    .kp-card,
    .kp-bag-shell,
    .kp-pkg-card,
    .kp-pay-box,
    .kp-kpi-card,
    .kp-metric-box,
    .kp-studio-banner {
        animation: none !important;
        opacity: 1 !important;
        transform: none !important;
    }
}
"""

# ==============================================================================
# EĞLENCELİ TEMA KATMANI: defter zemini, fosforlu vurgu, "pop" gölgeli kartlar
# ==============================================================================
FUN_THEME_CSS = """
:root {
    --fun-ink: #0F172A;
    --fun-yellow: #FDE047;
    --fun-yellow-deep: #FACC15;
    --fun-blue: #1D4ED8;
    --fun-paper: #FFFDF6;
}

/* Kareli defter zemini */
.stApp {
    background-color: var(--fun-paper);
    background-image:
        linear-gradient(rgba(29, 78, 216, 0.06) 1px, transparent 1px),
        linear-gradient(90deg, rgba(29, 78, 216, 0.06) 1px, transparent 1px);
    background-size: 26px 26px;
}
[data-testid="stHeader"] { background: rgba(255, 253, 246, 0.9); }

/* Başlıklarda yuvarlak, samimi yazı tipi */
.kp-hero-title, .kp-section-title, .kp-stat-num, .kp-mod-title, .kp-metric-num, .kp-page-strip-title {
    font-family: 'Baloo 2', 'Plus Jakarta Sans', sans-serif !important;
    letter-spacing: 0 !important;
}
.kp-hero-title { font-size: clamp(2.3rem, 4vw, 3.4rem); line-height: 1.05; }
.kp-section-title { font-size: 1.7rem; }

/* Fosforlu kalem vurgusu */
.kp-hero-title em {
    color: var(--fun-ink);
    background: linear-gradient(100deg, transparent 2%, var(--fun-yellow) 4%, var(--fun-yellow) 96%, transparent 98%);
    background-size: 100% 42%;
    background-repeat: no-repeat;
    background-position: 0 88%;
    padding: 0 0.15em;
}

/* Çıkartma gibi etiket */
.kp-eyebrow {
    background: var(--fun-yellow);
    color: var(--fun-ink);
    border: 2px solid var(--fun-ink);
    box-shadow: 2px 2px 0 var(--fun-ink);
    transform: rotate(-2deg);
    font-weight: 700;
}

/* Rakamlar: her biri kendi rengiyle */
.kp-stats > div:nth-child(1) .kp-stat-num { color: #1D4ED8; }
.kp-stats > div:nth-child(2) .kp-stat-num { color: #DB2777; }
.kp-stats > div:nth-child(3) .kp-stat-num { color: #059669; }
.kp-stat-num { font-size: 2.1rem; }

/* Butonlar: kalın çerçeve + bastıkça gömülen gölge */
.stButton > button, div[data-testid="stFormSubmitButton"] > button {
    border: 2px solid var(--fun-ink) !important;
    border-radius: 12px !important;
    box-shadow: 3px 3px 0 var(--fun-ink);
    font-weight: 700 !important;
    transition: transform 0.08s ease, box-shadow 0.08s ease, background-color 0.12s ease !important;
}
.stButton > button:hover, div[data-testid="stFormSubmitButton"] > button:hover {
    transform: translate(-1px, -1px);
    box-shadow: 4px 4px 0 var(--fun-ink);
    background-color: #FEF9C3 !important;
    color: var(--fun-ink) !important;
    border-color: var(--fun-ink) !important;
}
.stButton > button:active { transform: translate(2px, 2px); box-shadow: 1px 1px 0 var(--fun-ink); }
.stButton > button[kind="primary"]:hover { background: #1E40AF !important; color: #FFFFFF !important; }

/* Sol menü ve link görünümlü butonlar bu stilden muaf */
[data-testid="stSidebar"] .stButton > button,
[class*="st-key-mod_open_"] button, .st-key-prio_all button, [class*="st-key-landing_try_"] button, .st-key-landing_kvkk button {
    box-shadow: none !important;
    transform: none !important;
}
[data-testid="stSidebar"] .stButton > button { border: 0 !important; }
[data-testid="stSidebar"] .stButton > button[kind="primary"] {
    background: var(--fun-yellow) !important;
    color: var(--fun-ink) !important;
}
[class*="st-key-mod_open_"] button, .st-key-prio_all button, [class*="st-key-landing_try_"] button, .st-key-landing_kvkk button { border: 0 !important; }
[class*="st-key-mod_open_"] button:hover, .st-key-prio_all button:hover, [class*="st-key-landing_try_"] button:hover { background: transparent !important; }

/* Kartlar: kalın çerçeve + renkli pop gölge */
[class*="st-key-card_"] {
    border: 2px solid var(--fun-ink) !important;
    border-radius: 18px !important;
    background: #FFFFFF;
    box-shadow: 5px 5px 0 var(--fun-ink);
}
.kp-preview {
    position: relative;
    border: 2px solid var(--fun-ink);
    border-radius: 20px;
    box-shadow: 7px 7px 0 var(--fun-yellow-deep), 7px 7px 0 2px var(--fun-ink);
    transform: rotate(1.2deg);
    transition: transform 0.2s ease;
}
.kp-preview:hover { transform: rotate(0deg); }
.kp-preview-stat { border: 2px solid #E2E8F0; border-radius: 12px; }
.kp-preview-stat:nth-child(1) { background: #EFF6FF; }
.kp-preview-stat:nth-child(2) { background: #FDF2F8; }
.kp-preview-stat:nth-child(3) { background: #ECFDF5; }
.kp-meter { height: 10px; border-radius: 999px; border: 1.5px solid var(--fun-ink); background: #FFFFFF; }
.kp-meter > span { border-radius: 999px; }
.kp-sticker {
    position: absolute;
    top: -18px;
    right: -14px;
    background: #F472B6;
    color: #FFFFFF;
    font-weight: 800;
    font-size: 0.85rem;
    padding: 0.45rem 0.8rem;
    border: 2px solid var(--fun-ink);
    border-radius: 999px;
    box-shadow: 2px 2px 0 var(--fun-ink);
    transform: rotate(8deg);
    white-space: nowrap;
}

/* Modül ikonları: hover'da zıplama */
.kp-mod-icon {
    width: 52px; height: 52px; border-radius: 14px;
    border: 2px solid var(--fun-ink);
    box-shadow: 2px 2px 0 var(--fun-ink);
    transition: transform 0.2s ease;
}
[class*="st-key-card_mod_"]:hover .kp-mod-icon { transform: rotate(-8deg) scale(1.08); }

/* Sınıf hapları */
[class*="st-key-grade_chip_"] button { background: #FFFFFF; }

/* Öncelik numaraları */
.kp-prio-num { background: var(--fun-yellow); border: 2px solid var(--fun-ink); }

/* Geri sayım kutuları */
.kp-metric-box {
    border: 2px solid var(--fun-ink) !important;
    border-radius: 16px !important;
    box-shadow: 4px 4px 0 var(--fun-ink) !important;
}
.kp-metric-num { font-size: 2.2rem !important; color: var(--fun-blue) !important; }

/* Sol menü kota kutusu */
.kp-sidebar-meta { border: 2px solid var(--fun-ink); box-shadow: 3px 3px 0 var(--fun-ink); background: #FFFFFF; }

.kp-badge-on { border: 2px solid var(--fun-ink); box-shadow: 3px 3px 0 var(--fun-ink); transition: transform 0.15s ease; }
.kp-badge-on:hover { transform: rotate(-3deg) scale(1.04); }
.kp-user-card { border: 2px solid var(--fun-ink); box-shadow: 2px 2px 0 var(--fun-ink); }
.kp-avatar { background: var(--fun-yellow); color: var(--fun-ink); border: 2px solid var(--fun-ink); }
.kp-today-big, .kp-streak-num { font-family: 'Baloo 2', sans-serif !important; }
.kp-streak-hero { border: 2px solid var(--fun-ink); box-shadow: 4px 4px 0 #FB923C, 4px 4px 0 2px var(--fun-ink); }
.kp-banner { border: 2px solid var(--fun-ink) !important; }
.st-key-card_today { background: #F0FDF4 !important; }
.kp-celebrate { border: 2px solid var(--fun-ink); box-shadow: 4px 4px 0 #22C55E, 4px 4px 0 2px var(--fun-ink); transform: rotate(-0.6deg); }
.kp-celebrate b { font-family: 'Baloo 2', sans-serif; font-size: 1.3rem; }
.st-key-card_review { background: #FEFCE8 !important; }
.st-key-card_homework { background: #EFF6FF !important; }

@media (prefers-reduced-motion: reduce) {
    .kp-preview, .kp-eyebrow, .kp-sticker, .kp-mod-icon, .kp-badge-on { transform: none !important; transition: none !important; }
}
"""

# Her modül kartına kendi pastel zemini
FUN_THEME_CSS += "".join(
    f'.st-key-card_mod_{i} {{ background: {bg} !important; }}\n'
    for i, (_c, _t, _d, _p, bg, _fg) in enumerate(MODULE_CARDS)
)

css_to_inject = REFINED_OFFICE_CSS + ("\n" + FUN_THEME_CSS if FUN_THEME else "")
if EXTERNAL_CSS_PATH.exists():
    try:
        ext_css = EXTERNAL_CSS_PATH.read_text(encoding="utf-8").strip()
        if ext_css:
            css_to_inject = f"{css_to_inject}\n\n{ext_css}"
    except Exception:
        pass

st.markdown(f"<style>{css_to_inject}</style>", unsafe_allow_html=True)

# ==============================================================================
# AÇILIŞ ANİMASYONU (oturum başına bir kez)
# ==============================================================================
SPLASH_ENABLED = True
SPLASH_TITLE = "Kurşun Kalem"
SPLASH_SUBTITLE = "Dijital Ofis"

SPLASH_HTML = """
<style>
/* Açılışta sayfa kapsayıcısının dönüşümü kapatılır; yoksa tam ekran perde onun içine hapsolur */
.block-container {{ animation: none !important; transform: none !important; }}
.kp-splash {{
    position: fixed; inset: 0; z-index: 1000001;
    display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 0.6rem;
    background-color: #FFFDF6;
    background-image:
        linear-gradient(rgba(29, 78, 216, 0.07) 1px, transparent 1px),
        linear-gradient(90deg, rgba(29, 78, 216, 0.07) 1px, transparent 1px);
    background-size: 26px 26px;
    animation: kpSplashOut 0.55s cubic-bezier(0.7, 0, 0.84, 0) 2.45s forwards;
    pointer-events: none;
}}
.kp-splash-logo {{
    width: 92px; height: 92px; filter: drop-shadow(4px 4px 0 #0F172A);
    animation: kpSplashPop 0.55s cubic-bezier(0.34, 1.56, 0.64, 1) both;
}}
.kp-splash-word {{ position: relative; display: inline-block; margin-top: 0.4rem; }}
.kp-splash-title {{
    font-family: 'Baloo 2', 'Plus Jakarta Sans', sans-serif; font-weight: 800; color: #0F172A;
    font-size: clamp(2.6rem, 8vw, 4.6rem); line-height: 1; white-space: nowrap;
    clip-path: inset(-10% 100% -10% 0);
    animation: kpSplashWrite 1.05s cubic-bezier(0.45, 0, 0.25, 1) 0.45s forwards;
}}
.kp-splash-pencil {{
    position: absolute; top: -0.35em; left: 0; width: 0.9em; height: 0.9em;
    font-size: clamp(2.6rem, 8vw, 4.6rem); transform: translateX(-30%);
    opacity: 0;
    animation: kpSplashPencil 1.05s cubic-bezier(0.45, 0, 0.25, 1) 0.45s forwards;
}}
.kp-splash-mark {{
    height: 0.32em; font-size: clamp(2.6rem, 8vw, 4.6rem); margin-top: -0.38em;
    background: #FDE047; border-radius: 6px; transform-origin: left center; transform: scaleX(0); opacity: 0;
    position: relative; z-index: -1;
    animation: kpSplashMark 0.45s ease-out 1.4s forwards;
}}
.kp-splash-sub {{
    font-family: 'Plus Jakarta Sans', sans-serif; font-weight: 700; letter-spacing: 0.08em;
    color: #1D4ED8; font-size: clamp(1.05rem, 2.6vw, 1.35rem); margin-top: 0.35rem;
    opacity: 0; transform: translateY(8px);
    animation: kpSplashSub 0.45s ease-out 1.65s forwards;
}}
@keyframes kpSplashPop {{ from {{ transform: scale(0.3) rotate(-18deg); opacity: 0; }} to {{ transform: scale(1) rotate(0); opacity: 1; }} }}
@keyframes kpSplashWrite {{ to {{ clip-path: inset(-10% 0 -10% 0); }} }}
@keyframes kpSplashPencil {{
    0% {{ left: 0; opacity: 1; }}
    92% {{ left: 100%; opacity: 1; }}
    100% {{ left: 100%; opacity: 0; }}
}}
@keyframes kpSplashMark {{ from {{ opacity: 1; }} to {{ transform: scaleX(1); opacity: 1; }} }}
@keyframes kpSplashSub {{ to {{ opacity: 1; transform: translateY(0); }} }}
@keyframes kpSplashOut {{ to {{ opacity: 0; transform: translateY(-4%); visibility: hidden; }} }}
@media (prefers-reduced-motion: reduce) {{
    .kp-splash {{ animation: kpSplashOut 0.3s ease 0.6s forwards; }}
    .kp-splash-logo, .kp-splash-title, .kp-splash-mark, .kp-splash-sub {{ animation: none; clip-path: none; transform: none; opacity: 1; }}
    .kp-splash-pencil {{ display: none; }}
}}
</style>
<div class="kp-splash" aria-hidden="true">
    <img class="kp-splash-logo" src="{logo}" alt=""/>
    <div class="kp-splash-word">
        <div class="kp-splash-title">{title}</div>
        <svg class="kp-splash-pencil" viewBox="0 0 24 24" aria-hidden="true">
            <g transform="translate(4 20) rotate(45) scale(0.42)">
                <polygon points="0,0 -6,-11 6,-11" fill="#FDE7B0" stroke="#0F172A" stroke-width="2.6" stroke-linejoin="round"/>
                <polygon points="0,0 -2.6,-4.8 2.6,-4.8" fill="#0F172A"/>
                <rect x="-6" y="-33" width="12" height="22" fill="#FACC15" stroke="#0F172A" stroke-width="2.6"/>
                <rect x="-6" y="-38" width="12" height="5" fill="#CBD5E1" stroke="#0F172A" stroke-width="2.6"/>
                <rect x="-6" y="-45" width="12" height="7" rx="3" fill="#F472B6" stroke="#0F172A" stroke-width="2.6"/>
            </g>
        </svg>
        <div class="kp-splash-mark"></div>
    </div>
    <div class="kp-splash-sub">{subtitle}</div>
</div>
"""


def render_splash_once() -> None:
    """Oturumun ilk açılışında logolu, kalemle yazılan giriş animasyonunu gösterir."""
    if not SPLASH_ENABLED or st.session_state.get("_splash_done"):
        return
    st.session_state["_splash_done"] = True
    st.markdown(
        SPLASH_HTML.format(
            logo=LOGO_DATA_URI, title=html.escape(SPLASH_TITLE), subtitle=html.escape(SPLASH_SUBTITLE)
        ),
        unsafe_allow_html=True,
    )


render_splash_once()



def trigger_loader_bar() -> None:
    st.session_state["loader_nonce"] = uuid.uuid4().hex[:8]


def render_top_loader_bar() -> None:
    nonce = st.session_state.get("loader_nonce", "init")
    st.markdown(
        f'<div class="kp-top-loader" data-load-id="{html.escape(str(nonce))}"></div>',
        unsafe_allow_html=True,
    )


def render_loading_skeleton(placeholder: Any, lines: int = 3) -> None:
    bars = "".join(['<div class="kp-skel-line" style="width:94%;"></div>' for _ in range(max(2, lines))])
    placeholder.markdown(f'<div class="kp-skeleton-card">{bars}</div>', unsafe_allow_html=True)


# ==============================================================================
# YARDIMCI FONKSİYONLAR, KOTA YÖNETİMİ VE KALICI DURUM
# ==============================================================================
def _default_parent_stats() -> Dict[str, Any]:
    return {
        "student_name": "Elif Öztürk",
        "grade": "8. Sınıf (LGS)",
        "parent_name": "Selin Öztürk",
        "parent_phone": "+905551234567",
        "parent_email": "veli@ornek.com",
        "weekly_solved": 340,
        "correct": 275,
        "wrong": 45,
        "blank": 20,
        "mentor_completion": 88,
        "strong_topic": "Okuma ve Anlama (%84)",
        "weak_topic": "Sözel Mantık ve Muhakeme",
        "next_exam_date": (datetime.now() + timedelta(days=5)).strftime("%d.%m.%Y"),
        "next_exam_name": "Türkiye Geneli 2027 LGS Deneme Sınavı",
    }


def _default_parent_announcements() -> List[Dict[str, str]]:
    return []


def _default_activity_log() -> List[Dict[str, str]]:
    return []


def current_user() -> Optional[Dict[str, Any]]:
    return st.session_state.get("user")


def current_user_id() -> Optional[int]:
    u = current_user()
    return int(u["id"]) if u else None


def load_persisted_state() -> Dict[str, Any]:
    """Giriş yapan kullanıcının kayıtlı durumu (her kullanıcıya ayrı)."""
    uid = current_user_id()
    return db.load_state(uid) if uid else {}


def save_persisted_state() -> None:
    payload = {
        "backpack_cart": st.session_state.get("backpack_cart", {}),
        "selected_tier_id": st.session_state.get("selected_tier_id", SUBSCRIPTION_TIERS[1]["id"]),
        "used_question_quota": int(st.session_state.get("used_question_quota", 0)),
        "used_coach_quota": int(st.session_state.get("used_coach_quota", 0)),
        "remaining_question_quota": int(st.session_state.get("remaining_question_quota", 0)),
        "remaining_coach_quota": int(st.session_state.get("remaining_coach_quota", 0)),
        "parent_student_stats": st.session_state.get("parent_student_stats", _default_parent_stats()),
        "parent_announcements": st.session_state.get("parent_announcements", _default_parent_announcements()),
        "activity_log": st.session_state.get("activity_log", _default_activity_log()),
        "question_bank_2028": st.session_state.get("question_bank_2028", []),
        "error_notebook_2028": st.session_state.get("error_notebook_2028", []),
        "b2b_org_cache": st.session_state.get("b2b_org_cache"),
        "last_b2b_summary": st.session_state.get("last_b2b_summary", ""),
        "active_plan": st.session_state.get("active_plan"),
        "reviews_completed": int(st.session_state.get("reviews_completed", 0)),
        "photo_solves": int(st.session_state.get("photo_solves", 0)),
    }
    uid = current_user_id()
    if not uid:
        return
    try:
        db.save_state(uid, payload)
    except Exception:
        pass


def attach_logo_to_doc(doc: Document, width_in: float = 1.15) -> None:
    if not LOGO_PNG_BYTES:
        return
    try:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(4)
        p.add_run().add_picture(io.BytesIO(LOGO_PNG_BYTES), width=Inches(width_in))
    except Exception:
        pass


def resolve_gemini_api_key() -> Optional[str]:
    key: Optional[str] = None
    try:
        if "GEMINI_API_KEY" in st.secrets:
            key = str(st.secrets["GEMINI_API_KEY"]).strip()
    except Exception:
        key = None
    if not key:
        env_val = os.environ.get("GEMINI_API_KEY", "").strip()
        if env_val:
            key = env_val
    return key if key else None


def resolve_iyzico_config() -> Tuple[str, str, str]:
    api_key = ""
    secret_key = ""
    base_url = "https://sandbox-api.iyzipay.com"
    try:
        if "iyzico" in st.secrets:
            sec = st.secrets["iyzico"]
            api_key = str(sec.get("api_key", "")).strip()
            secret_key = str(sec.get("secret_key", "")).strip()
            base_url = str(sec.get("base_url", base_url)).strip()
    except Exception:
        pass
    if not api_key:
        for k_name in ("IYZICO_API_KEY", "IYSU_API_KEY"):
            val = os.environ.get(k_name, "").strip()
            if val:
                api_key = val
                break
            try:
                if k_name in st.secrets:
                    api_key = str(st.secrets[k_name]).strip()
                    if api_key:
                        break
            except Exception:
                pass
    if not secret_key:
        for s_name in ("IYZICO_SECRET_KEY", "IYSU_SECRET_KEY"):
            val = os.environ.get(s_name, "").strip()
            if val:
                secret_key = val
                break
            try:
                if s_name in st.secrets:
                    secret_key = str(st.secrets[s_name]).strip()
                    if secret_key:
                        break
            except Exception:
                pass
    env_url = os.environ.get("IYZICO_BASE_URL", "").strip()
    if env_url:
        base_url = env_url
    else:
        try:
            if "IYZICO_BASE_URL" in st.secrets and str(st.secrets["IYZICO_BASE_URL"]).strip():
                base_url = str(st.secrets["IYZICO_BASE_URL"]).strip()
        except Exception:
            pass
    return api_key, secret_key, base_url


def _secret(path: str, default: str = "") -> str:
    """secrets.toml'dan 'bolum.anahtar' biçiminde değer okur."""
    try:
        node: Any = st.secrets
        for part in path.split("."):
            node = node[part]
        return str(node).strip()
    except Exception:
        return default


def client_ip() -> str:
    try:
        fwd = st.context.headers.get("X-Forwarded-For", "")
        if fwd:
            return fwd.split(",")[0].strip()
        return st.context.headers.get("X-Real-Ip", "") or "127.0.0.1"
    except Exception:
        return "127.0.0.1"


def initialize_iyzico_checkout(
    buyer_name: str,
    buyer_surname: str,
    buyer_email: str,
    buyer_phone: str,
    amount_try: float,
    basket_title: str,
    identity_number: str = "",
) -> Dict[str, Any]:
    api_key, secret_key, base_url = resolve_iyzico_config()
    if not api_key or not secret_key:
        return {
            "status": "failure",
            "errorMessage": (
                "IYZICO_API_KEY ve IYZICO_SECRET_KEY tanımlı değil. "
                "Streamlit Secrets veya ortam değişkenlerine anahtarlarınızı ekleyin."
            ),
        }
    options = {"api_key": api_key, "secret_key": secret_key, "base_url": base_url}
    price_str = f"{max(1.0, float(amount_try)):.2f}"
    user = current_user() or {}
    full_name = f"{buyer_name.strip()} {buyer_surname.strip()}".strip() or user.get("name", "Müşteri")
    # iyzico T.C. kimlik no ister; kullanıcı girmezse secrets'taki varsayılan kullanılır.
    identity = identity_number.strip() or _secret("iyzico.identity_number", "11111111111")
    address = _secret("iyzico.address", "Türkiye")
    city = _secret("iyzico.city", "Istanbul")
    request = {
        "locale": "tr",
        "conversationId": f"KP28-{uuid.uuid4().hex[:12]}",
        "price": price_str,
        "paidPrice": price_str,
        "currency": "TRY",
        "basketId": f"B28-{uuid.uuid4().hex[:12]}",
        "paymentGroup": "PRODUCT",
        "callbackUrl": _secret("app_url", "https://kursunkalemdijitalofis.com.tr"),
        "enabledInstallments": [1, 2, 3, 6, 9, 12],
        "buyer": {
            "id": f"U{user.get('id', 0)}",
            "name": buyer_name.strip() or full_name,
            "surname": buyer_surname.strip() or "-",
            "gsmNumber": buyer_phone.strip() or "+900000000000",
            "email": buyer_email.strip() or user.get("email", ""),
            "identityNumber": identity,
            "registrationAddress": address,
            "ip": client_ip(),
            "city": city,
            "country": "Turkey",
        },
        "shippingAddress": {"contactName": full_name, "city": city, "country": "Turkey", "address": address},
        "billingAddress": {"contactName": full_name, "city": city, "country": "Turkey", "address": address},
        "basketItems": [
            {
                "id": "BI_KP28_101",
                "name": basket_title[:120],
                "category1": "Dijital Eğitim 2027",
                "itemType": "VIRTUAL",
                "price": price_str,
            }
        ],
    }
    checkout_form_initialize = iyzipay.CheckoutFormInitialize().create(request, options)
    raw_res = checkout_form_initialize.read().decode("utf-8")
    return json.loads(raw_res)


def retrieve_iyzico_checkout(token: str) -> Dict[str, Any]:
    """Ödeme formunun sonucunu iyzico'dan sorgular. paymentStatus == 'SUCCESS' ise ödeme alınmıştır."""
    api_key, secret_key, base_url = resolve_iyzico_config()
    options = {"api_key": api_key, "secret_key": secret_key, "base_url": base_url}
    res = iyzipay.CheckoutForm().retrieve({"locale": "tr", "token": token}, options)
    return json.loads(res.read().decode("utf-8"))


def get_gemini_client() -> genai.Client:
    api_key = resolve_gemini_api_key()
    if not api_key:
        raise RuntimeError(
            "GEMINI_API_KEY bulunamadı. Lütfen Streamlit Secrets veya ortam "
            "değişkenlerinde GEMINI_API_KEY tanımlayın."
        )
    return genai.Client(api_key=api_key)


def call_gemini_text(
    prompt: str,
    system_instruction: str,
    model: str = DEFAULT_MODEL,
) -> str:
    client = get_gemini_client()
    candidate_models = [DEFAULT_MODEL, FALLBACK_MODEL]
    if model and model not in (
        "gemini-2.5-pro",
        "gemini-3.1-pro-preview",
        "gemini-2.5-flash-lite",
    ):
        candidate_models.insert(0, model)

    deduped_models: List[str] = []
    for m in candidate_models:
        if m and m not in deduped_models:
            deduped_models.append(m)

    errors_by_model: List[str] = []
    for candidate_model in deduped_models:
        for attempt in range(3):
            try:
                response = client.models.generate_content(
                    model=candidate_model,
                    contents=prompt,
                    config=types.GenerateContentConfig(
                        system_instruction=system_instruction,
                        temperature=0.7,
                    ),
                )
                text = getattr(response, "text", None)
                if not text:
                    raise RuntimeError("Modelden boş yanıt döndü.")
                return text.strip()
            except Exception as exc:
                err_str = str(exc)
                errors_by_model.append(f"[{candidate_model}] {err_str}")
                if "404" in err_str or "NOT_FOUND" in err_str:
                    break
                if any(tok in err_str for tok in ("503", "UNAVAILABLE", "429", "RESOURCE_EXHAUSTED", "overloaded")):
                    time.sleep(1.2 * (attempt + 1))
                    continue
                break

    raise RuntimeError(f"Model çağrısı başarısız oldu: {' | '.join(errors_by_model[-2:])}")


def call_gemini_vision(image_bytes: bytes, mime_type: str, prompt: str, system_instruction: str) -> str:
    """Fotoğraftaki soruyu Gemini'ye gönderir (çok kipli istek)."""
    client = get_gemini_client()
    errors: List[str] = []
    for model_name in [DEFAULT_MODEL, FALLBACK_MODEL]:
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=[types.Part.from_bytes(data=image_bytes, mime_type=mime_type), prompt],
                config=types.GenerateContentConfig(system_instruction=system_instruction, temperature=0.4),
            )
            text = getattr(response, "text", None)
            if text:
                return text.strip()
            errors.append(f"[{model_name}] boş yanıt")
        except Exception as exc:
            errors.append(f"[{model_name}] {exc}")
    raise RuntimeError("Görsel çözüm başarısız: " + " | ".join(errors[-2:]))


# ==============================================================================
# ÖĞRENME KAYITLARI, SERİ & ROZETLER
# ==============================================================================
def user_progress(uid: int) -> Dict[str, Any]:
    days = db.activity_days(uid)
    active = [d["day"] for d in days if d["solved"] > 0]
    today_row = next((d for d in days if d["day"] == date.today().isoformat()), None)
    user = db.get_user(uid) or {}
    return {
        "streak": gamify.streak(active),
        "solved_today": int(today_row["solved"]) if today_row else 0,
        "goal": int(user.get("daily_goal", 20)),
        "total_solved": sum(int(d["solved"]) for d in days),
    }


def check_and_award_badges() -> List[str]:
    """Yeni kazanılan rozetleri verir, ekranda kutlar ve adlarını döner."""
    uid = current_user_id()
    if not uid:
        return []
    prog = user_progress(uid)
    last = db.recent_exam_results(uid, 1)
    ctx = {
        "streak": prog["streak"],
        "total_solved": prog["total_solved"],
        "exams": db.exam_count(uid),
        "last_exam_size": (last[0]["correct"] + last[0]["wrong"] + last[0]["blank"]) if last else 0,
        "last_exam_accuracy": last[0]["accuracy"] if last else 0,
        "reviews_completed": int(st.session_state.get("reviews_completed", 0)),
        "photo_solves": int(st.session_state.get("photo_solves", 0)),
        "assignments_done": len([a for a in db.student_assignments(uid) if a.get("submitted_at")]),
    }
    new_names: List[str] = []
    for code in gamify.earned_badges(ctx):
        if db.award_badge(uid, code):
            name = gamify.BADGES[code][0]
            new_names.append(name)
            st.toast(f"Yeni rozet: {name}!", icon=":material/military_tech:")
    if new_names:
        st.balloons()
    return new_names


def record_exam(correct: int, wrong: int, blank: int, source: str, assignment_id: Optional[int] = None) -> None:
    """Deneme sonucunu kalıcı kayıtlara işler: günlük aktivite, sonuç geçmişi, ödev teslimi, rozetler."""
    uid = current_user_id()
    if not uid:
        return
    db.record_practice(uid, correct, wrong, blank)
    db.add_exam_result(uid, source, correct, wrong, blank, assignment_id)
    graded = correct + wrong
    acc = 100.0 * correct / graded if graded else 0.0
    if graded >= 3 and acc >= 80:
        st.session_state["celebrate"] = "Tam isabet!" if acc >= 100 else "Harika iş!"
        st.balloons()
    if assignment_id:
        db.submit_assignment(assignment_id, uid, correct, wrong, blank)
    check_and_award_badges()


def badge_svg(color: str, earned: bool) -> str:
    fill = color if earned else "#E2E8F0"
    stroke = "#0F172A" if earned else "#94A3B8"
    return (
        f'<svg aria-hidden="true" width="44" height="44" viewBox="0 0 48 48">'
        f'<path d="M16 4h16l-4 12h-8z" fill="{stroke}" opacity="0.25"/>'
        f'<circle cx="24" cy="28" r="15" fill="{fill}" stroke="{stroke}" stroke-width="2.5"/>'
        f'<path d="M24 20l2.6 5.3 5.8.8-4.2 4.1 1 5.8-5.2-2.7-5.2 2.7 1-5.8-4.2-4.1 5.8-.8z" '
        f'fill="#FFFFFF" stroke="{stroke}" stroke-width="1.2" stroke-linejoin="round"/></svg>'
    )


def render_badge_grid(earned: Dict[str, str]) -> None:
    cells = []
    for code, (name, how, color) in gamify.BADGES.items():
        got = code in earned
        when = datetime.fromisoformat(earned[code]).strftime("%d.%m.%Y") if got else how
        cells.append(
            f'<div class="kp-badge {"kp-badge-on" if got else "kp-badge-off"}">{badge_svg(color, got)}'
            f'<div class="kp-badge-name">{html.escape(name)}</div>'
            f'<div class="kp-badge-sub">{html.escape(when)}</div></div>'
        )
    st.markdown(f'<div class="kp-badge-grid">{"".join(cells)}</div>', unsafe_allow_html=True)


def clear_page_ephemeral_state(target_page: str) -> None:
    if target_page != "Öğrenci Portalı & Çözüm":
        st.session_state.pop("last_exam_report", None)
    if target_page != "Öğrenci Röntgeni":
        st.session_state.pop("xray_last_scan", None)


def navigate_to(page_name: str) -> None:
    if page_name in PAGES:
        clear_page_ephemeral_state(page_name)
        st.session_state.active_page = page_name
        trigger_loader_bar()
        st.rerun()


def record_activity(module_name: str, summary: str) -> None:
    logs: List[Dict[str, str]] = st.session_state.setdefault("activity_log", [])
    logs.insert(
        0,
        {
            "time": datetime.now().strftime("%H:%M:%S"),
            "module": module_name,
            "summary": summary,
        },
    )
    st.session_state["activity_log"] = logs[:15]
    save_persisted_state()


def get_tier_by_id(tier_id: str) -> Dict[str, Any]:
    return next((t for t in SUBSCRIPTION_TIERS if t["id"] == tier_id), SUBSCRIPTION_TIERS[1])


def ensure_bag_initialized() -> None:
    if "backpack_cart" not in st.session_state or not isinstance(st.session_state["backpack_cart"], dict):
        default_id = st.session_state.get("selected_tier_id", SUBSCRIPTION_TIERS[1]["id"])
        st.session_state["backpack_cart"] = {default_id: 1}


def get_bag_items() -> Dict[str, int]:
    ensure_bag_initialized()
    cleaned: Dict[str, int] = {}
    for k, v in st.session_state["backpack_cart"].items():
        try:
            qty = int(v)
        except Exception:
            qty = 0
        if qty > 0:
            cleaned[str(k)] = qty
    st.session_state["backpack_cart"] = cleaned
    return cleaned


def get_bag_summary() -> Dict[str, Any]:
    bag = get_bag_items()
    entries: List[Dict[str, Any]] = []
    total_monthly_price = 0
    total_monthly_quota = 0
    total_coach_quota = 0
    total_count = 0

    for tid, qty in bag.items():
        tier = get_tier_by_id(tid)
        line_price = int(tier["monthly_price"]) * qty
        line_quota = int(tier["monthly_quota"]) * qty
        line_coach = int(tier["coach_quota"]) * qty

        total_monthly_price += line_price
        total_monthly_quota += line_quota
        total_coach_quota += line_coach
        total_count += qty

        entries.append(
            {
                "tier": tier,
                "qty": qty,
                "line_price": line_price,
                "line_quota": line_quota,
                "line_coach": line_coach,
            }
        )

    if entries:
        title = entries[0]["tier"]["name"] if len(entries) == 1 else f"Ofis Çantası ({total_count} Paket)"
        segment = entries[0]["tier"]["segment"] if len(entries) == 1 else "Karma / Çoklu Paket"
        badge = entries[0]["tier"]["badge"] if len(entries) == 1 else f"{len(entries)} Çeşit"
    else:
        title = "Boş Çanta"
        segment = "Seçim Bekleniyor"
        badge = "0 Paket"

    return {
        "id": "backpack_bundle",
        "name": title,
        "segment": segment,
        "badge": badge,
        "monthly_price": total_monthly_price,
        "monthly_quota": total_monthly_quota,
        "coach_quota": total_coach_quota,
        "total_count": total_count,
        "entries": entries,
    }


def is_dev_mode() -> bool:
    """Geliştirme modu: ödeme olmadan kota açan simülasyon butonu sadece burada görünür.
    Açmak için secrets.toml'a dev_mode = true ekle ya da KK_DEV_MODE=1 ortam değişkenini ver."""
    if os.environ.get("KK_DEV_MODE", "").strip() == "1":
        return True
    try:
        return bool(st.secrets.get("dev_mode", False))
    except Exception:
        return False


BILLING_MONTHS = {"Aylık Abonelik": 1, "3 Aylık (%10 İndirim)": 3, "Yıllık (%20 İndirim)": 12}


def get_active_plan() -> Dict[str, Any]:
    """Kotayı belirleyen paket. Çantadaki paketler değil, ödemesi doğrulanmış paket geçerlidir.
    Paket süresi bitince ücretsiz pakete dönülür; kullanım her 30 günde bir sıfırlanır."""
    if is_dev_mode():
        return get_bag_summary()
    plan = st.session_state.get("active_plan")
    if not isinstance(plan, dict):
        return dict(FREE_PLAN)
    now = datetime.now()
    try:
        if plan.get("expires_at") and now > datetime.fromisoformat(plan["expires_at"]):
            st.session_state["active_plan"] = None
            st.session_state["used_question_quota"] = 0
            st.session_state["used_coach_quota"] = 0
            return dict(FREE_PLAN)
        period_start = datetime.fromisoformat(plan["period_start"])
        if now >= period_start + timedelta(days=30):
            while now >= period_start + timedelta(days=30):
                period_start += timedelta(days=30)
            plan["period_start"] = period_start.isoformat(timespec="seconds")
            st.session_state["used_question_quota"] = 0
            st.session_state["used_coach_quota"] = 0
    except Exception:
        pass
    return plan


def activate_plan(summary: Dict[str, Any], billing_cycle: str) -> None:
    """Ödeme doğrulandıktan sonra çantadaki paketleri aktif paket yapar."""
    now = datetime.now()
    months = BILLING_MONTHS.get(billing_cycle, 1)
    st.session_state["active_plan"] = {
        "name": summary["name"],
        "segment": summary.get("segment", ""),
        "badge": summary.get("badge", ""),
        "monthly_price": int(summary["monthly_price"]),
        "monthly_quota": int(summary["monthly_quota"]),
        "coach_quota": int(summary["coach_quota"]),
        "total_count": int(summary.get("total_count", 1)),
        "billing_cycle": billing_cycle,
        "started_at": now.isoformat(timespec="seconds"),
        "period_start": now.isoformat(timespec="seconds"),
        "expires_at": (now + timedelta(days=30 * months)).isoformat(timespec="seconds"),
    }
    sync_plan_quotas(st.session_state["active_plan"], reset_usage=True)


def sync_plan_quotas(summary: Dict[str, Any], reset_usage: bool = False) -> None:
    if reset_usage:
        st.session_state["used_question_quota"] = 0
        st.session_state["used_coach_quota"] = 0

    used_q = max(0, int(st.session_state.get("used_question_quota", 0)))
    used_c = max(0, int(st.session_state.get("used_coach_quota", 0)))
    total_q = max(0, int(summary["monthly_quota"]))
    total_c = max(0, int(summary["coach_quota"]))

    if used_q > total_q:
        used_q = total_q
        st.session_state["used_question_quota"] = used_q
    if used_c > total_c:
        used_c = total_c
        st.session_state["used_coach_quota"] = used_c

    st.session_state["remaining_question_quota"] = max(0, total_q - used_q)
    st.session_state["remaining_coach_quota"] = max(0, total_c - used_c)
    save_persisted_state()


def cb_add_tier_to_bag(tier_id: str, qty: int = 1) -> None:
    bag = dict(get_bag_items())
    bag[tier_id] = bag.get(tier_id, 0) + max(1, int(qty))
    st.session_state["backpack_cart"] = bag
    st.session_state["selected_tier_id"] = tier_id
    sync_plan_quotas(get_active_plan(), reset_usage=False)
    tier = get_tier_by_id(tier_id)
    record_activity("Ofis Çantası", f"{tier['name']} çantaya eklendi")


def cb_update_bag_qty(tier_id: str, delta: int) -> None:
    bag = dict(get_bag_items())
    if tier_id not in bag:
        return
    new_qty = bag[tier_id] + delta
    tier = get_tier_by_id(tier_id)
    if new_qty <= 0:
        del bag[tier_id]
        record_activity("Ofis Çantası", f"{tier['name']} çantadan çıkarıldı")
    else:
        bag[tier_id] = new_qty
        record_activity("Ofis Çantası", f"{tier['name']} adedi {new_qty} yapıldı")
    st.session_state["backpack_cart"] = bag
    if bag:
        st.session_state["selected_tier_id"] = next(iter(bag.keys()))
    sync_plan_quotas(get_active_plan(), reset_usage=False)


def cb_clear_bag() -> None:
    st.session_state["backpack_cart"] = {}
    sync_plan_quotas(get_active_plan(), reset_usage=False)
    record_activity("Ofis Çantası", "Çantadaki tüm paketler silindi")


def consume_quota(kind: str, amount: int = 1) -> Tuple[bool, str]:
    summary = get_active_plan()
    sync_plan_quotas(summary, reset_usage=False)

    if kind == "question":
        rem = int(st.session_state.get("remaining_question_quota", 0))
        if rem < amount:
            return False, (
                f"Aktif paketinizde ({summary['name']}) yeterli soru üretim kotası kalmadı "
                f"(Kalan: {rem}, İstenen: {amount}). Paketler vitrininden paket ekleyebilirsiniz."
            )
        st.session_state["used_question_quota"] = int(st.session_state.get("used_question_quota", 0)) + amount
        sync_plan_quotas(summary, reset_usage=False)
        return True, ""

    rem_c = int(st.session_state.get("remaining_coach_quota", 0))
    if rem_c < amount:
        return False, (
            f"Aktif paketinizde ({summary['name']}) Dijimentor plan kotası doldu "
            f"(Kalan: {rem_c}). Paketler vitrininden paket ekleyebilirsiniz."
        )
    st.session_state["used_coach_quota"] = int(st.session_state.get("used_coach_quota", 0)) + amount
    sync_plan_quotas(summary, reset_usage=False)
    return True, ""


def refund_quota(kind: str, amount: int = 1) -> None:
    if kind == "question":
        st.session_state["used_question_quota"] = max(
            0, int(st.session_state.get("used_question_quota", 0)) - amount
        )
    else:
        st.session_state["used_coach_quota"] = max(
            0, int(st.session_state.get("used_coach_quota", 0)) - amount
        )
    sync_plan_quotas(get_active_plan(), reset_usage=False)


def add_to_question_bank(
    grade: str,
    branch: str,
    outcome: str,
    difficulty: str,
    content_text: str,
) -> None:
    bank: List[Dict[str, str]] = st.session_state.setdefault("question_bank_2028", [])
    bank.insert(
        0,
        {
            "id": uuid.uuid4().hex[:6].upper(),
            "date": datetime.now().strftime("%d.%m.%Y %H:%M"),
            "grade": grade,
            "branch": branch,
            "outcome": outcome,
            "difficulty": difficulty,
            "content": content_text,
        },
    )
    st.session_state["question_bank_2028"] = bank[:40]
    save_persisted_state()


def add_to_error_notebook(
    question_num: int,
    question_text: str,
    user_ans: str,
    correct_ans: str,
    elapsed_sec: float,
) -> None:
    nb: List[Dict[str, Any]] = st.session_state.setdefault("error_notebook_2028", [])
    now = datetime.now()
    diag = "Süre / Dikkat Baskısı" if elapsed_sec > 90 else "Çeldirici / Kavram İncelemesi"
    nb.insert(
        0,
        {
            "id": uuid.uuid4().hex[:6].upper(),
            "num": question_num,
            "date": now.strftime("%d.%m.%Y %H:%M"),
            "review_1d": (now + timedelta(days=1)).strftime("%d.%m.%Y"),
            "review_3d": (now + timedelta(days=3)).strftime("%d.%m.%Y"),
            "review_7d": (now + timedelta(days=7)).strftime("%d.%m.%Y"),
            "user_ans": user_ans,
            "correct_ans": correct_ans,
            "elapsed_sec": round(elapsed_sec, 1),
            "diagnosis": diag,
            "question_text": question_text,
            "stage": 0,
            "done": False,
        },
    )
    st.session_state["error_notebook_2028"] = nb[:50]
    save_persisted_state()


# ==============================================================================
# GÜÇLÜ SORU AYRIŞTIRICI & DOCX ÜRETİCİLERİ
# ==============================================================================
def extract_answer_key(raw_text: str) -> Dict[int, str]:
    if not raw_text.strip():
        return {}

    parts = re.split(
        r"(?:^|\n)#*\s*(?:Cevap\s*Anahtar\w*|Yanıt\s*Anahtar\w*)[^\n]*",
        raw_text,
        maxsplit=1,
        flags=re.IGNORECASE,
    )
    key_map: Dict[int, str] = {}

    patterns = [
        r"(?:Soru\s*)?(\d+)\s*[\.\-\:\)\]]\s*\**\s*([A-D])\b",
        r"(\d+)\s*\.\s*Soru\s*[\:\-\)\]]\s*\**\s*([A-D])\b",
        r"\b(\d+)\s*[-–—]\s*\**\s*([A-D])\b",
    ]

    if len(parts) > 1:
        for pat in patterns:
            for m in re.finditer(pat, parts[1], flags=re.IGNORECASE):
                q_num = int(m.group(1))
                ans_letter = m.group(2).upper()
                if q_num not in key_map:
                    key_map[q_num] = ans_letter

    q_blocks = re.split(
        r"(?=(?:^|\n)\s*(?:#+\s*)?(?:\*{0,2})\s*(?:Soru\s*\d+|\d+\s*[\.\)]))",
        parts[0],
        flags=re.IGNORECASE,
    )
    for blk in q_blocks:
        num_m = re.match(r"^\s*(?:#+\s*)?(?:\*{0,2})\s*(?:Soru\s*)?(\d+)\b", blk, flags=re.IGNORECASE)
        if not num_m:
            continue
        q_num = int(num_m.group(1))
        if q_num in key_map:
            continue
        ans_m = re.search(
            r"(?:Doğru\s*(?:Cevap|Yanıt)|Cevap|Yanıt)\s*[\:\-–]\s*\**\s*([A-D])\b",
            blk,
            flags=re.IGNORECASE,
        )
        if ans_m:
            key_map[q_num] = ans_m.group(1).upper()

    return key_map


def parse_quiz_questions(raw_text: str) -> List[Dict[str, Any]]:
    if not raw_text.strip():
        return []

    answer_map = extract_answer_key(raw_text)
    main_body = re.split(
        r"(?:^|\n)\s*#*\s*(?:Cevap\s*Anahtar\w*|Yanıt\s*Anahtar\w*)",
        raw_text,
        maxsplit=1,
        flags=re.IGNORECASE,
    )[0].strip()

    if not main_body:
        return []

    header_split_pattern = (
        r"(?=(?:^|\n)\s*(?:#{1,6}\s*)?(?:\*\*)?\s*(?:Soru\s+\d+|\d+\s*(?:\.\s*Soru|\)|\.\s+(?=[A-ZÇĞİÖŞÜ\"“'\*]))))"
    )
    raw_blocks = re.split(header_split_pattern, main_body, flags=re.IGNORECASE | re.MULTILINE)

    parsed: List[Dict[str, Any]] = []
    seen_nums: set[int] = set()

    for idx_b, block in enumerate(raw_blocks, start=1):
        text = block.strip()
        if len(text) < 15:
            continue

        num_match = re.match(
            r"^\s*(?:#{1,6}\s*)?(?:\*\*)?\s*(?:Soru\s+)?(\d+)\b",
            text,
            flags=re.IGNORECASE,
        )
        q_num = int(num_match.group(1)) if num_match else idx_b
        if q_num in seen_nums:
            q_num = max(seen_nums) + 1

        lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
        stem_lines: List[str] = []
        options: Dict[str, str] = {}
        current_opt: Optional[str] = None

        opt_pattern = re.compile(
            r"^\s*(?:[-*•]\s*)?(?:\*\*)?\s*([A-D])\s*[\)\.\:]\s*(?:\*\*)?\s*(.+)$",
            flags=re.IGNORECASE,
        )

        for ln in lines:
            if re.match(r"^\s*\[?(?:Redaksiyon|Çözüm|Cevap\s*Anahtar|Doğru\s*Cevap)", ln, flags=re.IGNORECASE):
                current_opt = None
                continue
            if re.match(r"^\s*[-_=]{3,}\s*$", ln):
                continue

            opt_m = opt_pattern.match(ln)
            if opt_m:
                letter = opt_m.group(1).upper()
                opt_val = re.sub(r"\*\*$", "", opt_m.group(2).strip()).strip()
                options[letter] = opt_val
                current_opt = letter
            elif current_opt:
                extra = re.sub(r"\*\*$", "", ln).strip()
                if extra:
                    options[current_opt] = f"{options[current_opt]} {extra}".strip()
            else:
                stem_lines.append(ln)

        if len(options) >= 2:
            ordered_opts = {k: options[k] for k in ["A", "B", "C", "D"] if k in options}
            for k, v in options.items():
                if k not in ordered_opts:
                    ordered_opts[k] = v

            seen_nums.add(q_num)
            parsed.append(
                {
                    "num": q_num,
                    "stem": "\n".join(stem_lines).strip(),
                    "options": ordered_opts,
                    "correct": answer_map.get(q_num),
                    "full_text": text,
                }
            )

    return parsed


def format_countdown(target: datetime) -> Tuple[int, int, int, int]:
    now = datetime.now()
    delta = target - now
    if delta.total_seconds() <= 0:
        return 0, 0, 0, 0
    total_seconds = int(delta.total_seconds())
    days = total_seconds // 86400
    hours = (total_seconds % 86400) // 3600
    minutes = (total_seconds % 3600) // 60
    seconds = total_seconds % 60
    return days, hours, minutes, seconds


def render_countdown_boxes(target: datetime) -> None:
    d, h, m, s = format_countdown(target)
    c1, c2, c3, c4 = st.columns(4)
    for col, val, lbl in [
        (c1, f"{d:03d}", "GÜN"),
        (c2, f"{h:02d}", "SAAT"),
        (c3, f"{m:02d}", "DAKİKA"),
        (c4, f"{s:02d}", "SANİYE"),
    ]:
        col.markdown(
            f"""
            <div class="kp-metric-box">
                <div class="kp-metric-num">{val}</div>
                <div class="kp-metric-lbl">{lbl}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )


def render_page_strip(page_name: str) -> None:
    title, desc = PAGE_DESCRIPTIONS.get(page_name, (page_name, "Kurşun Kalem Dijital Ofis 2027 çalışma alanı."))
    st.markdown(
        f"""
        <div class="kp-page-strip">
            <div>
                <h2 class="kp-page-strip-title">{html.escape(title)}</h2>
                <p class="kp-page-strip-desc">{html.escape(desc)}</p>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def build_test_docx(
    institution: str,
    branch: str,
    grade: str,
    learning_area: str,
    outcome: str,
    context_theme: str,
    generated_content: str,
) -> bytes:
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    attach_logo_to_doc(doc, width_in=1.15)
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_p.add_run(f"{institution.upper()} — DEĞERLENDİRME TESTİ")
    title_run.bold = True
    title_run.font.size = Pt(14)
    title_run.font.color.rgb = RGBColor(30, 41, 82)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_run = sub_p.add_run(f"Branş: {branch}  |  Sınıf: {grade}  |  Öğrenme Alanı: {learning_area}")
    sub_run.font.size = Pt(10)
    sub_run.font.color.rgb = RGBColor(71, 85, 105)

    meta_p = doc.add_paragraph()
    meta_run = meta_p.add_run(
        f"Hedef Kazanım: {outcome}\nBağlam / Tema: {context_theme}\nTarih: {datetime.now().strftime('%d.%m.%Y')}"
    )
    meta_run.font.size = Pt(9.5)
    meta_run.font.color.rgb = RGBColor(51, 65, 85)

    doc.add_paragraph("─" * 78)

    for line in generated_content.splitlines():
        cleaned = line.strip()
        if not cleaned:
            continue
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15
        run = p.add_run(cleaned)
        run.font.size = Pt(10.5)
        if re.match(r"^(Soru\s*\d+|Cevap\s*Anahtarı|Çözüm)", cleaned, re.IGNORECASE):
            run.bold = True
            run.font.color.rgb = RGBColor(30, 41, 82)

    footer_p = doc.add_paragraph()
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer_run = footer_p.add_run(f"\n{BRAND_NAME} • 2027")
    footer_run.italic = True
    footer_run.font.size = Pt(8.5)
    footer_run.font.color.rgb = RGBColor(100, 116, 139)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def build_coach_docx(title_meta: str, plan_text: str) -> bytes:
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    attach_logo_to_doc(doc, width_in=1.15)
    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = h.add_run(f"{BRAND_NAME.upper()} — DİJİMENTOR ÇALIŞMA PLANI")
    r.bold = True
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(30, 41, 82)

    m = doc.add_paragraph()
    mr = m.add_run(f"{title_meta}\nTarih: {datetime.now().strftime('%d.%m.%Y')}")
    mr.font.size = Pt(9.5)
    mr.font.color.rgb = RGBColor(71, 85, 105)

    doc.add_paragraph("─" * 78)

    for line in plan_text.splitlines():
        cleaned = line.strip()
        if not cleaned:
            continue
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15
        run = p.add_run(cleaned)
        run.font.size = Pt(10.5)
        if cleaned.startswith("#") or cleaned.endswith(":"):
            run.bold = True
            run.font.color.rgb = RGBColor(30, 41, 82)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def build_commercial_brief_docx(
    title: str,
    summary: Dict[str, Any],
    billing_cycle: str,
    installment: int,
    total_amount: int,
    extra_notes: str,
) -> bytes:
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    attach_logo_to_doc(doc, width_in=1.15)
    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = h.add_run(f"{BRAND_NAME.upper()} — OFİS ÇANTASI & ABONELİK ÖZETİ")
    r.bold = True
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(30, 41, 82)

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub.add_run(f"{title}  |  Tarih: {datetime.now().strftime('%d.%m.%Y')}").font.size = Pt(9.5)

    doc.add_paragraph("─" * 78)

    lines: List[str] = [
        f"Çanta Durumu: {summary['name']} ({summary['total_count']} Adet Paket)",
        f"Aylık Liste Toplamı: {summary['monthly_price']:,} ₺".replace(",", "."),
        f"Tercih Edilen Dönem: {billing_cycle}",
        f"iyzico Taksit Planı: {installment} Taksit",
        f"Hesaplanan Toplam Tutar: {total_amount:,} ₺".replace(",", "."),
        f"Toplam Aylık Soru Kotası: {summary['monthly_quota']:,} Soru".replace(",", "."),
        f"Toplam Dijimentor Plan Kotası: {summary['coach_quota']} Plan / Ay",
        "",
        "Çantadaki Paketler:",
    ]

    for entry in summary.get("entries", []):
        t = entry["tier"]
        qty = entry["qty"]
        lines.append(
            f"• {t['name']} x{qty} ({t['segment']}) — {entry['line_price']:,} ₺/ay | "
            f"{entry['line_quota']:,} soru/ay".replace(",", ".")
        )

    if extra_notes.strip():
        lines += ["", "Ek / Kurumsal Notlar:", extra_notes.strip()]

    for line in lines:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(3.5)
        run = p.add_run(line)
        run.font.size = Pt(10.5)
        if line.startswith(("Çanta Durumu", "Çantadaki Paketler", "Ek / Kurumsal Notlar")):
            run.bold = True
            run.font.color.rgb = RGBColor(30, 41, 82)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


def render_school_backpack_panel() -> Dict[str, Any]:
    summary = get_bag_summary()
    entries = summary["entries"]

    st.markdown('<div class="kp-bag-shell">', unsafe_allow_html=True)
    h_col, clr_col = st.columns([4, 1])
    with h_col:
        st.markdown("#### Kurşun Kalem Ofis Çantası")
        st.caption("Seçili paketler, adet kontrolleri, kota çubukları ve aylık toplam özet.")
    with clr_col:
        if entries:
            st.button(
                "Çantayı Boşalt",
                key="bag_clear_all",
                width="stretch",
                on_click=cb_clear_bag,
            )

    v_col, list_col = st.columns([1.1, 2.9])

    with v_col:
        chips_html = ""
        for entry in entries:
            safe_name = html.escape(entry["tier"]["name"])
            safe_qty = int(entry["qty"])
            chips_html += (
                f'<div style="background:#FFFFFF;color:#0F172A;border:1px solid #CBD5E1;'
                f'border-radius:7px;padding:0.22rem 0.5rem;font-size:0.74rem;font-weight:600;'
                f'margin-top:0.28rem;width:100%;text-align:center;">'
                f'{safe_name} (x{safe_qty})</div>'
            )
        if not chips_html:
            chips_html = '<div style="color:#E2E8F0;font-size:0.78rem;margin-top:0.35rem;text-align:center;">Çanta boş</div>'

        st.markdown(
            f"""
            <div class="kp-bag-visual">
                <img src="{BACKPACK_DATA_URI}" alt="Ofis Çantası"/>
                <div class="kp-bag-count-pill">{summary['total_count']} Paket</div>
                <div style="width:100%;margin-top:0.3rem;">{chips_html}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with list_col:
        if not entries:
            st.info("Ofis çantanız şu an boş. Aşağıdaki paketlerden **Çantaya Ekle** butonunu kullanabilirsiniz.")
        else:
            for entry in entries:
                t = entry["tier"]
                qty = entry["qty"]
                c_info, c_minus, c_plus, c_rem = st.columns([4.6, 0.75, 0.75, 1.15])
                with c_info:
                    st.markdown(
                        f"""
                        <div class="kp-bag-item-chip">
                            <div>
                                <div style="font-weight:700;font-size:0.88rem;color:#0F172A;">
                                    {html.escape(t['name'])}
                                    <span style="color:#64748B;font-size:0.76rem;font-weight:500;"> • {html.escape(t['segment'])}</span>
                                </div>
                                <div style="color:#475569;font-size:0.78rem;margin-top:0.12rem;">
                                    Adet: <b>{qty}</b> • Soru: <b>{entry['line_quota']:,}/ay</b> • Dijimentor: <b>{entry['line_coach']}/ay</b>
                                </div>
                            </div>
                            <div style="font-weight:700;font-size:0.92rem;color:#2563EB;white-space:nowrap;">
                                {entry['line_price']:,} ₺/ay
                            </div>
                        </div>
                        """.replace(",", "."),
                        unsafe_allow_html=True,
                    )
                with c_minus:
                    st.button("−", key=f"bag_minus_{t['id']}", width="stretch", on_click=cb_update_bag_qty, args=(t["id"], -1))
                with c_plus:
                    st.button("+", key=f"bag_plus_{t['id']}", width="stretch", on_click=cb_update_bag_qty, args=(t["id"], +1))
                with c_rem:
                    st.button("Çıkar", key=f"bag_del_{t['id']}", width="stretch", on_click=cb_update_bag_qty, args=(t["id"], -qty))

            st.markdown("<div style='margin-top:0.45rem;'></div>", unsafe_allow_html=True)
            m1, m2, m3 = st.columns(3)
            m1.metric("Aylık Toplam", f"{summary['monthly_price']:,} ₺".replace(",", "."))
            m2.metric("Aylık Soru Kotası", f"{summary['monthly_quota']:,}".replace(",", "."))
            m3.metric("Dijimentor Plan", f"{summary['coach_quota']} / Ay")

            plan_now = get_active_plan()
            rem_q = int(st.session_state.get("remaining_question_quota", plan_now["monthly_quota"]))
            rem_c = int(st.session_state.get("remaining_coach_quota", plan_now["coach_quota"]))
            q_pct = min(1.0, max(0.0, rem_q / max(1, int(plan_now["monthly_quota"]))))
            c_pct = min(1.0, max(0.0, rem_c / max(1, int(plan_now["coach_quota"]))))

            st.markdown("<div style='margin-top:0.55rem;'></div>", unsafe_allow_html=True)
            st.caption(f"Aktif paketin: **{plan_now['name']}** (çantadaki paketler ödeme doğrulanınca aktif olur)")
            pb1, pb2 = st.columns(2)
            with pb1:
                st.progress(q_pct, text=f"Kullanılabilir Soru Kotası: {rem_q:,} / {plan_now['monthly_quota']:,}".replace(",", "."))
            with pb2:
                st.progress(c_pct, text=f"Kullanılabilir Dijimentor: {rem_c} / {plan_now['coach_quota']} Plan")

    st.markdown("</div>", unsafe_allow_html=True)
    return summary


def render_iyzico_simulator(
    package_name: str,
    monthly_price: int,
    summary: Optional[Dict[str, Any]] = None,
) -> None:
    bag_summary = summary if summary is not None else get_bag_summary()
    effective_monthly = int(bag_summary["monthly_price"]) if bag_summary["total_count"] > 0 else int(monthly_price)
    display_label = f"{package_name} ({bag_summary['total_count']} Paket)" if bag_summary["total_count"] > 0 else package_name

    st.markdown('<div class="kp-pay-box">', unsafe_allow_html=True)
    st.markdown("#### iyzico Taksit & Ödeme Simülatörü")
    st.caption("Dönem avantajı ve taksit planına göre tutarı hesaplayın, iyzico ödeme formu başlatın veya DOCX özet indirin.")

    c1, c2, c3 = st.columns(3)
    with c1:
        billing_cycle = st.selectbox("Dönem Seçimi", ["Aylık Abonelik", "3 Aylık (%10 İndirim)", "Yıllık (%20 İndirim)"], key=f"cycle_{display_label}")
    with c2:
        installment = st.selectbox("Taksit Sayısı (iyzico)", [1, 2, 3, 6, 9, 12], index=2, key=f"inst_{display_label}")
    with c3:
        card_brand = st.selectbox("Kart Programı", ["Bonus", "Maximum", "World", "Axess", "Paraf", "Diğer Banka Kartı"], key=f"card_{display_label}")

    if "3 Aylık" in billing_cycle:
        total_amount = round(effective_monthly * 3 * 0.90)
    elif "Yıllık" in billing_cycle:
        total_amount = round(effective_monthly * 12 * 0.80)
    else:
        total_amount = effective_monthly

    monthly_installment = total_amount / installment

    st.markdown("<div style='margin:0.65rem 0;'></div>", unsafe_allow_html=True)
    m1, m2, m3 = st.columns(3)
    m1.metric("Toplam Tutar", f"{total_amount:,} ₺".replace(",", "."))
    m2.metric(f"{installment} Taksit Aylık", f"{monthly_installment:,.2f} ₺".replace(",", " "))
    m3.metric("Ödeme Altyapısı", f"iyzico • {card_brand}")

    st.markdown("<div style='margin:0.65rem 0;'></div>", unsafe_allow_html=True)
    _u = current_user() or {}
    _name_parts = str(_u.get("name", "")).split()
    b_col1, b_col2, b_col3, b_col4 = st.columns(4)
    with b_col1:
        buyer_name = st.text_input("Ad", value=" ".join(_name_parts[:-1]) or (_name_parts[0] if _name_parts else ""), key=f"iyz_name_{display_label}")
    with b_col2:
        buyer_surname = st.text_input("Soyad", value=_name_parts[-1] if len(_name_parts) > 1 else "", key=f"iyz_surname_{display_label}")
    with b_col3:
        buyer_email = st.text_input("E-posta", value=_u.get("email", ""), key=f"iyz_email_{display_label}")
    with b_col4:
        buyer_tc = st.text_input("T.C. kimlik no", max_chars=11, key=f"iyz_tc_{display_label}",
                                 help="iyzico fatura için ister. Bilgi sadece ödeme isteğine eklenir, saklanmaz.")

    btn_c1, btn_c2, btn_c3 = st.columns([1.8, 1.8, 1.1])
    with btn_c1:
        if st.button("iyzico Ödeme Formu Başlat (API)", key=f"iyz_live_btn_{display_label}", type="primary", width="stretch"):
            if total_amount <= 0:
                st.warning("Ödeme başlatmak için önce çantanıza en az bir paket ekleyin.")
            else:
                load_slot = st.empty()
                render_loading_skeleton(load_slot, lines=2)
                try:
                    res_dict = initialize_iyzico_checkout(
                        buyer_name=buyer_name,
                        buyer_surname=buyer_surname,
                        buyer_email=buyer_email,
                        buyer_phone="",
                        amount_try=float(total_amount),
                        basket_title=f"{BRAND_NAME} 2027 - {display_label} ({billing_cycle})",
                        identity_number=buyer_tc,
                    )
                except Exception as exc:
                    res_dict = {"status": "failure", "errorMessage": str(exc)}
                load_slot.empty()
                if res_dict.get("status") == "success" and res_dict.get("token"):
                    # Kota burada AÇILMAZ; ödeme iyzico'dan doğrulanınca açılır.
                    db.create_payment(res_dict["token"], current_user_id(), int(total_amount), billing_cycle, bag_summary)
                    record_activity("iyzico", f"{display_label} ({total_amount:,} ₺) ödeme formu açıldı".replace(",", "."))
                    st.session_state["iyzico_last_response"] = res_dict
                    st.success("Ödeme formu hazır. Ödemeyi tamamladıktan sonra aşağıdan doğrulayın.")
                else:
                    st.error(f"iyzico yanıtı: {res_dict.get('errorMessage', 'Bilinmeyen hata')}")

    with btn_c2:
        if is_dev_mode():
            if st.button("Simülasyon ile Kota Aktifleştir (geliştirici)", key=f"pay_btn_{display_label}", width="stretch"):
                activate_plan(bag_summary, billing_cycle)
                record_activity("Simülasyon", f"{display_label} ({total_amount:,} ₺) aktiflendi".replace(",", "."))
                st.success("Geliştirici modu: paket ödeme olmadan aktifleştirildi.")
        else:
            pending_pay = db.latest_pending_payment(current_user_id())
            if st.button("Ödemeyi tamamladım, doğrula", key=f"verify_btn_{display_label}",
                         width="stretch", disabled=pending_pay is None):
                try:
                    chk = retrieve_iyzico_checkout(pending_pay["token"])
                except Exception as exc:
                    chk = {"status": "failure", "errorMessage": str(exc)}
                paid_ok = (
                    chk.get("status") == "success"
                    and chk.get("paymentStatus") == "SUCCESS"
                    and float(chk.get("paidPrice") or 0) + 0.01 >= float(pending_pay["amount"])
                )
                if paid_ok:
                    db.mark_payment(pending_pay["token"], "success")
                    activate_plan(pending_pay["plan"], pending_pay["billing_cycle"])
                    record_activity("iyzico", f"Ödeme doğrulandı: {pending_pay['plan'].get('name', '')}")
                    st.session_state.pop("iyzico_last_response", None)
                    st.success("Ödemen doğrulandı, paketin aktif!")
                    st.balloons()
                elif chk.get("paymentStatus") == "FAILURE":
                    db.mark_payment(pending_pay["token"], "failed")
                    st.error("Ödeme başarısız görünüyor. Kartınızdan çekim yapılmadıysa tekrar deneyebilirsiniz.")
                else:
                    st.info(f"Ödeme henüz tamamlanmamış görünüyor. {chk.get('errorMessage', '')}")

    with btn_c3:
        brief_docx = build_commercial_brief_docx(
            title=f"{display_label} Özeti",
            summary=bag_summary,
            billing_cycle=billing_cycle,
            installment=int(installment),
            total_amount=int(total_amount),
            extra_notes=f"Kart Programı: {card_brand}",
        )
        st.download_button(
            label="Özet DOCX",
            data=brief_docx,
            file_name="kursun_kalem_ofis_2027_cantasi_ozet.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key=f"dl_brief_{display_label}",
            width="stretch",
        )

    iyz_res = st.session_state.get("iyzico_last_response")
    if isinstance(iyz_res, dict) and iyz_res.get("status") == "success":
        pay_url = iyz_res.get("paymentPageUrl")
        form_html = iyz_res.get("checkoutFormContent")
        if pay_url:
            st.markdown(
                f'<div style="margin-top:0.75rem;"><a href="{html.escape(str(pay_url))}" target="_blank">'
                f'<button style="background:#2563EB;color:#FFFFFF;border:none;border-radius:8px;'
                f'padding:0.55rem 1.15rem;font-weight:600;cursor:pointer;">'
                f'Güvenli iyzico Ödeme Sayfasını Aç</button></a></div>',
                unsafe_allow_html=True,
            )
        if form_html:
            embed_html(str(form_html), height=460)

    st.markdown("</div>", unsafe_allow_html=True)


def render_student_xray_panel() -> None:
    st.markdown(
        """
        <div class="kp-studio-banner">
            <div class="kp-studio-title">Öğrenci Röntgeni — Kazanım bazlı gerçek performans</div>
            <div class="kp-studio-sub">
                Teslim edilen ödevler ve çözülen denemelerden hesaplanır: hangi kazanımda ne kadar başarılı,
                sınıf ortalamasının neresinde, hangi ödevler eksik.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    student_id: Optional[int] = None
    student_name = ""
    teacher_scope: Optional[int] = None
    grade_hint = ""
    if ROLE == "ogretmen":
        classes = db.teacher_classes(USER["id"])
        if not classes:
            st.info("Röntgen için önce **Sınıflarım** sayfasından sınıf oluşturup öğrenci ekleyin.")
            return
        f1, f2 = st.columns(2)
        with f1:
            cls_map = {c["name"]: c for c in classes}
            cls_name = st.selectbox("Sınıf", list(cls_map.keys()), key="xray_cls")
        members = db.class_members(cls_map[cls_name]["id"])
        if not members:
            st.info("Bu sınıfta henüz öğrenci yok.")
            return
        with f2:
            mem_map = {m["name"]: m for m in members}
            student_name = st.selectbox("Öğrenci", list(mem_map.keys()), key="xray_stu")
        student_id = mem_map[student_name]["id"]
        teacher_scope = USER["id"]
        grade_hint = cls_map[cls_name].get("grade") or ""
    elif ROLE == "veli":
        student_id = LIVE_STUDENT_ID
        if not student_id:
            st.info("Röntgen için önce **Profilim** sayfasından çocuğunuzun hesabını bağlayın.")
            return
        child = db.get_user(student_id) or {}
        student_name = child.get("name", "")
        grade_hint = child.get("grade") or ""
    else:
        student_id = USER["id"]
        student_name = USER["name"]
        grade_hint = USER.get("grade") or ""

    xr = db.student_xray(student_id, teacher_scope)
    prog = user_progress(student_id)
    if not xr["done"] and not xr["exams"]:
        st.info(
            f"{student_name} için henüz yeterli veri yok. Ödev teslim edildikçe ve mini denemeler çözüldükçe "
            "röntgen kendiliğinden dolar."
        )
        if xr["assigned"]:
            st.caption(f"Atanmış ama teslim edilmemiş ödev: {len(xr['missing'])}")
        return

    m1, m2, m3, m4 = st.columns(4)
    if xr["avg_acc"] is not None:
        diff = round(xr["avg_acc"] - (xr["avg_class"] or xr["avg_acc"]), 1)
        m1.metric("Ödev doğruluğu", f"%{xr['avg_acc']}", f"{diff:+} puan sınıfa göre" if xr["avg_class"] is not None else None)
    else:
        m1.metric("Ödev doğruluğu", "—")
    m2.metric("Teslim edilen ödev", f"{xr['done']} / {xr['assigned']}",
              f"{len(xr['overdue'])} gecikmiş" if xr["overdue"] else None, delta_color="inverse")
    m3.metric("Çalışma serisi", f"{prog['streak']} gün")
    last_exam = xr["exams"][-1] if xr["exams"] else None
    m4.metric("Son deneme doğruluğu", f"%{round(float(last_exam['accuracy']), 1)}" if last_exam else "—")

    t_col, c_col = st.columns([3, 2])
    with t_col:
        st.markdown("#### Kazanım tablosu")
        if xr["outcomes"]:
            st.dataframe(
                [{"Kazanım": o["outcome"], "Branş": o["branch"] or "—", "Ödev": o["n"],
                  "Doğruluk (%)": o["acc"], "Sınıf ort. (%)": o["class_acc"],
                  "Fark": round(o["acc"] - o["class_acc"], 1), "Ort. net": o["net"]} for o in xr["outcomes"]],
                width="stretch", hide_index=True,
            )
            st.caption("En zayıf kazanım en üstte. Fark, aynı ödevlerdeki sınıf ortalamasına göredir.")
        else:
            st.caption("Teslim edilmiş ödev yok; kazanım tablosu ödevlerden oluşur.")
    with c_col:
        st.markdown("#### Doğruluk gidişatı")
        if len(xr["exams"]) >= 2:
            import pandas as pd

            trend = pd.DataFrame({"Doğruluk (%)": [float(e["accuracy"]) for e in xr["exams"]]},
                                 index=range(1, len(xr["exams"]) + 1))
            st.line_chart(trend, height=225, color="#2563EB")
            st.caption(f"Son {len(xr['exams'])} deneme/ödev, soldan sağa eskiden yeniye.")
        else:
            st.caption("Gidişat grafiği için en az iki deneme gerekli.")

    if xr["missing"]:
        with st.expander(f"Teslim edilmemiş ödevler ({len(xr['missing'])})"):
            for r in xr["missing"]:
                due = f" · son teslim {str(r['due_date'])[:10]}" if r.get("due_date") else ""
                late = " · **gecikmiş**" if r in xr["overdue"] else ""
                st.markdown(f"- {r['title']}{due}{late}")

    weak = [o for o in xr["outcomes"] if o["acc"] < 70 or o["acc"] + 10 < o["class_acc"]]
    st.markdown("#### Telafi önerisi")
    if not weak:
        st.success("Belirgin bir zayıf kazanım yok. Mevcut tempoyla devam; yeni kazanımlara geçilebilir.")
        return
    focus = weak[0]
    reasons = []
    if focus["acc"] < 70:
        reasons.append(f"doğruluk %{focus['acc']}")
    if focus["acc"] + 10 < focus["class_acc"]:
        reasons.append(f"sınıf ortalamasının {round(focus['class_acc'] - focus['acc'], 1)} puan altında")
    st.markdown(
        f'<div class="kp-card"><b>{html.escape(focus["outcome"])}</b> '
        f'({html.escape(focus["branch"] or "branş belirtilmemiş")}): {html.escape(", ".join(reasons))}. '
        f'Bu kazanımdan kısa bir tekrar testi ve konu anlatımı önerilir.'
        + (f'<br><span style="color:#64748B;font-size:0.85rem;">Diğer zayıf kazanımlar: '
           f'{html.escape(", ".join(o["outcome"] for o in weak[1:4]))}</span>' if len(weak) > 1 else "")
        + "</div>",
        unsafe_allow_html=True,
    )
    a1, a2 = st.columns(2)
    with a1:
        if ROLE != "veli" and st.button("Bu kazanım için soru üret", key="xray_to_factory", type="primary", width="stretch"):
            st.session_state["factory_focus_topic"] = focus["outcome"]
            if focus["branch"]:
                st.session_state["factory_branch"] = focus["branch"]
            if grade_hint in GRADE_LEVELS_2027:
                st.session_state["factory_grade"] = grade_hint
                st.session_state["factory_grade_sync"] = True
            st.session_state["fab_theme_inp"] = focus["outcome"]
            st.session_state["factory_sync_pending"] = True
            record_activity("Röntgen -> Fabrika", f"{focus['outcome']} Soru Fabrikasına aktarıldı")
            navigate_to("Soru & Redaksiyon Fabrikası")
    with a2:
        if ROLE == "ogrenci" and st.button("Dijimentor planına ekle", key="xray_to_mentor", width="stretch"):
            st.session_state["dk_weak"] = ", ".join(o["outcome"] for o in weak[:3])
            record_activity("Röntgen -> Dijimentor", "Zayıf kazanımlar mentorluğa aktarıldı")
            navigate_to("Dijimentor")


def render_parent_notification_center() -> None:
    st.markdown("---")
    st.markdown("### Veli Bilgilendirme & Otomatik Mesaj Merkezi")
    st.caption(
        "Veli bilgilendirme başlıkları altında otomatik mesaj taslağı oluşturun, "
        "duyuru akışına ekleyin ve kurumsal DOCX olarak indirin."
    )

    stats = st.session_state["parent_student_stats"]
    solved_safe = max(1, int(stats["weekly_solved"]))
    acc_pct = round((int(stats["correct"]) / solved_safe) * 100, 1)

    t_msg, t_feed, t_profile = st.tabs(["Mesaj Üreticisi", "Veli Duyuru Akışı", "Öğrenci & Veli Verileri"])

    with t_msg:
        c_left, c_right = st.columns([1.05, 1.35])
        with c_left:
            sel_cat = st.selectbox("Ana Konu Başlığı", list(PARENT_MESSAGE_TOPICS.keys()), key="pm_cat")
            subtopics = PARENT_MESSAGE_TOPICS[sel_cat]["subtopics"]
            sel_sub = st.selectbox("Alt Konu Başlığı", subtopics, key="pm_sub")
            channel = st.radio("Mesaj Formatı", ["WhatsApp / Kurumsal Mesaj", "Kısa SMS"], horizontal=True, key="pm_ch")

            raw_tpl = PARENT_MESSAGE_TOPICS[sel_cat]["templates"][sel_sub]
            msg_text = raw_tpl.format(
                parent_name=stats["parent_name"],
                student_name=stats["student_name"],
                grade=stats["grade"],
                parent_email=stats["parent_email"],
                weekly_solved=stats["weekly_solved"],
                correct_count=stats["correct"],
                wrong_count=stats["wrong"],
                blank_count=stats["blank"],
                accuracy=acc_pct,
                mentor_completion=stats["mentor_completion"],
                strong_topic=stats["strong_topic"],
                weak_topic=stats["weak_topic"],
                next_exam_date=stats["next_exam_date"],
                next_exam_name=stats["next_exam_name"],
                rem_questions=int(st.session_state.get("remaining_question_quota", 0)),
                rem_mentor=int(st.session_state.get("remaining_coach_quota", 0)),
            )

            if channel == "Kısa SMS":
                msg_text = (
                    f"{BRAND_NAME} 2027: Sn. {stats['parent_name']}, {stats['student_name']} "
                    f"({sel_sub}) haftalık {stats['weekly_solved']} soru, %{acc_pct} doğruluk, "
                    f"Dijimentor %{stats['mentor_completion']} uyum."
                )

            if st.button("Veli Duyuru Akışına Ekle", key="pm_add_feed", width="stretch"):
                st.session_state["parent_announcements"].insert(
                    0,
                    {
                        "category": sel_cat,
                        "subtopic": sel_sub,
                        "body": msg_text,
                        "date": datetime.now().strftime("%d.%m.%Y %H:%M"),
                    },
                )
                record_activity("Veli Mesajı", f"{sel_sub} duyurusu eklendi")
                st.success("Bilgilendirme mesajı Veli Duyuru Akışına eklendi.")

        with c_right:
            edited_msg = st.text_area("Hazır Veli Bilgilendirme Metni", value=msg_text, height=220, key="pm_body")
            clean_tel = re.sub(r"\D", "", stats["parent_phone"])
            encoded = urllib.parse.quote(edited_msg)
            wa_url = f"https://wa.me/{clean_tel}?text={encoded}"
            mail_url = (
                f"mailto:{stats['parent_email']}?subject={urllib.parse.quote(f'{BRAND_NAME} - {sel_sub}')}&body={encoded}"
            )

            l1, l2 = st.columns(2)
            l1.link_button("WhatsApp ile Gönder", wa_url, width="stretch")
            l2.link_button("E-Posta ile Gönder", mail_url, width="stretch")

            msg_docx = build_coach_docx(
                title_meta=f"Veli Bilgilendirme: {sel_cat} > {sel_sub} | Öğrenci: {stats['student_name']}",
                plan_text=edited_msg,
            )
            st.download_button(
                label="Veli Mesajını DOCX İndir",
                data=msg_docx,
                file_name="veli_bilgilendirme_mesaji_2027.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                width="stretch",
            )

    with t_feed:
        flt = st.selectbox("Kategoriye Göre Filtrele", ["Tümü"] + list(PARENT_MESSAGE_TOPICS.keys()), key="pm_flt")
        items = [a for a in st.session_state["parent_announcements"] if flt == "Tümü" or a["category"] == flt]
        for ann in items:
            st.markdown(
                f"""
                <div class="kp-card">
                    <div style="font-size:0.74rem;color:#64748B;font-weight:600;">
                        {html.escape(ann['category'])} • {html.escape(ann['subtopic'])} • {html.escape(ann['date'])}
                    </div>
                    <div style="font-size:0.88rem;color:#0F172A;margin-top:0.35rem;white-space:pre-line;">
                        {html.escape(ann['body'])}
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    with t_profile:
        with st.form("parent_profile_form"):
            p1, p2 = st.columns(2)
            with p1:
                st_n = st.text_input("Öğrenci Adı Soyadı", value=stats["student_name"])
                st_g = st.text_input("Sınıf / Düzey", value=stats["grade"])
                pr_n = st.text_input("Veli Adı Soyadı", value=stats["parent_name"])
                pr_ph = st.text_input("Veli Telefon", value=stats["parent_phone"])
                pr_em = st.text_input("Veli E-Posta", value=stats["parent_email"])
            with p2:
                w_s = st.number_input("Haftalık Çözülen Soru", min_value=0, value=int(stats["weekly_solved"]))
                c_c = st.number_input("Doğru Sayısı", min_value=0, value=int(stats["correct"]))
                w_c = st.number_input("Yanlış Sayısı", min_value=0, value=int(stats["wrong"]))
                m_c = st.slider("Dijimentor Plan Uyumu (%)", 0, 100, int(stats["mentor_completion"]))
                wk_t = st.text_input("Pekiştirilmesi Gereken Kazanım", value=stats["weak_topic"])
            if st.form_submit_button("Veli & Öğrenci Verilerini Güncelle", type="primary"):
                stats.update(
                    {
                        "student_name": st_n,
                        "grade": st_g,
                        "parent_name": pr_n,
                        "parent_phone": pr_ph,
                        "parent_email": pr_em,
                        "weekly_solved": int(w_s),
                        "correct": int(c_c),
                        "wrong": int(w_c),
                        "blank": max(0, int(w_s) - (int(c_c) + int(w_c))),
                        "mentor_completion": int(m_c),
                        "weak_topic": wk_t,
                    }
                )
                save_persisted_state()
                record_activity("Veli Paneli", f"{st_n} verileri güncellendi")
                st.success("Öğrenci ve veli verileri güncellendi.")


# ==============================================================================
# KVKK AYDINLATMA METNİ (geçici taslak; yayından önce hukukçuya kontrol ettirin)
# Köşeli parantezli alanları BURADAN doldurun; metin kayıt formunda, tanıtım sayfasının
# altında ve Profil sayfasında otomatik olarak bu bilgilerle gösterilir.
# ==============================================================================
KVKK_BILGI: Dict[str, str] = {
    "unvan": "[ŞİRKET UNVANI ya da ŞAHIS İŞLETMESİ SAHİBİNİN ADI SOYADI]",
    "adres": "[AÇIK ADRES]",
    "vergi_no": "[MERSİS NO ya da VERGİ NO]",
    "kep": "[VARSA KEP ADRESİ]",
    "eposta": SUPPORT_EMAIL,
    "telefon": SUPPORT_PHONE or "[VARSA TELEFON]",
    "hesap_silme_suresi": "[ör. 6 AY]",
    "guvenlik_kaydi_suresi": "[ör. 2 YIL]",
    "eposta_saglayici": "[E-POSTA SAĞLAYICISI, ör. Google Workspace]",
    "eposta_ulke": "[ÜLKE]",
    # Supabase projesini oluştururken seçtiğin bölge (Project Settings > General'da yazar)
    "supabase_bolge": "Almanya (Frankfurt veri merkezi)",
}
KVKK_GUNCELLEME = "27.11.2026"

KVKK_AYDINLATMA_MD = """
Kurşun Kalem Dijital Ofis olarak, 6698 sayılı Kişisel Verilerin Korunması Kanunu ("KVKK") uyarınca, platformumuzu kullanan öğrenci, veli ve öğretmenlerin kişisel verilerini nasıl işlediğimizi aşağıda açıklıyoruz. Bu metin bilgilendirme amaçlıdır; sizden onay istenmez.

#### 1. Veri sorumlusu

| Bilgi | Değer |
| --- | --- |
| Unvan | {unvan} |
| Adres | {adres} |
| MERSİS / Vergi no | {vergi_no} |
| E-posta | {eposta} |
| KEP adresi | {kep} |
| Telefon | {telefon} |

#### 2. İşlenen kişisel veriler

Yalnızca hizmeti sunmak için gereken verileri işleriz. Kart bilgileriniz bize ulaşmaz; doğrudan ödeme kuruluşu iyzico tarafından alınır.

| Veri kategorisi | Örnekler | Kimden |
| --- | --- | --- |
| Kimlik | Ad soyad | Öğrenci, veli, öğretmen |
| İletişim | E-posta adresi, (varsa) telefon | Tüm kullanıcılar |
| Hesap ve güvenlik | Kullanıcı rolü, şifrelenmiş parola (özet değeri), giriş kayıtları, IP adresi | Tüm kullanıcılar |
| Eğitim ve performans | Sınıf düzeyi, çözülen sorular, deneme sonuçları, hata defteri, çalışma serisi, rozetler, ödev teslimleri | Öğrenci |
| Sınıf bilgileri | Sınıf adı, sınıf kodu, atanan ödevler | Öğretmen |
| Veli bağlantısı | Velinin hangi öğrenciye bağlı olduğu | Veli |
| Yüklediğiniz içerik | "Fotoğrafla Çöz" ile gönderilen soru fotoğrafları, yazılan sorular ve notlar | Öğrenci |
| Müşteri işlem | Seçilen paket, ödeme tutarı, ödeme durumu, fatura için iyzico'ya iletilen ad soyad, e-posta ve T.C. kimlik no | Paket satın alan kullanıcı |

Fotoğraflar sunucularımızda saklanmaz; yalnızca çözüm üretmek için yapay zekâ hizmetine gönderilir. Özel nitelikli kişisel veri (sağlık, din vb.) toplamayız; lütfen fotoğraf ve notlarda bu tür bilgilere yer vermeyin.

#### 3. İşleme amaçları ve hukuki sebepler

Verilerinizi aşağıdaki amaçlarla ve KVKK'nın 5. maddesinin 2. fıkrasındaki hukuki sebeplere dayanarak işleriz. Bu işlemelerin hiçbiri açık rızaya dayanmadığı için sizden ayrıca açık rıza istenmez.

| Amaç | Hukuki sebep (KVKK m.5/2) |
| --- | --- |
| Hesap oluşturma, giriş ve hesap güvenliği | (c) Sözleşmenin kurulması ve ifası; (f) meşru menfaat |
| Soru üretimi, mini deneme, hata defteri, tekrar testleri, Dijimentor planı ve fotoğrafla çözüm hizmetinin sunulması | (c) Sözleşmenin ifası |
| Seri, günlük hedef ve rozetlerin hesaplanması | (c) Sözleşmenin ifası |
| Öğretmenin sınıfa ödev ataması ve sonuçları görmesi | (c) Sözleşmenin ifası |
| Bağlı veliye öğrencinin çalışma özetinin gösterilmesi ve haftalık rapor gönderilmesi | (c) Sözleşmenin ifası |
| Paket satışı, ödeme doğrulama, kota yönetimi, faturalandırma | (c) Sözleşmenin ifası; (ç) hukuki yükümlülük (vergi mevzuatı) |
| Destek taleplerinin yanıtlanması | (c) Sözleşmenin ifası |
| Hizmetin güvenliği, kötüye kullanımın önlenmesi, hataların giderilmesi | (f) Meşru menfaat |
| Resmî mercilerin talepleri ve olası uyuşmazlıklar | (ç) Hukuki yükümlülük; (e) bir hakkın tesisi, kullanılması veya korunması |

Verilerinizi reklam veya pazarlama amacıyla kullanmayız ve üçüncü kişilere satmayız.

#### 4. Toplama yöntemi

Veriler, platformu kullandığınızda elektronik ortamda ve otomatik yollarla toplanır: kayıt ve profil formları, çözdüğünüz denemeler, yüklediğiniz fotoğraflar, ödeme adımları ve destek yazışmaları aracılığıyla. Öğrencinin performans verileri, bağlı velinin ve sınıfın öğretmeninin ekranında gösterilmek üzere platform içinde üretilir.

#### 5. Verilerin aktarılması

Verilerinizi yalnızca hizmeti sunmak için gereken ölçüde, aşağıdaki alıcı gruplarına aktarırız.

**Yurt içi**

| Alıcı | Aktarılan veri | Amaç |
| --- | --- | --- |
| iyzico Ödeme Hizmetleri A.Ş. | Ad soyad, e-posta, T.C. kimlik no, IP adresi, ödeme tutarı | Ödemenin alınması ve doğrulanması |
| Bağlı veli | Öğrencinin çalışma özeti, deneme sonuçları, rozetleri | Veli takibi (öğrencinin paylaştığı bağlantı koduyla) |
| Sınıfın öğretmeni | Öğrencinin adı, ödev sonuçları, haftalık çözülen soru sayısı | Ödev ve sınıf takibi (öğrencinin katıldığı sınıf için) |
| Yetkili kamu kurum ve kuruluşları | Talep edilen veriler | Hukuki yükümlülük |

**Yurt dışı**

Platformun çalışması için yurt dışında bulunan altyapı hizmetleri kullanılır. Bu aktarımlar KVKK'nın 9. maddesi uyarınca, Kişisel Verileri Koruma Kurulu'nun ilan ettiği standart sözleşmeler imzalanarak ve Kurum'a bildirilerek yapılır.

| Alıcı | Ülke | Aktarılan veri | Amaç |
| --- | --- | --- | --- |
| Supabase Inc. | {supabase_bolge} | Hesap, eğitim, sınıf, ödeme kayıtları | Veritabanı barındırma |
| Snowflake Inc. (Streamlit Community Cloud) | ABD | Platform kullanılırken işlenen tüm veriler, IP adresi | Uygulama barındırma |
| Google LLC (Gemini API) | ABD | Soru üretim talepleri, fotoğraflar, çözüm istenen soru metinleri | Yapay zekâ ile soru ve çözüm üretimi |
| {eposta_saglayici} | {eposta_ulke} | Veli adı, e-posta adresi, haftalık rapor içeriği | Haftalık veli raporunun gönderilmesi |

#### 6. Saklama süresi

Verilerinizi işleme amacı sürdükçe saklar, süre dolunca siler, yok eder veya anonim hale getiririz.

| Veri | Saklama süresi |
| --- | --- |
| Hesap, eğitim ve sınıf verileri | Hesap açık kaldığı sürece; hesap kapatıldıktan sonra {hesap_silme_suresi} içinde silinir |
| Ödeme ve fatura kayıtları | İlgili vergi ve ticaret mevzuatında öngörülen süre boyunca |
| "Fotoğrafla Çöz" fotoğrafları | Platform tarafından saklanmaz; yapay zekâ hizmet sağlayıcısının kendi saklama koşulları geçerlidir |
| Güvenlik ve erişim kayıtları | {guvenlik_kaydi_suresi} |

#### 7. 18 yaşından küçük kullanıcılar

Platformu kullanan öğrencilerin çoğu 18 yaşından küçüktür. 18 yaşından küçük kullanıcılar platformu velilerinin bilgisi ve izniyle kullanmalıdır. Velilerin bu metni çocuklarıyla birlikte okumasını öneririz.

Veliler, çocuklarının hesabını bağlantı koduyla kendi hesaplarına bağlayarak çalışma özetini görebilir. Velisi olduğunuz çocuğunuz adına KVKK'dan doğan hakları kullanmak için de aşağıdaki başvuru yolunu izleyebilirsiniz.

#### 8. Haklarınız ve başvuru yolu

KVKK'nın 11. maddesi uyarınca şu haklara sahipsiniz:

1. Kişisel verilerinizin işlenip işlenmediğini öğrenme
2. İşlenmişse buna ilişkin bilgi talep etme
3. İşlenme amacını ve amacına uygun kullanılıp kullanılmadığını öğrenme
4. Yurt içinde veya yurt dışında aktarıldığı üçüncü kişileri bilme
5. Eksik veya yanlış işlenmişse düzeltilmesini isteme
6. KVKK'nın 7. maddesindeki şartlar çerçevesinde silinmesini veya yok edilmesini isteme
7. Düzeltme, silme ve yok etme işlemlerinin, verilerin aktarıldığı üçüncü kişilere bildirilmesini isteme
8. Yalnızca otomatik sistemlerle analiz edilmesi sonucu aleyhinize bir sonuç çıkmasına itiraz etme
9. Kanuna aykırı işleme nedeniyle zarara uğramanız hâlinde zararın giderilmesini talep etme

**Nasıl başvurulur:** Talebinizi yazılı olarak {adres} adresine, KEP adresimize ({kep}) ya da platformda kayıtlı e-posta adresinizden {eposta} adresine iletebilirsiniz. Başvurunuzda ad soyadınızı, iletişim bilgilerinizi ve talebinizi açıkça belirtin. Başvurular en geç 30 gün içinde ücretsiz yanıtlanır; işlem ayrıca bir maliyet gerektirirse Kurul'un belirlediği tarifedeki ücret alınabilir.

Hesabınızı ve verilerinizi silmek için de aynı adrese yazabilirsiniz.
"""


@st.dialog("KVKK Aydınlatma Metni", width="large")
def show_kvkk_dialog() -> None:
    render_kvkk_text()


def render_kvkk_text() -> None:
    """Aydınlatma metnini (doldurulmuş haliyle) ekrana basar."""
    st.caption(f"Son güncelleme: {KVKK_GUNCELLEME}")
    st.markdown(KVKK_AYDINLATMA_MD.format(**KVKK_BILGI))


# ==============================================================================
# GİRİŞ / KAYIT VE TANITIM SAYFASI (giriş yapmamış ziyaretçiler)
# ==============================================================================
def _database_url() -> str:
    """Supabase/Postgres adresi: Secrets'ta [database] url = "..." ya da DATABASE_URL = "..." """
    for path in ("database.url", "DATABASE_URL"):
        val = _secret(path, "")
        if val:
            return val
    return os.environ.get("DATABASE_URL", "") or os.environ.get("KK_DATABASE_URL", "")


db.configure(_database_url())


@st.cache_resource
def _init_database(url_marker: str) -> bool:
    db.init_db()
    return True


try:
    _init_database(db.backend())
except Exception as _db_exc:
    st.error(
        "Veritabanına bağlanılamadı. Streamlit Secrets'taki [database] url değerini kontrol edin "
        "(Supabase'te 'Session pooler' bağlantı adresini ve doğru şifreyi kullanın)."
    )
    st.caption(f"Teknik ayrıntı: {type(_db_exc).__name__}: {str(_db_exc).split('@')[-1][:300]}")
    st.stop()


def logout() -> None:
    for k in list(st.session_state.keys()):
        del st.session_state[k]
    st.rerun()


def _login_form() -> None:
    with st.form("login_form", border=False):
        email = st.text_input("E-posta", key="login_email")
        password = st.text_input("Şifre", type="password", key="login_pw")
        if st.form_submit_button("Giriş yap", type="primary", width="stretch"):
            user = db.authenticate(email, password)
            if user:
                st.session_state["user"] = user
                st.rerun()
            else:
                st.error("E-posta ya da şifre hatalı.")
    _forgot_password()


def _smtp_cfg() -> Dict[str, Any]:
    try:
        return dict(st.secrets.get("smtp", {}))
    except Exception:
        return {}


def _forgot_password() -> None:
    with st.expander("Şifremi unuttum"):
        cfg = _smtp_cfg()
        if not reports.smtp_ready(cfg):
            st.caption(
                "E-postayla şifre sıfırlama şu an kapalı. Hesabınızı öğretmeniniz açtıysa şifrenizi o sıfırlayabilir; "
                + (f"diğer durumlarda {SUPPORT_EMAIL} adresine yazın." if SUPPORT_EMAIL else "diğer durumlarda bize ulaşın.")
            )
            return
        st.caption("Hesabınızı öğretmeniniz açtıysa (e-postanız yoksa) şifrenizi öğretmeniniz sıfırlar.")
        fp_email = st.text_input("Kayıtlı e-postanız", key="fp_email")
        if st.button("Sıfırlama kodu gönder", key="fp_send", width="stretch"):
            try:
                made = db.create_reset_code(fp_email)
                if made:
                    user_fp, code = made
                    reports.send_email(
                        cfg, user_fp["email"], "Kurşun Kalem Dijital Ofis • Şifre sıfırlama kodu",
                        f"Merhaba {user_fp['name']},\n\nŞifre sıfırlama kodunuz: {code}\n\n"
                        f"Kod {db.RESET_TTL_MIN} dakika geçerlidir. Bu isteği siz yapmadıysanız bu e-postayı yok sayın; "
                        "şifreniz değişmez.\n\nKurşun Kalem Dijital Ofis",
                    )
                st.session_state["fp_sent"] = True
                st.success("Bu e-posta kayıtlıysa kod gönderildi. Gelen kutunuzu (ve gereksiz klasörünü) kontrol edin.")
            except ValueError as exc:
                st.warning(str(exc))
            except Exception:
                st.error("E-posta gönderilemedi. Biraz sonra tekrar deneyin.")
        if st.session_state.get("fp_sent"):
            with st.form("fp_form", border=False):
                fp_code = st.text_input("E-postadaki 8 haneli kod", max_chars=8)
                fp_pw = st.text_input("Yeni şifre (en az 8 karakter)", type="password")
                fp_pw2 = st.text_input("Yeni şifre (tekrar)", type="password")
                if st.form_submit_button("Şifreyi değiştir", type="primary", width="stretch"):
                    if fp_pw != fp_pw2:
                        st.error("Şifreler aynı değil.")
                    else:
                        try:
                            db.reset_password_with_code(fp_email, fp_code, fp_pw)
                            st.session_state.pop("fp_sent", None)
                            st.success("Şifreniz değişti. Yeni şifrenizle giriş yapabilirsiniz.")
                        except ValueError as exc:
                            st.error(str(exc))


def _register_form() -> None:
    if st.session_state.get("reg_role_pending"):
        st.session_state["reg_role"] = st.session_state.pop("reg_role_pending")
    role_label = st.radio(
        "Hesap türü", list(db.ROLE_LABELS.values()), horizontal=True, key="reg_role"
    )
    role = next(k for k, v in db.ROLE_LABELS.items() if v == role_label)
    with st.form("register_form", border=False):
        name = st.text_input("Ad soyad")
        email = st.text_input("E-posta")
        password = st.text_input("Şifre (en az 8 karakter)", type="password")
        grade = None
        link_code = ""
        if role == "ogrenci":
            grade = st.selectbox("Sınıfın", GRADE_LEVELS_2027, index=3)
        elif role == "veli":
            link_code = st.text_input("Öğrenci bağlantı kodu (isteğe bağlı)", help="Çocuğunuzun Profil sayfasında yazan 6 haneli kod.")
        with st.expander("KVKK Aydınlatma Metni (okumak için tıklayın)", icon=":material/policy:"):
            render_kvkk_text()
        # KVKK 2026/347 ilke kararı: aydınlatma metni için onay değil, okunduğuna dair beyan alınır
        read_notice = st.checkbox("KVKK Aydınlatma Metni'ni okudum, bilgi edindim.")
        minor_ok = True
        if role == "ogrenci":
            minor_ok = st.checkbox("18 yaşından küçüksem bu platformu velimin bilgisi ve izniyle kullanıyorum.")
        if st.form_submit_button("Hesap oluştur", type="primary", width="stretch"):
            if not read_notice:
                st.error("Devam etmek için Aydınlatma Metni'ni okuduğunuzu işaretleyin.")
                return
            if not minor_ok:
                st.error("18 yaşından küçükseniz velinizin bilgisi ve izni gereklidir; ilgili kutuyu işaretleyin.")
                return
            try:
                user = db.create_user(email, name, password, role, grade)
                if role == "veli" and link_code.strip():
                    try:
                        db.link_parent(user["id"], link_code)
                        user = db.get_user(user["id"])
                    except ValueError as exc:
                        st.warning(f"Hesap açıldı ama öğrenci bağlanamadı: {exc}")
                st.session_state["user"] = user
                st.rerun()
            except ValueError as exc:
                st.error(str(exc))


def landing_stats_html() -> str:
    if LANDING_STATS:
        items = "".join(
            f'<div><div class="kp-stat-num">{html.escape(n)}</div><div class="kp-stat-lbl">{html.escape(l)}</div></div>'
            for n, l in LANDING_STATS[:3]
        )
    else:
        facts = [("5–12", "sınıf düzeyi"), ("3", "rol: öğrenci · veli · öğretmen"), ("1 dk", "ücretsiz kayıt")]
        items = "".join(
            f'<div><div class="kp-stat-num">{n}</div><div class="kp-stat-lbl">{l}</div></div>' for n, l in facts
        )
    return f'<div class="kp-stats">{items}</div>'


# Açılış sayfası: kurumsal görünüm (içerideki eğlenceli tema bundan etkilenmez)
LANDING_CSS = """
.stApp { background: #FFFFFF !important; background-image: none !important; }
[data-testid="stHeader"] { background: rgba(255,255,255,0.92) !important; }
.block-container { max-width: min(1240px, 94vw) !important; }
.lp, .lp * { font-family: 'Inter', system-ui, sans-serif; }
.lp h1, .lp h2, .lp h3, .lp-h { font-family: 'Plus Jakarta Sans', 'Inter', sans-serif !important; color: #0F172A; letter-spacing: -0.02em; }

/* Butonlar: düz, ince çerçeve */
.stButton > button, div[data-testid="stFormSubmitButton"] > button, .stLinkButton > a, .stDownloadButton > button {
    border: 1px solid #CBD5E1 !important; border-radius: 8px !important; box-shadow: none !important;
    transform: none !important; font-weight: 600 !important; min-height: 2.6rem;
}
.stButton > button:hover, div[data-testid="stFormSubmitButton"] > button:hover, .stLinkButton > a:hover {
    background: #F8FAFC !important; color: #0F172A !important; border-color: #94A3B8 !important;
    transform: none !important; box-shadow: none !important;
}
.stButton > button[kind="primary"], div[data-testid="stFormSubmitButton"] > button[kind="primaryFormSubmit"] {
    background: #1D4ED8 !important; border-color: #1D4ED8 !important; color: #FFFFFF !important;
}
.stButton > button[kind="primary"]:hover, div[data-testid="stFormSubmitButton"] > button[kind="primaryFormSubmit"]:hover {
    background: #1E40AF !important; border-color: #1E40AF !important; color: #FFFFFF !important;
}
[class*="st-key-landing_try_"] button, .st-key-landing_kvkk button, .st-key-lp_kvkk_trust button {
    border: 0 !important; background: transparent !important; color: #1D4ED8 !important;
    padding-left: 0 !important; min-height: 0; justify-content: flex-start;
}
[class*="st-key-landing_try_"] button:hover, .st-key-landing_kvkk button:hover, .st-key-lp_kvkk_trust button:hover {
    background: transparent !important; color: #1E40AF !important; text-decoration: underline;
}

/* Kartlar: ince çerçeve, gölge yok */
[class*="st-key-card_"], [class*="st-key-lp_"][data-testid="stVerticalBlockBorderWrapper"],
[class*="st-key-lp_"] {
    border-radius: 12px !important; box-shadow: none !important; transform: none !important;
}
[class*="st-key-card_"] { border: 1px solid #E2E8F0 !important; background: #FFFFFF !important; }
.st-key-card_auth { box-shadow: 0 12px 32px -12px rgba(15, 23, 42, 0.18) !important; }

/* Sekmeler ve giriş alanları */
.stTabs [data-baseweb="tab"] { font-weight: 600; }
.stTextInput input { border-radius: 8px !important; }

.lp-nav { display: flex; align-items: center; justify-content: space-between; padding: 0.4rem 0 1.1rem 0;
          border-bottom: 1px solid #EEF2F7; margin-bottom: 2.2rem; }
.lp-brand { display: flex; align-items: center; gap: 10px; font-weight: 700; color: #0F172A; font-size: 1.02rem; }
.lp-brand img { width: 32px; height: 32px; border-radius: 8px; }
.lp-nav-links { display: flex; gap: 1.6rem; font-size: 0.9rem; }
.lp-nav-links a { color: #475569 !important; text-decoration: none !important; }
.lp-nav-links a:hover { color: #0F172A !important; }

.lp-eyebrow { display: inline-block; font-size: 0.78rem; font-weight: 600; color: #1D4ED8; background: #EFF6FF;
              border: 1px solid #DBEAFE; padding: 0.3rem 0.7rem; border-radius: 999px; letter-spacing: 0.01em; }
.lp-hero-title { font-size: clamp(2.1rem, 3.6vw, 3.1rem); line-height: 1.1; font-weight: 800; margin: 1rem 0 1rem 0; }
.lp-hero-title span { color: #1D4ED8; }
.lp-lead { font-size: 1.1rem; line-height: 1.65; color: #475569; max-width: 40rem; margin: 0 0 1.3rem 0; }
.lp-checks { list-style: none; padding: 0; margin: 0 0 1.5rem 0; display: grid; gap: 0.45rem; }
.lp-checks li { color: #334155; font-size: 0.95rem; display: flex; gap: 0.55rem; align-items: center; }
.lp-checks li::before { content: ""; width: 18px; height: 18px; flex: none; border-radius: 50%; background: #DCFCE7
    url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%2316A34A' stroke-width='3' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpath d='M20 6 9 17l-5-5'/%3E%3C/svg%3E") center/12px no-repeat; }

.lp-section { margin-top: 4.5rem; }
.lp-kicker { font-size: 0.8rem; font-weight: 700; color: #1D4ED8; text-transform: none; letter-spacing: 0.04em; margin-bottom: 0.5rem; }
.lp-h2 { font-size: clamp(1.6rem, 2.4vw, 2.1rem); font-weight: 800; line-height: 1.2; margin: 0 0 0.6rem 0; }
.lp-sub { color: #64748B; font-size: 1rem; line-height: 1.6; max-width: 46rem; margin: 0 0 1.6rem 0; }
.lp-center { text-align: center; }
.lp-center .lp-sub { margin-left: auto; margin-right: auto; }

/* Ürün önizlemesi */
.lp-shot { border: 1px solid #E2E8F0; border-radius: 14px; background: #FFFFFF; overflow: hidden;
           box-shadow: 0 24px 48px -24px rgba(15, 23, 42, 0.25); }
.lp-shot-bar { display: flex; align-items: center; gap: 6px; padding: 0.6rem 0.9rem; background: #F8FAFC; border-bottom: 1px solid #E2E8F0; }
.lp-shot-bar i { width: 9px; height: 9px; border-radius: 50%; background: #CBD5E1; display: inline-block; }
.lp-shot-bar span { margin-left: 0.6rem; font-size: 0.78rem; color: #64748B; }
.lp-shot-body { padding: 1.1rem 1.2rem 1.2rem 1.2rem; }
.lp-shot-kpis { display: grid; grid-template-columns: repeat(3, 1fr); gap: 0.7rem; margin-bottom: 1rem; }
.lp-kpi { border: 1px solid #EEF2F7; border-radius: 10px; padding: 0.65rem 0.75rem; }
.lp-kpi b { display: block; font-size: 1.35rem; color: #0F172A; font-family: 'Plus Jakarta Sans', sans-serif; }
.lp-kpi small { color: #64748B; font-size: 0.75rem; }
.lp-row { display: grid; grid-template-columns: 1fr 120px 44px; gap: 0.8rem; align-items: center; padding: 0.55rem 0;
          border-top: 1px solid #F1F5F9; font-size: 0.86rem; color: #334155; }
.lp-row:first-of-type { border-top: 0; }
.lp-bar { height: 8px; background: #F1F5F9; border-radius: 999px; overflow: hidden; }
.lp-bar > span { display: block; height: 100%; border-radius: 999px; }
.lp-row em { font-style: normal; font-weight: 700; text-align: right; color: #0F172A; }
.lp-pending { margin-top: 0.9rem; display: flex; justify-content: space-between; align-items: center; background: #FFFBEB;
              border: 1px solid #FDE68A; border-radius: 10px; padding: 0.6rem 0.8rem; font-size: 0.85rem; color: #92400E; }
.lp-pending b { color: #78350F; }
.lp-note { font-size: 0.75rem; color: #94A3B8; margin-top: 0.6rem; text-align: right; }

/* Adımlar */
.lp-steps { display: grid; grid-template-columns: repeat(3, 1fr); gap: 1.2rem; }
.lp-step { border: 1px solid #E2E8F0; border-radius: 12px; padding: 1.3rem 1.3rem 1.2rem 1.3rem; background: #FFFFFF; }
.lp-step-num { width: 30px; height: 30px; border-radius: 8px; background: #EFF6FF; color: #1D4ED8; font-weight: 800;
               display: flex; align-items: center; justify-content: center; font-size: 0.9rem; margin-bottom: 0.8rem; }
.lp-step h3 { font-size: 1.05rem; font-weight: 700; margin: 0 0 0.35rem 0; }
.lp-step p { color: #64748B; font-size: 0.92rem; line-height: 1.55; margin: 0; }

/* Özellikler */
.lp-feat-ico { width: 38px; height: 38px; border-radius: 10px; background: #F1F5F9; color: #1D4ED8;
               display: flex; align-items: center; justify-content: center; margin-bottom: 0.75rem; }
.lp-feat-ico svg { width: 20px; height: 20px; }
.lp-feat-title { font-weight: 700; color: #0F172A; font-size: 1rem; margin-bottom: 0.3rem; font-family: 'Plus Jakarta Sans', sans-serif; }
.lp-feat-desc { color: #64748B; font-size: 0.9rem; line-height: 1.55; min-height: 4.2em; }

/* Fiyatlar */
.lp-price-seg { font-size: 0.8rem; font-weight: 700; color: #1D4ED8; }
.lp-price-name { font-size: 1.2rem; font-weight: 800; color: #0F172A; margin: 0.25rem 0 0.4rem 0; font-family: 'Plus Jakarta Sans', sans-serif; }
.lp-price-val { font-size: 1.9rem; font-weight: 800; color: #0F172A; font-family: 'Plus Jakarta Sans', sans-serif; }
.lp-price-val small { font-size: 0.9rem; font-weight: 500; color: #64748B; }
.lp-price-desc { color: #64748B; font-size: 0.9rem; margin: 0.3rem 0 0.9rem 0; min-height: 2.8em; }
.lp-price-list { list-style: none; padding: 0; margin: 0 0 1rem 0; display: grid; gap: 0.45rem; }
.lp-price-list li { font-size: 0.9rem; color: #334155; padding-left: 1.4rem; position: relative; }
.lp-price-list li::before { content: "✓"; position: absolute; left: 0; color: #16A34A; font-weight: 800; }
.st-key-card_price_teacher { border: 2px solid #1D4ED8 !important; }
.lp-ribbon { display: inline-block; background: #1D4ED8; color: #FFFFFF; font-size: 0.72rem; font-weight: 700;
             padding: 0.2rem 0.55rem; border-radius: 999px; margin-left: 0.4rem; vertical-align: middle; }

/* Güven */
.lp-trust { display: grid; grid-template-columns: repeat(4, 1fr); gap: 1rem; }
.lp-trust > div { border-top: 2px solid #E2E8F0; padding-top: 0.9rem; }
.lp-trust b { display: block; color: #0F172A; font-size: 0.95rem; margin-bottom: 0.3rem; }
.lp-trust span { color: #64748B; font-size: 0.88rem; line-height: 1.5; }

.lp-cta { background: #0F172A; border-radius: 16px; padding: 2.4rem 2.2rem; color: #E2E8F0; }
.lp-cta h2 { color: #FFFFFF !important; font-size: 1.7rem; font-weight: 800; margin: 0 0 0.4rem 0; }
.lp-cta p { color: #94A3B8; margin: 0; }
.lp-footer { border-top: 1px solid #EEF2F7; margin-top: 3.5rem; padding-top: 1.3rem; color: #94A3B8; font-size: 0.85rem;
             display: flex; justify-content: space-between; flex-wrap: wrap; gap: 0.6rem; }
.lp-footer a { color: #64748B !important; }
[data-testid="stExpander"] details { border-radius: 10px !important; border-color: #E2E8F0 !important; }

@media (max-width: 900px) {
    .lp-steps, .lp-trust { grid-template-columns: 1fr; }
    .lp-nav-links { display: none; }
    .lp-shot-kpis { grid-template-columns: 1fr 1fr 1fr; }
}
"""

LP_ICONS = {
    "bank": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M9 11l3 3L22 4"/><path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/></svg>',
    "class": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>',
    "chart": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 3v18h18"/><path d="M7 15l4-4 3 3 5-6"/></svg>',
    "mail": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="4" width="20" height="16" rx="2"/><path d="m22 7-10 6L2 7"/></svg>',
    "print": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 9V2h12v7"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><rect x="6" y="14" width="12" height="8"/></svg>',
    "student": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 10 12 5 2 10l10 5 10-5z"/><path d="M6 12v5c3 2 9 2 12 0v-5"/></svg>',
}

LANDING_FEATURES: List[Tuple[str, str, str, str]] = [
    ("bank", "Soru bankası ve öğretmen onayı",
     "Yapay zekâ kazanıma uygun soruları taslak olarak hazırlar. Düzeltir, onaylarsınız; bankanız büyür.",
     "Soru Bankam|Soru & Redaksiyon Fabrikası"),
    ("class", "Sınıf ve ödev yönetimi",
     "Öğrenciler sınıf koduyla katılır ya da Excel listesiyle toplu eklenir. Ödeve konu anlatım videosu eklenebilir.",
     "Sınıflarım|Öğrenci Portalı & Çözüm"),
    ("chart", "Kazanım bazlı rapor",
     "Her öğrencinin hangi kazanımda ne kadar başarılı olduğunu sınıf ortalamasıyla birlikte görürsünüz.",
     "Öğrenci Röntgeni"),
    ("mail", "Veli raporu",
     "Haftalık özet velilere sizin ve kurumunuzun adıyla e-posta ya da WhatsApp ile gider.",
     "Sınıflarım|Veli Takip Paneli"),
    ("print", "Basılı kitapçık ve optik form",
     "Onaylı sorulardan tek tıkla Word kitapçık ve optik form çıkarın; sınıfta kâğıtla uygulayın.",
     "Soru & Redaksiyon Fabrikası"),
    ("student", "Öğrenci uygulaması",
     "Fotoğrafla soru çözümü, mini deneme, hata defteri, tekrar testleri, seri ve rozetler.",
     "Fotoğrafla Çöz|Öğrenci Portalı & Çözüm"),
]

LANDING_FAQ: List[Tuple[str, str]] = [
    ("Ücretli mi?",
     "Pilot döneminde öğretmenler için ücretsiz deneme sunuyoruz; kart bilgisi istenmez. Kurumlar için "
     "öğrenci sayısına göre teklif hazırlıyoruz."),
    ("Yapay zekânın ürettiği sorular doğru mu?",
     "Yapay zekâ hata yapabilir. Bu yüzden her soru önce taslak olarak gelir; siz okuyup düzeltmeden ve "
     "onaylamadan öğrencilere gitmez."),
    ("Öğrenciler sisteme nasıl katılır?",
     "Sınıf kodunu paylaşırsınız ya da öğrenci listesini yapıştırırsınız. E-postası olmayan öğrencilere giriş "
     "bilgisi ve geçici şifre üretilir, siz dağıtırsınız."),
    ("Veriler nerede saklanır?",
     "Bulut veritabanında, şifreli bağlantıyla saklanır; şifreler tek yönlü olarak şifrelenir. Hangi verinin "
     "neden işlendiği KVKK Aydınlatma Metni'nde yazar."),
    ("Kurum olarak nasıl kullanırız?",
     "Öğretmenleriniz tek tek ücretsiz başlayabilir. Kurum genelinde kullanım, ortak soru hakkı ve kurum logolu "
     "çıktılar için bize yazın; birlikte bir pilot planlayalım."),
]


def _demo_mailto() -> str:
    if not SUPPORT_EMAIL:
        return ""
    subject = urllib.parse.quote("Kurum demosu talebi")
    body = urllib.parse.quote(
        "Merhaba,\n\nKurum adı:\nŞehir:\nÖğrenci sayısı:\nÖğretmen sayısı:\nTelefon:\n\n"
        "Kurşun Kalem Dijital Ofis için demo görüşmesi talep ediyoruz."
    )
    return f"mailto:{SUPPORT_EMAIL}?subject={subject}&body={body}"


def _start_teacher_trial() -> None:
    st.session_state["reg_role_pending"] = db.ROLE_LABELS["ogretmen"]
    st.session_state["post_login_page"] = ["Soru Bankam"]
    st.session_state["post_login_title"] = "Öğretmen paneli"
    st.rerun()


def render_landing() -> None:
    st.markdown(
        '<style>[data-testid="stSidebar"], [data-testid="stSidebarCollapsedControl"] {display:none;}'
        + LANDING_CSS + "</style>",
        unsafe_allow_html=True,
    )
    demo_link = _demo_mailto()
    st.markdown(
        f"""
        <div class="lp lp-nav">
            <div class="lp-brand"><img src="{LOGO_DATA_URI}" alt=""/>Kurşun Kalem Dijital Ofis</div>
            <div class="lp-nav-links">
                <a href="#nasil-calisir">Nasıl çalışır</a>
                <a href="#ozellikler">Özellikler</a>
                <a href="#fiyatlar">Fiyatlar</a>
                <a href="#sss">Sık sorulanlar</a>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ---------- Üst bölüm ----------
    hero_l, hero_r = st.columns([1.3, 1.0], gap="large", vertical_alignment="center")
    with hero_l:
        st.markdown(
            """
            <div class="lp">
                <span class="lp-eyebrow">Öğretmenler ve etüt merkezleri için</span>
                <h1 class="lp-hero-title">Soru hazırlayın, ödev verin, <span>veliye rapor gönderin.</span></h1>
                <p class="lp-lead">
                    Yapay zekâ kazanıma uygun soruları taslak olarak hazırlar; siz kontrol edip onaylarsınız.
                    Öğrencileriniz onaylı soruları çözer, sonuçlar kazanım bazında raporlanır, veliler haftalık
                    özeti sizin imzanızla alır.
                </p>
                <ul class="lp-checks">
                    <li>Kurulum yok, tarayıcıdan ve telefondan çalışır</li>
                    <li>5–12. sınıf, LGS ve YKS 2027 hazırlığı</li>
                    <li>Pilot döneminde ücretsiz deneme, kart bilgisi gerekmez</li>
                </ul>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if LANDING_STATS:
            st.markdown(landing_stats_html(), unsafe_allow_html=True)
        b1, b2, _sp = st.columns([1.25, 1.1, 0.6])
        with b1:
            if st.button("Öğretmen olarak ücretsiz dene", key="landing_teacher", type="primary", width="stretch"):
                _start_teacher_trial()
        with b2:
            if demo_link:
                st.link_button("Kurumum için demo iste", demo_link, width="stretch")
    with hero_r:
        intent_title = st.session_state.get("post_login_title")
        with st.container(border=True, key="card_auth"):
            st.markdown(
                '<div class="lp"><h3 class="lp-h" style="font-size:1.25rem;margin:0 0 0.2rem 0;">Hesabınıza girin</h3>'
                '<p style="color:#64748B;font-size:0.9rem;margin:0 0 0.4rem 0;">Öğretmen, öğrenci ve veli girişi. '
                'Hesabınız yoksa ücretsiz deneme için kayıt olun.</p></div>',
                unsafe_allow_html=True,
            )
            if intent_title:
                st.markdown(
                    f'<div class="kp-banner kp-banner-review"><b>{html.escape(intent_title)}</b> için önce ücretsiz '
                    f'hesabını aç ya da giriş yap; ardından doğrudan oraya götüreceğiz.</div>',
                    unsafe_allow_html=True,
                )
                embed_html(
                    "<script>try{const d=window.parent.document;"
                    "const m=d.querySelector('[data-testid=stMain]')||d.querySelector('section.main');"
                    "if(m){m.scrollTo({top:0,behavior:'smooth'});}window.parent.scrollTo({top:0,behavior:'smooth'});}catch(e){}</script>",
                    height=1,
                )
                t_reg, t_login = st.tabs(["Ücretsiz deneme", "Giriş yap"])
            else:
                t_login, t_reg = st.tabs(["Giriş yap", "Ücretsiz deneme"])
            with t_login:
                _login_form()
            with t_reg:
                _register_form()

    # ---------- Ürün önizlemesi ----------
    st.markdown('<div class="lp-section"></div>', unsafe_allow_html=True)
    pv_l, pv_r = st.columns([1.0, 1.25], gap="large", vertical_alignment="center")
    with pv_l:
        st.markdown(
            """
            <div class="lp">
                <div class="lp-kicker">SINIF RAPORU</div>
                <h2 class="lp-h2">Hangi öğrenci hangi kazanımda zorlanıyor, bir bakışta görün.</h2>
                <p class="lp-sub">Teslim edilen her ödev, kazanım bazında sınıf raporuna işlenir. En zayıf kazanım
                en üstte durur; tek tıkla o kazanım için yeni soru üretip ödev olarak atarsınız.</p>
                <ul class="lp-checks">
                    <li>Öğrenci bazında doğruluk ve sınıf ortalamasıyla fark</li>
                    <li>Teslim edilmeyen ve geciken ödevler</li>
                    <li>Telafi önerisi ve tek tıkla yeni test</li>
                </ul>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with pv_r:
        rows = [("Üslü ifadeler", 48, "#DC2626"), ("Kareköklü ifadeler", 61, "#EA580C"),
                ("Veri analizi", 74, "#CA8A04"), ("Çarpanlar ve katlar", 86, "#16A34A")]
        rows_html = "".join(
            f'<div class="lp-row"><span>{name}</span><div class="lp-bar"><span style="width:{pct}%;background:{col}"></span></div>'
            f"<em>%{pct}</em></div>"
            for name, pct, col in rows
        )
        st.markdown(
            f"""
            <div class="lp lp-shot">
                <div class="lp-shot-bar"><i></i><i></i><i></i><span>Sınıflarım · 8-A Matematik · Sınıf analizi</span></div>
                <div class="lp-shot-body">
                    <div class="lp-shot-kpis">
                        <div class="lp-kpi"><b>28</b><small>öğrenci</small></div>
                        <div class="lp-kpi"><b>%67</b><small>ortalama doğruluk</small></div>
                        <div class="lp-kpi"><b>12</b><small>teslim edilen ödev</small></div>
                    </div>
                    {rows_html}
                    <div class="lp-pending"><span><b>6 soru</b> onayınızı bekliyor</span><span>Soru Bankam →</span></div>
                    <div class="lp-note">Örnek ekran; rakamlar temsilidir.</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    # ---------- Nasıl çalışır ----------
    st.markdown(
        """
        <div class="lp lp-section lp-center" id="nasil-calisir">
            <div class="lp-kicker">NASIL ÇALIŞIR</div>
            <h2 class="lp-h2">Üç adımda ilk ödeviniz hazır</h2>
            <p class="lp-sub">Hazırlık saatler değil dakikalar sürer; kontrol her zaman sizde kalır.</p>
        </div>
        <div class="lp lp-steps">
            <div class="lp-step"><div class="lp-step-num">1</div><h3>Soruları üretin</h3>
                <p>Sınıf düzeyi, branş ve kazanımı seçin. Yapay zekâ çeldiricileriyle birlikte taslak soruları hazırlar.</p></div>
            <div class="lp-step"><div class="lp-step-num">2</div><h3>Kontrol edin, onaylayın</h3>
                <p>Her soruyu okuyun, gerekirse düzeltin, doğru cevabı işaretleyin. Onaylı sorular bankanızda birikir.</p></div>
            <div class="lp-step"><div class="lp-step-num">3</div><h3>Ödev verin, raporu alın</h3>
                <p>Seçtiğiniz soruları sınıfa atayın ya da yazdırın. Sonuçlar kazanım bazında rapora, özeti velilere gider.</p></div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ---------- Özellikler ----------
    st.markdown(
        """
        <div class="lp lp-section lp-center" id="ozellikler">
            <div class="lp-kicker">ÖZELLİKLER</div>
            <h2 class="lp-h2">Öğretmenin ihtiyacı olan her şey, tek yerde</h2>
            <p class="lp-sub">Öğretmen, öğrenci ve veli aynı sistemi kullanır; herkes yalnızca kendi ekranını görür.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    for row_start in range(0, len(LANDING_FEATURES), 3):
        cols = st.columns(3, gap="medium")
        for j, (ico, f_title, f_desc, f_page) in enumerate(LANDING_FEATURES[row_start:row_start + 3]):
            idx = row_start + j
            with cols[j]:
                with st.container(border=True, key=f"card_feat_{idx}"):
                    st.markdown(
                        f'<div class="lp"><div class="lp-feat-ico">{LP_ICONS[ico]}</div>'
                        f'<div class="lp-feat-title">{html.escape(f_title)}</div>'
                        f'<div class="lp-feat-desc">{html.escape(f_desc)}</div></div>',
                        unsafe_allow_html=True,
                    )
                    if st.button("Ücretsiz dene →", key=f"landing_try_{idx}"):
                        st.session_state["post_login_page"] = f_page.split("|")
                        st.session_state["post_login_title"] = f_title
                        st.rerun()

    # ---------- Fiyatlar ----------
    st.markdown(
        """
        <div class="lp lp-section lp-center" id="fiyatlar">
            <div class="lp-kicker">FİYATLAR</div>
            <h2 class="lp-h2">Önce ücretsiz deneyin, sonra karar verin</h2>
            <p class="lp-sub">Pilot döneminde öğretmen hesabı ücretsizdir. Kurumlar için öğrenci sayısına göre teklif hazırlıyoruz.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    student_tiers = [t for t in SUBSCRIPTION_TIERS if t.get("segment") in ("Öğrenci", "YKS", "Bireysel", "Veli", "Aile")] or SUBSCRIPTION_TIERS
    min_student = min(int(t["monthly_price"]) for t in student_tiers)
    pc1, pc2, pc3 = st.columns(3, gap="medium")
    with pc1:
        with st.container(border=True, key="card_price_teacher"):
            st.markdown(
                """
                <div class="lp">
                    <div class="lp-price-seg">Öğretmen <span class="lp-ribbon">Önerilen</span></div>
                    <div class="lp-price-name">Ücretsiz Deneme</div>
                    <div class="lp-price-val">0 ₺ <small>pilot dönemi boyunca</small></div>
                    <div class="lp-price-desc">Tek öğretmen ve sınıfları için tüm öğretmen araçları.</div>
                    <ul class="lp-price-list">
                        <li>Soru üretimi ve soru bankası</li>
                        <li>Sınırsız sınıf ve öğrenci</li>
                        <li>Kazanım raporu ve öğrenci röntgeni</li>
                        <li>Veli raporu, kitapçık ve optik form</li>
                    </ul>
                </div>
                """,
                unsafe_allow_html=True,
            )
            if st.button("Ücretsiz denemeyi başlat", key="lp_price_teacher", type="primary", width="stretch"):
                _start_teacher_trial()
    with pc2:
        with st.container(border=True, key="card_price_org"):
            st.markdown(
                """
                <div class="lp">
                    <div class="lp-price-seg">Etüt merkezi ve okul</div>
                    <div class="lp-price-name">Kurum</div>
                    <div class="lp-price-val">Teklif <small>öğrenci sayısına göre</small></div>
                    <div class="lp-price-desc">Birden çok öğretmen ve şube için ortak kullanım.</div>
                    <ul class="lp-price-list">
                        <li>Öğretmen paketindeki her şey</li>
                        <li>Kurum adıyla veli raporları</li>
                        <li>Toplu öğrenci hesabı açma</li>
                        <li>Kuruluma ve eğitime destek</li>
                    </ul>
                </div>
                """,
                unsafe_allow_html=True,
            )
            if demo_link:
                st.link_button("Teklif ve demo iste", demo_link, width="stretch")
    with pc3:
        with st.container(border=True, key="card_price_student"):
            st.markdown(
                f"""
                <div class="lp">
                    <div class="lp-price-seg">Öğrenci ve veli</div>
                    <div class="lp-price-name">Bireysel</div>
                    <div class="lp-price-val">{min_student:,} ₺ <small>/ ay'dan başlayan</small></div>
                    <div class="lp-price-desc">Öğretmeni sistemi kullanmayan öğrenciler için.</div>
                    <ul class="lp-price-list">
                        <li>İlk {FREE_PLAN["monthly_quota"]} soru ücretsiz</li>
                        <li>Fotoğrafla soru çözümü</li>
                        <li>Mini deneme ve hata defteri</li>
                        <li>Haftalık çalışma planı</li>
                    </ul>
                </div>
                """.replace(f"{min_student:,}", f"{min_student:,}".replace(",", ".")),
                unsafe_allow_html=True,
            )
            if st.button("Paketleri incele", key="lp_price_student", width="stretch"):
                st.session_state["post_login_page"] = ["Paketler & Abonelik Vitrini"]
                st.session_state["post_login_title"] = "Paketler"
                st.rerun()

    # ---------- Güven ----------
    st.markdown(
        """
        <div class="lp lp-section">
            <div class="lp-trust">
                <div><b>Öğretmen onayı</b><span>Yapay zekânın ürettiği hiçbir soru, siz onaylamadan öğrenciye gitmez.</span></div>
                <div><b>KVKK'ya uygun işleyiş</b><span>Hangi verinin neden işlendiği aydınlatma metninde açıkça yazar.</span></div>
                <div><b>Güvenli saklama</b><span>Şifreli bağlantı; şifreler tek yönlü şifrelenerek saklanır.</span></div>
                <div><b>Herkes kendi ekranında</b><span>Öğretmen yalnızca kendi sınıfını, veli yalnızca kendi çocuğunu görür.</span></div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    _k1, _k2 = st.columns([1, 5])
    with _k1:
        if st.button("KVKK Aydınlatma Metni →", key="lp_kvkk_trust"):
            show_kvkk_dialog()

    # ---------- SSS ----------
    st.markdown(
        """
        <div class="lp lp-section lp-center" id="sss">
            <div class="lp-kicker">SIK SORULANLAR</div>
            <h2 class="lp-h2">Aklınıza takılanlar</h2>
        </div>
        """,
        unsafe_allow_html=True,
    )
    _fq_l, fq_mid, _fq_r = st.columns([1, 4, 1])
    with fq_mid:
        for q_text, a_text in LANDING_FAQ:
            with st.expander(q_text):
                st.markdown(a_text)

    # ---------- Son çağrı ----------
    st.markdown('<div class="lp-section"></div>', unsafe_allow_html=True)
    with st.container(key="lp_cta_wrap"):
        cta_l, cta_r = st.columns([2.2, 1], vertical_alignment="center")
        with cta_l:
            st.markdown(
                '<div class="lp lp-cta"><h2>Bu hafta ilk ödevinizi verin.</h2>'
                '<p>Kayıt bir dakika sürer. Kart bilgisi gerekmez, istediğiniz zaman bırakabilirsiniz.</p></div>',
                unsafe_allow_html=True,
            )
        with cta_r:
            if st.button("Öğretmen olarak ücretsiz dene", key="lp_cta_teacher", type="primary", width="stretch"):
                _start_teacher_trial()
            if demo_link:
                st.link_button("Kurumum için demo iste", demo_link, width="stretch")

    contact = " · ".join(
        x for x in [
            f'<a href="mailto:{SUPPORT_EMAIL}">{SUPPORT_EMAIL}</a>' if SUPPORT_EMAIL else "",
            _support_phone_html(),
        ] if x
    )
    st.markdown(
        f'<div class="lp lp-footer"><span>© {datetime.now().year} Kurşun Kalem Dijital Ofis</span>'
        f"<span>{contact}</span></div>",
        unsafe_allow_html=True,
    )
    _f1, _f2, _f3 = st.columns([2, 1, 2])
    with _f2:
        if st.button("KVKK Aydınlatma Metni", key="landing_kvkk", width="stretch"):
            show_kvkk_dialog()


if not current_user():
    render_landing()
    st.stop()

# Kullanıcı bilgisini her açılışta tazele (veli bağlantısı, hedef vb. değişmiş olabilir)
_fresh_user = db.get_user(current_user_id())
if not _fresh_user:
    logout()
st.session_state["user"] = _fresh_user
USER: Dict[str, Any] = _fresh_user
ROLE: str = USER["role"]
ALLOWED_PAGES: List[str] = ROLE_PAGES.get(ROLE, PAGES)

# ==============================================================================
# SESSION STATE BAŞLATMA & KALICI VERİ YÜKLEME
# ==============================================================================
_persisted = load_persisted_state()

_intent = st.session_state.pop("post_login_page", None)
st.session_state.pop("post_login_title", None)
if _intent:
    _target = next((pg for pg in _intent if pg in ALLOWED_PAGES), None)
    if _target:
        st.session_state.active_page = _target
if "active_page" not in st.session_state or st.session_state.active_page not in ALLOWED_PAGES:
    st.session_state.active_page = "Ana Sayfa"

if "loader_nonce" not in st.session_state:
    trigger_loader_bar()

render_top_loader_bar()

if "generated_test_text" not in st.session_state:
    st.session_state.generated_test_text = ""
if "dijimentor_plan" not in st.session_state:
    st.session_state.dijimentor_plan = ""
if "selected_tier_id" not in st.session_state:
    st.session_state.selected_tier_id = _persisted.get("selected_tier_id", SUBSCRIPTION_TIERS[1]["id"])
if "backpack_cart" not in st.session_state and isinstance(_persisted.get("backpack_cart"), dict):
    st.session_state["backpack_cart"] = dict(_persisted["backpack_cart"])

ensure_bag_initialized()
if "active_plan" not in st.session_state:
    st.session_state["active_plan"] = _persisted.get("active_plan")
for _k in ("reviews_completed", "photo_solves"):
    if _k not in st.session_state:
        st.session_state[_k] = int(_persisted.get(_k, 0) or 0)
bag_summary_global = get_active_plan()

if "used_question_quota" not in st.session_state:
    if _persisted.get("used_question_quota") is not None:
        st.session_state["used_question_quota"] = int(_persisted["used_question_quota"])
    elif _persisted.get("remaining_question_quota") is not None:
        st.session_state["used_question_quota"] = max(
            0, int(bag_summary_global["monthly_quota"]) - int(_persisted["remaining_question_quota"])
        )
    else:
        st.session_state["used_question_quota"] = 0

if "used_coach_quota" not in st.session_state:
    if _persisted.get("used_coach_quota") is not None:
        st.session_state["used_coach_quota"] = int(_persisted["used_coach_quota"])
    elif _persisted.get("remaining_coach_quota") is not None:
        st.session_state["used_coach_quota"] = max(
            0, int(bag_summary_global["coach_quota"]) - int(_persisted["remaining_coach_quota"])
        )
    else:
        st.session_state["used_coach_quota"] = 0

sync_plan_quotas(bag_summary_global, reset_usage=False)

if "live_tick" not in st.session_state:
    st.session_state.live_tick = False
if "portal_question_draft" not in st.session_state:
    st.session_state.portal_question_draft = ""
if "portal_exam_source" not in st.session_state:
    st.session_state.portal_exam_source = ""
if "activity_log" not in st.session_state:
    p_logs = _persisted.get("activity_log")
    st.session_state.activity_log = p_logs if isinstance(p_logs, list) and p_logs else _default_activity_log()
if "parent_student_stats" not in st.session_state:
    p_stats = _persisted.get("parent_student_stats")
    merged_stats = _default_parent_stats()
    if isinstance(p_stats, dict):
        merged_stats.update(p_stats)
    st.session_state.parent_student_stats = merged_stats
if "parent_announcements" not in st.session_state:
    p_ann = _persisted.get("parent_announcements")
    st.session_state.parent_announcements = p_ann if isinstance(p_ann, list) and p_ann else _default_parent_announcements()
if "question_bank_2028" not in st.session_state:
    st.session_state["question_bank_2028"] = _persisted.get("question_bank_2028", []) if isinstance(_persisted.get("question_bank_2028"), list) else []
if "error_notebook_2028" not in st.session_state:
    st.session_state["error_notebook_2028"] = _persisted.get("error_notebook_2028", []) if isinstance(_persisted.get("error_notebook_2028"), list) else []
if "b2b_org_cache" not in st.session_state and isinstance(_persisted.get("b2b_org_cache"), dict):
    st.session_state["b2b_org_cache"] = _persisted["b2b_org_cache"]
if "last_b2b_summary" not in st.session_state and _persisted.get("last_b2b_summary"):
    st.session_state["last_b2b_summary"] = _persisted["last_b2b_summary"]


def refresh_live_stats() -> Optional[int]:
    """Öğrenci kartlarındaki haftalık rakamları gerçek kayıtlardan doldurur.
    Öğrenci kendi verisini, bağlı veli çocuğunun verisini görür. Hedef öğrencinin id'sini döner."""
    target = USER["id"] if ROLE == "ogrenci" else USER.get("parent_of")
    if not target:
        return None
    stu = db.get_user(int(target)) or {}
    ws = db.weekly_stats(int(target))
    stats = st.session_state.parent_student_stats
    stats.update(
        {
            "student_name": stu.get("name", stats["student_name"]),
            "grade": stu.get("grade") or stats["grade"],
            "weekly_solved": ws["weekly_solved"],
            "correct": ws["correct"],
            "wrong": ws["wrong"],
            "blank": ws["blank"],
        }
    )
    if ROLE == "veli":
        stats["parent_name"] = USER["name"]
        stats["parent_email"] = USER["email"]
        child_state = db.load_state(int(target)).get("parent_student_stats") or {}
        for key in ("strong_topic", "weak_topic", "mentor_completion", "next_exam_date", "next_exam_name"):
            if key in child_state:
                stats[key] = child_state[key]
    return int(target)


LIVE_STUDENT_ID = refresh_live_stats()

api_key_ready = resolve_gemini_api_key() is not None

# ==============================================================================
# KENAR ÇUBUĞU
# ==============================================================================
def _fmt(n: int) -> str:
    """1500 -> '1.500'"""
    return f"{int(n):,}".replace(",", ".")


with st.sidebar:
    st.markdown(
        f"""
        <div class="kp-sidebar-brand">
            <img src="{LOGO_DATA_URI}" alt=""/>
            <div>
                <div class="kp-sidebar-brand-title">{BRAND_SHORT}</div>
                <div class="kp-sidebar-brand-sub">Dijital Ofis</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Kullanıcı kartı + seri
    _prog_side = user_progress(USER["id"]) if ROLE == "ogrenci" else None
    _streak_html = (
        f'<span class="kp-streak-pill" title="Üst üste çalışılan gün">{_prog_side["streak"]} gün seri</span>'
        if _prog_side else ""
    )
    st.markdown(
        f"""
        <div class="kp-user-card">
            <div class="kp-avatar">{html.escape(USER["name"][:1].upper())}</div>
            <div style="min-width:0;flex:1;">
                <div class="kp-user-name">{html.escape(USER["name"])}</div>
                <div class="kp-user-role">{db.ROLE_LABELS.get(ROLE, ROLE)}{(" · " + html.escape(USER["grade"])) if USER.get("grade") else ""}</div>
                {_streak_html}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Gruplanmış, ikonlu menü (sadece bu rolün görebildiği sayfalar)
    for group_title, all_items in NAV_GROUPS:
        items = [it for it in all_items if it[0] in ALLOWED_PAGES]
        if not items:
            continue
        if group_title:
            st.markdown(f'<div class="kp-nav-title">{html.escape(tr_upper(group_title))}</div>', unsafe_allow_html=True)
        for p_name, p_label, p_icon in items:
            is_current = st.session_state.active_page == p_name
            if st.button(
                p_label,
                key=f"nav_{p_name}",
                icon=p_icon,
                type="primary" if is_current else "secondary",
                width="stretch",
            ):
                clear_page_ephemeral_state(p_name)
                st.session_state.active_page = p_name
                trigger_loader_bar()
                st.rerun()

    # Paket & kota kartı (ilerleme çubuklu)
    total_q = int(bag_summary_global["monthly_quota"])
    total_c = int(bag_summary_global["coach_quota"])
    rem_q = int(st.session_state.remaining_question_quota)
    rem_c = int(st.session_state.remaining_coach_quota)
    q_pct = round(100 * rem_q / total_q) if total_q else 0
    c_pct = round(100 * rem_c / total_c) if total_c else 0
    status_color = "#16A34A" if api_key_ready else "#D97706"
    status_text = "Sistem aktif" if api_key_ready else "API anahtarı bekleniyor"

    st.markdown(
        f"""
        <div class="kp-sidebar-meta">
            <div class="kp-meta-head"><span>Aktif paketin</span></div>
            <div class="kp-meta-sub">{bag_summary_global['total_count']} paket · {_fmt(bag_summary_global['monthly_price'])} ₺/ay</div>
            <div class="kp-meter-label"><span>Soru kotası</span><span>{_fmt(rem_q)} / {_fmt(total_q)}</span></div>
            <div class="kp-meter"><span style="width:{q_pct}%"></span></div>
            <div class="kp-meter-label"><span>Dijimentor</span><span>{rem_c} / {total_c}</span></div>
            <div class="kp-meter"><span style="width:{c_pct}%"></span></div>
            <div class="kp-meter-label" style="margin-top:0.6rem;">
                <span>Soru havuzu: <b>{len(st.session_state['question_bank_2028'])}</b></span>
                <span>Hata defteri: <b>{len(st.session_state['error_notebook_2028'])}</b></span>
            </div>
            <div class="kp-status"><span class="kp-status-dot" style="background:{status_color}"></span>{status_text}</div>
            <div class="kp-status" style="border-top:0;padding-top:0;margin-top:0.3rem;">
                <span class="kp-status-dot" style="background:{'#16A34A' if db.backend() == 'postgres' else '#D97706'}"></span>
                {'Veriler kalıcı (Supabase)' if db.backend() == 'postgres' else 'Veriler geçici: veritabanı bağlı değil'}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if st.button("Çıkış yap", key="logout_btn", icon=":material/logout:", width="stretch"):
        logout()

    # Tehlikeli işlem: göz önünden kaldırıldı, "Ayarlar" altında
    with st.expander("Ayarlar", icon=":material/settings:"):
        if not st.session_state.get("confirm_reset_session", False):
            if st.button("Verilerimi sıfırla", key="reset_session_btn", width="stretch"):
                st.session_state["confirm_reset_session"] = True
                st.rerun()
        else:
            st.warning("Soru havuzun, hata defterin ve çantan silinsin mi? Paketin ve rozetlerin korunur.")
            rc1, rc2 = st.columns(2)
            with rc1:
                if st.button("Evet", key="confirm_reset_yes", type="primary", width="stretch"):
                    keep_plan = st.session_state.get("active_plan")
                    keep_user = st.session_state.get("user")
                    db.save_state(USER["id"], {
                        "active_plan": keep_plan,
                        "used_question_quota": int(st.session_state.get("used_question_quota", 0)),
                        "used_coach_quota": int(st.session_state.get("used_coach_quota", 0)),
                    })
                    for k in list(st.session_state.keys()):
                        del st.session_state[k]
                    st.session_state["user"] = keep_user
                    st.rerun()
            with rc2:
                if st.button("İptal", key="confirm_reset_no", width="stretch"):
                    st.session_state["confirm_reset_session"] = False
                    st.rerun()

# ==============================================================================
# ÜST ŞERİT
# ==============================================================================
st.markdown(
    f"""
    <div class="kp-topbar">
        <span class="kp-topbar-brand">Kurşun Kalem Dijital Ofis</span>
        <span>{_support_line_html()}</span>
    </div>
    """,
    unsafe_allow_html=True,
)

current_page = st.session_state.active_page
if current_page != "Ana Sayfa":
    render_page_strip(current_page)

TR_DAYS = ["Pzt", "Sal", "Çar", "Per", "Cum", "Cmt", "Paz"]


def start_review_exam(items: List[Dict[str, Any]]) -> None:
    """Hata defterindeki vadesi gelen sorularla bir tekrar denemesi başlatır."""
    text, mapping = gamify.build_review_exam(items)
    st.session_state["portal_exam_source"] = text
    st.session_state["review_map"] = mapping
    st.session_state["active_assignment"] = None
    st.session_state.pop("last_exam_report", None)
    record_activity("Aralıklı Tekrar", f"{len(items)} soruluk tekrar testi başlatıldı")
    navigate_to("Öğrenci Portalı & Çözüm")


def render_lesson_video(url: str) -> None:
    """YouTube/Vimeo/mp4 bağlantısını oynatır; desteklenmeyen adreste bağlantı butonu gösterir."""
    if not url:
        return
    low = url.lower()
    if any(h in low for h in ("youtube.com/", "youtu.be/", "vimeo.com/")) or low.split("?")[0].endswith((".mp4", ".webm")):
        try:
            st.video(url)
            return
        except Exception:
            pass
    st.link_button("Konu anlatım videosunu aç", url, width="stretch")


def start_assignment(assignment: Dict[str, Any]) -> None:
    st.session_state["portal_exam_source"] = assignment["content"]
    st.session_state["active_assignment"] = {"id": assignment["id"], "title": assignment["title"],
                                             "video_url": assignment.get("video_url") or ""}
    st.session_state["review_map"] = {}
    st.session_state.pop("last_exam_report", None)
    navigate_to("Öğrenci Portalı & Çözüm")


def render_week_chart(user_id: int, height: int = 200) -> None:
    import pandas as pd

    by_day = db.weekly_stats(user_id)["by_day"]
    today = date.today()
    days = [today - timedelta(days=i) for i in range(6, -1, -1)]
    counts = [int(by_day.get(d.isoformat(), 0)) for d in days]
    if sum(counts) == 0:
        st.markdown(
            f'<div class="kp-empty-chart" style="height:{height}px;">'
            '<div><b>Bu hafta henüz soru çözülmedi.</b><br>'
            'İlk mini denemeyi çözünce çubuklar burada belirecek.</div></div>',
            unsafe_allow_html=True,
        )
        return
    df = pd.DataFrame(
        {
            "Gün": ["Bugün" if d == today else TR_DAYS[d.weekday()] for d in days],
            "Çözülen soru": counts,
        }
    )
    df["Gün"] = pd.Categorical(df["Gün"], categories=list(df["Gün"]), ordered=True)
    st.bar_chart(df, x="Gün", y="Çözülen soru", color="#1D4ED8", height=height)


def render_week_card(stats_now: Dict[str, Any], show_sticker: bool = True) -> None:
    solved_now = max(1, int(stats_now["weekly_solved"]))
    graded_now = int(stats_now["correct"]) + int(stats_now["wrong"])
    acc_now = round(int(stats_now["correct"]) / graded_now * 100) if graded_now else 0
    lgs_days_left = format_countdown(LGS_TARGET)[0]
    sticker_html = (
        f'<span class="kp-sticker">LGS\'ye {lgs_days_left} gün!</span>'
        if FUN_THEME and show_sticker and lgs_days_left > 0 else ""
    )
    st.markdown(
        f"""
        <div class="kp-preview">
            {sticker_html}
            <div class="kp-preview-head">
                <div><small>Son 7 gün</small><strong>{html.escape(str(stats_now['student_name']))}</strong></div>
                <span class="kp-pill">{html.escape(str(stats_now['grade']))}</span>
            </div>
            <div class="kp-preview-stats">
                <div class="kp-preview-stat"><b>{_fmt(int(stats_now['weekly_solved']))}</b><span>çözülen soru</span></div>
                <div class="kp-preview-stat"><b>%{acc_now}</b><span>doğruluk</span></div>
                <div class="kp-preview-stat"><b>%{int(stats_now['mentor_completion'])}</b><span>plan uyumu</span></div>
            </div>
            <div class="kp-meter" style="margin-bottom:0.9rem;"><span style="width:{int(stats_now['mentor_completion'])}%"></span></div>
            <div class="kp-preview-row"><span>Güçlü alan</span><span>{html.escape(str(stats_now['strong_topic']))}</span></div>
            <div class="kp-preview-row"><span>Pekiştirilecek</span><span>{html.escape(str(stats_now['weak_topic']))}</span></div>
            <div class="kp-preview-row"><span>Sıradaki deneme</span><span>{html.escape(str(stats_now['next_exam_date']))}</span></div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_countdown_section() -> None:
    head_col, toggle_col = st.columns([4, 1], vertical_alignment="center")
    with head_col:
        st.markdown('<h2 class="kp-section-title">2027 LGS\'ye kalan süre</h2>', unsafe_allow_html=True)
    with toggle_col:
        live_counter = st.toggle("Canlı sayaç", value=st.session_state.live_tick)
        st.session_state.live_tick = live_counter
    render_countdown_boxes(LGS_TARGET)
    yks_d, yks_h, _, _ = format_countdown(YKS_TARGET)
    st.caption(f"LGS: 13 Haziran 2027 Pazar · YKS (TYT): 19 Haziran 2027 Cumartesi — YKS'ye kalan: {yks_d} gün {yks_h} saat. Tarihler MEB ve ÖSYM 2027 takvimlerinden; sayaç saati tahminidir.")


def render_modules_section() -> None:
    cards = [m for m in MODULE_CARDS if m[3] in ALLOWED_PAGES]
    if not cards:
        return
    st.markdown(
        '<h2 class="kp-section-title">Modüller</h2>'
        '<p class="kp-section-sub">Her biri tek başına çalışır, birlikte daha güçlüdür.</p>',
        unsafe_allow_html=True,
    )
    mod_cols = st.columns(len(cards))
    for col, (code, m_title, m_desc, m_page, m_bg, m_fg) in zip(mod_cols, cards):
        idx = MODULE_CARDS.index((code, m_title, m_desc, m_page, m_bg, m_fg))
        with col:
            with st.container(border=True, key=f"card_mod_{idx}"):
                st.markdown(
                    f"""
                    <div class="kp-mod-icon" style="background:{'#FFFFFF' if FUN_THEME else m_bg};color:{m_fg};">{MODULE_ICONS.get(code, html.escape(code))}</div>
                    <div class="kp-mod-title">{html.escape(m_title)}</div>
                    <div class="kp-mod-desc">{html.escape(m_desc)}</div>
                    """,
                    unsafe_allow_html=True,
                )
                if st.button("Aç →", key=f"mod_open_{idx}"):
                    navigate_to(m_page)


def gap() -> None:
    st.markdown('<div class="kp-gap"></div>', unsafe_allow_html=True)


def greeting(sub: str) -> None:
    first = USER["name"].split()[0]
    hour = datetime.now().hour
    hello = "Günaydın" if 5 <= hour < 12 else "İyi akşamlar" if hour >= 18 else "Merhaba"
    st.markdown(
        f'<h1 class="kp-hero-title" style="font-size:clamp(1.9rem,3vw,2.6rem);margin-bottom:0.3rem;">'
        f'{hello}, <em>{html.escape(first)}</em>!</h1><p class="kp-hero-lead">{html.escape(sub)}</p>',
        unsafe_allow_html=True,
    )


def extract_explanation(full_text: str) -> str:
    """Soru bloğundaki [Çözüm ...] kısmını ayıklar (redaksiyon notu hariç)."""
    m = re.search(r"[\*\[]*\s*Çözüm[^\n\]:]*\]?\**\s*:?\s*(.*?)(?=\n\s*[\*\[]*\s*Redaksiyon|\Z)",
                  full_text, flags=re.IGNORECASE | re.DOTALL)
    return m.group(1).strip().strip("*").strip() if m else ""


def _strip_q_header(stem: str) -> str:
    """'Soru 3' / '**3.**' gibi numara başlığını kökten ayıklar."""
    out = re.sub(r"^\s*(?:#{1,6}\s*)?(?:\*\*)?\s*(?:Soru\s*\d+|\d+\s*(?:\.\s*Soru|[\.\)]))\s*[\.:\-–]?\s*(?:\*\*)?\s*",
                 "", stem, count=1, flags=re.IGNORECASE)
    return out.strip() or stem.strip()


def add_text_to_bank(content: str, grade: str, branch: str, outcome: str, difficulty: str = "") -> Tuple[int, int]:
    """Üretilen test metnini sorulara bölüp öğretmenin bankasına TASLAK olarak ekler. (eklenen, cevapsız)"""
    parsed = parse_quiz_questions(content)
    items = [
        {"stem": _strip_q_header(q["stem"]), "options": q["options"], "correct": q.get("correct"),
         "explanation": extract_explanation(q.get("full_text", ""))}
        for q in parsed
    ]
    ids = db.add_bank_questions(USER["id"], items, grade, branch, outcome, difficulty, status="taslak")
    missing = sum(1 for it in items if not it.get("correct"))
    return len(ids), missing


def teacher_brand() -> Tuple[str, str]:
    """(kurum adı, imza) — öğretmen profilinden, boşsa varsayılanlar."""
    prof = db.get_teacher_profile(USER["id"]) if ROLE == "ogretmen" else {}
    return (prof.get("org_name") or BRAND_NAME, prof.get("signature") or USER["name"])


def render_assign_popover(content: str, grade: str, branch: str, outcome: str, key: str, difficulty: str = "") -> None:
    """Öğretmen için: test önce Soru Bankası'na taslak olarak gider; ödev yalnızca onaylı sorulardan atanır."""
    if ROLE != "ogretmen":
        return
    if st.button("Soru Bankama taslak ekle", key=f"to_bank_{key}", type="primary",
                 width="stretch", icon=":material/inventory_2:"):
        added, missing = add_text_to_bank(content, grade, branch, outcome, difficulty)
        if not added:
            st.error("Bu metinde soru bulunamadı (Soru 1 ... A) B) C) D) biçimi gerekli).")
        else:
            msg = f"{added} soru taslak olarak eklendi. Soru Bankam'da kontrol edip onaylayın."
            if missing:
                msg += f" {missing} sorunun cevabı okunamadı; onaydan önce işaretleyin."
            record_activity("Soru Bankam", f"{added} taslak soru eklendi")
            st.session_state["bank_flash"] = msg
            navigate_to("Soru Bankam")


def render_class_assign_form(content: str, grade: str, branch: str, outcome: str, key: str) -> None:
    """Onaylı sorulardan oluşan metni sınıfa ödev olarak atar."""
    classes = db.teacher_classes(USER["id"])
    with st.popover("Sınıfa ödev ata", width="stretch"):
        if not classes:
            st.info("Önce **Sınıflarım** sayfasından bir sınıf oluşturun.")
            return
        cls_map = {f"{c['name']} ({c['member_count']} öğrenci)": c["id"] for c in classes}
        cls_label = st.selectbox("Sınıf", list(cls_map.keys()), key=f"as_cls_{key}")
        title = st.text_input("Ödev başlığı", value=f"{branch} · {outcome[:60]}", key=f"as_title_{key}")
        due = st.date_input("Son teslim", value=date.today() + timedelta(days=7), key=f"as_due_{key}")
        video = st.text_input(
            "Konu anlatım videosu (isteğe bağlı)", key=f"as_vid_{key}",
            placeholder="https://www.youtube.com/watch?v=...",
            help="Öğrenci soruları çözmeden önce bu videoyu görür. Kendi videonuz ya da YouTube bağlantısı olabilir.",
        )
        if st.button("Ata", key=f"as_btn_{key}", type="primary", width="stretch"):
            if not parse_quiz_questions(content):
                st.error("Bu metinde çözülebilir soru bulunamadı (Soru 1 ... A) B) C) D) biçimi gerekli).")
                return
            try:
                db.create_assignment(cls_map[cls_label], title, content, grade, branch, outcome, due.isoformat(), video)
            except ValueError as exc:
                st.error(str(exc))
                return
            record_activity("Ödev", f"{title} → {cls_label}")
            st.success("Ödev atandı. Öğrenciler portaldaki Ödevlerim sekmesinde görecek.")


def parse_student_list(text: str, upload: Any = None) -> List[Dict[str, str]]:
    """Serbest metin ya da CSV/XLSX'ten [{"name", "email"}] listesi."""
    raw_rows: List[List[str]] = []
    if upload is not None:
        import pandas as pd

        try:
            if upload.name.lower().endswith(".xlsx"):
                df = pd.read_excel(upload, header=None, dtype=str)
            else:
                data = upload.getvalue().decode("utf-8-sig", errors="replace")
                sep = ";" if data.count(";") > data.count(",") else ","
                df = pd.read_csv(io.StringIO(data), header=None, dtype=str, sep=sep)
            raw_rows += [[str(v) if str(v) != "nan" else "" for v in row] for row in df.fillna("").values.tolist()]
        except Exception as exc:
            st.error(f"Dosya okunamadı: {exc}")
    for line in (text or "").splitlines():
        if line.strip():
            raw_rows.append(re.split(r"[;,\t]", line))
    out: List[Dict[str, str]] = []
    seen: set = set()
    for cells in raw_rows:
        cells = [c.strip() for c in cells if c is not None]
        if not cells or not cells[0]:
            continue
        name = cells[0]
        if name.lower() in ("ad soyad", "ad", "isim", "öğrenci", "ogrenci", "name"):
            continue  # başlık satırı
        email = next((c for c in cells[1:] if "@" in c), "")
        k = (name.lower(), email.lower())
        if k in seen:
            continue
        seen.add(k)
        out.append({"name": name, "email": email})
    return out


def render_bank_editor(q: Dict[str, Any]) -> None:
    """Taslak sorunun düzenleme kartı: düzelt, cevabı işaretle, onayla ya da sil."""
    qid = q["id"]
    letters = sorted(set(["A", "B", "C", "D"]) | set(q["options"].keys()))
    with st.container(border=True, key=f"card_bq_{qid}"):
        st.caption(" · ".join(x for x in [f"#{qid}", q.get("grade"), q.get("branch"), q.get("difficulty")] if x))
        with st.form(f"bq_form_{qid}", border=False):
            outcome = st.text_input("Kazanım", value=q.get("outcome") or "", key=f"bq_out_{qid}")
            stem = st.text_area("Soru kökü", value=q["stem"], height=120, key=f"bq_stem_{qid}")
            opt_cols = st.columns(2)
            opts: Dict[str, str] = {}
            for i, letter in enumerate(letters):
                with opt_cols[i % 2]:
                    opts[letter] = st.text_input(f"{letter})", value=q["options"].get(letter, ""), key=f"bq_{qid}_{letter}")
            cur = q.get("correct")
            correct = st.radio("Doğru cevap", letters, index=letters.index(cur) if cur in letters else None,
                               horizontal=True, key=f"bq_ok_{qid}")
            if cur not in letters:
                st.warning("Cevap anahtarı okunamadı. Doğru şıkkı işaretleyin.")
            expl = st.text_area("Çözüm / açıklama (öğrenciye gösterilmez, sizin notunuz)",
                                value=q.get("explanation") or "", height=80, key=f"bq_ex_{qid}")
            b1, b2, b3 = st.columns(3)
            approve = b1.form_submit_button("Kaydet ve onayla", type="primary", width="stretch")
            save = b2.form_submit_button("Sadece kaydet", width="stretch")
            remove = b3.form_submit_button("Sil", width="stretch")
        if approve or save:
            try:
                db.update_bank_question(USER["id"], qid, stem, opts, correct, expl, outcome)
                if approve:
                    db.set_bank_status(USER["id"], [qid], "onayli")
                    st.toast(f"#{qid} onaylandı.")
                st.rerun()
            except ValueError as exc:
                st.error(str(exc))
        if remove:
            db.delete_bank_questions(USER["id"], [qid])
            st.rerun()


def _bank_meta(chosen: List[Dict[str, Any]]) -> Tuple[str, str, str]:
    grades = list(dict.fromkeys(q.get("grade") or "" for q in chosen))
    branches = list(dict.fromkeys(q.get("branch") or "" for q in chosen))
    outcomes = list(dict.fromkeys(q.get("outcome") or "" for q in chosen if q.get("outcome")))
    grade = grades[0] if len(grades) == 1 else "Karma"
    branch = branches[0] if len(branches) == 1 else "Karma"
    outcome = outcomes[0] if len(outcomes) == 1 else "; ".join(outcomes)[:120]
    return grade, branch, outcome


def booklet_download(content: str, title: str, meta: str, key: str) -> None:
    questions = parse_quiz_questions(content)
    if not questions:
        st.button("Optik formlu kitapçık", key=f"bk_dis_{key}", disabled=True, width="stretch",
                  help="Çözülebilir soru bulunamadı.")
        return
    st.download_button(
        "Optik formlu kitapçık",
        data=reports.build_exam_booklet_docx(questions, title, meta, LOGO_PNG_BYTES),
        file_name="kursun_kalem_deneme_kitapcigi.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        key=f"bk_{key}",
        icon=":material/print:",
        width="stretch",
    )


def load_into_portal(content: str) -> None:
    st.session_state["portal_exam_source"] = content
    st.session_state["review_map"] = {}
    st.session_state["active_assignment"] = None
    st.session_state.pop("last_exam_report", None)
    navigate_to("Öğrenci Portalı & Çözüm")


# ==============================================================================
# 1. ANA SAYFA (role göre)
# ==============================================================================
if current_page == "Ana Sayfa":
    stats_now = st.session_state["parent_student_stats"]

    # ---------------- ÖĞRENCİ ----------------
    if ROLE == "ogrenci":
        prog = user_progress(USER["id"])
        due = gamify.due_reviews(st.session_state.get("error_notebook_2028", []))
        pending = [a for a in db.student_assignments(USER["id"]) if not a.get("submitted_at")]
        greeting("Bugün de biraz ilerleyelim. Küçük adımlar, büyük netler.")

        t1, t2, t3 = st.columns(3, gap="medium")
        with t1:
            with st.container(border=True, key="card_today"):
                goal_pct = min(100, round(100 * prog["solved_today"] / max(1, prog["goal"])))
                done_txt = "Hedef tamam!" if goal_pct >= 100 else f"{prog['goal'] - prog['solved_today']} soru kaldı"
                st.markdown(
                    f"""
                    <div class="kp-today-label">BUGÜNKÜ HEDEFİN</div>
                    <div class="kp-today-big">{prog['solved_today']}<span> / {prog['goal']} soru</span></div>
                    <div class="kp-meter" style="margin:0.6rem 0 0.4rem;"><span style="width:{goal_pct}%;background:#16A34A"></span></div>
                    <div class="kp-today-sub">{done_txt} · <b>{prog['streak']} gün</b> seri</div>
                    """,
                    unsafe_allow_html=True,
                )
                if st.button("Mini deneme çöz", key="home_go_exam", type="primary", width="stretch"):
                    navigate_to("Öğrenci Portalı & Çözüm")
        with t2:
            with st.container(border=True, key="card_review"):
                st.markdown(
                    f"""
                    <div class="kp-today-label">TEKRAR ZAMANI</div>
                    <div class="kp-today-big">{len(due)}<span> soru</span></div>
                    <div class="kp-today-sub" style="margin-top:0.55rem;">
                    {"Hata defterinden bugün tekrar etmen gerekenler hazır." if due else "Bugün tekrar edilecek soru yok. Harika!"}
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                if st.button("Tekrar testini başlat", key="home_review", width="stretch", disabled=not due):
                    start_review_exam(due)
        with t3:
            with st.container(border=True, key="card_homework"):
                st.markdown(
                    f"""
                    <div class="kp-today-label">ÖDEVLERİN</div>
                    <div class="kp-today-big">{len(pending)}<span> bekleyen</span></div>
                    <div class="kp-today-sub" style="margin-top:0.55rem;">
                    {html.escape(pending[0]["title"] + " · " + pending[0]["class_name"]) if pending else "Bekleyen ödevin yok."}
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                if st.button("Ödevi çöz" if pending else "Sınıfa katıl", key="home_homework", width="stretch"):
                    if pending:
                        start_assignment(pending[0])
                    else:
                        st.session_state["portal_tab_hint"] = "Ödevlerim"
                        navigate_to("Öğrenci Portalı & Çözüm")

        gap()
        w_l, w_r = st.columns([1.2, 1.0], gap="large", vertical_alignment="center")
        with w_l:
            st.markdown('<h2 class="kp-section-title" style="font-size:1.3rem;">Son 7 gün</h2>', unsafe_allow_html=True)
            render_week_chart(USER["id"], height=240)
        with w_r:
            render_week_card(stats_now)

        gap()
        # Sınıf seçici
        with st.container(border=True, key="card_grade"):
            sel_l, sel_r = st.columns([1.3, 1.0], vertical_alignment="center")
            with sel_l:
                st.markdown(
                    '<h2 class="kp-section-title">Ne çalışmak istiyorsun?</h2>'
                    '<p class="kp-section-sub" style="margin-bottom:0;">Sınıfını ya da sınavını seç, uygun modül açılsın.</p>',
                    unsafe_allow_html=True,
                )
            with sel_r:
                group_names = list(GRADE_GROUPS.keys())
                selected_group = st.segmented_control(
                    "Kademe", options=group_names, default=group_names[0],
                    key="grade_group_sel", label_visibility="collapsed",
                ) or group_names[0]
            g_idx = group_names.index(selected_group)
            chip_cols = st.columns(4)
            for idx, chip_label in enumerate(GRADE_GROUPS[selected_group]):
                chip = CHIP_BY_LABEL[chip_label]
                if chip["page"] not in ALLOWED_PAGES:
                    continue
                with chip_cols[idx % 4]:
                    if st.button(chip["label"], key=f"grade_chip_{g_idx}_{idx}", width="stretch"):
                        st.session_state["factory_grade"] = chip["grade"]
                        st.session_state["factory_grade_sync"] = True
                        st.session_state["dk_exam"] = chip["exam"]
                        record_activity("Hızlı Kademe", f"{chip['label']} seçildi")
                        navigate_to(chip["page"])

        gap()
        render_modules_section()

        gap()
        # Önerilen odak konuları: önce öğrencinin gerçek zayıf kazanımları, veri yoksa genel öneriler
        _xr_home = db.student_xray(USER["id"])
        _my_weak = [o for o in _xr_home["outcomes"] if o["acc"] < 80][:4]
        if _my_weak:
            focus_items = [
                {"topic": o["outcome"], "branch": o["branch"] or "—", "grade": USER.get("grade") or "",
                 "action": f"{o['n']} ödevde ortalama %{o['acc']} doğru", "pct": o["acc"]}
                for o in _my_weak
            ]
            focus_sub = "Senin ödev sonuçlarına göre en çok tekrar isteyen kazanımlar."
        else:
            focus_items = [{**p, "pct": None} for p in ONCELIKLI_EKSIK_SIRASI]
            focus_sub = "Genel öneri: sınavda sık çıkan, öğrencilerin çok zorlandığı konular. Ödev çözdükçe sana özel olur."
        pr_l, pr_r = st.columns([4, 1], vertical_alignment="bottom")
        with pr_l:
            st.markdown(
                '<h2 class="kp-section-title">Önerilen odak konuları</h2>'
                f'<p class="kp-section-sub">{focus_sub}</p>',
                unsafe_allow_html=True,
            )
        with pr_r:
            if st.button("Tüm analizi gör →", key="prio_all"):
                navigate_to("Öğrenci Röntgeni")
        with st.container(border=True, key="card_prio"):
            for idx, p_item in enumerate(focus_items):
                if idx:
                    st.markdown('<div class="kp-prio-divider"></div>', unsafe_allow_html=True)
                c_num, c_txt, c_meter, c_btn = st.columns([0.45, 5, 2.2, 1.2], vertical_alignment="center")
                with c_num:
                    st.markdown(f'<div class="kp-prio-num">{idx + 1}</div>', unsafe_allow_html=True)
                with c_txt:
                    st.markdown(
                        f'<div class="kp-prio-title">{html.escape(p_item["topic"])}</div>'
                        f'<div class="kp-prio-sub">{html.escape(p_item["branch"])} · {html.escape(p_item["action"])}</div>',
                        unsafe_allow_html=True,
                    )
                with c_meter:
                    pct = p_item.get("pct")
                    if pct is not None:
                        bar_color = "#DC2626" if pct < 50 else "#EA580C" if pct < 60 else "#CA8A04"
                        st.markdown(
                            f'<div class="kp-meter-label"><span>Senin doğrun</span><b style="color:#0F172A">%{pct}</b></div>'
                            f'<div class="kp-meter"><span style="width:{pct}%;background:{bar_color}"></span></div>',
                            unsafe_allow_html=True,
                        )
                    else:
                        st.markdown('<span class="kp-tag-warn">Genel öneri</span>', unsafe_allow_html=True)
                with c_btn:
                    if st.button("Çalış", key=f"prio_btn_{idx}", width="stretch"):
                        st.session_state["factory_focus_topic"] = p_item["topic"]
                        if p_item["branch"] != "—":
                            st.session_state["factory_branch"] = p_item["branch"]
                        if p_item["grade"] in GRADE_LEVELS_2027:
                            st.session_state["factory_grade"] = p_item["grade"]
                            st.session_state["factory_grade_sync"] = True
                        st.session_state["fab_theme_inp"] = p_item["topic"]
                        st.session_state["factory_sync_pending"] = True
                        record_activity("Öncelikli Akış", f"{p_item['topic']} seçildi")
                        navigate_to("Soru & Redaksiyon Fabrikası")

        gap()
        b_l, b_r = st.columns([2, 1], gap="large")
        with b_l:
            earned = db.user_badges(USER["id"])
            st.markdown(
                f'<h2 class="kp-section-title" style="font-size:1.3rem;">Rozetlerin '
                f'<span style="font-size:1rem;color:#64748B;">({len(earned)}/{len(gamify.BADGES)})</span></h2>',
                unsafe_allow_html=True,
            )
            render_badge_grid(earned)
        with b_r:
            st.markdown('<h2 class="kp-section-title" style="font-size:1.3rem;">Son aktiviteler</h2>', unsafe_allow_html=True)
            logs = st.session_state.get("activity_log", [])[:5]
            if not logs:
                st.caption("Henüz bir hareket yok. İlk mini denemeni çözerek başla!")
            for item in logs:
                st.markdown(
                    f"""
                    <div style="padding:0.6rem 0;border-bottom:1px solid #E2E8F0;">
                        <div style="font-size:0.78rem;color:#64748B;">{html.escape(item['time'])} · {html.escape(item['module'])}</div>
                        <div style="font-size:0.9rem;color:#0F172A;margin-top:0.1rem;">{html.escape(item['summary'])}</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

        gap()
        render_countdown_section()

    # ---------------- VELİ ----------------
    elif ROLE == "veli":
        if not LIVE_STUDENT_ID:
            greeting("Çocuğunuzun hesabını bağlayın, ilerlemesini buradan takip edin.")
            with st.container(border=True, key="card_link"):
                st.markdown(
                    '<h2 class="kp-section-title" style="font-size:1.3rem;">Öğrenci bağla</h2>'
                    '<p class="kp-section-sub">Bağlantı kodu, çocuğunuzun <b>Profilim &amp; Rozetler</b> sayfasında yazar.</p>',
                    unsafe_allow_html=True,
                )
                lc1, lc2 = st.columns([3, 1], vertical_alignment="bottom")
                with lc1:
                    code_in = st.text_input("Bağlantı kodu", max_chars=6, key="home_link_code")
                with lc2:
                    if st.button("Bağla", key="home_link_btn", type="primary", width="stretch"):
                        try:
                            child = db.link_parent(USER["id"], code_in)
                            st.success(f"{child['name']} hesabınıza bağlandı.")
                            st.rerun()
                        except ValueError as exc:
                            st.error(str(exc))
        else:
            child = db.get_user(LIVE_STUDENT_ID) or {}
            child_prog = user_progress(LIVE_STUDENT_ID)
            greeting(f"{child.get('name', 'Öğrenciniz')} bu hafta nasıl gidiyor, birlikte bakalım.")
            v1, v2, v3 = st.columns(3)
            v1.metric("Çalışma serisi", f"{child_prog['streak']} gün")
            v2.metric("Bugün çözülen", f"{child_prog['solved_today']} / {child_prog['goal']}")
            v3.metric("Kazanılan rozet", f"{len(db.user_badges(LIVE_STUDENT_ID))} / {len(gamify.BADGES)}")
            gap()
            w_l, w_r = st.columns([1.2, 1.0], gap="large", vertical_alignment="center")
            with w_l:
                st.markdown('<h2 class="kp-section-title" style="font-size:1.3rem;">Son 7 gün</h2>', unsafe_allow_html=True)
                render_week_chart(LIVE_STUDENT_ID, height=240)
            with w_r:
                render_week_card(stats_now, show_sticker=False)
            if st.button("Veli Takip Paneline git →", key="home_parent_panel", type="primary"):
                navigate_to("Veli Takip Paneli")
        gap()
        render_countdown_section()

    # ---------------- ÖĞRETMEN ----------------
    else:
        classes = db.teacher_classes(USER["id"])
        greeting("Sınıflarınız ve ödevleriniz tek ekranda.")
        _bc = db.bank_counts(USER["id"])
        if _bc["taslak"]:
            hb1, hb2 = st.columns([4, 1], vertical_alignment="center")
            with hb1:
                st.markdown(
                    f'<div class="kp-banner kp-banner-hw"><b>{_bc["taslak"]} soru</b> onayınızı bekliyor. '
                    f'Onaylı soru sayınız: <b>{_bc["onayli"]}</b>.</div>',
                    unsafe_allow_html=True,
                )
            with hb2:
                if st.button("İncele", key="home_bank", type="primary", width="stretch"):
                    navigate_to("Soru Bankam")
        if not classes:
            with st.container(border=True, key="card_first_class"):
                st.markdown(
                    '<h2 class="kp-section-title" style="font-size:1.3rem;">İlk sınıfınızı oluşturun</h2>'
                    '<p class="kp-section-sub">Sınıf kodunu öğrencilerinizle paylaşın; ödev atayın, sonuçları anında görün.</p>',
                    unsafe_allow_html=True,
                )
                if st.button("Sınıf oluştur", key="home_new_class", type="primary"):
                    navigate_to("Sınıflarım")
        else:
            cls_cols = st.columns(min(4, len(classes)))
            for i, cls in enumerate(classes[:4]):
                ov = db.class_overview(cls["id"])
                avgs = [r["avg_accuracy"] for r in ov["by_student"] if r["avg_accuracy"] is not None]
                with cls_cols[i]:
                    with st.container(border=True, key=f"card_cls_{cls['id']}"):
                        st.markdown(
                            f"""
                            <div class="kp-today-label">{html.escape(tr_upper(cls.get('grade') or 'Sınıf'))}</div>
                            <div class="kp-today-big">{html.escape(cls['name'])}</div>
                            <div class="kp-today-sub" style="margin-top:0.4rem;">
                                {cls['member_count']} öğrenci · {ov['total_assignments']} ödev<br>
                                Ortalama doğruluk: <b>{('%' + str(round(sum(avgs) / len(avgs)))) if avgs else '—'}</b><br>
                                Sınıf kodu: <b>{cls['code']}</b>
                            </div>
                            """,
                            unsafe_allow_html=True,
                        )
                        if st.button("Aç", key=f"home_cls_{cls['id']}", width="stretch"):
                            st.session_state["teacher_class_sel"] = cls["id"]
                            navigate_to("Sınıflarım")
        gap()
        render_modules_section()
        gap()
        render_countdown_section()

    if st.session_state.live_tick:
        time.sleep(1)
        st.rerun()

# ==============================================================================
# 2. SORU & REDAKSİYON FABRİKASI (5–12. SINIF MÜFREDAT + KADEMELİ SERİ + HAVUZ)
# ==============================================================================
elif current_page == "Soru & Redaksiyon Fabrikası":
    tab_gen, tab_bank = st.tabs(["Soru Üretim & Redaksiyon Stüdyosu", f"Soru Havuzu ({len(st.session_state['question_bank_2028'])})"])

    with tab_gen:
        if st.session_state.get("factory_grade") in GRADE_LEVELS_2027 and st.session_state.get("factory_grade_sync"):
            st.session_state["fab_grade_sel"] = st.session_state["factory_grade"]
            st.session_state["factory_grade_sync"] = False

        cur_grade_init = st.session_state.get("fab_grade_sel", "8. Sınıf (LGS)")
        if cur_grade_init not in GRADE_LEVELS_2027:
            cur_grade_init = "8. Sınıf (LGS)"

        if st.session_state.get("factory_sync_pending"):
            st.session_state["fab_grade_sel"] = st.session_state.get("factory_grade", cur_grade_init)
            g_tmp = st.session_state["fab_grade_sel"]
            b_list = get_branches_for_grade_2027(g_tmp)
            target_b = st.session_state.get("factory_branch")
            st.session_state["fab_branch_sel"] = target_b if target_b in b_list else b_list[0]
            a_list = get_areas_for_grade_branch_2027(g_tmp, st.session_state["fab_branch_sel"])
            st.session_state["fab_area_sel"] = a_list[0]
            o_list = get_outcomes_for_2027(g_tmp, st.session_state["fab_branch_sel"], a_list[0])
            focus_t = st.session_state.get("factory_focus_topic")
            if focus_t and focus_t not in o_list:
                st.session_state["fab_custom_outcome_chk"] = True
                st.session_state["fab_custom_outcome_inp"] = focus_t
            else:
                st.session_state["fab_outcome_sel"] = focus_t if focus_t in o_list else o_list[0]
            if focus_t:
                st.session_state["fab_theme_inp"] = focus_t
            st.session_state["factory_sync_pending"] = False

        if st.session_state.get("factory_focus_topic"):
            st.info(f"Aktarılan odak konu: **{st.session_state['factory_focus_topic']}**")

        with st.expander("5–12. Sınıf Müfredat, Zorluk & Kalite Parametreleri", expanded=True):
            r1c1, r1c2, r1c3 = st.columns(3)
            with r1c1:
                institution = st.text_input("Kurum / Ofis Adı", value=BRAND_NAME, key="fab_inst_inp")
                grade = st.selectbox("Sınıf Seviyesi (5–12)", GRADE_LEVELS_2027, key="fab_grade_sel")
                branches = get_branches_for_grade_2027(grade)
                if st.session_state.get("fab_branch_sel") not in branches:
                    st.session_state["fab_branch_sel"] = branches[0]
                branch = st.selectbox("Branş Seçimi", branches, key="fab_branch_sel")
            with r1c2:
                areas = get_areas_for_grade_branch_2027(grade, branch)
                if st.session_state.get("fab_area_sel") not in areas:
                    st.session_state["fab_area_sel"] = areas[0]
                learning_area = st.selectbox("Öğrenme Alanı / Tema", areas, key="fab_area_sel")

                use_custom_outcome = st.checkbox("Özel / Manuel Kazanım veya Konu Gir", key="fab_custom_outcome_chk")
                if use_custom_outcome:
                    outcome = st.text_input(
                        "Özel Kazanım / Zümre Konusu",
                        value=st.session_state.get("fab_custom_outcome_inp", ""),
                        placeholder="Örn: Paragrafta yardımcı düşünce & çeldirici kurgusu",
                        key="fab_custom_outcome_inp",
                    ) or "Özel Kazanım Odaklı Soru Üretimi"
                else:
                    outcomes = get_outcomes_for_2027(grade, branch, learning_area)
                    if st.session_state.get("fab_outcome_sel") not in outcomes:
                        st.session_state["fab_outcome_sel"] = outcomes[0]
                    outcome = st.selectbox("Hedef Kazanım", outcomes, key="fab_outcome_sel")

                if "fab_theme_inp" not in st.session_state:
                    st.session_state["fab_theme_inp"] = st.session_state.get("factory_focus_topic") or "2027 DİJİTAL OFİS & HAYAT"
                context_theme = st.text_input("Bağlam / Tema / Senaryo", key="fab_theme_inp")
            with r1c3:
                production_mode = st.radio(
                    "Üretim Modu",
                    ["Standart Test Paketi", "Kademeli Zorluk Serisi (Kolay → Orta → Zor)"],
                    key="fab_prod_mode",
                )
                question_count = st.slider("Soru Adedi", 1, 10, 3, key="fab_q_cnt")
                if "fab_diff_sld" not in st.session_state:
                    # Uyarlanabilir zorluk: öğrencinin son denemelerine göre başlangıç seviyesi
                    accs = [r["accuracy"] for r in db.recent_exam_results(USER["id"], 3)] if ROLE == "ogrenci" else []
                    sug, why = gamify.suggest_difficulty(accs)
                    st.session_state["fab_diff_sld"] = sug
                    st.session_state["fab_diff_why"] = why if accs else ""
                difficulty = st.select_slider(
                    "Temel Zorluk Düzeyi",
                    options=gamify.DIFFICULTY_LEVELS,
                    key="fab_diff_sld",
                )
                if st.session_state.get("fab_diff_why"):
                    st.caption(st.session_state["fab_diff_why"])
                distractor_mode = st.selectbox(
                    "Çeldirici & Kök Stratejisi",
                    [
                        "Muhakeme & Çeldirici Tuzağı Dengeli",
                        "Olumsuz Kök ('Değildir / Ulaşılamaz') Odaklı",
                        "Öncüllü & Bağlamsal Analiz Odaklı",
                    ],
                    index=0,
                    key="fab_dist_sel",
                )
                auto_review = st.checkbox("Çift Katmanlı Redaksiyon & Çeldirici Kalite Denetimi", value=True, key="fab_rev_chk")

        if st.button("Kurumsal Soru Paketi Üret", type="primary", key="fab_gen_btn"):
            quota_ok, quota_msg = consume_quota("question", int(question_count))
            if not quota_ok:
                st.error(quota_msg)
            else:
                mode_instruction = (
                    f"Her soruyu '{difficulty}' seviyesinde üret."
                    if production_mode == "Standart Test Paketi"
                    else "Soruları Kolay (pekiştirme), Orta (kazanım) ve Zor (2027 LGS/YKS muhakeme) olarak kademelendir."
                )
                sys_prompt = (
                    "Sen Türkiye Yüzyılı Maarif Modeli, 2027 LGS ve 2027 YKS ölçme-değerlendirme standartlarında "
                    "uzman bir kıdemli test editörüsün. Soruların özgün, dilbilgisi ve noktalama açısından "
                    "kusursuz, çeldiricilerinin uzunluk ve dilbilgisi paralelliği açısından dengeli olmasını sağla. "
                    "Her soruyu 'Soru 1', 'Soru 2'... şeklinde numaralandır ve şıkları 'A)', 'B)', 'C)', 'D)' "
                    "olarak yaz. Her sorunun altında kısa bir [Çözüm & Şık Gerekçesi] ve redaksiyon açıksa "
                    "[Redaksiyon & Çeldirici Kalite Notu] ekle; en sonda mutlaka 'Cevap Anahtarı' başlığı altında "
                    "'1-A', '2-C' formatında net anahtar ver."
                )
                user_prompt = (
                    f"Kurum: {institution}\nSınıf Seviyesi: {grade}\nBranş: {branch}\n"
                    f"Öğrenme Alanı: {learning_area}\nHedef Kazanım: {outcome}\n"
                    f"Bağlam / Tema / Senaryo: {context_theme}\nZorluk Kurgusu: {mode_instruction}\n"
                    f"Çeldirici & Kök Stratejisi: {distractor_mode}\nSoru Adedi: {question_count}\n"
                    f"Redaksiyon & Kalite Denetimi: {'Aktif' if auto_review else 'Pasif'}\n\n"
                    "Lütfen numaralandırılmış çoktan seçmeli (Soru 1, Soru 2... ve A, B, C, D şıkları) özgün sorular üret."
                )
                skel_slot = st.empty()
                render_loading_skeleton(skel_slot, lines=4)
                try:
                    result = call_gemini_text(prompt=user_prompt, system_instruction=sys_prompt)
                    skel_slot.empty()
                    st.session_state.generated_test_text = result
                    st.session_state["last_fab_meta"] = {
                        "grade": grade,
                        "branch": branch,
                        "outcome": outcome,
                        "difficulty": difficulty if production_mode == "Standart Test Paketi" else "Kademeli Seri",
                    }
                    record_activity("Soru Fabrikası", f"{grade} • {branch} • {question_count} soru üretildi")
                    st.success(
                        f"Soru paketi oluşturuldu. Kalan soru kotası: "
                        f"{int(st.session_state.remaining_question_quota):,}".replace(",", ".")
                    )
                except Exception as exc:
                    skel_slot.empty()
                    refund_quota("question", int(question_count))
                    st.error(f"Soru üretim hatası oluştu: {exc}")

        if st.session_state.generated_test_text:
            st.markdown("---")
            st.markdown("### Üretilen Test, Çözüm Gerekçeleri ve Redaksiyon Çıktısı")
            st.markdown(st.session_state.generated_test_text)

            docx_bytes = build_test_docx(
                institution=institution,
                branch=branch,
                grade=grade,
                learning_area=learning_area,
                outcome=outcome,
                context_theme=context_theme,
                generated_content=st.session_state.generated_test_text,
            )
            act_c1, act_c2, act_c3, act_c4 = st.columns(4)
            with act_c1:
                st.download_button(
                    label="Testi DOCX İndir",
                    data=docx_bytes,
                    file_name=f"kursun_kalem_ofis_2027_{branch.lower().replace(' ', '_')}_test.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    width="stretch",
                )
            with act_c2:
                if st.button("Soru Havuzuna Kaydet", width="stretch"):
                    add_to_question_bank(
                        grade=grade,
                        branch=branch,
                        outcome=outcome,
                        difficulty=difficulty,
                        content_text=st.session_state.generated_test_text,
                    )
                    record_activity("Soru Havuzu", f"{grade} • {branch} testi havuza kaydedildi")
                    st.toast("Test Soru Havuzuna kaydedildi.")
                    st.rerun()
            with act_c3:
                if st.button("Mini Deneme Olarak Çöz", width="stretch"):
                    record_activity("Köprü", "Test Öğrenci Portalı mini denemesine aktarıldı")
                    load_into_portal(st.session_state.generated_test_text)
            with act_c4:
                if st.button("İlk Soruyu Çözüm Rehberine Gönder", width="stretch"):
                    parsed_q = parse_quiz_questions(st.session_state.generated_test_text)
                    st.session_state["portal_question_draft"] = (
                        parsed_q[0]["full_text"] if parsed_q else st.session_state.generated_test_text[:1000]
                    )
                    navigate_to("Öğrenci Portalı & Çözüm")
            act_c5, act_c6, _sp = st.columns([1, 1, 2])
            with act_c5:
                booklet_download(
                    st.session_state.generated_test_text,
                    f"{institution} Deneme",
                    f"{grade} · {branch} · {outcome}",
                    "gen",
                )
            with act_c6:
                render_assign_popover(st.session_state.generated_test_text, grade, branch, outcome, "gen",
                                      str(st.session_state.get("last_fab_meta", {}).get("difficulty", "")))

    with tab_bank:
        bank_items: List[Dict[str, str]] = st.session_state.get("question_bank_2028", [])
        if not bank_items:
            st.info("Soru havuzunda henüz kayıtlı test bulunmuyor. Üretim sekmesinden **Soru Havuzuna Kaydet** butonunu kullanabilirsiniz.")
        else:
            combined_bank_text = "\n\n".join(
                [f"=== [{b['grade']} | {b['branch']} | {b['difficulty']}] {b['outcome']} ===\n{b['content']}" for b in bank_items]
            )
            bundle_docx = build_test_docx(
                institution=BRAND_NAME,
                branch="Karma Zümre Havuzu",
                grade="5–12. Sınıf",
                learning_area="Kayıtlı Soru Havuzu",
                outcome="Havuzdaki Seçili Kazanımlar",
                context_theme="2027 Kurumsal Soru Kitapçığı",
                generated_content=combined_bank_text,
            )
            st.download_button(
                label=f"Tüm Havuzu ({len(bank_items)} Kayıt) Toplu DOCX İndir",
                data=bundle_docx,
                file_name="kursun_kalem_2027_soru_havuzu.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
            for b_item in bank_items:
                with st.expander(f"[{b_item['id']}] {b_item['grade']} • {b_item['branch']} ({b_item['difficulty']}) — {b_item['date']}"):
                    st.caption(f"Kazanım: {b_item['outcome']}")
                    st.markdown(b_item["content"])
                    bc1, bc2, bc3 = st.columns(3)
                    with bc1:
                        if st.button("Bu Testi Portala Aktar", key=f"load_bank_{b_item['id']}", width="stretch"):
                            load_into_portal(b_item["content"])
                    with bc2:
                        booklet_download(b_item["content"], "Deneme", f"{b_item['grade']} · {b_item['branch']}", b_item["id"])
                    with bc3:
                        render_assign_popover(b_item["content"], b_item["grade"], b_item["branch"], b_item["outcome"], b_item["id"],
                                          b_item.get("difficulty", ""))

# ==============================================================================
# 3. ÖĞRENCİ PORTALI & ÇÖZÜM (+ SÜRE TAKİBİ + HATA DEFTERİ)
# ==============================================================================
elif current_page == "Öğrenci Portalı & Çözüm":
    _due_now = gamify.due_reviews(st.session_state.get("error_notebook_2028", []))
    _pending_hw = [a for a in db.student_assignments(USER["id"]) if not a.get("submitted_at")]
    if st.session_state.pop("portal_tab_hint", None) == "Ödevlerim":
        st.info("Ödevlerini ve sınıfa katılma alanını **Ödevlerim** sekmesinde bulabilirsin.")
    tab_exam, tab_hw, tab_hint, tab_err = st.tabs([
        "Mini Deneme",
        f"Ödevlerim ({len(_pending_hw)})",
        "Adımlı Çözüm Asistanı",
        f"Hata Defteri & Tekrar ({len(_due_now)})",
    ])

    with tab_exam:
        active_hw = st.session_state.get("active_assignment")
        review_map: Dict[int, str] = st.session_state.get("review_map") or {}
        if active_hw:
            # Ödevde metin (ve cevap anahtarı) gizli kalır, öğrenci değiştiremez
            st.markdown(
                f'<div class="kp-banner kp-banner-hw"><b>Ödev:</b> {html.escape(active_hw["title"])}</div>',
                unsafe_allow_html=True,
            )
            exam_source = st.session_state.get("portal_exam_source", "")
            if active_hw.get("video_url"):
                with st.expander("Önce konu anlatımını izle", expanded=True, icon=":material/play_circle:"):
                    render_lesson_video(active_hw["video_url"])
            if st.button("Ödevden çık", key="leave_hw"):
                st.session_state["active_assignment"] = None
                st.session_state["portal_exam_source"] = ""
                st.rerun()
        else:
            if review_map:
                st.markdown(
                    '<div class="kp-banner kp-banner-review"><b>Tekrar testi:</b> doğru yaptığın sorular '
                    'bir sonraki tekrar aşamasına geçer, yanlışların takvimi baştan başlar.</div>',
                    unsafe_allow_html=True,
                )
            with st.expander("Soru metni (Soru Fabrikasından gelir ya da buraya yapıştırabilirsin)",
                             expanded=not st.session_state.get("portal_exam_source")):
                exam_source = st.text_area(
                    "Deneme Soru Metni",
                    value=st.session_state.get("portal_exam_source") or st.session_state.get("generated_test_text", ""),
                    height=140,
                    placeholder="Soru 1) ... A) ... B) ... C) ... D) ... Cevap Anahtarı: 1-B",
                    label_visibility="collapsed",
                )
            if exam_source != st.session_state.get("portal_exam_source"):
                st.session_state["review_map"] = {}
                review_map = {}
            st.session_state["portal_exam_source"] = exam_source
        questions = parse_quiz_questions(exam_source)
        exam_key = uuid.uuid5(uuid.NAMESPACE_OID, exam_source or "-").hex[:10]

        if not questions:
            st.info(
                "Etkileşimli şık işaretleyebilmek için **Soru & Redaksiyon Fabrikası** sayfasından "
                "bir test üretin veya 'Soru 1 ... A) ... B) ... C) ... D) ...' biçiminde sorular yapıştırın."
            )
        else:
            if "exam_start_ts" not in st.session_state:
                st.session_state["exam_start_ts"] = time.time()

            st.markdown(f"#### Aktif Mini Deneme ({len(questions)} Soru)")
            st.caption("Boş bırakmak isterseniz **Boş Bırak** seçeneğini işaretleyebilir veya seçim yapmayabilirsiniz; boş sorular yanlışa sayılmaz.")
            user_answers: Dict[int, Optional[str]] = {}
            q_durations: Dict[int, float] = {}

            for idx, q in enumerate(questions):
                with st.container(border=True):
                    st.markdown(q["stem"])
                    c_opt, c_sec = st.columns([3.4, 1.0])
                    with c_opt:
                        opts = list(q["options"].keys()) + ["BOŞ"]
                        chosen = st.radio(
                            f"Soru {q['num']} Yanıtın",
                            options=opts,
                            index=None,
                            format_func=lambda k, o_map=q["options"]: "Boş Bırak" if k == "BOŞ" else f"{k}) {o_map[k]}",
                            horizontal=True,
                            key=f"mini_q_{exam_key}_{idx}",
                        )
                        user_answers[idx] = None if chosen in (None, "BOŞ") else chosen
                    with c_sec:
                        q_durations[idx] = float(
                            st.number_input(
                                f"Süre (sn) • S{q['num']}",
                                min_value=5,
                                max_value=300,
                                value=65,
                                step=5,
                                key=f"mini_dur_{exam_key}_{idx}",
                            )
                        )

            if st.button("Denemeyi Bitir ve Netimi Hesapla", type="primary"):
                # Aynı deneme ikinci kez notlanırsa kalıcı kayıtlara (defter, tekrar, sonuç) tekrar yazma
                already_graded = st.session_state.get("graded_exam_key") == exam_key
                correct_cnt = 0
                wrong_cnt = 0
                blank_cnt = 0
                ungraded_cnt = 0
                wrong_questions: List[Dict[str, Any]] = []
                result_rows: List[Dict[str, str]] = []
                total_sec = 0.0

                for idx, q in enumerate(questions):
                    u_ans = user_answers[idx]
                    c_ans = q.get("correct")
                    sec_val = q_durations.get(idx, 60.0)
                    total_sec += sec_val

                    if u_ans is None:
                        blank_cnt += 1
                        status_txt = "Boş"
                    elif not c_ans:
                        ungraded_cnt += 1
                        status_txt = "Anahtar Yok"
                    elif u_ans == c_ans:
                        correct_cnt += 1
                        status_txt = "Doğru"
                    else:
                        wrong_cnt += 1
                        wrong_questions.append(q)
                        status_txt = "Yanlış"
                        if q["num"] not in review_map and not already_graded:
                            add_to_error_notebook(
                                question_num=q["num"],
                                question_text=q["full_text"],
                                user_ans=u_ans,
                                correct_ans=c_ans,
                                elapsed_sec=sec_val,
                            )

                    # Tekrar testindeyse hata defteri kaydını ilerlet / sıfırla
                    if q["num"] in review_map and c_ans and u_ans is not None and not already_graded:
                        nb_item = next(
                            (it for it in st.session_state["error_notebook_2028"] if str(it.get("id")) == review_map[q["num"]]),
                            None,
                        )
                        if nb_item is not None:
                            gamify.advance_review(nb_item, u_ans == c_ans)
                            if nb_item.get("done"):
                                st.session_state["reviews_completed"] = int(st.session_state.get("reviews_completed", 0)) + 1

                    result_rows.append(
                        {
                            "Soru": f"Soru {q['num']}",
                            "Senin Yanıtın": u_ans or "Boş",
                            "Doğru Cevap": c_ans or "Belirtilmedi",
                            "Süre (sn)": f"{int(sec_val)} sn",
                            "Durum": status_txt,
                        }
                    )

                graded_total = correct_cnt + wrong_cnt
                net_score = max(0.0, correct_cnt - (wrong_cnt / 3.0))
                accuracy = (correct_cnt / graded_total * 100.0) if graded_total > 0 else 0.0
                avg_sec = total_sec / max(1, len(questions))

                st.session_state["last_exam_report"] = {
                    "correct": correct_cnt,
                    "wrong": wrong_cnt,
                    "blank": blank_cnt,
                    "net": net_score,
                    "accuracy": accuracy,
                    "avg_sec": avg_sec,
                    "rows": result_rows,
                    "first_wrong": wrong_questions[0]["full_text"] if wrong_questions else "",
                }
                record_activity(
                    "Ödev" if active_hw else "Tekrar Testi" if review_map else "Mini Deneme",
                    f"{correct_cnt}D {wrong_cnt}Y {blank_cnt}B • {net_score:.2f} Net • Ort {avg_sec:.0f} sn/soru",
                )
                # Aynı denemeyi iki kez kaydetme (buton tekrar tıklanırsa)
                if not already_graded:
                    st.session_state["graded_exam_key"] = exam_key
                    record_exam(
                        correct_cnt, wrong_cnt, blank_cnt,
                        source="odev" if active_hw else "tekrar" if review_map else "mini",
                        assignment_id=int(active_hw["id"]) if active_hw else None,
                    )
                    if active_hw:
                        st.success("Ödevin teslim edildi, öğretmenin sonucunu görebilir.")
                        st.session_state["active_assignment"] = None
                if review_map:
                    st.session_state["review_map"] = {}
                save_persisted_state()

            report = st.session_state.get("last_exam_report")
            if report:
                st.markdown("---")
                cheer = st.session_state.pop("celebrate", None)
                if cheer:
                    st.markdown(
                        f'<div class="kp-celebrate"><b>{html.escape(cheer)}</b>'
                        f'<span>%{report["accuracy"]:.0f} doğrulukla bitirdin. Serin büyümeye devam ediyor!</span></div>',
                        unsafe_allow_html=True,
                    )
                r1, r2, r3, r4 = st.columns(4)
                r1.metric("Doğru / Boş", f"{report['correct']}D / {report['blank']}B")
                r2.metric("Yanlış Sayısı", f"{report['wrong']} Soru")
                r3.metric("Hesaplanan Net (3Y=1D)", f"{report['net']:.2f} Net")
                r4.metric("Ort. Soru Süresi", f"{report['avg_sec']:.0f} sn (%{report['accuracy']:.1f})")

                st.dataframe(report["rows"], width="stretch", hide_index=True)

                if report.get("first_wrong"):
                    wa_c1, wa_c2 = st.columns(2)
                    with wa_c1:
                        if st.button("Yanlış Soruyu Çözüm Rehberine Aktar", width="stretch"):
                            st.session_state["portal_question_draft"] = report["first_wrong"]
                            st.toast("Yanlış soru Adımlı Çözüm Asistanı sekmesine aktarıldı.")
                            st.rerun()
                    with wa_c2:
                        if st.button("Yanlış Sorudan Benzer Soru Üret", type="primary", width="stretch"):
                            snippet = re.sub(r"\s+", " ", report["first_wrong"])[:110]
                            st.session_state["factory_focus_topic"] = f"Telafi: {snippet}"
                            st.session_state["fab_theme_inp"] = f"Telafi: {snippet}"
                            record_activity("Telafi Akışı", "Yanlış sorudan benzer soru üretimine geçildi")
                            navigate_to("Soru & Redaksiyon Fabrikası")

    with tab_hw:
        my_classes = db.student_classes(USER["id"])
        j1, j2 = st.columns([3, 1], vertical_alignment="bottom")
        with j1:
            join_code = st.text_input("Sınıf kodu ile katıl", max_chars=6, placeholder="Öğretmeninin verdiği 6 haneli kod", key="join_code_in")
        with j2:
            if st.button("Katıl", key="join_class_btn", width="stretch"):
                try:
                    cls = db.join_class(USER["id"], join_code)
                    st.success(f"{cls['name']} sınıfına katıldın.")
                    st.rerun()
                except ValueError as exc:
                    st.error(str(exc))
        if my_classes:
            st.caption("Sınıfların: " + " · ".join(f"{c['name']} ({c['teacher_name']})" for c in my_classes))
        hw_all = db.student_assignments(USER["id"])
        if not hw_all:
            st.info("Henüz ödevin yok. Bir sınıfa katıldığında öğretmeninin atadığı testler burada görünür.")
        for hw in hw_all:
            with st.container(border=True):
                h1, h2 = st.columns([4, 1], vertical_alignment="center")
                with h1:
                    due_txt = f" · Son teslim: {datetime.fromisoformat(hw['due_date']).strftime('%d.%m.%Y')}" if hw.get("due_date") else ""
                    status = (
                        f"Teslim edildi · %{hw['my_accuracy']} doğruluk · {hw['my_net']} net"
                        if hw.get("submitted_at") else "Bekliyor"
                    )
                    st.markdown(f"**{hw['title']}** · {hw['class_name']}{due_txt}")
                    st.caption(status + (" · Konu anlatım videosu var" if hw.get("video_url") else ""))
                with h2:
                    label = "Tekrar çöz" if hw.get("submitted_at") else "Çöz"
                    if st.button(label, key=f"hw_start_{hw['id']}", type="secondary" if hw.get("submitted_at") else "primary",
                                 width="stretch"):
                        start_assignment(hw)

    with tab_hint:
        col_l, col_r = st.columns([2, 1])
        with col_l:
            default_q = st.session_state.get("portal_question_draft", "")
            student_question = st.text_area(
                "Çözümünü İncelemek İstediğin Soru Metni",
                value=default_q,
                height=160,
                placeholder="Soru kökünü, verilen öncülleri ve şıkları buraya yapıştır...",
            )
        with col_r:
            hint_mode = st.radio(
                "Yardım Düzeyi",
                ["Önce İpucu Ver (Kendim Çözeyim)", "Adım Adım Çözümü Göster"],
                index=0,
            )
            student_attempt = st.text_input("Senin Tahminin / Şıkkın", placeholder="Örn: C şıkkı, çünkü...")

        if st.button("Çözüm Rehberi Oluştur", type="primary", key="btn_solve_guide"):
            if not student_question.strip():
                st.warning("Lütfen önce bir soru metni girin.")
            else:
                sys_inst = (
                    "Sen sabırlı, pedagojik ve teşvik edici bir eğitim mentorusun. "
                    "Öğrenciye doğrudan ezberletmek yerine mantığını kavratacak şekilde açıkla."
                )
                prompt = (
                    f"Soru:\n{student_question}\n\n"
                    f"Öğrencinin Denemesi: {student_attempt or 'Belirtilmedi'}\n"
                    f"Yardım Düzeyi: {hint_mode}\n\n"
                    "Lütfen Türkçe olarak net ve anlaşılır yanıt ver."
                )
                skel_hint = st.empty()
                render_loading_skeleton(skel_hint, lines=3)
                try:
                    ans = call_gemini_text(prompt=prompt, system_instruction=sys_inst)
                    skel_hint.empty()
                    record_activity("Çözüm Asistanı", f"{hint_mode} üretildi")
                    st.markdown("#### Çözüm Analizi")
                    st.markdown(ans)
                except Exception as exc:
                    skel_hint.empty()
                    st.error(f"Çözüm üretilirken hata oluştu: {exc}")

    with tab_err:
        err_list: List[Dict[str, Any]] = st.session_state.get("error_notebook_2028", [])
        open_items = [it for it in err_list if not it.get("done")]
        done_items = [it for it in err_list if it.get("done")]
        if not err_list:
            st.info("Hata defterinde henüz kayıt bulunmuyor. Mini denemede yanlış işaretlenen sorular otomatik buraya eklenir.")
        else:
            rv1, rv2 = st.columns([3, 1], vertical_alignment="center")
            with rv1:
                st.markdown(
                    f"**Bugün tekrar edilecek: {len(_due_now)} soru** · Açık: {len(open_items)} · "
                    f"Tamamen öğrenilen: {len(done_items)}"
                )
                st.caption("Aralıklı tekrar: her soru 1., 3. ve 7. günde tekrar sorulur. Üçünü de doğru yaparsan defterden düşer.")
            with rv2:
                if st.button("Tekrar testini başlat", key="err_review_start", type="primary",
                             width="stretch", disabled=not _due_now):
                    start_review_exam(_due_now)
            show_done = st.toggle("Öğrenilenleri de göster", value=False, key="err_show_done")
            for item in (err_list if show_done else open_items):
                with st.container(border=True):
                    stage = int(item.get("stage", 0))
                    stage_txt = "Öğrenildi" if item.get("done") else f"Tekrar aşaması {stage}/3"
                    st.markdown(
                        f"**[{item['id']}] Soru {item['num']}** • {stage_txt} • Teşhis: **{item['diagnosis']}** • "
                        f"Senin: `{item['user_ans']}` / Doğru: `{item['correct_ans']}` • Süre: `{item['elapsed_sec']} sn`"
                    )
                    st.caption(
                        f"Tekrar Hatırlatmaları → 1. Gün: {item['review_1d']} | "
                        f"3. Gün: {item['review_3d']} | 7. Gün: {item['review_7d']}"
                    )
                    st.markdown(item["question_text"])
                    if st.button("Çözüm Rehberinde Aç", key=f"err_solve_{item['id']}"):
                        st.session_state["portal_question_draft"] = item["question_text"]
                        st.toast("Soru Adımlı Çözüm Asistanına aktarıldı.")

# ==============================================================================
# 4. DİJİMENTOR (2027 + AŞAMALI KAMP PLANI)
# ==============================================================================
elif current_page == "Dijimentor":
    st.markdown(
        """
        <div class="kp-studio-banner">
            <div class="kp-studio-title">Dijimentor — Akıllı Sınav, Faz Planı &amp; Performans Mentorluğu</div>
            <div class="kp-studio-sub">
                2027 hedef sınavına, haftalık çalışma kapasitene ve çeldirici/zaman analizine göre
                aşamalı kamp ve haftalık aksiyon planı üretir.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    current_hours = int(st.session_state.get("dk_hours", 16))
    daily_avg = current_hours / 7.0
    k_c1, k_c2, k_c3, k_c4 = st.columns(4)
    k_c1.markdown(
        f"""
        <div class="kp-kpi-card">
            <div class="kp-kpi-label">MENTORLUK MOTORU</div>
            <div class="kp-kpi-value">Dijimentor 2027</div>
            <div class="kp-kpi-sub">Kalan Plan: {int(st.session_state.remaining_coach_quota)} Hak</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    k_c2.markdown(
        f"""
        <div class="kp-kpi-card">
            <div class="kp-kpi-label">GÜNLÜK ORTALAMA TEMPO</div>
            <div class="kp-kpi-value">{daily_avg:.1f} Saat</div>
            <div class="kp-kpi-sub">Haftalık Toplam: {current_hours} Saat</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    k_c3.markdown(
        """
        <div class="kp-kpi-card">
            <div class="kp-kpi-label">2027 KAMP FAZLARI</div>
            <div class="kp-kpi-value">3 Aşama</div>
            <div class="kp-kpi-sub">Kazanım • Branş Hızı • Genel Deneme</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    k_c4.markdown(
        """
        <div class="kp-kpi-card">
            <div class="kp-kpi-label">ÇIKTI FORMATI</div>
            <div class="kp-kpi-value">Plan + DOCX</div>
            <div class="kp-kpi-sub">İndirilebilir Çalışma Rehberi</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.write("")
    st.markdown("#### Hızlı 2027 Mentorluk Senaryoları")
    p_cols = st.columns(len(DIJIMENTOR_PRESETS) + 1)
    for idx, (preset_title, preset_data) in enumerate(DIJIMENTOR_PRESETS.items()):
        with p_cols[idx]:
            if st.button(preset_title, key=f"preset_{idx}", width="stretch"):
                st.session_state["dk_exam"] = preset_data["exam"]
                st.session_state["dk_hours"] = preset_data["hours"]
                st.session_state["dk_goal"] = preset_data["goal"]
                st.session_state["dk_focus"] = preset_data["focus"]
                st.session_state["dk_weak"] = preset_data["weak"]
                st.session_state["dk_status"] = preset_data["status"]
                st.session_state["dk_active_preset"] = preset_title
                st.session_state["factory_focus_topic"] = preset_data["topic"]
                st.session_state["factory_branch"] = preset_data["branch"]
                st.session_state["factory_grade"] = preset_data["grade"]
                st.session_state["fab_theme_inp"] = preset_data["topic"]
                st.session_state["factory_sync_pending"] = True
                st.session_state["factory_grade_sync"] = True
                record_activity("Dijimentor Yönlendirme", f"{preset_title} seçildi ve Soru Fabrikasına yönlendirildi")
                navigate_to("Soru & Redaksiyon Fabrikası")
    with p_cols[-1]:
        if st.button("Planı Temizle", key="clear_dk_plan", width="stretch"):
            st.session_state.dijimentor_plan = ""
            st.session_state["dk_active_preset"] = ""
            st.rerun()

    st.markdown("---")
    st.markdown("#### Öğrenci 2027 Hedef ve Çalışma Profili")

    exam_opts = ["2027 LGS", "2027 TYT - AYT", "5-7. Sınıf Akademik Gelişim"]
    stored_exam = st.session_state.get("dk_exam", "2027 LGS")
    if stored_exam not in exam_opts:
        stored_exam = "2027 LGS"

    r1_a, r1_b, r1_c = st.columns(3)
    with r1_a:
        target_exam = st.selectbox(
            "Hedef Sınav / Kademe",
            exam_opts,
            index=exam_opts.index(stored_exam),
            key="dk_exam_sel",
        )
        st.session_state["dk_exam"] = target_exam

        weekly_hours = st.slider("Haftalık Çalışma Süresi (Saat)", min_value=4, max_value=45, value=current_hours, key="dk_hours_sld")
        st.session_state["dk_hours"] = weekly_hours

    with r1_b:
        target_goal = st.text_input(
            "Hedef Puan / Net / Okul Bandı",
            value=st.session_state.get("dk_goal", "Nitelikli lise / yüksek net ve sıfıra yakın dikkat hatası"),
            key="dk_goal_inp",
        )
        st.session_state["dk_goal"] = target_goal

        focus_opts = [
            "Haftalık Dengeli Çalışma Programı",
            "Deneme Net Artırma & Zaman Yönetimi",
            "Dikkat ve Çeldirici Tuzağı Azaltma",
            "Sınav Kaygısı & Motivasyon Rutini",
        ]
        stored_focus = st.session_state.get("dk_focus", focus_opts[0])
        if stored_focus not in focus_opts:
            stored_focus = focus_opts[0]
        coach_focus = st.selectbox(
            "Mentorluk Odak Alanı",
            focus_opts,
            index=focus_opts.index(stored_focus),
            key="dk_focus_sel",
        )
        st.session_state["dk_focus"] = coach_focus

    with r1_c:
        weak_topics = st.text_area(
            "Zorlandığın Dersler / Kazanımlar",
            value=st.session_state.get("dk_weak", "Sözel mantık ve muhakeme, üslü ifadeler, dikkat kaynaklı işlem hataları"),
            height=82,
            key="dk_weak_area",
        )
        st.session_state["dk_weak"] = weak_topics

        recent_score = st.text_input(
            "Mevcut Deneme / Süre Durumu",
            value=st.session_state.get("dk_status", "Orta seviye, süreyi yetiştirmede zorlanıyorum"),
            key="dk_status_inp",
        )
        st.session_state["dk_status"] = recent_score

    q_hrs = round(weekly_hours * 0.45, 1)
    r_hrs = round(weekly_hours * 0.30, 1)
    a_hrs = round(weekly_hours * 0.25, 1)
    st.markdown(
        f"""
        <div style="margin-top:0.65rem;padding-top:0.75rem;border-top:1px solid #E2E8F0;">
            <div style="font-size:0.75rem;font-weight:600;color:#475569;letter-spacing:0.04em;">
                ÖNERİLEN HAFTALIK KAPASİTE DAĞILIMI ({weekly_hours} SAAT)
            </div>
            <div class="kp-dist-bar">
                <div style="width:45%;background:#2563EB;"></div>
                <div style="width:30%;background:#5A52E8;"></div>
                <div style="width:25%;background:#94A3B8;"></div>
            </div>
            <div class="kp-dist-legend">
                <span>Soru &amp; Muhakeme: <b>{q_hrs} sa</b> (%45)</span>
                <span>Konu Tekrarı: <b>{r_hrs} sa</b> (%30)</span>
                <span>Deneme &amp; Analiz: <b>{a_hrs} sa</b> (%25)</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.write("")
    include_routine = st.checkbox("2027 Aşamalı Kamp Fazları (Kazanım → Branş → Genel Deneme) ve hafta sonu kontrol listesi dahil et", value=True)

    if st.button("Dijimentor 2027 Planı Oluştur", type="primary"):
        quota_ok, quota_msg = consume_quota("coach", 1)
        if not quota_ok:
            st.error(quota_msg)
        else:
            sys_coach = (
                "Sen 'Dijimentor' adlı, 2027 sınav dönemine hazırlanan öğrenciler için çalışan üst düzey, "
                "net, empatik ve veri odaklı bir dijital eğitim mentorusun. Yanıtını başlıklarla, "
                "madde işaretleriyle ve uygulanabilir zaman bloklarıyla kurumsal bir mentorluk "
                "raporu gibi düzenle. Şu bölümleri net oluştur:\n"
                "1. Teşhis ve 2027 Aşamalı Kamp Yol Haritası (Faz 1 Kazanım, Faz 2 Branş Hızı, Faz 3 Genel Deneme)\n"
                "2. Pazartesi–Pazar Haftalık Çalışma Planı (saat ve odak ders dağılımı)\n"
                "3. Çeldirici & Zaman Yönetimi Protokolü\n"
                "4. Haftalık Başarı KPI'ları ve Kontrol Listesi"
            )
            prompt_coach = (
                f"Hedef Sınav / Kademe: {target_exam}\nHedef Bandı: {target_goal}\n"
                f"Haftalık Çalışma Süresi: {weekly_hours} saat\n"
                f"Zorlanılan Dersler / Kazanımlar: {weak_topics}\n"
                f"Mentorluk Odak Alanı: {coach_focus}\n"
                f"Mevcut Deneme / Süre Durumu: {recent_score}\n"
                f"Aşamalı Kamp & Analiz Protokolü: {'Evet' if include_routine else 'Özet'}\n\n"
                "Lütfen öğrenciye doğrudan uygulanabilir, motive edici ve net bir Türkçe 2027 planı sun."
            )
            skel_coach = st.empty()
            render_loading_skeleton(skel_coach, lines=4)
            try:
                plan = call_gemini_text(prompt=prompt_coach, system_instruction=sys_coach)
                skel_coach.empty()
                st.session_state.dijimentor_plan = plan
                record_activity("Dijimentor", f"{target_exam} • {weekly_hours} saatlik plan üretildi")
                st.success("Dijimentor 2027 planı başarıyla oluşturuldu.")
            except Exception as exc:
                skel_coach.empty()
                refund_quota("coach", 1)
                st.error(f"Dijimentor planı oluşturulurken hata oluştu: {exc}")

    if st.session_state.dijimentor_plan:
        st.markdown("---")
        st.markdown("### Kişiselleştirilmiş Dijimentor 2027 Mentorluk Raporu")
        st.markdown(st.session_state.dijimentor_plan)

        meta_summary = (
            f"Hedef: {target_exam} | Hedef Bandı: {target_goal} | "
            f"Haftalık Süre: {weekly_hours} Saat | Odak: {coach_focus}"
        )
        docx_coach = build_coach_docx(meta_summary, st.session_state.dijimentor_plan)
        txt_report = (
            f"{BRAND_NAME}\n{meta_summary}\nTarih: {datetime.now().strftime('%d.%m.%Y')}\n"
            f"{'=' * 60}\n\n{st.session_state.dijimentor_plan}"
        )

        d_col1, d_col2, d_col3 = st.columns(3)
        with d_col1:
            st.download_button(
                label="Mentorluk Planını DOCX İndir",
                data=docx_coach,
                file_name="dijimentor_2027_calisma_plani.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                width="stretch",
            )
        with d_col2:
            st.download_button(
                label="Planı Metin (.txt) İndir",
                data=txt_report.encode("utf-8"),
                file_name="dijimentor_2027_calisma_plani.txt",
                mime="text/plain",
                width="stretch",
            )
        with d_col3:
            if st.button("Zayıf Konulardan Soru Üret", width="stretch"):
                first_w = weak_topics.split(",")[0].strip()
                st.session_state["factory_focus_topic"] = first_w
                st.session_state["fab_theme_inp"] = first_w
                navigate_to("Soru & Redaksiyon Fabrikası")

# ==============================================================================
# 5. VELİ TAKİP PANELİ (+ BU HAFTA NE DEĞİŞTİ KARTLARI)
# ==============================================================================
elif current_page == "Veli Takip Paneli":
    stats_now = st.session_state["parent_student_stats"]
    if not LIVE_STUDENT_ID:
        st.info("Önce çocuğunuzun hesabını bağlayın. Bağlantı kodu, öğrencinin Profilim & Rozetler sayfasında yazar.")
        lk1, lk2 = st.columns([3, 1], vertical_alignment="bottom")
        with lk1:
            vp_code = st.text_input("Bağlantı kodu", max_chars=6, key="vp_link_code")
        with lk2:
            if st.button("Bağla", key="vp_link_btn", type="primary", width="stretch"):
                try:
                    db.link_parent(USER["id"], vp_code)
                    st.rerun()
                except ValueError as exc:
                    st.error(str(exc))
        st.stop()

    this_week = db.weekly_stats(LIVE_STUDENT_ID)
    last_week = db.weekly_stats(LIVE_STUDENT_ID, date.today() - timedelta(days=7))
    child_prog = user_progress(LIVE_STUDENT_ID)
    v1, v2, v3, v4 = st.columns(4)
    v1.metric("Çalışma serisi", f"{child_prog['streak']} gün")
    v2.metric(
        "Bu hafta çözülen", f"{this_week['weekly_solved']} soru",
        f"{this_week['weekly_solved'] - last_week['weekly_solved']:+d} geçen haftaya göre",
    )
    v3.metric(
        "Doğruluk", f"%{this_week['accuracy']}",
        f"{this_week['accuracy'] - last_week['accuracy']:+.1f} puan" if last_week["weekly_solved"] else None,
    )
    v4.metric("Çalıştığı gün", f"{this_week['active_days']} / 7")

    st.markdown("#### Bu Hafta Ne Değişti?")
    w_c1, w_c2, w_c3 = st.columns(3)
    w_c1.markdown(
        f"""
        <div class="kp-card">
            <div style="font-size:0.72rem;font-weight:700;color:#2563EB;">GÜÇLÜ KAZANIM</div>
            <div style="font-weight:800;color:#1E2952;margin-top:0.25rem;">{html.escape(str(stats_now['strong_topic']))}</div>
            <div style="font-size:0.8rem;color:#64748B;margin-top:0.2rem;">İsabet oranı istikrarlı biçimde yükseldi.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    w_c2.markdown(
        f"""
        <div class="kp-card">
            <div style="font-size:0.72rem;font-weight:700;color:#D97706;">RİSKLİ KAZANIM</div>
            <div style="font-weight:800;color:#1E2952;margin-top:0.25rem;">{html.escape(str(stats_now['weak_topic']))}</div>
            <div style="font-size:0.8rem;color:#64748B;margin-top:0.2rem;">Hata Defteri ve telafi seti ile takip ediliyor.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    w_c3.markdown(
        f"""
        <div class="kp-card">
            <div style="font-size:0.72rem;font-weight:700;color:#059669;">SIRADAKİ DENEME</div>
            <div style="font-weight:800;color:#1E2952;margin-top:0.25rem;">{html.escape(str(stats_now['next_exam_date']))}</div>
            <div style="font-size:0.8rem;color:#64748B;margin-top:0.2rem;">{html.escape(str(stats_now['next_exam_name']))}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    g_col1, g_col2 = st.columns(2)
    with g_col1:
        st.markdown("#### Son 7 gün çözülen soru")
        render_week_chart(LIVE_STUDENT_ID, height=210)
    with g_col2:
        st.markdown("#### Son denemeler (doğruluk %)")
        recent = list(reversed(db.recent_exam_results(LIVE_STUDENT_ID, 8)))
        if recent:
            import pandas as pd
            df_r = pd.DataFrame(
                {"Deneme": [f"{i + 1}. ({r['created_at'][5:10]})" for i, r in enumerate(recent)],
                 "Doğruluk": [r["accuracy"] for r in recent]}
            )
            df_r["Deneme"] = pd.Categorical(df_r["Deneme"], categories=list(df_r["Deneme"]), ordered=True)
            st.bar_chart(df_r, x="Deneme", y="Doğruluk", color="#1D4ED8", height=210)
        else:
            st.caption("Henüz tamamlanmış deneme yok.")

    st.markdown("#### Rozetler")
    render_badge_grid(db.user_badges(LIVE_STUDENT_ID))

    # Haftalık raporu e-postayla gönder
    st.markdown("#### Haftalık rapor")
    try:
        smtp_cfg = dict(st.secrets.get("smtp", {}))
    except Exception:
        smtp_cfg = {}
    _teachers = db.student_teachers(LIVE_STUDENT_ID)
    report_body = reports.weekly_report_text(
        (_teachers[0].get("org_name") if _teachers else None) or BRAND_NAME,
        USER["name"], stats_now["student_name"], this_week, child_prog["streak"], [],
        str(stats_now.get("weak_topic", "")),
        teacher_line=", ".join(f"{t.get('signature') or t['name']} ({t['class_name']})" for t in _teachers),
    )
    with st.expander("Raporu önizle"):
        st.text(report_body)
    rp1, rp2 = st.columns(2)
    with rp1:
        if st.button("Raporu e-postama gönder", key="send_weekly_report", type="primary",
                     width="stretch", disabled=not reports.smtp_ready(smtp_cfg)):
            try:
                reports.send_email(smtp_cfg, USER["email"], f"{BRAND_NAME} • {stats_now['student_name']} haftalık özet", report_body)
                db.log_report_sent(USER["id"], db.week_key())
                st.success(f"Rapor {USER['email']} adresine gönderildi.")
            except Exception as exc:
                st.error(f"E-posta gönderilemedi: {exc}")
    with rp2:
        st.link_button(
            "WhatsApp ile paylaş",
            "https://wa.me/?text=" + urllib.parse.quote(report_body),
            width="stretch",
        )
    if not reports.smtp_ready(smtp_cfg):
        st.caption("E-posta gönderimi için secrets.toml'a [smtp] bölümü eklenmeli. Her pazar otomatik gönderim için weekly_report.py betiğini zamanlayın.")

    render_parent_notification_center()

# ==============================================================================
# 6. ÖĞRENCİ RÖNTGENİ
# ==============================================================================
elif current_page == "Öğrenci Röntgeni":
    render_student_xray_panel()

# ==============================================================================
# 7. B2B LİSANS & KURUM (+ DENEME TAKVİMİ)
# ==============================================================================
elif current_page == "B2B Lisans & Kurum":
    tab_b2b_form, tab_b2b_zumre = st.tabs(["Kurumsal Lisans Teklif Stüdyosu", "Zümre Soru Dağıtımı & Kurumsal Rapor"])

    with tab_b2b_form:
        with st.form("b2b_inquiry_form"):
            c1, c2 = st.columns(2)
            with c1:
                org_name = st.text_input("Kurum / Ofis Adı", value="Kurşun Kalem Akademi Ofisi")
                city = st.text_input("İl / İlçe", value="İstanbul")
                teacher_seats = st.number_input("Öğretmen / Editör Koltuk Sayısı", min_value=1, max_value=500, value=15)
            with c2:
                student_seats = st.number_input("Öğrenci Hesap Sayısı", min_value=10, max_value=10000, value=250)
                contact_name = st.text_input("Yetkili Adı Soyadı")
                contact_email = st.text_input("Kurumsal E-posta")

            submitted = st.form_submit_button("2027 Kurumsal Lisans Teklifi Hazırla", type="primary")
            if submitted:
                if not org_name.strip():
                    st.warning("Lütfen kurum adını girin.")
                else:
                    st.session_state["last_b2b_summary"] = (
                        f"Dönem: 2027 | Kurum: {org_name} ({city}) | Yetkili: {contact_name or '-'} | "
                        f"E-posta: {contact_email or '-'} | Koltuk: {teacher_seats} Öğretmen, "
                        f"{student_seats} Öğrenci"
                    )
                    st.session_state["b2b_org_cache"] = {
                        "org": org_name,
                        "teachers": int(teacher_seats),
                        "students": int(student_seats),
                    }
                    record_activity("B2B Lisans", f"{org_name} teklif taslağı hazırlandı")
                    st.success(
                        f"**{org_name}** ({city}) için 2027 dönemi **{teacher_seats} öğretmen** ve "
                        f"**{student_seats} öğrenci** kapasiteli kurumsal lisans talebi kaydedildi."
                    )

        if st.session_state.get("last_b2b_summary"):
            cache_for_doc = st.session_state.get("b2b_org_cache", {"org": org_name, "students": int(student_seats)})
            b2b_tier = get_tier_by_id("b2b_kurs" if cache_for_doc.get("students", 250) <= 300 else "b2b_kampus")
            b2b_summary = {
                "name": b2b_tier["name"],
                "total_count": 1,
                "monthly_price": b2b_tier["monthly_price"],
                "monthly_quota": b2b_tier["monthly_quota"],
                "coach_quota": b2b_tier["coach_quota"],
                "entries": [
                    {
                        "tier": b2b_tier,
                        "qty": 1,
                        "line_price": b2b_tier["monthly_price"],
                        "line_quota": b2b_tier["monthly_quota"],
                        "line_coach": b2b_tier["coach_quota"],
                    }
                ],
            }
            b2b_docx = build_commercial_brief_docx(
                title=f"{cache_for_doc.get('org', 'Kurumsal')} 2027 Lisans Teklif Özeti",
                summary=b2b_summary,
                billing_cycle="Yıllık (%20 İndirim)",
                installment=6,
                total_amount=round(b2b_tier["monthly_price"] * 12 * 0.80),
                extra_notes=st.session_state["last_b2b_summary"],
            )
            st.download_button(
                label="Kurumsal Lisans Teklifini DOCX İndir",
                data=b2b_docx,
                file_name="kursun_kalem_ofis_2027_kurumsal_teklif.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

    with tab_b2b_zumre:
        org_cache = st.session_state.get("b2b_org_cache", {"org": "Kurşun Kalem Akademi Ofisi", "teachers": 15, "students": 250})
        recommended_tier = get_tier_by_id("b2b_kurs" if org_cache["students"] <= 300 else "b2b_kampus")
        total_q = recommended_tier["monthly_quota"]

        bk1, bk2, bk3, bk4 = st.columns(4)
        bk1.metric("Kurum Profili", org_cache["org"])
        bk2.metric("Aylık Soru Havuzu", f"{total_q:,}".replace(",", "."))
        bk3.metric("Öğretmen Başına Kota", f"{total_q // max(1, org_cache['teachers']):,} Soru".replace(",", "."))
        bk4.metric("Öğrenci Başına Test", f"{total_q // max(1, org_cache['students'])} Soru / Ay")

        st.markdown("#### Branş Zümreleri Arası Soru Kotası Dağılımı")
        dist_ratios = {
            "Türkçe & Edebiyat Zümresi": 0.28,
            "Matematik & Geometri Zümresi": 0.30,
            "Fen Bilimleri Zümresi": 0.22,
            "Sosyal Bilimler & İnkılap Zümresi": 0.12,
            "Yabancı Dil (İngilizce) Zümresi": 0.08,
        }
        zumre_rows = []
        chart_map: Dict[str, int] = {}
        for z_name, ratio in dist_ratios.items():
            allocated = int(round(total_q * ratio))
            short_lbl = z_name.split(" ")[0]
            chart_map[short_lbl] = allocated
            zumre_rows.append(
                {
                    "Zümre Birimi": z_name,
                    "Pay (%)": f"%{int(ratio * 100)}",
                    "Aylık Soru Kotası": f"{allocated:,}".replace(",", "."),
                    "Haftalık Deneme Hedefi": f"{max(1, allocated // 80)} Kitapçık",
                    "Redaksiyon Standardı": "Maarif Modeli • Çift Kontrol",
                }
            )

        zc1, zc2 = st.columns([3, 2])
        with zc1:
            st.dataframe(zumre_rows, width="stretch", hide_index=True)
        with zc2:
            st.bar_chart(chart_map, color="#2563EB", height=225)

        st.markdown("#### Kurum İçi Deneme & Kazanım Tarama Planı")
        st.dataframe(
            [
                {"Hafta": "1. Hafta", "Kademe": "8. Sınıf (LGS)", "Odak": "Sözel-Sayısal Muhakeme Tarama", "Durum": "Planlandı"},
                {"Hafta": "2. Hafta", "Kademe": "5–7. Sınıf", "Odak": "Maarif Modeli Ara Kazanım Testi", "Durum": "Planlandı"},
                {"Hafta": "3. Hafta", "Kademe": "9–12. Sınıf (YKS)", "Odak": "TYT-AYT Hız & Çeldirici Denemesi", "Durum": "Taslak"},
            ],
            width="stretch",
            hide_index=True,
        )

# ==============================================================================
# 8. PAKETLER & ABONELİK VİTRİNİ
# ==============================================================================
elif current_page == "Paketler & Abonelik Vitrini":
    total_pkg_count = len(SUBSCRIPTION_TIERS)

    st.markdown(
        """
        <div class="kp-value-strip">
            <div class="kp-value-strip-item"><span class="kp-value-dot"></span> Türkiye Yüzyılı Maarif Modeli 2027 Soru Havuzu</div>
            <div class="kp-value-strip-item"><span class="kp-value-dot"></span> Dijimentor Haftalık Akıllı Çalışma Planı</div>
            <div class="kp-value-strip-item"><span class="kp-value-dot"></span> iyzico Güvencesiyle 12 Taksite Varan Ödeme</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    current_bag = get_bag_items()
    active_bag_summary = render_school_backpack_panel()

    st.markdown("#### Segment Filtresi")
    selected_segment = st.segmented_control(
        "Segment Filtresi",
        options=["Tümü", "Öğrenci", "YKS", "Veli", "Kurum / B2B"],
        default="Tümü",
        label_visibility="collapsed",
    )

    filtered_tiers = (
        SUBSCRIPTION_TIERS
        if not selected_segment or selected_segment == "Tümü"
        else [t for t in SUBSCRIPTION_TIERS if t["segment"] == selected_segment]
    )

    for row_start in range(0, len(filtered_tiers), 3):
        cols = st.columns(3)
        for idx, tier in enumerate(filtered_tiers[row_start : row_start + 3]):
            with cols[idx]:
                in_bag_qty = current_bag.get(tier["id"], 0)
                bag_tag = f" • Çantada: {in_bag_qty}" if in_bag_qty > 0 else ""
                is_featured = tier["badge"] in ("En Çok Tercih Edilen", "VIP", "Kampüs")
                card_cls = "kp-pkg-card kp-pkg-featured" if is_featured else "kp-pkg-card"
                feats_li = "".join([f"<li>• {html.escape(feat)}</li>" for feat in tier["features"]])
                st.markdown(
                    f"""
                    <div class="{card_cls}">
                        <div class="kp-pkg-top">
                            <span class="kp-pkg-seg">{html.escape(tier["segment"])}</span>
                            <span class="kp-pkg-badge">{html.escape(tier["badge"])}{html.escape(bag_tag)}</span>
                        </div>
                        <div class="kp-pkg-title">{html.escape(tier["name"])}</div>
                        <div class="kp-pkg-price">
                            {tier["monthly_price"]:,} ₺ <span>/ ay</span>
                        </div>
                        <div class="kp-pkg-desc">{html.escape(tier["desc"])}</div>
                        <ul class="kp-pkg-feats">{feats_li}</ul>
                    </div>
                    """.replace(",", "."),
                    unsafe_allow_html=True,
                )
                st.button(
                    "Çantaya Ekle",
                    key=f"select_{tier['id']}",
                    type="primary" if in_bag_qty > 0 else "secondary",
                    width="stretch",
                    on_click=cb_add_tier_to_bag,
                    args=(tier["id"], 1),
                )

    st.markdown("---")
    st.markdown(f"#### {total_pkg_count} Paket Karşılaştırma Matrisi")
    matrix_rows = [
        {
            "Paket Adı": t["name"],
            "Segment": t["segment"],
            "Etiket": t["badge"],
            "Çantadaki Adet": current_bag.get(t["id"], 0),
            "Aylık Soru Kotası": f"{t['monthly_quota']:,}".replace(",", "."),
            "Dijimentor Plan": f"{t['coach_quota']}/ay",
            "Aylık Ücret (₺)": f"{t['monthly_price']:,}".replace(",", "."),
            "Öne Çıkan Kapsam": t["features"][0],
        }
        for t in SUBSCRIPTION_TIERS
    ]
    st.dataframe(matrix_rows, width="stretch", hide_index=True)

    st.markdown("---")
    render_iyzico_simulator(
        package_name=active_bag_summary["name"],
        monthly_price=active_bag_summary["monthly_price"],
        summary=active_bag_summary,
    )

# ==============================================================================
# FOTOĞRAFLA ÇÖZ (öğrenci)
# ==============================================================================
elif current_page == "Fotoğrafla Çöz":
    ph1, ph2 = st.columns([1.2, 1])
    with ph1:
        src_mode = st.segmented_control("Kaynak", ["Dosya yükle", "Kamera"], default="Dosya yükle",
                                        key="photo_src", label_visibility="collapsed")
        img_file = (
            st.camera_input("Soruyu çek", key="photo_cam")
            if src_mode == "Kamera"
            else st.file_uploader("Soru fotoğrafı", type=["png", "jpg", "jpeg", "webp"], key="photo_up")
        )
    with ph2:
        photo_mode = st.radio("Yardım düzeyi", ["Önce ipucu ver", "Adım adım çöz"], key="photo_mode")
        photo_note = st.text_input("Not (isteğe bağlı)", placeholder="Örn: C dedim ama yanlış çıktı", key="photo_note")
    if img_file is not None:
        st.image(img_file, width=360)
    if st.button("Çözümü göster", type="primary", key="photo_solve_btn", disabled=img_file is None):
        img_bytes = img_file.getvalue()
        if len(img_bytes) > 8 * 1024 * 1024:
            st.error("Fotoğraf 8 MB'tan büyük olamaz.")
        else:
            ok_q, q_msg = consume_quota("question", 1)
            if not ok_q:
                st.error(q_msg)
            else:
                slot = st.empty()
                render_loading_skeleton(slot, lines=3)
                try:
                    answer = call_gemini_vision(
                        img_bytes,
                        getattr(img_file, "type", None) or "image/jpeg",
                        (
                            f"Fotoğraftaki soruyu oku. Öğrencinin sınıfı: {USER.get('grade') or 'belirtilmedi'}. "
                            f"Yardım düzeyi: {photo_mode}. Öğrencinin notu: {photo_note or '-'}.\n"
                            "Önce soruyu kısaca yeniden yaz, sonra yardım düzeyine uygun şekilde açıkla. "
                            "Soru okunamıyorsa bunu söyle ve daha net bir fotoğraf iste."
                        ),
                        (
                            "Sen sabırlı, teşvik edici bir Türk eğitim mentorusun. LGS/YKS müfredatına hâkimsin. "
                            "'Önce ipucu ver' seçildiyse cevabı söyleme, sadece yol göster. Türkçe yanıt ver."
                        ),
                    )
                    slot.empty()
                    st.session_state["photo_solves"] = int(st.session_state.get("photo_solves", 0)) + 1
                    save_persisted_state()
                    record_activity("Fotoğrafla Çöz", f"{photo_mode} üretildi")
                    st.markdown("#### Çözüm")
                    st.markdown(answer)
                    check_and_award_badges()
                except Exception as exc:
                    slot.empty()
                    refund_quota("question", 1)
                    st.error(f"Çözüm üretilemedi: {exc}")

# ==============================================================================
# 9. SINIFLARIM (ÖĞRETMEN): sınıf oluşturma, ödevler, sınıf karşılaştırma paneli
# ==============================================================================
elif current_page == "Sınıflarım":
    my_classes = db.teacher_classes(USER["id"])

    with st.expander("Yeni sınıf oluştur", expanded=not my_classes, icon=":material/add:"):
        nc1, nc2, nc3 = st.columns([2, 2, 1], vertical_alignment="bottom")
        with nc1:
            new_cls_name = st.text_input("Sınıf adı", placeholder="Örn: 8-A Türkçe", key="new_cls_name")
        with nc2:
            new_cls_grade = st.selectbox("Sınıf düzeyi", GRADE_LEVELS_2027, index=3, key="new_cls_grade")
        with nc3:
            if st.button("Oluştur", key="new_cls_btn", type="primary", width="stretch"):
                try:
                    created = db.create_class(USER["id"], new_cls_name, new_cls_grade)
                    st.session_state["teacher_class_sel"] = created["id"]
                    record_activity("Sınıflarım", f"{created['name']} oluşturuldu")
                    st.rerun()
                except ValueError as exc:
                    st.error(str(exc))

    if not my_classes:
        st.info("Henüz sınıfınız yok. Yukarıdan ilk sınıfınızı oluşturun.")
        st.stop()

    cls_by_id = {c["id"]: c for c in my_classes}
    default_cls = st.session_state.get("teacher_class_sel")
    cls_ids = list(cls_by_id.keys())
    sel_id = st.selectbox(
        "Sınıf",
        cls_ids,
        index=cls_ids.index(default_cls) if default_cls in cls_ids else 0,
        format_func=lambda cid: f"{cls_by_id[cid]['name']} · {cls_by_id[cid]['member_count']} öğrenci",
        key="teacher_class_pick",
    )
    st.session_state["teacher_class_sel"] = sel_id
    cls = cls_by_id[sel_id]

    st.markdown(
        f'<div class="kp-banner kp-banner-hw">Öğrencileriniz <b>Öğrenci Portalı → Ödevlerim</b> sekmesinde '
        f'bu kodla katılır: <span class="kp-code">{cls["code"]}</span></div>',
        unsafe_allow_html=True,
    )

    t_over, t_assign, t_members, t_parents = st.tabs(["Sınıf analizi", "Ödevler", "Öğrenciler", "Veli raporları"])

    with t_over:
        ov = db.class_overview(sel_id)
        subs_rows = [r for r in ov["by_student"] if r["avg_accuracy"] is not None]
        o1, o2, o3 = st.columns(3)
        o1.metric("Öğrenci", cls["member_count"])
        o2.metric("Ödev", ov["total_assignments"])
        o3.metric(
            "Sınıf ortalaması",
            f"%{round(sum(r['avg_accuracy'] for r in subs_rows) / len(subs_rows))}" if subs_rows else "—",
        )
        if not ov["by_outcome"]:
            st.caption("Ödev teslimleri geldikçe kazanım bazlı karşılaştırma burada oluşur.")
        else:
            import pandas as pd

            st.markdown("#### Kazanım bazlı sınıf başarısı")
            st.caption("En düşükten yükseğe: üstteki kazanımlar sınıfça tekrar edilmeli.")
            df_o = pd.DataFrame(ov["by_outcome"])
            df_o["Kazanım"] = df_o["outcome"].str.slice(0, 60)
            df_o["Kazanım"] = pd.Categorical(df_o["Kazanım"], categories=list(dict.fromkeys(df_o["Kazanım"])), ordered=True)
            st.bar_chart(df_o, x="Kazanım", y="avg_accuracy", color="#1D4ED8", height=260, horizontal=True)
            weakest = ov["by_outcome"][0]
            wc1, wc2 = st.columns([3, 1], vertical_alignment="center")
            with wc1:
                st.markdown(f"En zayıf kazanım: **{weakest['outcome']}** (%{weakest['avg_accuracy']})")
            with wc2:
                if st.button("Telafi testi üret", key="cls_remedial", type="primary", width="stretch"):
                    st.session_state["factory_focus_topic"] = weakest["outcome"]
                    st.session_state["fab_theme_inp"] = weakest["outcome"]
                    if weakest.get("branch"):
                        st.session_state["factory_branch"] = weakest["branch"]
                    if cls.get("grade"):
                        st.session_state["factory_grade"] = cls["grade"]
                        st.session_state["factory_grade_sync"] = True
                    st.session_state["factory_sync_pending"] = True
                    navigate_to("Soru & Redaksiyon Fabrikası")

        st.markdown("#### Öğrenci bazlı durum")
        st.dataframe(
            [
                {
                    "Öğrenci": r["name"],
                    "Teslim": f"{r['done']} / {ov['total_assignments']}",
                    "Ort. doğruluk (%)": r["avg_accuracy"] if r["avg_accuracy"] is not None else "—",
                    "Ort. net": r["avg_net"] if r["avg_net"] is not None else "—",
                }
                for r in ov["by_student"]
            ],
            width="stretch",
            hide_index=True,
        )

    with t_assign:
        st.caption("Ödevler yalnızca onayladığınız sorulardan oluşur: **Soru Bankam → Onaylı sorular** sekmesinden seçip atayın.")
        ac1, ac2, _ = st.columns([1, 1, 2])
        with ac1:
            if st.button("Soru Bankam'a git", key="cls_to_bank", icon=":material/inventory_2:", width="stretch"):
                navigate_to("Soru Bankam")
        with ac2:
            if st.button("Yeni soru üret", key="cls_to_factory", icon=":material/edit_note:", width="stretch"):
                if cls.get("grade"):
                    st.session_state["factory_grade"] = cls["grade"]
                    st.session_state["factory_grade_sync"] = True
                navigate_to("Soru & Redaksiyon Fabrikası")
        assignments = db.class_assignments(sel_id)
        if not assignments:
            st.info("Bu sınıfa henüz ödev atanmadı.")
        for a in assignments:
            due_txt = datetime.fromisoformat(a["due_date"]).strftime("%d.%m.%Y") if a.get("due_date") else "—"
            with st.expander(f"{a['title']} · {a['submission_count']}/{cls['member_count']} teslim · son gün {due_txt}"):
                res = db.assignment_results(a["id"])
                st.dataframe(
                    [
                        {
                            "Öğrenci": r["name"],
                            "Durum": "Teslim etti" if r["submitted_at"] else "Bekliyor",
                            "D/Y/B": f"{r['correct']}/{r['wrong']}/{r['blank']}" if r["submitted_at"] else "—",
                            "Net": r["net"] if r["submitted_at"] else "—",
                            "Doğruluk (%)": r["accuracy"] if r["submitted_at"] else "—",
                        }
                        for r in res
                    ],
                    width="stretch",
                    hide_index=True,
                )
                if a.get("video_url"):
                    st.caption(f"Konu anlatım videosu: {a['video_url']}")
                booklet_download(a["content"], a["title"], f"{cls['name']} · {a.get('branch') or ''}", f"as_{a['id']}")

    with t_members:
        members = db.class_members(sel_id)
        with st.expander("Toplu öğrenci ekle (liste ya da Excel/CSV)", icon=":material/group_add:",
                         expanded=not members):
            st.caption(
                "Her satıra bir öğrenci yazın: **Ad Soyad** ya da **Ad Soyad, e-posta**. E-postası olmayan öğrenciye "
                "giriş adresi ve geçici şifre üretilir; listeyi indirip öğrencilere dağıtın."
            )
            bulk_txt = st.text_area("Öğrenci listesi", height=140, key=f"bulk_txt_{sel_id}",
                                    placeholder="Elif Öztürk\nEmir Demir, emir@ornek.com")
            bulk_file = st.file_uploader("ya da dosya yükleyin (1. sütun ad soyad, 2. sütun e-posta)",
                                         type=["csv", "xlsx"], key=f"bulk_file_{sel_id}")
            bulk_ok = st.checkbox(
                "Bu öğrencilerin (ve 18 yaş altıysa velilerinin) KVKK kapsamında bilgilendirildiğini, hesapları "
                "kurumum adına açtığımı beyan ederim.",
                key=f"bulk_ok_{sel_id}",
            )
            if st.button("Hesapları oluştur ve sınıfa ekle", key=f"bulk_btn_{sel_id}", type="primary",
                         disabled=not bulk_ok):
                rows = parse_student_list(bulk_txt, bulk_file)
                if not rows:
                    st.error("Listede öğrenci bulunamadı.")
                elif len(rows) > 200:
                    st.error("Tek seferde en fazla 200 öğrenci eklenebilir.")
                else:
                    with st.spinner(f"{len(rows)} hesap hazırlanıyor..."):
                        st.session_state[f"bulk_result_{sel_id}"] = db.bulk_create_students(
                            USER["id"], sel_id, rows, cls.get("grade"))
                    record_activity("Sınıflarım", f"{cls['name']}: {len(rows)} öğrenci toplu eklendi")
                    st.rerun()
            bulk_res = st.session_state.get(f"bulk_result_{sel_id}")
            if bulk_res:
                made = [r for r in bulk_res if r["password"]]
                st.success(f"{len(made)} yeni hesap açıldı, {len(bulk_res) - len(made)} satır ayrıca işlendi.")
                st.dataframe(
                    [{"Öğrenci": r["name"], "Giriş e-postası": r["email"], "Geçici şifre": r["password"] or "—",
                      "Durum": r["status"]} for r in bulk_res],
                    width="stretch", hide_index=True,
                )
                csv_buf = io.StringIO()
                csv_buf.write("Ad Soyad;Giriş e-postası;Geçici şifre;Durum\n")
                for r in bulk_res:
                    csv_buf.write(f"{r['name']};{r['email']};{r['password']};{r['status']}\n")
                st.download_button(
                    "Giriş bilgilerini indir (Excel'de açılır)",
                    data=csv_buf.getvalue().encode("utf-8-sig"),
                    file_name=f"{cls['name']}_giris_bilgileri.csv".replace(" ", "_"),
                    mime="text/csv", key=f"bulk_dl_{sel_id}", icon=":material/download:",
                )
                st.caption("Şifreler yalnızca şimdi görünür; güvenli biçimde dağıttıktan sonra bu listeyi silin.")
                if st.button("Listeyi kapat", key=f"bulk_close_{sel_id}"):
                    st.session_state.pop(f"bulk_result_{sel_id}", None)
                    st.rerun()

        if not members:
            st.info(f"Henüz öğrenci yok. Sınıf kodunu paylaşın: {cls['code']} ya da yukarıdan toplu ekleyin.")
        else:
            st.dataframe(
                [
                    {
                        "Öğrenci": m["name"],
                        "Giriş": m["email"],
                        "Sınıf": m.get("grade") or "—",
                        "Seri (gün)": user_progress(m["id"])["streak"],
                        "Bu hafta soru": db.weekly_stats(m["id"])["weekly_solved"],
                    }
                    for m in members
                ],
                width="stretch",
                hide_index=True,
            )
            gen_members = [m for m in members if m["email"].endswith(db.GENERATED_EMAIL_DOMAIN)]
            if gen_members:
                with st.popover("Öğrenci şifresini sıfırla"):
                    st.caption("Yalnızca sizin açtığınız (kendi e-postası olmayan) hesaplar. "
                               "Kendi e-postası olan öğrenci girişteki 'Şifremi unuttum' bölümünü kullanır.")
                    rs_map = {f"{m['name']} ({m['email']})": m["id"] for m in gen_members}
                    rs_sel = st.selectbox("Öğrenci", list(rs_map.keys()), key=f"rs_sel_{sel_id}")
                    if st.button("Yeni geçici şifre üret", key=f"rs_btn_{sel_id}", type="primary"):
                        try:
                            new_pw = db.teacher_reset_student_password(USER["id"], rs_map[rs_sel])
                            st.success(f"Yeni şifre: **{new_pw}** — öğrenciye iletin; Profilim'den değiştirebilir.")
                        except ValueError as exc:
                            st.error(str(exc))
            with st.popover("Öğrenciyi sınıftan çıkar"):
                rm_map = {f"{m['name']} ({m['email']})": m["id"] for m in members}
                rm_sel = st.selectbox("Öğrenci", list(rm_map.keys()), key=f"rm_sel_{sel_id}")
                if st.button("Çıkar", key=f"rm_btn_{sel_id}", type="primary"):
                    db.remove_member(USER["id"], sel_id, rm_map[rm_sel])
                    st.rerun()

    with t_parents:
        parents = db.class_parents(sel_id)
        org_name, signature = teacher_brand()
        st.caption(
            f"Raporlar **{signature} · {org_name}** imzasıyla gider. İmza ve kurum adını **Profilim** sayfasından "
            "değiştirebilirsiniz. Veli, çocuğunun Profil sayfasındaki bağlantı koduyla hesabını bağlamalıdır."
        )
        if not parents:
            st.info("Bu sınıfta hesabını bağlamış veli yok henüz.")
        else:
            try:
                smtp_cfg_t = dict(st.secrets.get("smtp", {}))
            except Exception:
                smtp_cfg_t = {}
            reports_by_parent = []
            for pr in parents:
                wk = db.weekly_stats(pr["student_id"])
                body = reports.weekly_report_text(
                    org_name, pr["parent_name"], pr["student_name"], wk,
                    user_progress(pr["student_id"])["streak"], [], "",
                    teacher_line=f"{signature} ({cls['name']})",
                )
                reports_by_parent.append((pr, wk, body))
            st.dataframe(
                [{"Öğrenci": pr["student_name"], "Veli": pr["parent_name"], "Bu hafta soru": wk["weekly_solved"],
                  "Doğruluk (%)": wk["accuracy"], "Rapor bu hafta": "Gönderildi" if db.report_already_sent(pr["parent_id"], db.week_key()) else "—"}
                 for pr, wk, _ in reports_by_parent],
                width="stretch", hide_index=True,
            )
            if st.button(f"Tüm velilere e-posta gönder ({len(parents)})", key=f"mail_all_{sel_id}", type="primary",
                         disabled=not reports.smtp_ready(smtp_cfg_t)):
                sent, failed = 0, 0
                for pr, _, body in reports_by_parent:
                    if db.report_already_sent(pr["parent_id"], db.week_key()):
                        continue
                    try:
                        reports.send_email(smtp_cfg_t, pr["parent_email"],
                                           f"{org_name} • {pr['student_name']} haftalık özet", body)
                        db.log_report_sent(pr["parent_id"], db.week_key())
                        sent += 1
                    except Exception:
                        failed += 1
                st.success(f"{sent} rapor gönderildi." + (f" {failed} gönderim başarısız." if failed else ""))
            if not reports.smtp_ready(smtp_cfg_t):
                st.caption("Toplu e-posta için Secrets'a [smtp] bölümü eklenmeli. O zamana kadar WhatsApp ile tek tek gönderebilirsiniz.")
            for pr, _, body in reports_by_parent:
                with st.expander(f"{pr['student_name']} → {pr['parent_name']}"):
                    st.text(body)
                    st.link_button("WhatsApp ile gönder", "https://wa.me/?text=" + urllib.parse.quote(body))

# ==============================================================================
# 11. SORU BANKAM (ÖĞRETMEN: TASLAK → ONAY → ÖDEV)
# ==============================================================================
elif current_page == "Soru Bankam":
    if st.session_state.get("bank_flash"):
        st.success(st.session_state.pop("bank_flash"))
    counts = db.bank_counts(USER["id"])
    flt = db.bank_filters(USER["id"])
    m1, m2, m3 = st.columns(3)
    m1.metric("Onay bekleyen", counts["taslak"])
    m2.metric("Onaylı soru", counts["onayli"])
    with m3:
        if st.button("Yeni soru üret", key="bank_to_factory", icon=":material/edit_note:", width="stretch"):
            navigate_to("Soru & Redaksiyon Fabrikası")
    f1, f2, f3 = st.columns([1, 1, 2])
    with f1:
        f_grade = st.selectbox("Sınıf düzeyi", ["Tümü"] + flt["grades"], key="bk_f_grade")
    with f2:
        f_branch = st.selectbox("Branş", ["Tümü"] + flt["branches"], key="bk_f_branch")
    with f3:
        f_search = st.text_input("Ara (soru kökü ya da kazanım)", key="bk_f_search")
    fg = None if f_grade == "Tümü" else f_grade
    fb = None if f_branch == "Tümü" else f_branch

    t_draft, t_ok, t_new = st.tabs(
        [f"Onay bekleyen ({counts['taslak']})", f"Onaylı sorular ({counts['onayli']})", "Elle soru yaz"]
    )

    with t_draft:
        drafts = db.list_bank_questions(USER["id"], "taslak", fg, fb, f_search)
        if not drafts:
            st.info("Onay bekleyen soru yok. Soru & Redaksiyon'da test üretip **Soru Bankama taslak ekle** deyin.")
        else:
            st.caption(
                "Yapay zekâ bazen yanlış cevap anahtarı ya da hatalı şık üretebilir. Her soruyu okuyup "
                "düzeltin; öğrencilere yalnızca onayladığınız sorular gider."
            )
            ready = [q["id"] for q in drafts if q.get("correct")]
            if ready:
                with st.popover(f"Cevabı işaretli {len(ready)} soruyu toplu onayla"):
                    st.markdown("Bu soruları **okuduğunuzu ve kontrol ettiğinizi** onaylıyor musunuz?")
                    if st.button("Evet, hepsini onayla", key="bank_bulk_ok", type="primary"):
                        db.set_bank_status(USER["id"], ready, "onayli")
                        st.rerun()
            for q in drafts[:25]:
                render_bank_editor(q)
            if len(drafts) > 25:
                st.caption(f"İlk 25 taslak gösteriliyor ({len(drafts)} toplam). Onayladıkça sıradakiler gelir.")

    with t_ok:
        approved = db.list_bank_questions(USER["id"], "onayli", fg, fb, f_search)
        if not approved:
            st.info("Henüz onaylı soru yok. Taslakları kontrol edip onaylayın ya da **Elle soru yaz** sekmesini kullanın.")
        else:
            shown = approved[:100]
            sa1, sa2, _ = st.columns([1, 1, 2])
            with sa1:
                if st.button("Görünenlerin tümünü seç", key="bank_sel_all", width="stretch"):
                    for q in shown:
                        st.session_state[f"bsel_{q['id']}"] = True
            with sa2:
                if st.button("Seçimi temizle", key="bank_sel_none", width="stretch"):
                    for q in shown:
                        st.session_state[f"bsel_{q['id']}"] = False
            for q in shown:
                ck, body = st.columns([0.05, 0.95], vertical_alignment="top")
                with ck:
                    st.checkbox("Seç", key=f"bsel_{q['id']}", label_visibility="collapsed")
                with body:
                    with st.expander(f"#{q['id']} · {(q.get('branch') or '')} · {q['stem'][:90]}"):
                        if q.get("outcome"):
                            st.caption(f"Kazanım: {q['outcome']}")
                        st.markdown(q["stem"])
                        st.markdown("\n".join(
                            f"- **{k})** {v}" + ("  ✓" if k == q.get("correct") else "") for k, v in q["options"].items()
                        ))
                        if q.get("explanation"):
                            st.caption(f"Çözüm notu: {q['explanation']}")
                        if st.button("Düzenlemek için taslağa al", key=f"bq_back_{q['id']}"):
                            db.set_bank_status(USER["id"], [q["id"]], "taslak")
                            st.rerun()
            if len(approved) > 100:
                st.caption(f"İlk 100 soru gösteriliyor ({len(approved)} toplam). Filtreleri kullanın.")

            chosen = [q for q in shown if st.session_state.get(f"bsel_{q['id']}")]
            st.markdown("---")
            if not chosen:
                st.caption("Ödev ya da kitapçık için yukarıdan soru seçin.")
            else:
                st.markdown(f"**Seçili {len(chosen)} soru**")
                sel_text = db.questions_to_text(chosen)
                c_grade, c_branch, c_outcome = _bank_meta(chosen)
                org_name, _sig = teacher_brand()
                x1, x2, x3 = st.columns(3)
                with x1:
                    render_class_assign_form(sel_text, c_grade, c_branch, c_outcome, "bank")
                with x2:
                    booklet_download(sel_text, f"{org_name} Deneme", f"{c_grade} · {c_branch}", "bank")
                with x3:
                    with st.popover("Seçilenleri sil", width="stretch"):
                        st.markdown(f"{len(chosen)} soru kalıcı olarak silinecek.")
                        if st.button("Sil", key="bank_del_sel", type="primary"):
                            db.delete_bank_questions(USER["id"], [q["id"] for q in chosen])
                            for q in chosen:
                                st.session_state.pop(f"bsel_{q['id']}", None)
                            st.rerun()

    with t_new:
        st.caption("Kendi yazdığınız sorular doğrudan onaylı olarak bankaya eklenir.")
        with st.form("bank_manual", clear_on_submit=True):
            n1, n2 = st.columns(2)
            with n1:
                mn_grade = st.selectbox("Sınıf düzeyi", GRADE_LEVELS_2027, index=3, key="mn_grade")
            with n2:
                mn_branch = st.text_input("Branş", placeholder="Örn: Matematik", key="mn_branch")
            mn_outcome = st.text_input("Kazanım", key="mn_outcome")
            mn_stem = st.text_area("Soru kökü", height=120, key="mn_stem")
            mo_cols = st.columns(2)
            mn_opts: Dict[str, str] = {}
            for i, letter in enumerate(["A", "B", "C", "D"]):
                with mo_cols[i % 2]:
                    mn_opts[letter] = st.text_input(f"{letter})", key=f"mn_{letter}")
            mn_correct = st.radio("Doğru cevap", ["A", "B", "C", "D"], index=None, horizontal=True, key="mn_correct")
            mn_expl = st.text_area("Çözüm notu (isteğe bağlı)", height=70, key="mn_expl")
            if st.form_submit_button("Bankaya ekle", type="primary"):
                filled = {k: v for k, v in mn_opts.items() if v.strip()}
                if not mn_stem.strip() or len(filled) < 2:
                    st.error("Soru kökü ve en az iki şık gerekli.")
                elif not mn_correct or mn_correct not in filled:
                    st.error("Doğru cevabı (dolu bir şık) işaretleyin.")
                else:
                    db.add_bank_questions(
                        USER["id"],
                        [{"stem": mn_stem, "options": filled, "correct": mn_correct, "explanation": mn_expl}],
                        mn_grade, mn_branch.strip(), mn_outcome.strip(), "", status="onayli",
                    )
                    record_activity("Soru Bankam", "Elle soru eklendi")
                    st.success("Soru onaylı olarak eklendi.")

# ==============================================================================
# 10. PROFİLİM & ROZETLER
# ==============================================================================
elif current_page == "Profilim & Rozetler":
    pf_l, pf_r = st.columns([1, 1], gap="large")
    with pf_l:
        with st.container(border=True, key="card_profile"):
            st.markdown(
                f'<h2 class="kp-section-title" style="font-size:1.3rem;">{html.escape(USER["name"])}</h2>'
                f'<p class="kp-section-sub">{html.escape(USER["email"])} · {db.ROLE_LABELS.get(ROLE, ROLE)}</p>',
                unsafe_allow_html=True,
            )
            if ROLE == "ogrenci":
                cur_grade = USER.get("grade") if USER.get("grade") in GRADE_LEVELS_2027 else GRADE_LEVELS_2027[3]
                new_grade = st.selectbox("Sınıfın", GRADE_LEVELS_2027, index=GRADE_LEVELS_2027.index(cur_grade), key="pf_grade")
                new_goal = st.slider("Günlük soru hedefin", 5, 100, int(USER.get("daily_goal", 20)), step=5, key="pf_goal")
                if st.button("Kaydet", key="pf_save", type="primary"):
                    db.update_grade(USER["id"], new_grade)
                    db.set_daily_goal(USER["id"], new_goal)
                    st.success("Profilin güncellendi.")
                    st.rerun()
                st.markdown(
                    f'<div class="kp-banner kp-banner-hw" style="margin-top:0.8rem;">Veli bağlantı kodun: '
                    f'<span class="kp-code">{USER.get("link_code") or "—"}</span><br>'
                    f'<span style="font-size:0.82rem;">Velin bu kodla kendi hesabından ilerlemeni takip edebilir.</span></div>',
                    unsafe_allow_html=True,
                )
            elif ROLE == "veli":
                child = db.get_user(int(USER["parent_of"])) if USER.get("parent_of") else None
                st.markdown(f"Bağlı öğrenci: **{child['name'] if child else 'yok'}**")
                pv_code = st.text_input("Yeni bağlantı kodu", max_chars=6, key="pf_link_code")
                if st.button("Öğrenciyi bağla / değiştir", key="pf_link_btn"):
                    try:
                        db.link_parent(USER["id"], pv_code)
                        st.rerun()
                    except ValueError as exc:
                        st.error(str(exc))
            elif ROLE == "ogretmen":
                tprof = db.get_teacher_profile(USER["id"])
                t_org = st.text_input("Kurum / okul adı", value=tprof.get("org_name") or "",
                                      placeholder="Örn: Başarı Etüt Merkezi", key="pf_org")
                t_sig = st.text_input("Rapor imzası", value=tprof.get("signature") or USER["name"],
                                      placeholder="Örn: Ayşe Yılmaz – Matematik Öğretmeni", key="pf_sig")
                st.caption("Veli raporlarında ve basılı kitapçıklarda bu bilgiler kullanılır.")
                if st.button("Kaydet", key="pf_t_save", type="primary"):
                    db.save_teacher_profile(USER["id"], t_org, t_sig)
                    st.success("Kurum ve imza bilgileri kaydedildi.")
            plan_now = get_active_plan()
            exp = plan_now.get("expires_at")
            st.caption(
                f"Aktif paket: {plan_now['name']}"
                + (f" · bitiş {datetime.fromisoformat(exp).strftime('%d.%m.%Y')}" if exp else "")
            )
    with pf_r:
        target_id = USER["id"] if ROLE != "veli" else LIVE_STUDENT_ID
        if target_id and ROLE != "ogretmen":
            prog = user_progress(int(target_id))
            st.markdown(
                f"""
                <div class="kp-streak-hero">
                    <div class="kp-streak-num">{prog['streak']}</div>
                    <div>
                        <div class="kp-mod-title">günlük seri</div>
                        <div class="kp-mod-desc" style="min-height:0;">Toplam {_fmt(prog['total_solved'])} soru çözüldü.
                        Seriyi korumak için her gün en az bir deneme çöz.</div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
    with st.expander("Şifremi değiştir", icon=":material/lock:"):
        with st.form("pw_change", border=False, clear_on_submit=True):
            pw_old = st.text_input("Mevcut şifre", type="password")
            pw_new = st.text_input("Yeni şifre (en az 8 karakter)", type="password")
            pw_new2 = st.text_input("Yeni şifre (tekrar)", type="password")
            if st.form_submit_button("Değiştir", type="primary"):
                if pw_new != pw_new2:
                    st.error("Yeni şifreler aynı değil.")
                else:
                    try:
                        db.change_password(USER["id"], pw_old, pw_new)
                        st.success("Şifreniz değişti.")
                    except ValueError as exc:
                        st.error(str(exc))
    with st.expander("KVKK Aydınlatma Metni", icon=":material/policy:"):
        render_kvkk_text()
    if ROLE != "ogretmen":
        target_id = USER["id"] if ROLE != "veli" else LIVE_STUDENT_ID
        if target_id:
            earned_now = db.user_badges(int(target_id))
            st.markdown(
                f'<h2 class="kp-section-title" style="margin-top:1.5rem;">Rozetler '
                f'<span style="font-size:1rem;color:#64748B;">({len(earned_now)}/{len(gamify.BADGES)})</span></h2>',
                unsafe_allow_html=True,
            )
            render_badge_grid(earned_now)
