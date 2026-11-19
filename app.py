import os
import re
import io
import csv
import json
import sqlite3
import secrets
import random
import urllib.request
from datetime import datetime, timedelta
from functools import wraps

from flask import (
    Flask, request, redirect, url_for, session,
    flash, render_template, abort, g, Response, make_response
)
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from docx import Document
import openpyxl


# ============================================================
# CONFIGURATION
# ============================================================

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
if os.path.exists("/data"):
    DATABASE = os.path.join("/data", "portal.db")
    UPLOAD_FOLDER = os.path.join("/data", "uploads")
else:
    DATABASE = os.path.join(BASE_DIR, "portal.db")
    UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "CHANGE-THIS-TO-A-LONG-RANDOM-SECRET-KEY-12345")
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # 25MB max file upload

SESSIONS = [
    ("8:00 AM", 1),
    ("11:00 AM", 2),
    ("2:00 PM", 3),
]

REPORT_REASONS = [
    "Course not registered",
    "Invalid matriculation number",
    "System malfunctioning",
]

MAT_PATTERN = re.compile(r"\bnou\d{9}\b", re.I)
COURSE_PATTERN = re.compile(r"\b[A-Z]{3}\s*\d{3}\b|\b[A-Z]{3}\s*[0-9OISL]{3}\b", re.I)
TIME_PATTERN = re.compile(r"\b(?:\d{4}-\d{2}-\d{2}\s+)?\d{1,2}(?::\d{2})?\s*[AP]M\b|\b\d{1,2}:\d{2}(?:\s*[AP]M)?\b", re.I)

MAT_HEAL_PATTERN = re.compile(r"\bnou[A-Z0-9]{9}\b", re.I)
COURSE_HEAL_PATTERN = re.compile(r"\b[A-Z]{3}\s*\d{3}\b|\b[A-Z]{3}\s*[0-9OISL]{3}\b", re.I)


def intelligent_cleanse_mat(raw):
    raw_lower = raw.strip().lower()
    if not raw_lower.startswith("nou"):
        m = MAT_HEAL_PATTERN.search(raw_lower)
        if m:
            raw_lower = m.group(0).lower()

    if raw_lower.startswith("nou"):
        prefix = "nou"
        digits_part = raw_lower[3:].upper()
        digits_cleaned = (
            digits_part
            .replace("O", "0")
            .replace("I", "1")
            .replace("L", "1")
            .replace("S", "5")
        )
        healed = prefix + digits_cleaned
        if re.fullmatch(r"nou\d{9}", healed):
            return healed, True, (healed != raw_lower)
    return raw_lower, False, False


def intelligent_cleanse_course(raw):
    raw = raw.strip().upper().replace(" ", "")
    m = re.search(r"([A-Z]{3})([A-Z0-9]{3})", raw)
    if m:
        letters = m.group(1)
        digits = m.group(2).replace("O", "0").replace("I", "1").replace("L", "1").replace("S", "5")
        healed = letters + digits
        if re.fullmatch(r"[A-Z]{3}\d{3}", healed):
            return healed, True, (healed != raw)
    return raw, False, False


def intelligent_session_bucket(time_str):
    """
    Intelligently maps any time string or header to one of the 3 canonical work sessions:
    - 8:00 AM
    - 11:00 AM
    - 2:00 PM
    """
    if not time_str:
        return "8:00 AM"

    ts = time_str.strip().upper().replace(" ", "")
    m = re.search(r"(\d{1,2})(?::(\d{2}))?(AM|PM)?", ts)
    if not m:
        return "8:00 AM"

    h = int(m.group(1))
    is_pm = m.group(3) == "PM"
    if is_pm and h < 12:
        h += 12
    if not is_pm and h == 12 and "AM" in ts:
        h = 0

    if h < 11:
        return "8:00 AM"
    elif 11 <= h < 14:
        return "11:00 AM"
    else:
        return "2:00 PM"


def normalize_time_key(time_str):
    if not time_str:
        return "99:99"
    ts = time_str.strip().upper().replace(" ", "")
    m_ampm = re.match(r"^(\d{1,2})(?::(\d{2}))?(AM|PM)$", ts)
    if m_ampm:
        h = int(m_ampm.group(1))
        m = m_ampm.group(2) if m_ampm.group(2) else "00"
        is_pm = m_ampm.group(3) == "PM"
        if is_pm and h < 12:
            h += 12
        if not is_pm and h == 12:
            h = 0
        return f"{h:02d}:{m}"

    m_24 = re.match(r"^(\d{1,2}):(\d{2})$", ts)
    if m_24:
        h = int(m_24.group(1))
        m = m_24.group(2)
        return f"{h:02d}:{m}"

    return ts



# ============================================================
# DATABASE MANAGEMENT (Per-request lifecycle via Flask `g`)
# ============================================================

def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(DATABASE, check_same_thread=False)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA foreign_keys = ON")
    return g.db


@app.teardown_appcontext
def close_db(exception=None):
    db = g.pop("db", None)
    if db is not None:
        db.close()


def standalone_db():
    connection = sqlite3.connect(DATABASE, check_same_thread=False)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys = ON")
    return connection


@app.after_request
def add_header(response):
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response


def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def init_db():
    con = standalone_db()
    con.executescript("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        email TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        role TEXT NOT NULL DEFAULT 'user',
        approved INTEGER NOT NULL DEFAULT 0,
        active INTEGER NOT NULL DEFAULT 1,
        created_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS work_sessions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE NOT NULL,
        sequence INTEGER UNIQUE NOT NULL,
        opened_once INTEGER NOT NULL DEFAULT 0,
        is_open INTEGER NOT NULL DEFAULT 0,
        is_current INTEGER NOT NULL DEFAULT 0,
        opened_at TEXT,
        closed_at TEXT
    );

    CREATE TABLE IF NOT EXISTS jobs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        session_id INTEGER NOT NULL,
        mat_no TEXT NOT NULL,
        course_code TEXT NOT NULL,
        student_name TEXT DEFAULT '',
        signature_id TEXT DEFAULT '',
        department TEXT DEFAULT '',
        source TEXT DEFAULT '',
        status TEXT NOT NULL DEFAULT 'available',
        is_priority INTEGER NOT NULL DEFAULT 0,
        created_at TEXT NOT NULL,
        UNIQUE(session_id, mat_no, course_code),
        FOREIGN KEY(session_id) REFERENCES work_sessions(id)
    );

    CREATE TABLE IF NOT EXISTS allocations (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        job_id INTEGER NOT NULL,
        user_id INTEGER NOT NULL,
        status TEXT NOT NULL DEFAULT 'allocated',
        score REAL,
        ip_address TEXT DEFAULT '',
        timer_started_at TEXT,
        submitted_at TEXT,
        created_at TEXT NOT NULL,
        shown_at TEXT,
        started_at TEXT,
        finished_at TEXT,
        manual INTEGER NOT NULL DEFAULT 0,
        admin_id INTEGER,
        replacement_for INTEGER,
        FOREIGN KEY(job_id) REFERENCES jobs(id),
        FOREIGN KEY(user_id) REFERENCES users(id),
        FOREIGN KEY(admin_id) REFERENCES users(id),
        FOREIGN KEY(replacement_for) REFERENCES allocations(id)
    );

    CREATE TABLE IF NOT EXISTS reports (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        allocation_id INTEGER NOT NULL,
        job_id INTEGER NOT NULL,
        user_id INTEGER NOT NULL,
        reason TEXT NOT NULL,
        explanation TEXT NOT NULL,
        status TEXT NOT NULL DEFAULT 'open',
        replacement_allocation_id INTEGER,
        created_at TEXT NOT NULL,
        resolved_at TEXT,
        FOREIGN KEY(allocation_id) REFERENCES allocations(id),
        FOREIGN KEY(job_id) REFERENCES jobs(id),
        FOREIGN KEY(user_id) REFERENCES users(id)
    );

    CREATE TABLE IF NOT EXISTS notifications (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        message TEXT NOT NULL,
        is_read INTEGER NOT NULL DEFAULT 0,
        created_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS audit_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        admin_id INTEGER,
        action TEXT NOT NULL,
        details TEXT NOT NULL,
        created_at TEXT NOT NULL
    );
    """)

    for col, col_def in [
        ("ip_address", "TEXT DEFAULT ''"),
        ("timer_started_at", "TEXT"),
        ("submitted_at", "TEXT"),
    ]:
        try:
            con.execute(f"ALTER TABLE allocations ADD COLUMN {col} {col_def}")
            con.commit()
        except sqlite3.OperationalError:
            pass

    try:
        con.execute("ALTER TABLE jobs ADD COLUMN is_priority INTEGER NOT NULL DEFAULT 0")
        con.commit()
    except sqlite3.OperationalError:
        pass

    for name, sequence in SESSIONS:
        con.execute("""
            INSERT OR IGNORE INTO work_sessions(name, sequence)
            VALUES (?, ?)
        """, (name, sequence))

    con.commit()
    con.close()


def audit(admin_id, action, details):
    con = get_db()
    con.execute("""
        INSERT INTO audit_log(admin_id, action, details, created_at)
        VALUES (?, ?, ?, ?)
    """, (admin_id, action, details, now()))
    con.commit()


def notify_admin(message):
    con = get_db()
    con.execute("""
        INSERT INTO notifications(message, created_at)
        VALUES (?, ?)
    """, (message, now()))
    con.commit()


# ============================================================
# SECURITY & AUTHENTICATION HELPERS
# ============================================================

def get_notification_action(message):
    msg_lower = message.lower()
    if "registration" in msg_lower or "awaiting approval" in msg_lower:
        return {"url": url_for("admin_users"), "label": "Review & Approve Users", "icon": "bi-people", "btn": "btn-success"}
    elif "report" in msg_lower or "malfunction" in msg_lower:
        return {"url": url_for("admin_reports"), "label": "View & Resolve Reports", "icon": "bi-flag", "btn": "btn-danger"}
    elif "upload" in msg_lower or "record" in msg_lower:
        return {"url": url_for("admin_jobs"), "label": "View Uploaded Jobs", "icon": "bi-table", "btn": "btn-primary"}
    elif "session" in msg_lower:
        return {"url": url_for("admin_dashboard"), "label": "Manage Sessions", "icon": "bi-sliders", "btn": "btn-warning text-dark"}
    elif "allocation" in msg_lower:
        return {"url": url_for("admin_allocations"), "label": "View Allocations", "icon": "bi-list-task", "btn": "btn-info text-dark"}
    else:
        return {"url": url_for("admin_dashboard"), "label": "Admin Panel", "icon": "bi-arrow-right", "btn": "btn-outline-secondary"}


def csrf_token():
    if "csrf" not in session:
        session["csrf"] = secrets.token_hex(32)
    return session["csrf"]


def check_csrf():
    supplied = request.form.get("csrf_token", "")
    current_csrf = session.get("csrf")
    if not supplied or not current_csrf or supplied != current_csrf:
        session["csrf"] = secrets.token_hex(32)
        return False
    return True


@app.before_request
def protect_post_requests():
    if request.method == "POST":
        if not check_csrf():
            flash("Security session token updated. Please submit your request again.", "warning")
            return redirect(request.referrer or url_for("dashboard"))


@app.context_processor
def inject_globals():
    return {
        "csrf_token": csrf_token,
        "current_user": current_user,
        "current_year": datetime.now().year,
        "get_notification_action": get_notification_action,
    }


def current_user():
    user_id = session.get("user_id")
    if not user_id:
        return None
    con = get_db()
    user = con.execute(
        "SELECT * FROM users WHERE id=?", (user_id,)
    ).fetchone()
    return user


def login_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        user = current_user()
        if not user:
            return redirect(url_for("login"))
        if not user["active"]:
            session.clear()
            flash("Your account has been suspended.", "danger")
            return redirect(url_for("login"))
        if user["role"] == "user" and not user["approved"]:
            session.clear()
            flash("Your account is awaiting administrator approval.", "warning")
            return redirect(url_for("login"))
        return fn(*args, **kwargs)
    return wrapper


def admin_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        user = current_user()
        if not user or user["role"] != "admin":
            abort(403)
        if not user["active"]:
            abort(403)
        return fn(*args, **kwargs)
    return wrapper


# ============================================================
# GENERAL DATABASE HELPERS
# ============================================================

def get_session_by_id(con, session_id):
    return con.execute(
        "SELECT * FROM work_sessions WHERE id=?", (session_id,)
    ).fetchone()


def current_allocation_session(con):
    return con.execute("""
        SELECT * FROM work_sessions
        WHERE is_current=1 AND is_open=1
        LIMIT 1
    """).fetchone()


def available_job(con, session_id):
    return con.execute("""
        SELECT j.* FROM jobs j
        JOIN work_sessions s ON s.id = j.session_id
        WHERE j.session_id=? AND s.is_open=1 AND j.status='available'
          AND j.id NOT IN (
              SELECT job_id FROM allocations WHERE status != 'reported'
          )
        ORDER BY j.is_priority DESC, RANDOM()
        LIMIT 1
    """, (session_id,)).fetchone()


def create_allocation(con, job_id, user_id, manual=False,
                      admin_id=None, replacement_for=None):
    job = con.execute(
        "SELECT * FROM jobs WHERE id=?", (job_id,)
    ).fetchone()
    if not job or job["status"] != "available":
        raise ValueError("This job is not available.")

    # Double check no active allocation exists for this job
    existing = con.execute("SELECT 1 FROM allocations WHERE job_id=? AND status != 'reported'", (job_id,)).fetchone()
    if existing:
        con.execute("UPDATE jobs SET status='assigned' WHERE id=?", (job_id,))
        raise ValueError("This job has already been allocated.")

    con.execute("""
        INSERT INTO allocations(
            job_id,user_id,status,created_at,manual,admin_id,replacement_for
        ) VALUES(?,?,'allocated',?,?,?,?)
    """, (
        job_id, user_id, now(), int(manual), admin_id, replacement_for
    ))
    allocation_id = con.execute(
        "SELECT last_insert_rowid()"
    ).fetchone()[0]
    con.execute(
        "UPDATE jobs SET status='assigned' WHERE id=?", (job_id,)
    )
    return allocation_id


def allocate_jobs_to_user(con, session_id, user_id, requested_count, manual=False, admin_id=None):
    session_row = con.execute("SELECT is_open FROM work_sessions WHERE id=?", (session_id,)).fetchone()
    if not session_row or not session_row["is_open"]:
        raise ValueError("This session table is closed. Allocations have stopped.")

    total_units_allocated = 0
    total_jobs_allocated = 0
    multi_course_mats_allocated = 0
    MAX_MULTI_COURSE_MATS = 2

    while total_units_allocated < requested_count and total_jobs_allocated < requested_count:
        if multi_course_mats_allocated < MAX_MULTI_COURSE_MATS:
            mat_row = con.execute("""
                SELECT j.mat_no FROM jobs j
                JOIN work_sessions s ON s.id = j.session_id
                WHERE j.session_id=? AND s.is_open=1 AND j.status='available'
                  AND j.id NOT IN (
                      SELECT job_id FROM allocations WHERE status != 'reported'
                  )
                GROUP BY j.mat_no
                ORDER BY MAX(j.is_priority) DESC, RANDOM()
                LIMIT 1
            """, (session_id,)).fetchone()
        else:
            mat_row = con.execute("""
                SELECT j.mat_no FROM jobs j
                JOIN work_sessions s ON s.id = j.session_id
                WHERE j.session_id=? AND s.is_open=1 AND j.status='available'
                  AND j.id NOT IN (
                      SELECT job_id FROM allocations WHERE status != 'reported'
                  )
                GROUP BY j.mat_no
                HAVING COUNT(*) = 1
                ORDER BY MAX(j.is_priority) DESC, RANDOM()
                LIMIT 1
            """, (session_id,)).fetchone()

            if not mat_row:
                mat_row = con.execute("""
                    SELECT j.mat_no FROM jobs j
                    JOIN work_sessions s ON s.id = j.session_id
                    WHERE j.session_id=? AND s.is_open=1 AND j.status='available'
                      AND j.id NOT IN (
                          SELECT job_id FROM allocations WHERE status != 'reported'
                      )
                    GROUP BY j.mat_no
                    ORDER BY MAX(j.is_priority) DESC, RANDOM()
                    LIMIT 1
                """, (session_id,)).fetchone()

        if not mat_row:
            break

        mat_no = mat_row["mat_no"]

        already_allocated_mat = con.execute("""
            SELECT 1 FROM allocations a JOIN jobs j ON j.id = a.job_id
            WHERE j.session_id = ? AND j.mat_no = ? AND a.status != 'reported'
        """, (session_id, mat_no)).fetchone()

        if already_allocated_mat:
            con.execute("UPDATE jobs SET status='assigned' WHERE session_id=? AND mat_no=?", (session_id, mat_no))
            continue

        mat_jobs = con.execute("""
            SELECT j.* FROM jobs j
            JOIN work_sessions s ON s.id = j.session_id
            WHERE j.session_id=? AND s.is_open=1 AND j.mat_no=? AND j.status='available'
              AND j.id NOT IN (
                  SELECT job_id FROM allocations WHERE status != 'reported'
              )
            ORDER BY j.id
        """, (session_id, mat_no)).fetchall()

        if not mat_jobs:
            continue

        is_multi_course = len(mat_jobs) > 1

        if is_multi_course and multi_course_mats_allocated >= MAX_MULTI_COURSE_MATS:
            con.execute("UPDATE jobs SET status='assigned' WHERE session_id=? AND mat_no=?", (session_id, mat_no))
            continue

        if total_jobs_allocated + len(mat_jobs) > requested_count:
            if is_multi_course:
                con.execute("UPDATE jobs SET status='assigned' WHERE session_id=? AND mat_no=?", (session_id, mat_no))
                continue
            else:
                mat_jobs = mat_jobs[:(requested_count - total_jobs_allocated)]

        allocated_in_round = 0
        for mj in mat_jobs:
            existing = con.execute("""
                SELECT 1 FROM allocations WHERE job_id = ? AND status != 'reported'
            """, (mj["id"],)).fetchone()
            if existing:
                con.execute("UPDATE jobs SET status='assigned' WHERE id=?", (mj["id"],))
                continue

            create_allocation(con, mj["id"], user_id, manual=manual, admin_id=admin_id)
            total_jobs_allocated += 1
            allocated_in_round += 1

        if allocated_in_round > 0:
            total_units_allocated += 1
            if is_multi_course:
                multi_course_mats_allocated += 1

        remaining_unique = con.execute("""
            SELECT COUNT(DISTINCT j.mat_no) AS n FROM jobs j
            JOIN work_sessions s ON s.id = j.session_id
            WHERE j.session_id=? AND s.is_open=1 AND j.status='available'
              AND j.id NOT IN (SELECT job_id FROM allocations WHERE status != 'reported')
        """, (session_id,)).fetchone()["n"]

        if remaining_unique == 0 or total_units_allocated >= requested_count or total_jobs_allocated >= requested_count:
            break

    return total_jobs_allocated


@app.route("/admin/job/<int:job_id>/toggle-priority", methods=["POST"])
@login_required
@admin_required
def admin_toggle_priority(job_id):
    con = get_db()
    job = con.execute("SELECT * FROM jobs WHERE id=?", (job_id,)).fetchone()
    if not job:
        abort(404)

    new_val = 0 if job["is_priority"] else 1
    con.execute("UPDATE jobs SET is_priority=? WHERE id=?", (new_val, job_id))
    con.commit()

    status_label = "Priority" if new_val else "Standard"
    audit(current_user()["id"], "TOGGLE_PRIORITY", f"Job #{job_id} ({job['mat_no']} / {job['course_code']}) set to {status_label}")
    flash(f"Job successfully set to {status_label}.", "success")
    return redirect(request.referrer or url_for("admin_jobs"))


def scored_job_count(con, user_id):
    row = con.execute("""
        SELECT COUNT(*) AS total FROM allocations
        WHERE user_id=? AND score IS NOT NULL
    """, (user_id,)).fetchone()
    return row["total"]


def user_allocation_rows(con, user_id):
    return con.execute("""
        SELECT a.*, j.mat_no, j.course_code, j.student_name,
               j.signature_id, j.department, j.source,
               s.name AS session_name,
               (SELECT COUNT(DISTINCT j2.course_code)
                FROM allocations a2 JOIN jobs j2 ON j2.id = a2.job_id
                WHERE j2.session_id = j.session_id AND j2.mat_no = j.mat_no
                  AND a2.user_id = a.user_id AND a2.status != 'reported') AS multi_count
        FROM allocations a
        JOIN jobs j ON j.id=a.job_id
        JOIN work_sessions s ON s.id=j.session_id
        WHERE a.user_id=? AND a.status IN ('allocated', 'shown', 'started')
        ORDER BY a.id DESC
    """, (user_id,)).fetchall()


def user_completed_allocation_rows(con, user_id):
    return con.execute("""
        SELECT a.*, j.mat_no, j.course_code, j.student_name,
               j.signature_id, j.department, j.source,
               s.name AS session_name
        FROM allocations a
        JOIN jobs j ON j.id=a.job_id
        JOIN work_sessions s ON s.id=j.session_id
        WHERE a.user_id=? AND a.status IN ('finished', 'scored')
        ORDER BY a.id DESC
    """, (user_id,)).fetchall()


def user_report_rows(con, user_id):
    return con.execute("""
        SELECT r.*, j.mat_no, j.course_code, s.name AS session_name
        FROM reports r
        JOIN jobs j ON j.id = r.job_id
        JOIN work_sessions s ON s.id = j.session_id
        WHERE r.user_id = ?
        ORDER BY r.id DESC
    """, (user_id,)).fetchall()


# ============================================================
# MULTI-FORMAT RECORD IMPORT PARSER (.txt, .docx, .csv, .xlsx)
# ============================================================

def parse_records_with_gemini(raw_text):
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key or not raw_text.strip():
        return None

    prompt = f"""
You are an AI data extraction assistant for an examination marking portal.
Extract all exam job records from the following text and organize them into structured JSON.

Rules:
1. Matriculation Number must start with lowercase "nou" followed by 9 digits (e.g. nou261113275). Heal any OCR mistakes (e.g. letter O -> digit 0).
2. Course Code must be 3 letters followed by 3 digits (e.g. PCR111, GST101, NSC209).
3. Assign each record to one of the 3 work session times: "8:00 AM", "11:00 AM", or "2:00 PM".
4. Extract Student Name, Department, and Signature ID if present.

Return ONLY a raw valid JSON array of objects with keys:
"mat_no", "course_code", "student_name", "department", "signature_id", "session_time".

Text to parse:
{raw_text}
"""
    try:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key}"
        headers = {"Content-Type": "application/json"}
        payload = json.dumps({"contents": [{"parts": [{"text": prompt}]}]}).encode("utf-8")

        req = urllib.request.Request(url, data=payload, headers=headers, method="POST")
        with urllib.request.urlopen(req, timeout=15) as response:
            res_json = json.loads(response.read().decode())
            text = res_json['candidates'][0]['content']['parts'][0]['text']
            text = re.sub(r"^```(?:json)?\s*", "", text.strip(), flags=re.I)
            text = re.sub(r"\s*```$", "", text.strip())
            dict_items = json.loads(text)
            if isinstance(dict_items, list):
                return dict_items
    except Exception as e:
        print("Gemini API call warning:", e)
    return None


def parse_records_from_rows(rows_data, raw_text=""):
    """
    Intelligent AI Record Parser & Arranger:
    1. Invokes Gemini AI if GEMINI_API_KEY environment variable is configured.
    2. Fallback to built-in AI Heuristic Engine (tracks section headers, OCR healing, session bucketing).
    """
    if raw_text:
        gemini_items = parse_records_with_gemini(raw_text)
        if gemini_items:
            records = []
            corrected_count = 0
            for item in gemini_items:
                mat, mat_valid, mat_corrected = intelligent_cleanse_mat(str(item.get("mat_no", "")))
                course, course_valid, course_corrected = intelligent_cleanse_course(str(item.get("course_code", "")))
                if mat_valid and course_valid:
                    if mat_corrected or course_corrected:
                        corrected_count += 1
                    session_bucket = intelligent_session_bucket(str(item.get("session_time", "")))
                    records.append({
                        "mat_no": mat,
                        "course_code": course,
                        "student_name": str(item.get("student_name", "")).strip(),
                        "signature_id": str(item.get("signature_id", "")).strip(),
                        "department": str(item.get("department", "")).strip(),
                        "time_str": str(item.get("session_time", "")).strip(),
                        "session_bucket": session_bucket,
                        "source": f"{mat} {course} {item.get('session_time', '')}",
                    })
            records.sort(key=lambda r: (normalize_time_key(r["time_str"]), r["mat_no"], r["course_code"]))
            unique = {}
            for r in records:
                unique[(r["mat_no"], r["course_code"])] = r
            return list(unique.values()), corrected_count

    # Built-in AI Heuristic Parser
    records = []
    current_time_str = ""
    corrected_count = 0

    for item in rows_data:
        if isinstance(item, dict):
            raw_mat = str(item.get("mat_no", ""))
            raw_course = str(item.get("course_code", ""))
            student_name = str(item.get("student_name", "")).strip()
            department = str(item.get("department", "")).strip()
            signature_id = str(item.get("signature_id", "")).strip()
            time_str = str(item.get("time_str", "")).strip()
            source = str(item.get("source", ""))
        else:
            line = str(item).strip()
            if not line:
                continue

            times = TIME_PATTERN.findall(line)
            mats = MAT_HEAL_PATTERN.findall(line)
            courses = COURSE_HEAL_PATTERN.findall(line.upper())

            # Section header detection (e.g. "8:30AM", "11AM", "2PM")
            if times and not mats and not courses:
                current_time_str = times[0].strip()
                continue

            if not mats or not courses:
                continue

            raw_mat = mats[0]
            raw_course = courses[0]
            time_str = times[0].strip() if times else current_time_str

            clean_line = line
            clean_line = MAT_HEAL_PATTERN.sub("", clean_line)
            clean_line = COURSE_HEAL_PATTERN.sub("", clean_line)
            if time_str:
                clean_line = clean_line.replace(time_str, "")
            clean_line = re.sub(r"[|,\-:]+", " ", clean_line).strip()

            student_name = clean_line
            department = ""
            signature_id = ""
            source = line

        # AI Smart Healing & Validation
        mat, mat_valid, mat_corrected = intelligent_cleanse_mat(raw_mat)
        course, course_valid, course_corrected = intelligent_cleanse_course(raw_course)

        if not mat_valid or not course_valid:
            continue

        if mat_corrected or course_corrected:
            corrected_count += 1

        session_bucket = intelligent_session_bucket(time_str)

        records.append({
            "mat_no": mat,
            "course_code": course,
            "student_name": student_name,
            "signature_id": signature_id,
            "department": department,
            "time_str": time_str,
            "session_bucket": session_bucket,
            "source": source,
        })

    # Sort records chronologically by Session Bucket -> Mat No -> Course Code
    records.sort(key=lambda r: (normalize_time_key(r["time_str"]), r["mat_no"], r["course_code"]))

    unique = {}
    for r in records:
        unique[(r["mat_no"], r["course_code"])] = r

    final_list = list(unique.values())
    return final_list, corrected_count


def extract_upload(file):
    filename = secure_filename(file.filename or "")
    extension = os.path.splitext(filename)[1].lower()

    if extension == ".txt":
        content = file.read().decode("utf-8-sig", errors="replace")
        return parse_records_from_rows(content.splitlines())

    if extension == ".docx":
        document = Document(io.BytesIO(file.read()))
        lines = [p.text for p in document.paragraphs if p.text.strip()]
        for table in document.tables:
            for row in table.rows:
                lines.append(" | ".join(cell.text for cell in row.cells))
        return parse_records_from_rows(lines)

    if extension == ".csv":
        content = file.read().decode("utf-8-sig", errors="replace")
        reader = csv.reader(io.StringIO(content))
        rows = list(reader)
        if not rows:
            return []

        header = [c.strip().lower() for c in rows[0]]
        mat_idx = -1
        course_idx = -1
        name_idx = -1
        dept_idx = -1
        sig_idx = -1

        for idx, col in enumerate(header):
            if "mat" in col or "matric" in col:
                mat_idx = idx
            elif "course" in col or "code" in col:
                course_idx = idx
            elif "name" in col or "student" in col:
                name_idx = idx
            elif "dept" in col or "department" in col:
                dept_idx = idx
            elif "sig" in col or "signature" in col:
                sig_idx = idx

        if mat_idx != -1 and course_idx != -1:
            dict_records = []
            for row in rows[1:]:
                if not row or len(row) <= max(mat_idx, course_idx):
                    continue
                dict_records.append({
                    "mat_no": row[mat_idx],
                    "course_code": row[course_idx],
                    "student_name": row[name_idx] if name_idx != -1 and name_idx < len(row) else "",
                    "department": row[dept_idx] if dept_idx != -1 and dept_idx < len(row) else "",
                    "signature_id": row[sig_idx] if sig_idx != -1 and sig_idx < len(row) else "",
                    "source": " | ".join(row),
                })
            return parse_records_from_rows(dict_records)
        else:
            lines = [" | ".join(r) for r in rows]
            return parse_records_from_rows(lines)

    if extension == ".xlsx":
        workbook = openpyxl.load_workbook(io.BytesIO(file.read()), data_only=True)
        sheet = workbook.active
        rows = []
        for row in sheet.iter_rows(values_only=True):
            if any(row):
                rows.append([str(c) if c is not None else "" for c in row])

        if not rows:
            return []

        header = [c.strip().lower() for c in rows[0]]
        mat_idx = -1
        course_idx = -1
        name_idx = -1
        dept_idx = -1
        sig_idx = -1

        for idx, col in enumerate(header):
            if "mat" in col or "matric" in col:
                mat_idx = idx
            elif "course" in col or "code" in col:
                course_idx = idx
            elif "name" in col or "student" in col:
                name_idx = idx
            elif "dept" in col or "department" in col:
                dept_idx = idx
            elif "sig" in col or "signature" in col:
                sig_idx = idx

        if mat_idx != -1 and course_idx != -1:
            dict_records = []
            for row in rows[1:]:
                if not row or len(row) <= max(mat_idx, course_idx):
                    continue
                dict_records.append({
                    "mat_no": row[mat_idx],
                    "course_code": row[course_idx],
                    "student_name": row[name_idx] if name_idx != -1 and name_idx < len(row) else "",
                    "department": row[dept_idx] if dept_idx != -1 and dept_idx < len(row) else "",
                    "signature_id": row[sig_idx] if sig_idx != -1 and sig_idx < len(row) else "",
                    "source": " | ".join(row),
                })
            return parse_records_from_rows(dict_records)
        else:
            lines = [" | ".join(r) for r in rows]
            return parse_records_from_rows(lines)

    raise ValueError("Only .docx, .txt, .csv, and .xlsx files are supported.")


# ============================================================
# PUBLIC ROUTES (Setup, Auth, Landing)
# ============================================================

@app.route("/")
def index():
    if current_user():
        return redirect(url_for("dashboard"))
    return redirect(url_for("login"))


@app.route("/setup", methods=["GET", "POST"])
def setup():
    con = get_db()
    count = con.execute(
        "SELECT COUNT(*) AS n FROM users WHERE role='admin'"
    ).fetchone()["n"]

    if count > 0 and request.method == "GET":
        return render_template("setup.html", disabled=True, user=current_user())

    if request.method == "POST":
        if count > 0:
            flash("Initial admin account already exists. Setup is disabled.", "danger")
            return redirect(url_for("login"))

        username = request.form.get("username", "").strip()
        email = request.form.get("email", "").strip().lower()
        password = request.form.get("password", "")

        if len(password) < 12:
            flash("Use a password with at least 12 characters.", "warning")
            return redirect(url_for("setup"))

        try:
            con.execute("""
                INSERT INTO users(username,email,password_hash,role,approved,created_at)
                VALUES(?,?,?,'admin',1,?)
            """, (username, email, generate_password_hash(password), now()))
            con.commit()
            flash("Administrator created successfully. Please log in.", "success")
        except sqlite3.IntegrityError:
            flash("Username or email already exists.", "danger")

        return redirect(url_for("login"))

    return render_template("setup.html", disabled=False, user=current_user())


@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        email = request.form.get("email", "").strip().lower()
        password = request.form.get("password", "")

        if len(username) < 3 or len(password) < 12:
            flash("Username must be at least 3 characters and password at least 12.", "danger")
            return redirect(url_for("register"))

        con = get_db()
        try:
            con.execute("""
                INSERT INTO users(username,email,password_hash,created_at)
                VALUES(?,?,?,?)
            """, (username, email, generate_password_hash(password), now()))
            con.commit()
            notify_admin("New user registration awaiting approval: " + username)
            flash("Registration received. Please wait for administrator approval.", "success")
            return redirect(url_for("login"))
        except sqlite3.IntegrityError:
            flash("That username or email is already registered.", "danger")

    return render_template("register.html", user=current_user())


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        identity = request.form.get("identity", "").strip()
        password = request.form.get("password", "")

        con = get_db()
        user = con.execute("""
            SELECT * FROM users
            WHERE username=? OR email=?
        """, (identity, identity.lower())).fetchone()

        if not user or not check_password_hash(user["password_hash"], password):
            flash("Incorrect login credentials.", "danger")
            return redirect(url_for("login"))

        if not user["active"]:
            flash("Your account has been suspended.", "danger")
            return redirect(url_for("login"))

        if user["role"] == "user" and not user["approved"]:
            flash("Your account is awaiting administrator approval.", "warning")
            return redirect(url_for("login"))

        session.clear()
        session["user_id"] = user["id"]
        session["csrf"] = secrets.token_hex(32)
        flash(f"Welcome back, {user['username']}!", "success")
        return redirect(url_for("dashboard"))

    return render_template("login.html", user=current_user())


@app.route("/logout")
def logout():
    session.clear()
    flash("You have logged out.", "info")
    return redirect(url_for("login"))


# ============================================================
# USER DASHBOARD & WORKFLOW
# ============================================================

@app.route("/dashboard")
@login_required
def dashboard():
    user = current_user()

    if user["role"] == "admin":
        return redirect(url_for("admin_dashboard"))

    con = get_db()
    allocations = user_allocation_rows(con, user["id"])
    completed_allocations = user_completed_allocation_rows(con, user["id"])
    user_reports = user_report_rows(con, user["id"])
    open_sessions = con.execute("SELECT * FROM work_sessions WHERE is_open=1 ORDER BY sequence").fetchall()
    scored = scored_job_count(con, user["id"])

    return render_template(
        "dashboard.html",
        user=user,
        allocations=allocations,
        completed_allocations=completed_allocations,
        user_reports=user_reports,
        open_sessions=open_sessions,
        scored=scored
    )


def allocate_jobs_mixed_sessions(con, user_id, requested_count, manual=False, admin_id=None):
    open_sessions = con.execute("""
        SELECT * FROM work_sessions WHERE is_open=1 ORDER BY sequence
    """).fetchall()

    if not open_sessions:
        raise ValueError("There are no open allocation sessions.")

    total_allocated = 0
    attempts = 0
    while total_allocated < requested_count and attempts < (requested_count * 3):
        attempts += 1
        available_sessions = []
        for s in open_sessions:
            cnt = con.execute("""
                SELECT COUNT(*) AS n FROM jobs WHERE session_id=? AND status='available'
            """, (s["id"],)).fetchone()["n"]
            if cnt > 0:
                available_sessions.append(s)

        if not available_sessions:
            break

        target_session = random.choice(available_sessions)
        allocated_in_session = allocate_jobs_to_user(con, target_session["id"], user_id, 1, manual=manual, admin_id=admin_id)
        if allocated_in_session > 0:
            total_allocated += allocated_in_session

    return total_allocated


@app.route("/request-jobs", methods=["POST"])
@login_required
def request_jobs():
    user = current_user()
    if user["role"] != "user":
        abort(403)

    con = get_db()
    open_sessions = con.execute("SELECT * FROM work_sessions WHERE is_open=1").fetchall()

    if not open_sessions:
        flash("There are no open allocation sessions.", "danger")
        return redirect(url_for("dashboard"))

    # Restriction: User can hold up to 8 active unsubmitted jobs at a time
    active_unsubmitted = con.execute("""
        SELECT COUNT(*) AS n FROM allocations a
        JOIN jobs j ON j.id = a.job_id
        WHERE a.user_id = ? AND a.status IN ('allocated', 'shown', 'started')
    """, (user["id"],)).fetchone()["n"]

    if active_unsubmitted >= 8:
        flash("You have reached the limit of 8 active unsubmitted jobs. Please submit some before requesting new ones.", "warning")
        return redirect(url_for("dashboard"))

    try:
        quantity = int(request.form.get("quantity", "0"))
    except ValueError:
        quantity = 0

    if quantity < 1 or quantity > 6:
        flash("You cannot request more than 6 jobs at a time (maximum 6 jobs per request).", "warning")
        return redirect(url_for("dashboard"))

    # Rate limit: max 6 jobs within 20 minutes interval
    twenty_mins_ago = datetime.now() - timedelta(minutes=20)
    twenty_mins_ago_str = twenty_mins_ago.strftime("%Y-%m-%d %H:%M:%S")

    recent_allocated_count = con.execute("""
        SELECT COUNT(*) AS n FROM allocations
        WHERE user_id = ? AND created_at >= ? AND status != 'reported'
    """, (user["id"], twenty_mins_ago_str)).fetchone()["n"]

    if recent_allocated_count + quantity > 6:
        max_allowed = max(0, 6 - recent_allocated_count)
        flash(f"Rate limit: You have received {recent_allocated_count} job(s) in the last 20 minutes. You can request at most {max_allowed} more job(s) right now.", "warning")
        return redirect(url_for("dashboard"))

    if active_unsubmitted + quantity > 8:
        flash(f"You currently have {active_unsubmitted} active job(s). You can request at most {8 - active_unsubmitted} more job(s) (limit is 8 active).", "warning")
        return redirect(url_for("dashboard"))

    allocated = 0
    try:
        con.execute("BEGIN IMMEDIATE")
        allocated = allocate_jobs_mixed_sessions(con, user["id"], quantity, manual=False)
        con.commit()
    except Exception as e:
        con.rollback()
        flash(f"Error allocating jobs: {e}", "danger")
        return redirect(url_for("dashboard"))

    if allocated > 0:
        if allocated < quantity:
            flash(f"Partial allocation: Only {allocated} job(s) were available and allocated (that is all that remains across open sessions).", "warning")
        else:
            flash(f"Successfully allocated {allocated} job(s) mixed across open session tables.", "success")
    else:
        flash("No available jobs remain in the open sessions.", "warning")

    return redirect(url_for("dashboard"))


def get_owned_allocation(con, allocation_id, user_id):
    row = con.execute("""
        SELECT a.*, j.session_id, j.mat_no, j.course_code
        FROM allocations a JOIN jobs j ON j.id=a.job_id
        WHERE a.id=? AND a.user_id=?
    """, (allocation_id, user_id)).fetchone()
    if not row:
        abort(404)
    return row


@app.route("/allocation/<int:allocation_id>/show", methods=["POST"])
@login_required
def show_allocation(allocation_id):
    user = current_user()
    con = get_db()
    row = get_owned_allocation(con, allocation_id, user["id"])

    if row["status"] != "allocated":
        flash("This job cannot be shown in its current state.", "warning")
        return redirect(url_for("dashboard"))

    ip_address = request.form.get("ip_address", "").strip()
    if not ip_address:
        flash("Please enter an IP number or station identifier to open this job.", "danger")
        return redirect(url_for("dashboard"))

    # Once IP/number is put, start the timer immediately (status='started')
    con.execute("""
        UPDATE allocations SET status='started', shown_at=?, started_at=?, timer_started_at=?, ip_address=?
        WHERE id=? AND user_id=? AND status='allocated'
    """, (now(), now(), now(), ip_address, allocation_id, user["id"]))
    con.commit()
    flash(f"Job opened with IP/Station '{ip_address}' and timer started!", "success")
    return redirect(url_for("dashboard"))


@app.route("/allocation/<int:allocation_id>/submit", methods=["POST"])
@login_required
def submit_allocation(allocation_id):
    user = current_user()
    con = get_db()
    row = get_owned_allocation(con, allocation_id, user["id"])

    if row["status"] != "started":
        flash("Only an active started job can be submitted.", "warning")
        return redirect(url_for("dashboard"))

    # 30 minute threshold check (elapsed >= 30 minutes / 1800 seconds)
    if row["timer_started_at"]:
        start_dt = datetime.strptime(row["timer_started_at"], "%Y-%m-%d %H:%M:%S")
        elapsed_secs = (datetime.now() - start_dt).total_seconds()
        if elapsed_secs < 1800: # 30 mins = 1800 seconds
            flash("End Exam / Submit is only available after 30 minutes have elapsed.", "warning")
            return redirect(url_for("dashboard"))

    try:
        score = float(request.form["score"])
    except (ValueError, KeyError):
        score = -1

    if not 0 <= score <= 100:
        flash("Score must be between 0 and 100.", "danger")
        return redirect(url_for("dashboard"))

    con.execute("""
        UPDATE allocations SET status='finished', finished_at=?, submitted_at=?, score=?
        WHERE id=? AND user_id=? AND status='started'
    """, (now(), now(), score, allocation_id, user["id"]))
    con.commit()
    flash("Job ended and submitted successfully! You may now request new jobs.", "success")
    return redirect(url_for("dashboard"))


@app.route("/allocation/<int:allocation_id>/replace-done", methods=["POST"])
@login_required
def replace_done_allocation(allocation_id):
    user = current_user()
    con = get_db()
    row = get_owned_allocation(con, allocation_id, user["id"])

    if row["status"] not in ("shown", "started"):
        flash("You can request a replacement for an active or started job.", "warning")
        return redirect(url_for("dashboard"))

    try:
        con.execute("BEGIN IMMEDIATE")

        con.execute("""
            UPDATE allocations SET status='finished', finished_at=?, score=100
            WHERE id=? AND user_id=?
        """, (now(), allocation_id, user["id"]))

        con.execute("UPDATE jobs SET status='finished' WHERE id=?", (row["job_id"],))

        replacement = available_job(con, row["session_id"])
        replacement_id = None
        if replacement:
            replacement_id = create_allocation(
                con, replacement["id"], user["id"], manual=False,
                replacement_for=allocation_id
            )
            flash(f"Job marked as already done. Replacement job #{replacement_id} automatically assigned to you!", "success")
        else:
            flash("Job marked as done, but no other available jobs remain in this session.", "info")

        con.commit()
    except Exception as e:
        con.rollback()
        flash(f"Error processing replacement: {e}", "danger")

    return redirect(url_for("dashboard"))


@app.route("/allocation/<int:allocation_id>/finish", methods=["POST"])
@login_required
def finish_allocation(allocation_id):
    user = current_user()
    con = get_db()
    row = get_owned_allocation(con, allocation_id, user["id"])

    if row["status"] != "started":
        flash("Only a started job can be marked finished.", "warning")
        return redirect(url_for("dashboard"))

    con.execute("""
        UPDATE allocations SET status='finished', finished_at=?
        WHERE id=? AND user_id=? AND status='started'
    """, (now(), allocation_id, user["id"]))
    con.commit()
    flash("Work marked as finished. Please enter your score.", "success")
    return redirect(url_for("dashboard"))


@app.route("/allocation/<int:allocation_id>/score", methods=["POST"])
@login_required
def score_allocation(allocation_id):
    user = current_user()
    con = get_db()
    row = get_owned_allocation(con, allocation_id, user["id"])

    if row["status"] != "finished":
        flash("Only finished jobs can be scored.", "warning")
        return redirect(url_for("dashboard"))

    try:
        score = float(request.form["score"])
    except (ValueError, KeyError):
        score = -1

    if not 0 <= score <= 100:
        flash("Score must be between 0 and 100.", "danger")
        return redirect(url_for("dashboard"))

    con.execute("""
        UPDATE allocations SET score=?
        WHERE id=? AND user_id=? AND status='finished'
    """, (score, allocation_id, user["id"]))
    con.commit()
    flash("Score saved successfully.", "success")
    return redirect(url_for("dashboard"))


@app.route("/allocation/<int:allocation_id>/report", methods=["POST"])
@login_required
def report_allocation(allocation_id):
    user = current_user()
    con = get_db()
    row = get_owned_allocation(con, allocation_id, user["id"])

    if row["status"] not in ("shown", "started"):
        flash("You can report a job only after it has been shown.", "warning")
        return redirect(url_for("dashboard"))

    reason = request.form.get("reason", "").strip()
    explanation = request.form.get("explanation", "").strip()

    if reason not in REPORT_REASONS or not explanation:
        flash("Please choose a valid report reason and provide an explanation.", "danger")
        return redirect(url_for("dashboard"))

    try:
        con.execute("BEGIN IMMEDIATE")

        con.execute("""
            INSERT INTO reports(
                allocation_id,job_id,user_id,reason,explanation,created_at
            ) VALUES(?,?,?,?,?,?)
        """, (
            allocation_id, row["job_id"], user["id"],
            reason, explanation, now()
        ))
        report_id = con.execute("SELECT last_insert_rowid()").fetchone()[0]

        if reason == "System malfunctioning":
            con.execute("UPDATE jobs SET status='available' WHERE id=?", (row["job_id"],))
            con.execute("UPDATE allocations SET status='paused' WHERE id=?", (allocation_id,))
            message = (
                f"SYSTEM MALFUNCTION report #{report_id}: "
                f"{user['username']} | {row['mat_no']} | "
                f"{row['course_code']} | {explanation}. "
                "Affected job released for reassignment."
            )
        else:
            con.execute("UPDATE jobs SET status='review' WHERE id=?", (row["job_id"],))
            con.execute("UPDATE allocations SET status='reported' WHERE id=?", (allocation_id,))

            replacement = available_job(con, row["session_id"])
            replacement_id = None
            if replacement:
                replacement_id = create_allocation(
                    con, replacement["id"], user["id"],
                    replacement_for=allocation_id
                )

            con.execute("UPDATE reports SET replacement_allocation_id=? WHERE id=?", (replacement_id, report_id))

            if replacement_id:
                message = (
                    f"Report #{report_id}: {user['username']} reported "
                    f"{row['mat_no']} / {row['course_code']} ({reason}). "
                    f"Replacement allocation #{replacement_id} issued."
                )
            else:
                message = (
                    f"Report #{report_id}: {user['username']} reported "
                    f"{row['mat_no']} / {row['course_code']} ({reason}). "
                    "No replacement currently available."
                )

        con.execute("INSERT INTO notifications(message,created_at) VALUES(?,?)", (message, now()))
        con.commit()
    except Exception:
        con.rollback()
        raise

    flash("Your report has been sent to the administrator.", "success")
    return redirect(url_for("dashboard"))


# ============================================================
# ADMIN DASHBOARD & CONTROLS
# ============================================================

@app.route("/admin")
@login_required
@admin_required
def admin_dashboard():
    con = get_db()

    users_count = con.execute("SELECT COUNT(*) AS n FROM users WHERE role='user'").fetchone()["n"]
    pending_count = con.execute("SELECT COUNT(*) AS n FROM users WHERE role='user' AND approved=0").fetchone()["n"]
    jobs_count = con.execute("SELECT COUNT(*) AS n FROM jobs").fetchone()["n"]
    open_reports_count = con.execute("SELECT COUNT(*) AS n FROM reports WHERE status='open'").fetchone()["n"]
    notifications = con.execute("SELECT * FROM notifications ORDER BY id DESC LIMIT 20").fetchall()
    sessions = con.execute("SELECT * FROM work_sessions ORDER BY sequence").fetchall()

    session_summary = []
    for s in sessions:
        count = con.execute("SELECT COUNT(*) AS n FROM jobs WHERE session_id=?", (s["id"],)).fetchone()["n"]
        session_summary.append({"session": s, "count": count})

    return render_template(
        "admin/dashboard.html",
        user=current_user(),
        users_count=users_count,
        pending_count=pending_count,
        jobs_count=jobs_count,
        open_reports_count=open_reports_count,
        notifications=notifications,
        sessions=sessions,
        session_summary=session_summary
    )


@app.route("/admin/session/<int:sid>/open", methods=["POST"])
@login_required
@admin_required
def open_session(sid):
    admin = current_user()
    con = get_db()
    target = get_session_by_id(con, sid)

    if not target:
        abort(404)

    earlier = con.execute("""
        SELECT COUNT(*) AS n FROM work_sessions
        WHERE sequence < ? AND opened_once=0
    """, (target["sequence"],)).fetchone()["n"]

    if earlier:
        flash("You must open earlier sessions first in sequence.", "warning")
        return redirect(url_for("admin_dashboard"))

    con.execute("""
        UPDATE work_sessions SET is_open=1, opened_once=1,
        is_current=0, opened_at=?
        WHERE id=?
    """, (now(), sid))
    con.execute("UPDATE work_sessions SET is_current=1 WHERE id=?", (sid,))
    con.commit()

    audit(admin["id"], "OPEN_SESSION", target["name"])
    notify_admin(f"{target['name']} session opened by {admin['username']}.")
    flash(f"{target['name']} session is now open and set as current target.", "success")
    return redirect(url_for("admin_dashboard"))


@app.route("/admin/session/<int:sid>/close", methods=["POST"])
@login_required
@admin_required
def close_session(sid):
    admin = current_user()
    con = get_db()
    target = get_session_by_id(con, sid)

    if not target:
        abort(404)

    con.execute("""
        UPDATE work_sessions
        SET is_open=0, is_current=0, closed_at=?
        WHERE id=?
    """, (now(), sid))

    remaining = con.execute("""
        SELECT id FROM work_sessions
        WHERE is_open=1
        ORDER BY sequence DESC LIMIT 1
    """).fetchone()

    if remaining:
        con.execute("UPDATE work_sessions SET is_current=1 WHERE id=?", (remaining["id"],))

    con.commit()

    audit(admin["id"], "CLOSE_SESSION", target["name"])
    notify_admin(f"{target['name']} session closed by {admin['username']}.")
    flash(f"{target['name']} session closed.", "info")
    return redirect(url_for("admin_dashboard"))


@app.route("/admin/users")
@login_required
@admin_required
def admin_users():
    con = get_db()
    users = con.execute("SELECT * FROM users ORDER BY role DESC, created_at DESC").fetchall()

    users_with_work = []
    for u in users:
        active_allocations = con.execute("""
            SELECT a.*, j.mat_no, j.course_code, s.name AS session_name
            FROM allocations a
            JOIN jobs j ON j.id = a.job_id
            JOIN work_sessions s ON s.id = j.session_id
            WHERE a.user_id = ? AND a.status IN ('allocated', 'shown', 'started')
            ORDER BY a.id DESC
        """, (u["id"],)).fetchall()

        users_with_work.append({
            "user": u,
            "active_allocations": active_allocations
        })

    return render_template("admin/users.html", user=current_user(), users_with_work=users_with_work)


@app.route("/admin/user/<int:uid>/allocations")
@login_required
@admin_required
def admin_user_allocations(uid):
    con = get_db()
    user_target = con.execute("SELECT * FROM users WHERE id=?", (uid,)).fetchone()
    if not user_target:
        abort(404)

    rows = con.execute("""
        SELECT a.*, j.mat_no, j.course_code, j.student_name,
               j.signature_id, j.department, j.source,
               s.name AS session_name
        FROM allocations a
        JOIN jobs j ON j.id = a.job_id
        JOIN work_sessions s ON s.id = j.session_id
        WHERE a.user_id = ?
        ORDER BY a.id DESC
    """, (uid,)).fetchall()

    return render_template("admin/user_allocations.html", user=current_user(), user_target=user_target, rows=rows)


@app.route("/admin/user/<int:uid>/approve", methods=["POST"])
@login_required
@admin_required
def approve_user(uid):
    con = get_db()
    con.execute("UPDATE users SET approved=1 WHERE id=? AND role='user'", (uid,))
    con.commit()
    audit(current_user()["id"], "APPROVE_USER", f"User ID: {uid}")
    flash("User account approved.", "success")
    return redirect(url_for("admin_users"))


@app.route("/admin/users/approve-all", methods=["POST"])
@login_required
@admin_required
def admin_approve_all_users():
    con = get_db()
    try:
        con.execute("UPDATE users SET approved=1 WHERE role='user' AND approved=0")
        con.commit()
        audit(current_user()["id"], "APPROVE_ALL_USERS", "Approved all pending user sign-ups.")
        flash("All pending user accounts have been successfully approved.", "success")
    except Exception as e:
        con.rollback()
        flash(f"Error approving all users: {e}", "danger")

    return redirect(url_for("admin_users"))


@app.route("/admin/user/<int:uid>/toggle", methods=["POST"])
@login_required
@admin_required
def toggle_user(uid):
    con = get_db()
    con.execute("""
        UPDATE users SET active=CASE WHEN active=1 THEN 0 ELSE 1 END
        WHERE id=? AND role='user'
    """, (uid,))
    con.commit()
    audit(current_user()["id"], "TOGGLE_USER", f"User ID: {uid}")
    flash("User account status updated.", "info")
    return redirect(url_for("admin_users"))


@app.route("/admin/upload", methods=["GET", "POST"])
@login_required
@admin_required
def upload_records():
    con = get_db()
    sessions = con.execute("SELECT * FROM work_sessions ORDER BY sequence").fetchall()
    session_map = {s["name"]: s["id"] for s in sessions}
    has_gemini_key = bool(os.environ.get("GEMINI_API_KEY"))

    if request.method == "POST":
        file = request.files.get("file")
        pasted_text = request.form.get("pasted_text", "").strip()
        already_done_text = request.form.get("already_done_text", "").strip()
        auto_allocate = request.form.get("auto_allocate") == "1"
        selected_session_val = request.form.get("session_id", "auto")

        try:
            if pasted_text:
                records, corrections = parse_records_from_rows(pasted_text.splitlines(), raw_text=pasted_text)
            elif file and file.filename:
                records, corrections = extract_upload(file)
            else:
                raise ValueError("Please paste job records into the chat box or select a file to upload.")

            if not records:
                raise ValueError(
                    "No valid matriculation/course pairs were found. "
                    "Ensure records contain matriculation numbers (e.g. nou123456789) and course codes (e.g. GST101)."
                )

            done_set = set()
            if already_done_text:
                try:
                    done_records, _ = parse_records_from_rows(already_done_text.splitlines(), raw_text=already_done_text)
                    for dr in done_records:
                        done_set.add((dr["mat_no"], dr["course_code"]))
                except Exception:
                    pass

                for line in already_done_text.splitlines():
                    line_lower = line.strip().lower()
                    if not line_lower:
                        continue
                    mats = re.findall(r'nou\d{9}', line_lower)
                    courses = re.findall(r'[a-z]{3}\d{3}', line_lower)
                    if mats and courses:
                        for m in mats:
                            for c in courses:
                                done_set.add((m, c.upper()))

            added = 0
            duplicates = 0
            bucket_counts = {"8:00 AM": 0, "11:00 AM": 0, "2:00 PM": 0}

            forced_sid = None
            if selected_session_val != "auto" and str(selected_session_val).isdigit():
                forced_sid = int(selected_session_val)

            for r in records:
                bucket = r["session_bucket"]
                bucket_counts[bucket] = bucket_counts.get(bucket, 0) + 1

                if forced_sid is not None:
                    target_sid = forced_sid
                else:
                    target_sid = session_map.get(bucket, sessions[0]["id"] if sessions else 1)

                is_done = (r["mat_no"], r["course_code"]) in done_set
                job_status = 'finished' if is_done else 'available'

                try:
                    con.execute("""
                        INSERT INTO jobs(
                            session_id,mat_no,course_code,student_name,
                            signature_id,department,source,status,created_at
                        ) VALUES(?,?,?,?,?,?,?,?,?)
                    """, (
                        target_sid, r["mat_no"], r["course_code"],
                        r["student_name"], r["signature_id"],
                        r["department"], r["source"], job_status, now()
                    ))
                    added += 1
                except sqlite3.IntegrityError:
                    if is_done:
                        con.execute("""
                            UPDATE jobs SET status='finished'
                            WHERE session_id=? AND mat_no=? AND course_code=?
                        """, (target_sid, r["mat_no"], r["course_code"]))
                    duplicates += 1

            con.commit()

            if auto_allocate and added > 0:
                try:
                    approved_users = con.execute("""
                        SELECT id FROM users WHERE role='user' AND approved=1 AND active=1
                    """).fetchall()
                    if approved_users:
                        open_sessions = con.execute("SELECT id FROM work_sessions WHERE is_open=1").fetchall()
                        for os_row in open_sessions:
                            sid = os_row["id"]
                            for u in approved_users:
                                try:
                                    allocate_jobs_to_user(con, sid, u["id"], 5, manual=False)
                                except Exception:
                                    pass
                except Exception:
                    pass

            audit(current_user()["id"], "UPLOAD_RECORDS",
                  f"Added {added}, duplicates {duplicates}, corrections {corrections}")
            notify_admin(f"Record upload: {added} added, {duplicates} duplicates skipped.")

            engine_label = "Gemini AI" if has_gemini_key else "Smart AI Engine"
            breakdown_msg = (
                f"✨ {engine_label} organized {added} records into 8:00 AM ({bucket_counts.get('8:00 AM', 0)}), "
                f"11:00 AM ({bucket_counts.get('11:00 AM', 0)}), and 2:00 PM ({bucket_counts.get('2:00 PM', 0)}). "
            )
            if auto_allocate and added > 0:
                breakdown_msg += "Jobs were also automatically allocated to available users! "
            if corrections > 0:
                breakdown_msg += f"Healed & fixed {corrections} OCR typos/formatting errors! "
            if duplicates > 0:
                breakdown_msg += f"Skipped {duplicates} duplicate entries."

            flash(breakdown_msg, "success")
        except Exception as error:
            flash(str(error), "danger")

        return redirect(url_for("upload_records"))

    session_data = []
    for s in sessions:
        rows = con.execute("""
            SELECT j.*, a.id AS allocation_id, a.status AS allocation_status,
                   a.score, a.manual, a.created_at AS allocated_at,
                   a.started_at, a.finished_at,
                   u.username AS assigned_user,
                   (SELECT COUNT(DISTINCT j2.course_code)
                    FROM allocations a2 JOIN jobs j2 ON j2.id = a2.job_id
                    WHERE j2.session_id = j.session_id AND j2.mat_no = j.mat_no
                      AND a2.user_id = a.user_id AND a2.status != 'reported') AS multi_count
            FROM jobs j
            LEFT JOIN allocations a ON a.id = (
                SELECT MAX(id) FROM allocations WHERE job_id = j.id
            )
            LEFT JOIN users u ON u.id = a.user_id
            WHERE j.session_id = ?
            ORDER BY j.id DESC
        """, (s["id"],)).fetchall()
        session_data.append({"session": s, "rows": rows})

    return render_template("admin/upload.html", user=current_user(), sessions=sessions, has_gemini_key=has_gemini_key, session_data=session_data)


@app.route("/admin/jobs")
@login_required
@admin_required
def admin_jobs():
    con = get_db()
    rows = con.execute("""
        SELECT j.*, s.name AS session_name,
               a.id AS allocation_id, a.status AS allocation_status,
               a.score, a.manual, a.created_at AS allocated_at,
               a.started_at, a.finished_at,
               u.username AS assigned_user,
               (SELECT COUNT(DISTINCT j2.course_code)
                FROM allocations a2 JOIN jobs j2 ON j2.id = a2.job_id
                WHERE j2.session_id = j.session_id AND j2.mat_no = j.mat_no
                  AND a2.user_id = a.user_id AND a2.status != 'reported') AS multi_count
        FROM jobs j
        JOIN work_sessions s ON s.id = j.session_id
        LEFT JOIN allocations a ON a.id = (
            SELECT MAX(id) FROM allocations WHERE job_id = j.id
        )
        LEFT JOIN users u ON u.id = a.user_id
        ORDER BY j.id DESC
    """).fetchall()
    return render_template("admin/jobs.html", user=current_user(), rows=rows)


@app.route("/admin/session-tables")
@login_required
@admin_required
def admin_session_tables():
    con = get_db()
    sessions = con.execute("SELECT * FROM work_sessions ORDER BY sequence").fetchall()
    approved_users = con.execute("""
        SELECT id, username FROM users WHERE role='user' AND approved=1 AND active=1 ORDER BY username
    """).fetchall()

    session_data = []
    for s in sessions:
        rows = con.execute("""
            SELECT j.*, a.id AS allocation_id, a.status AS allocation_status,
                   a.score, a.manual, a.created_at AS allocated_at,
                   a.started_at, a.finished_at,
                   u.username AS assigned_user
            FROM jobs j
            LEFT JOIN allocations a ON a.id = (
                SELECT MAX(id) FROM allocations WHERE job_id = j.id
            )
            LEFT JOIN users u ON u.id = a.user_id
            WHERE j.session_id = ?
            ORDER BY j.id DESC
        """, (s["id"],)).fetchall()

        total_jobs = len(rows)
        finished_jobs = sum(1 for r in rows if r["allocation_status"] in ['finished', 'scored'] or r["score"] is not None)
        assigned_jobs = sum(1 for r in rows if r["allocation_status"] in ['allocated', 'shown', 'started'])
        available_jobs = sum(1 for r in rows if r["status"] == 'available')

        session_data.append({
            "session": s,
            "rows": rows,
            "total_jobs": total_jobs,
            "finished_jobs": finished_jobs,
            "assigned_jobs": assigned_jobs,
            "available_jobs": available_jobs
        })

    return render_template("admin/session_tables.html", user=current_user(), session_data=session_data, approved_users=approved_users)


@app.route("/admin/allocation/<int:allocation_id>/release", methods=["POST"])
@login_required
@admin_required
def admin_release_allocation(allocation_id):
    con = get_db()
    alloc = con.execute("""
        SELECT a.*, j.id AS job_id, j.mat_no, j.course_code
        FROM allocations a JOIN jobs j ON j.id = a.job_id
        WHERE a.id=?
    """, (allocation_id,)).fetchone()
    if not alloc:
        abort(404)

    job_id = alloc["job_id"]

    try:
        con.execute("BEGIN IMMEDIATE")
        con.execute("UPDATE allocations SET status='reported' WHERE id=?", (allocation_id,))
        con.execute("UPDATE jobs SET status='available' WHERE id=?", (job_id,))
        con.commit()
        audit(current_user()["id"], "RELEASE_JOB", f"Released job #{job_id} ({alloc['mat_no']} / {alloc['course_code']}) back to unallocated pool.")
        flash("Job released back to unallocated pool awaiting allocation.", "success")
    except Exception as e:
        con.rollback()
        flash(f"Error releasing job: {e}", "danger")

    return redirect(request.referrer or url_for("admin_session_tables"))


@app.route("/admin/allocation/<int:allocation_id>/reallocate", methods=["POST"])
@login_required
@admin_required
def admin_reallocate_allocation(allocation_id):
    con = get_db()
    alloc = con.execute("""
        SELECT a.*, j.id AS job_id, j.session_id, j.mat_no, j.course_code
        FROM allocations a JOIN jobs j ON j.id = a.job_id
        WHERE a.id=?
    """, (allocation_id,)).fetchone()
    if not alloc:
        abort(404)

    new_user_id = request.form.get("new_user_id")
    if not new_user_id:
        flash("Please select a user to reallocate this job to.", "warning")
        return redirect(request.referrer or url_for("admin_session_tables"))

    try:
        new_user_id = int(new_user_id)
        con.execute("BEGIN IMMEDIATE")
        con.execute("UPDATE allocations SET status='reported' WHERE id=?", (allocation_id,))
        create_allocation(con, alloc["job_id"], new_user_id, manual=True, admin_id=current_user()["id"])
        con.commit()
        audit(current_user()["id"], "REALLOCATE_JOB", f"Reallocated job #{alloc['job_id']} to user ID {new_user_id}")
        flash("Job successfully reallocated to the selected user.", "success")
    except Exception as e:
        con.rollback()
        flash(f"Error reallocating job: {e}", "danger")

    return redirect(request.referrer or url_for("admin_session_tables"))


@app.route("/admin/job/<int:job_id>/delete", methods=["POST"])
@login_required
@admin_required
def admin_delete_job(job_id):
    con = get_db()
    job = con.execute("SELECT * FROM jobs WHERE id=?", (job_id,)).fetchone()
    if not job:
        abort(404)

    try:
        con.execute("BEGIN IMMEDIATE")
        con.execute("DELETE FROM reports WHERE job_id=?", (job_id,))
        con.execute("DELETE FROM allocations WHERE job_id=?", (job_id,))
        con.execute("DELETE FROM jobs WHERE id=?", (job_id,))
        con.commit()
        audit(current_user()["id"], "DELETE_JOB", f"Deleted job ID {job_id} ({job['mat_no']} / {job['course_code']})")
        flash("Job record and its allocations deleted entirely.", "success")
    except Exception as e:
        con.rollback()
        flash(f"Error deleting job: {e}", "danger")

    return redirect(request.referrer or url_for("upload_records"))


@app.route("/admin/job/<int:job_id>/edit", methods=["POST"])
@login_required
@admin_required
def admin_edit_job(job_id):
    con = get_db()
    job = con.execute("SELECT * FROM jobs WHERE id=?", (job_id,)).fetchone()
    if not job:
        abort(404)

    mat_no = request.form.get("mat_no", "").strip().lower()
    course_code = request.form.get("course_code", "").strip().upper()
    student_name = request.form.get("student_name", "").strip()
    department = request.form.get("department", "").strip()

    if not re.fullmatch(r"nou\d{9}", mat_no):
        flash("Matriculation number must be nou followed by 9 digits.", "danger")
        return redirect(request.referrer or url_for("upload_records"))
    if not re.fullmatch(r"[A-Z]{3}\d{3}", course_code):
        flash("Course code must be 3 letters followed by 3 digits.", "danger")
        return redirect(request.referrer or url_for("upload_records"))

    try:
        con.execute("""
            UPDATE jobs
            SET mat_no=?, course_code=?, student_name=?, department=?
            WHERE id=?
        """, (mat_no, course_code, student_name, department, job_id))
        con.commit()
        audit(current_user()["id"], "EDIT_JOB", f"Edited job ID {job_id} ({mat_no} / {course_code})")
        flash("Job record updated successfully.", "success")
    except sqlite3.IntegrityError:
        flash("A job with this matriculation number and course code already exists in this session.", "danger")
    except Exception as e:
        flash(f"Error updating job: {e}", "danger")

    return redirect(request.referrer or url_for("upload_records"))


@app.route("/admin/session/<int:sid>/clear", methods=["POST"])
@login_required
@admin_required
def admin_clear_session(sid):
    con = get_db()
    target = get_session_by_id(con, sid)
    if not target:
        abort(404)

    try:
        con.execute("BEGIN IMMEDIATE")
        job_ids = [row["id"] for row in con.execute("SELECT id FROM jobs WHERE session_id=?", (sid,)).fetchall()]
        if job_ids:
            placeholders = ",".join(["?"] * len(job_ids))
            con.execute(f"DELETE FROM reports WHERE job_id IN ({placeholders})", job_ids)
            con.execute(f"DELETE FROM allocations WHERE job_id IN ({placeholders})", job_ids)
            con.execute("DELETE FROM jobs WHERE session_id=?", (sid,))
        con.commit()
        audit(current_user()["id"], "CLEAR_SESSION", f"Cleared all jobs in session {target['name']}")
        flash(f"All records in {target['name']} section cleared successfully. You can now re-upload.", "success")
    except Exception as e:
        con.rollback()
        flash(f"Error clearing session section: {e}", "danger")

    return redirect(request.referrer or url_for("upload_records"))


@app.route("/admin/clear-all-records", methods=["POST"])
@login_required
@admin_required
def admin_clear_all_records():
    con = get_db()
    try:
        con.execute("BEGIN IMMEDIATE")
        con.execute("DELETE FROM reports")
        con.execute("DELETE FROM allocations")
        con.execute("DELETE FROM jobs")
        con.commit()
        audit(current_user()["id"], "CLEAR_ALL_RECORDS", "Deleted all uploaded records and allocations across all sessions.")
        flash("All uploaded records and allocation history have been completely deleted.", "success")
    except Exception as e:
        con.rollback()
        flash(f"Error deleting all records: {e}", "danger")

    return redirect(request.referrer or url_for("upload_records"))


@app.route("/admin/end-all-allocations", methods=["POST"])
@login_required
@admin_required
def admin_end_all_allocations():
    con = get_db()
    try:
        con.execute("BEGIN IMMEDIATE")

        active_allocs = con.execute("""
            SELECT a.id, a.job_id FROM allocations a
            WHERE a.status IN ('allocated', 'shown', 'started')
        """).fetchall()

        for a in active_allocs:
            con.execute("""
                UPDATE allocations SET status='finished', finished_at=?, submitted_at=?, score=100
                WHERE id=?
            """, (now(), now(), a["id"]))
            con.execute("UPDATE jobs SET status='finished' WHERE id=?", (a["job_id"],))

        con.commit()
        audit(current_user()["id"], "END_ALL_ALLOCATIONS", f"Force ended {len(active_allocs)} active user allocations.")
        flash(f"Successfully ended and closed {len(active_allocs)} active user allocation(s). Users are now cleared.", "success")
    except Exception as e:
        con.rollback()
        flash(f"Error ending all allocations: {e}", "danger")

    return redirect(request.referrer or url_for("admin_dashboard"))


@app.route("/admin/session/<int:sid>/add-job", methods=["POST"])
@login_required
@admin_required
def admin_add_job(sid):
    con = get_db()
    target = get_session_by_id(con, sid)
    if not target:
        abort(404)

    mat_no = request.form.get("mat_no", "").strip().lower()
    course_code = request.form.get("course_code", "").strip().upper()
    student_name = request.form.get("student_name", "").strip()
    department = request.form.get("department", "").strip()

    if not re.fullmatch(r"nou\d{9}", mat_no):
        flash("Matriculation number must be nou followed by 9 digits.", "danger")
        return redirect(request.referrer or url_for("upload_records"))
    if not re.fullmatch(r"[A-Z]{3}\d{3}", course_code):
        flash("Course code must be 3 letters followed by 3 digits.", "danger")
        return redirect(request.referrer or url_for("upload_records"))

    try:
        con.execute("""
            INSERT INTO jobs(session_id, mat_no, course_code, student_name, department, status, created_at)
            VALUES(?, ?, ?, ?, ?, 'available', ?)
        """, (sid, mat_no, course_code, student_name, department, now()))
        con.commit()
        audit(current_user()["id"], "ADD_JOB", f"Added job {mat_no} / {course_code} to session {target['name']}")
        flash("New record added to table successfully.", "success")
    except sqlite3.IntegrityError:
        flash("This matriculation number and course code already exists in this session.", "danger")
    except Exception as e:
        flash(f"Error adding job: {e}", "danger")

    return redirect(request.referrer or url_for("upload_records"))


@app.route("/admin/allocation/<int:allocation_id>/remove-replace", methods=["POST"])
@login_required
@admin_required
def admin_remove_and_replace(allocation_id):
    con = get_db()
    alloc = con.execute("""
        SELECT a.*, j.session_id, j.mat_no, j.course_code
        FROM allocations a JOIN jobs j ON j.id = a.job_id
        WHERE a.id=?
    """, (allocation_id,)).fetchone()
    if not alloc:
        abort(404)

    user_id = alloc["user_id"]
    session_id = alloc["session_id"]
    job_id = alloc["job_id"]

    try:
        con.execute("BEGIN IMMEDIATE")
        con.execute("UPDATE allocations SET status='reported' WHERE id=?", (allocation_id,))
        con.execute("UPDATE jobs SET status='available' WHERE id=?", (job_id,))

        replacement = available_job(con, session_id)
        replacement_id = None
        if replacement:
            replacement_id = create_allocation(
                con, replacement["id"], user_id, manual=True,
                admin_id=current_user()["id"], replacement_for=allocation_id
            )
            flash(f"Record removed from user and replacement job #{replacement_id} automatically assigned.", "success")
        else:
            flash("Record removed from user, but no other available jobs remain in this session for auto-reassignment.", "warning")

        con.commit()
        audit(current_user()["id"], "REMOVE_AND_REPLACE", f"Allocation #{allocation_id} removed, replacement #{replacement_id} assigned.")
    except Exception as e:
        con.rollback()
        flash(f"Error in reassigning: {e}", "danger")

    return redirect(url_for("admin_jobs"))


@app.route("/admin/allocation/<int:allocation_id>/remove-no-action", methods=["POST"])
@login_required
@admin_required
def admin_remove_no_action(allocation_id):
    con = get_db()
    alloc = con.execute("""
        SELECT a.*, j.id AS job_id, j.mat_no, j.course_code
        FROM allocations a JOIN jobs j ON j.id = a.job_id
        WHERE a.id=?
    """, (allocation_id,)).fetchone()
    if not alloc:
        abort(404)

    job_id = alloc["job_id"]

    try:
        con.execute("BEGIN IMMEDIATE")
        con.execute("UPDATE allocations SET status='paused' WHERE id=?", (allocation_id,))
        con.execute("UPDATE jobs SET status='available' WHERE id=?", (job_id,))
        con.commit()
        audit(current_user()["id"], "REMOVE_NO_ACTION", f"Allocation #{allocation_id} unassigned, job released to available.")
        flash("Record unassigned from user and job released back to available with no further action.", "success")
    except Exception as e:
        con.rollback()
        flash(f"Error: {e}", "danger")

    return redirect(url_for("admin_jobs"))


@app.route("/admin/bulk-allocate", methods=["POST"])
@login_required
@admin_required
def admin_bulk_allocate():
    con = get_db()
    try:
        user_id = int(request.form["user_id"])
        session_id = int(request.form["session_id"])
        quantity = int(request.form["quantity"])

        if quantity < 1 or quantity > 100:
            raise ValueError("Please specify between 1 and 100 jobs to allocate.")

        target = get_session_by_id(con, session_id)
        if not target or not target["is_open"]:
            raise ValueError("Please select an open session.")

        user = con.execute("""
            SELECT * FROM users WHERE id=? AND role='user'
            AND approved=1 AND active=1
        """, (user_id,)).fetchone()
        if not user:
            raise ValueError("Please select an approved active user.")

        allocated = 0
        con.execute("BEGIN IMMEDIATE")
        allocated = allocate_jobs_to_user(con, session_id, user_id, quantity, manual=True, admin_id=current_user()["id"])
        con.commit()

        audit(current_user()["id"], "BULK_ALLOCATION",
              f"Bulk allocated {allocated} jobs to user {user['username']} for {target['name']}")
        notify_admin(f"Admin bulk allocated {allocated} jobs to {user['username']} for {target['name']}.")
        flash(f"Successfully bulk allocated {allocated} job(s) to {user['username']}.", "success")
    except Exception as error:
        con.rollback()
        flash(str(error), "danger")

    return redirect(url_for("admin_allocations"))


@app.route("/admin/allocations", methods=["GET", "POST"])
@login_required
@admin_required
def admin_allocations():
    con = get_db()
    users = con.execute("""
        SELECT id,username FROM users
        WHERE role='user' AND approved=1 AND active=1
        ORDER BY username
    """).fetchall()
    sessions = con.execute("""
        SELECT * FROM work_sessions WHERE is_open=1 ORDER BY sequence
    """).fetchall()

    if request.method == "POST":
        try:
            user_id = int(request.form["user_id"])
            session_id = int(request.form["session_id"])
            mat_no = request.form["mat_no"].strip().lower()
            course = request.form["course_code"].strip().upper()

            if not re.fullmatch(r"nou\d{9}", mat_no):
                raise ValueError("Matriculation number must be nou followed by 9 digits.")
            if not re.fullmatch(r"[A-Z]{3}\d{3}", course):
                raise ValueError("Course code must be 3 letters followed by 3 digits.")

            target = get_session_by_id(con, session_id)
            if not target or not target["is_open"]:
                raise ValueError("Please select an open session.")

            user = con.execute("""
                SELECT * FROM users WHERE id=? AND role='user'
                AND approved=1 AND active=1
            """, (user_id,)).fetchone()
            if not user:
                raise ValueError("Please select an approved active user.")

            job = con.execute("""
                SELECT * FROM jobs
                WHERE session_id=? AND mat_no=? AND course_code=?
            """, (session_id, mat_no, course)).fetchone()

            if not job:
                con.execute("""
                    INSERT INTO jobs(session_id,mat_no,course_code,status,created_at)
                    VALUES(?,?,?,'available',?)
                """, (session_id, mat_no, course, now()))
                job = con.execute(
                    "SELECT * FROM jobs WHERE id=last_insert_rowid()"
                ).fetchone()

            if job["status"] != "available":
                raise ValueError("This job is already allocated or under review.")

            con.execute("BEGIN IMMEDIATE")
            allocation_id = create_allocation(
                con, job["id"], user_id, manual=True,
                admin_id=current_user()["id"]
            )
            con.commit()

            audit(current_user()["id"], "MANUAL_ALLOCATION",
                  f"Allocation #{allocation_id}: {user['username']} / "
                  f"{target['name']} / {mat_no} / {course}")
            notify_admin(
                f"Manual allocation #{allocation_id}: {mat_no} / {course} "
                f"assigned to {user['username']} for {target['name']}."
            )
            flash("Manual allocation completed successfully.", "success")
            return redirect(url_for("admin_allocations"))

        except Exception as error:
            con.rollback()
            flash(str(error), "danger")

    rows = con.execute("""
        SELECT a.*,u.username,j.mat_no,j.course_code,s.name AS session_name
        FROM allocations a
        JOIN users u ON u.id=a.user_id
        JOIN jobs j ON j.id=a.job_id
        JOIN work_sessions s ON s.id=j.session_id
        ORDER BY a.id DESC LIMIT 300
    """).fetchall()

    unallocated_jobs = con.execute("""
        SELECT j.*, s.name AS session_name
        FROM jobs j
        JOIN work_sessions s ON s.id=j.session_id
        WHERE j.status='available'
        ORDER BY j.id DESC LIMIT 300
    """).fetchall()

    return render_template("admin/allocations.html", user=current_user(), users=users, sessions=sessions, rows=rows, unallocated_jobs=unallocated_jobs)


@app.route("/admin/reports")
@login_required
@admin_required
def admin_reports():
    con = get_db()
    reports = con.execute("""
        SELECT r.*,u.username,j.mat_no,j.course_code,s.name AS session_name,
               a.status AS allocation_status
        FROM reports r
        JOIN users u ON u.id=r.user_id
        JOIN jobs j ON j.id=r.job_id
        JOIN work_sessions s ON s.id=j.session_id
        JOIN allocations a ON a.id=r.allocation_id
        ORDER BY r.id DESC
    """).fetchall()

    reporting_users = con.execute("""
        SELECT DISTINCT u.id, u.username
        FROM reports r
        JOIN users u ON u.id = r.user_id
        ORDER BY u.username
    """).fetchall()

    return render_template("admin/reports.html", user=current_user(), reports=reports, reporting_users=reporting_users)


@app.route("/admin/export-reports-csv")
@login_required
@admin_required
def export_reports_csv():
    con = get_db()
    reports = con.execute("""
        SELECT r.id, u.username, s.name AS session_name, j.mat_no, j.course_code,
               r.reason, r.explanation, r.status, r.created_at
        FROM reports r
        JOIN users u ON u.id = r.user_id
        JOIN jobs j ON j.id = r.job_id
        JOIN work_sessions s ON s.id = j.session_id
        ORDER BY r.id DESC
    """).fetchall()

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["Report ID", "User", "Session", "Matriculation No", "Course Code", "Reason", "Explanation", "Status", "Created At"])
    for r in reports:
        writer.writerow([r["id"], r["username"], r["session_name"], r["mat_no"], r["course_code"], r["reason"], r["explanation"], r["status"], r["created_at"]])

    return Response(
        output.getvalue(),
        mimetype="text/csv",
        headers={"Content-Disposition": "attachment; filename=reported_cases_export.csv"}
    )


@app.route("/admin/user/<int:uid>/export-reports-csv")
@login_required
@admin_required
def export_user_reports_csv(uid):
    con = get_db()
    user_target = con.execute("SELECT * FROM users WHERE id=?", (uid,)).fetchone()
    if not user_target:
        abort(404)

    reports = con.execute("""
        SELECT r.id, s.name AS session_name, j.mat_no, j.course_code,
               r.reason, r.explanation, r.status, r.created_at
        FROM reports r
        JOIN jobs j ON j.id = r.job_id
        JOIN work_sessions s ON s.id = j.session_id
        WHERE r.user_id = ?
        ORDER BY r.id DESC
    """, (uid,)).fetchall()

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["Report ID", "Session", "Matriculation No", "Course Code", "Reason", "Explanation", "Status", "Created At"])
    for r in reports:
        writer.writerow([r["id"], r["session_name"], r["mat_no"], r["course_code"], r["reason"], r["explanation"], r["status"], r["created_at"]])

    safe_username = secure_filename(user_target["username"])
    return Response(
        output.getvalue(),
        mimetype="text/csv",
        headers={"Content-Disposition": f"attachment; filename=reports_user_{safe_username}.csv"}
    )


@app.route("/admin/report/<int:rid>/resolve", methods=["POST"])
@login_required
@admin_required
def resolve_report(rid):
    con = get_db()
    con.execute("""
        UPDATE reports SET status='resolved',resolved_at=?
        WHERE id=? AND status='open'
    """, (now(), rid))
    con.commit()
    audit(current_user()["id"], "RESOLVE_REPORT", f"Report ID: {rid}")
    flash("Report marked as resolved.", "success")
    return redirect(url_for("admin_reports"))


@app.route("/admin/audit-log")
@login_required
@admin_required
def admin_audit_log():
    con = get_db()
    logs = con.execute("""
        SELECT l.*, u.username AS admin_username
        FROM audit_log l
        LEFT JOIN users u ON u.id = l.admin_id
        ORDER BY l.id DESC LIMIT 300
    """).fetchall()
    return render_template("admin/audit_log.html", user=current_user(), logs=logs)


@app.route("/admin/export.csv")
@login_required
@admin_required
def export_csv():
    con = get_db()
    rows = con.execute("""
        SELECT u.username,s.name AS session_name,j.mat_no,j.course_code,
               j.student_name,j.signature_id,j.department,j.source,
               a.status,a.score,a.manual,a.created_at,a.started_at,a.finished_at
        FROM allocations a
        JOIN users u ON u.id=a.user_id
        JOIN jobs j ON j.id=a.job_id
        JOIN work_sessions s ON s.id=j.session_id
        ORDER BY a.id
    """).fetchall()

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow([
        "Username", "Session", "Matriculation Number", "Course Code",
        "Student Name", "Signature ID", "Department", "Source",
        "Status", "Score", "Manual Allocation", "Allocated At",
        "Started At", "Finished At"
    ])
    for row in rows:
        writer.writerow(list(row))

    return Response(
        output.getvalue(),
        mimetype="text/csv",
        headers={"Content-Disposition": "attachment; filename=allocations.csv"}
    )


@app.route("/admin/export-completed-csv")
@login_required
@admin_required
def export_completed_csv():
    con = get_db()
    rows = con.execute("""
        SELECT j.mat_no, j.course_code, j.student_name, s.name AS session_name, a.score, a.finished_at
        FROM allocations a
        JOIN jobs j ON j.id = a.job_id
        JOIN work_sessions s ON s.id = j.session_id
        WHERE a.status IN ('finished', 'scored')
        ORDER BY a.id DESC
    """).fetchall()

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["Matriculation Number", "Course Code", "Student Name", "Session", "Score", "Finished At"])
    for r in rows:
        writer.writerow([r["mat_no"], r["course_code"], r["student_name"] or "", r["session_name"], r["score"] or "", r["finished_at"] or ""])

    return Response(
        output.getvalue(),
        mimetype="text/csv",
        headers={"Content-Disposition": "attachment; filename=completed_jobs_export.csv"}
    )


@app.route("/admin/session/<int:sid>/export-completed-csv")
@login_required
@admin_required
def export_session_completed_csv(sid):
    con = get_db()
    target = get_session_by_id(con, sid)
    if not target:
        abort(404)

    rows = con.execute("""
        SELECT j.mat_no, j.course_code, j.student_name, u.username AS completed_by, a.score, a.finished_at
        FROM allocations a
        JOIN jobs j ON j.id = a.job_id
        JOIN users u ON u.id = a.user_id
        WHERE j.session_id = ? AND a.status IN ('finished', 'scored')
        ORDER BY a.id DESC
    """, (sid,)).fetchall()

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["Matriculation Number", "Course Code", "Student Name", "Completed By", "Score", "Finished At"])
    for r in rows:
        writer.writerow([r["mat_no"], r["course_code"], r["student_name"] or "", r["completed_by"], r["score"] or "", r["finished_at"] or ""])

    safe_name = secure_filename(target["name"])
    return Response(
        output.getvalue(),
        mimetype="text/csv",
        headers={"Content-Disposition": f"attachment; filename=completed_jobs_session_{safe_name}.csv"}
    )


# ============================================================
# ERROR HANDLERS & INITIALIZATION
# ============================================================

@app.errorhandler(400)
def bad_request(error):
    msg = getattr(error, 'description', 'Invalid form request.')
    flash(msg, "warning")
    return redirect(request.referrer or url_for("dashboard"))


@app.errorhandler(403)
def forbidden(error):
    return render_template("errors/403.html", user=current_user()), 403


@app.errorhandler(404)
def not_found(error):
    return render_template("errors/404.html", user=current_user()), 404


@app.errorhandler(413)
def file_too_large(error):
    flash("Uploaded file is too large. Maximum size allowed is 25MB.", "danger")
    return redirect(url_for("upload_records"))


# @app.errorhandler(500)
# def internal_error(error):
#     flash("An internal server error occurred. Please try again.", "danger")
#     return redirect(url_for("dashboard"))


init_db()

if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
