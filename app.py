#!/usr/bin/env python3
"""OAK Builders — Bidding Tracker Web App"""
import os, sys, secrets, json, hashlib
from datetime import datetime, timedelta
from pathlib import Path
from functools import wraps

sys.path.insert(0, os.path.dirname(__file__))

from flask import (Flask, render_template, request, redirect, url_for, jsonify,
                   session, flash, g, send_from_directory)
from models import (get_db, init_db, get_project_summary, get_doc_readiness,
                    safe_float, safe_date_str,
                    STATUS_OPTIONS, DOC_STATUS_OPTIONS, MILESTONE_STATUS_OPTIONS,
                    DOCUMENTS, MILESTONES)

# Statuses that go to the separate "archived" table and don't get auto-scanned
ARCHIVED_STATUSES = ('LOST', 'NOT BIDDING', 'FOLLOWING UP', 'NOPE')

# ── Cloud vs Local mode ──
CLOUD_MODE = os.environ.get("CLOUD_MODE", "0") == "1"

app = Flask(__name__)
app.permanent_session_lifetime = timedelta(days=7)

# ── Config ──
BIDDING_DIR = Path(os.environ.get("BIDDING_DIR", str(Path(__file__).parent.parent.resolve())))
CONFIG_FILE = Path(__file__).parent / "config.json"

def load_config():
    if CONFIG_FILE.exists():
        with open(CONFIG_FILE) as f:
            return json.load(f)
    return {}

def save_config(cfg):
    with open(CONFIG_FILE, "w") as f:
        json.dump(cfg, f, indent=2)

def get_or_create_secret_key():
    """Persist secret key in config so sessions survive restarts."""
    cfg = load_config()
    if "secret_key" not in cfg:
        cfg["secret_key"] = secrets.token_hex(32)
        save_config(cfg)
    return cfg["secret_key"]

app.secret_key = os.environ.get("SECRET_KEY") or get_or_create_secret_key()

# ── HTTPS / Proxy support for cloud hosting (Render, Railway, etc.) ──
if CLOUD_MODE:
    from werkzeug.middleware.proxy_fix import ProxyFix
    app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)
    app.config['SESSION_COOKIE_SECURE'] = True
    app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'

def hash_password(pw):
    return hashlib.sha256(pw.encode()).hexdigest()

def is_setup_done():
    cfg = load_config()
    return bool(cfg.get("password_hash"))


# ── Auth ──
def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not is_setup_done():
            return redirect(url_for("setup"))
        if not session.get("logged_in"):
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated


@app.route("/setup", methods=["GET", "POST"])
def setup():
    if is_setup_done():
        return redirect(url_for("login"))
    if request.method == "POST":
        pw = request.form.get("password", "")
        if len(pw) < 4:
            flash("Password must be at least 4 characters.", "error")
            return render_template("setup.html")
        cfg = load_config()
        cfg["password_hash"] = hash_password(pw)
        cfg["created_at"] = datetime.now().isoformat()
        save_config(cfg)
        session.permanent = True
        session["logged_in"] = True
        flash("Setup complete! Welcome to your Bidding Tracker.", "success")
        return redirect(url_for("dashboard"))
    return render_template("setup.html")


@app.route("/login", methods=["GET", "POST"])
def login():
    if not is_setup_done():
        return redirect(url_for("setup"))
    if request.method == "POST":
        pw = request.form.get("password", "")
        cfg = load_config()
        if hash_password(pw) == cfg.get("password_hash"):
            session.permanent = True
            session["logged_in"] = True
            return redirect(url_for("dashboard"))
        flash("Incorrect password.", "error")
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/reset-password", methods=["GET", "POST"])
def reset_password():
    """Reset password — accessible by URL only (no link on the site)."""
    if request.method == "POST":
        pw = request.form.get("password", "")
        if len(pw) < 4:
            flash("Password must be at least 4 characters.", "error")
            return render_template("setup.html")
        cfg = load_config()
        cfg["password_hash"] = hash_password(pw)
        save_config(cfg)
        session.clear()
        flash("Password updated. Please log in with your new password.", "success")
        return redirect(url_for("login"))
    return render_template("setup.html")


# ── Dashboard ──
@app.route("/")
@login_required
def dashboard():
    conn = get_db()
    projects = conn.execute("SELECT * FROM projects ORDER BY deadline ASC").fetchall()
    conn.close()

    enriched = []
    for p in projects:
        ep = get_project_summary(p)
        ep['doc_readiness'] = get_doc_readiness(p['id'])
        enriched.append(ep)

    # Sort: deadline ascending (NULLs last), then readiness descending
    enriched.sort(key=lambda p: (
        0 if p.get('deadline') else 1,          # projects with deadlines first
        p.get('deadline') or '9999-12-31',       # earliest deadline first
        -(p.get('readiness') or 0),              # higher readiness first
    ))

    # Split into active and archived
    active_projects = [p for p in enriched if p['status'] not in ARCHIVED_STATUSES]
    archived_projects = [p for p in enriched if p['status'] in ARCHIVED_STATUSES]

    # Read filters from query string
    filter_status = request.args.get('status', '')
    filter_urgency = request.args.get('urgency', '')

    # Apply filters to active projects only
    filtered = active_projects
    if filter_status:
        filtered = [p for p in filtered if p['status'] == filter_status]
    if filter_urgency:
        filtered = [p for p in filtered if p['urgency'] == filter_urgency]

    # Stats (always based on full active list, not filtered)
    urgent = [p for p in active_projects if p['urgency'] == 'urgent']
    expired = [p for p in active_projects if p['urgency'] == 'expired']
    total_pipeline = sum(p['bid_price'] for p in active_projects if p['bid_price'] > 0)

    # Status counts (all projects)
    status_counts = {}
    for p in enriched:
        s = p['status']
        status_counts[s] = status_counts.get(s, 0) + 1

    # Unique statuses/urgencies for filter dropdowns (active only)
    active_statuses = sorted(set(p['status'] for p in active_projects))
    active_urgencies = sorted(set(p['urgency'] for p in active_projects))

    return render_template("dashboard.html",
        projects=filtered, archived_projects=archived_projects,
        active_count=len(active_projects),
        urgent_count=len(urgent), expired_count=len(expired),
        total_pipeline=total_pipeline, status_counts=status_counts,
        filter_status=filter_status, filter_urgency=filter_urgency,
        active_statuses=active_statuses, active_urgencies=active_urgencies,
        cloud_mode=CLOUD_MODE,
        now=datetime.now())


# ── Project Detail ──
@app.route("/project/<int:pid>")
@login_required
def project_detail(pid):
    conn = get_db()
    project = conn.execute("SELECT * FROM projects WHERE id = ?", (pid,)).fetchone()
    if not project:
        flash("Project not found.", "error")
        return redirect(url_for("dashboard"))

    docs = conn.execute("SELECT * FROM documents WHERE project_id = ? ORDER BY id", (pid,)).fetchall()
    miles = conn.execute("SELECT * FROM milestones WHERE project_id = ? ORDER BY id", (pid,)).fetchall()
    logs = conn.execute("SELECT * FROM activity_log WHERE project_id = ? ORDER BY created_at DESC LIMIT 20", (pid,)).fetchall()
    conn.close()

    ep = get_project_summary(project)
    ep['doc_readiness'] = get_doc_readiness(pid)

    return render_template("project.html",
        project=ep, docs=docs, milestones=miles, logs=logs,
        status_options=STATUS_OPTIONS, doc_status_options=DOC_STATUS_OPTIONS,
        milestone_status_options=MILESTONE_STATUS_OPTIONS)


# ── API: Update Project ──
@app.route("/api/project/<int:pid>", methods=["POST"])
@login_required
def update_project(pid):
    conn = get_db()
    data = request.get_json() if request.is_json else request.form.to_dict()

    allowed_fields = ['name', 'deadline', 'win_score', 'readiness', 'agency', 'status',
                      'scope', 'contact_info', 'notes', 'labor', 'materials', 'equipment',
                      'subcontractors', 'overhead', 'profit_pct']
    updates = []
    values = []
    for field in allowed_fields:
        if field in data:
            updates.append(f"{field} = ?")
            val = data[field]
            if field in ('win_score', 'readiness', 'labor', 'materials', 'equipment',
                         'subcontractors', 'overhead', 'profit_pct'):
                val = safe_float(val)
            elif field == 'deadline':
                val = safe_date_str(val)
            values.append(val)

    if updates:
        updates.append("updated_at = ?")
        values.append(datetime.now().isoformat())
        values.append(pid)
        conn.execute(f"UPDATE projects SET {', '.join(updates)} WHERE id = ?", values)
        conn.execute("INSERT INTO activity_log (project_id, action, detail) VALUES (?, ?, ?)",
                     (pid, "update", f"Updated: {', '.join(data.keys())}"))
        conn.commit()

    conn.close()
    if request.is_json:
        return jsonify({"ok": True})
    return redirect(url_for("project_detail", pid=pid))


# ── API: Update Document Status ──
@app.route("/api/doc/<int:doc_id>", methods=["POST"])
@login_required
def update_doc(doc_id):
    conn = get_db()
    data = request.get_json() if request.is_json else request.form.to_dict()
    status = data.get("status")
    if status and status in DOC_STATUS_OPTIONS:
        conn.execute("UPDATE documents SET status = ? WHERE id = ?", (status, doc_id))
        doc = conn.execute("SELECT * FROM documents WHERE id = ?", (doc_id,)).fetchone()
        if doc:
            conn.execute("INSERT INTO activity_log (project_id, action, detail) VALUES (?, ?, ?)",
                         (doc['project_id'], "doc_update", f"{doc['doc_name']}: {status}"))
        conn.commit()
    # Allow updating source/notes from the web UI
    source = data.get("source_file")
    if source is not None:
        conn.execute("UPDATE documents SET source_file = ? WHERE id = ?", (source, doc_id))
        conn.commit()
    notes = data.get("notes")
    if notes is not None:
        conn.execute("UPDATE documents SET notes = ? WHERE id = ?", (notes, doc_id))
        conn.commit()
    conn.close()
    return jsonify({"ok": True})


# ── API: Update Milestone ──
@app.route("/api/milestone/<int:mid>", methods=["POST"])
@login_required
def update_milestone(mid):
    conn = get_db()
    data = request.get_json() if request.is_json else request.form.to_dict()
    if "status" in data:
        conn.execute("UPDATE milestones SET status = ? WHERE id = ?", (data["status"], mid))
        conn.commit()
    if "milestone_date" in data:
        conn.execute("UPDATE milestones SET milestone_date = ? WHERE id = ?",
                     (safe_date_str(data["milestone_date"]), mid))
        conn.commit()
    conn.close()
    return jsonify({"ok": True})


# ── API: Add Project ──
@app.route("/api/project/new", methods=["POST"])
@login_required
def add_project():
    conn = get_db()
    data = request.get_json() if request.is_json else request.form.to_dict()
    name = data.get("name", "New Project")

    cursor = conn.execute("""
        INSERT INTO projects (name, folder_name, sheet_name, status)
        VALUES (?, ?, ?, 'NEED MORE INFO')
    """, (name, name, name[:31]))
    pid = cursor.lastrowid

    for doc_name in DOCUMENTS:
        conn.execute("INSERT INTO documents (project_id, doc_name) VALUES (?, ?)", (pid, doc_name))
    for mile_name in MILESTONES:
        conn.execute("INSERT INTO milestones (project_id, milestone_name) VALUES (?, ?)", (pid, mile_name))

    conn.execute("INSERT INTO activity_log (project_id, action, detail) VALUES (?, ?, ?)",
                 (pid, "created", f"New project: {name}"))
    conn.commit()
    conn.close()

    if request.is_json:
        return jsonify({"ok": True, "id": pid})
    return redirect(url_for("project_detail", pid=pid))


# ── API: Dashboard data (for live refresh) ──
@app.route("/api/dashboard")
@login_required
def api_dashboard():
    conn = get_db()
    projects = conn.execute("SELECT * FROM projects ORDER BY deadline ASC").fetchall()
    conn.close()
    result = []
    for p in projects:
        ep = get_project_summary(p)
        ep['doc_readiness'] = get_doc_readiness(p['id'])
        result.append(ep)
    # Sort: deadline ascending (NULLs last), then readiness descending
    result.sort(key=lambda p: (
        0 if p.get('deadline') else 1,
        p.get('deadline') or '9999-12-31',
        -(p.get('readiness') or 0),
    ))
    return jsonify(result)


# ── Run folder scan (reuse existing logic) ──
@app.route("/api/scan", methods=["POST"])
@login_required
def run_scan():
    """Trigger a document scan of project folders."""
    if CLOUD_MODE:
        return jsonify({"ok": True, "changes": [], "message": "Folder scan is disabled in cloud mode."})
    from scanner import scan_all_projects
    changes = scan_all_projects(BIDDING_DIR)
    return jsonify({"ok": True, "changes": changes})


# ── Always init DB on import (needed for Gunicorn) ──
init_db()

if __name__ == "__main__":
    # Check if migration needed
    conn = get_db()
    count = conn.execute("SELECT COUNT(*) as c FROM projects").fetchone()['c']
    conn.close()
    if count == 0:
        print("No projects in database. Run migrate_from_excel.py first.")
        print(f"  python {os.path.join(os.path.dirname(__file__), 'migrate_from_excel.py')}")

    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=os.environ.get("DEBUG", "0") == "1")
