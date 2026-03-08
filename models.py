"""OAK Builders Bidding Tracker — Database Models & Schema"""
import sqlite3
import os
from datetime import datetime, date

# Support persistent disk on cloud (Render mounts to /data)
_default_db = os.path.join(os.path.dirname(__file__), "bidding_tracker.db")
_env_db = os.environ.get("DATABASE_PATH", "")

if _env_db:
    # Try to use the configured path; fall back to local if directory can't be created
    try:
        os.makedirs(os.path.dirname(_env_db) or ".", exist_ok=True)
        DB_PATH = _env_db
    except OSError:
        print(f"WARNING: Cannot create directory for {_env_db}, falling back to local DB")
        DB_PATH = _default_db
else:
    DB_PATH = _default_db

DOCUMENTS = [
    "Bid Bond / Bid Security", "Performance Bond", "Payment Bond",
    "Certificate of Insurance", "Bid Form / Proposal Form", "Non-Collusion Affidavit",
    "MBE/WBE Compliance", "Safety Plan / OSHA Logs", "Financial Statements",
    "References / Past Projects", "Subcontractor List", "Project Schedule",
]

MILESTONES = [
    "Pre-Bid Meeting", "Site Visit", "RFI Deadline", "Bid Submission",
    "Award Announcement", "Contract Signing", "Notice to Proceed", "Project Completion",
]

STATUS_OPTIONS = ["YES", "NOPE", "NEED MORE INFO", "MAYBE", "PREPARING", "SUBMITTED", "NOT BIDDING", "WIN", "LOST", "FOLLOWING UP"]
DOC_STATUS_OPTIONS = ["Pending", "In Progress", "Done", "N/A"]
MILESTONE_STATUS_OPTIONS = ["Upcoming", "Complete", "Missed", "N/A"]


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def init_db():
    conn = get_db()
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS projects (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sheet_name TEXT UNIQUE,
            folder_name TEXT,
            name TEXT NOT NULL,
            deadline TEXT,
            win_score REAL DEFAULT 50,
            readiness REAL DEFAULT 10,
            agency TEXT DEFAULT 'TBD',
            status TEXT DEFAULT 'NEED MORE INFO',
            scope TEXT DEFAULT '',
            contact_info TEXT DEFAULT '',
            notes TEXT DEFAULT '',
            labor REAL DEFAULT 0,
            materials REAL DEFAULT 0,
            equipment REAL DEFAULT 0,
            subcontractors REAL DEFAULT 0,
            overhead REAL DEFAULT 0,
            profit_pct REAL DEFAULT 15,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS documents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL,
            doc_name TEXT NOT NULL,
            status TEXT DEFAULT 'Pending',
            source_file TEXT DEFAULT '',
            notes TEXT DEFAULT '',
            FOREIGN KEY (project_id) REFERENCES projects(id) ON DELETE CASCADE,
            UNIQUE(project_id, doc_name)
        );

        CREATE TABLE IF NOT EXISTS milestones (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL,
            milestone_name TEXT NOT NULL,
            milestone_date TEXT,
            status TEXT DEFAULT 'Upcoming',
            FOREIGN KEY (project_id) REFERENCES projects(id) ON DELETE CASCADE,
            UNIQUE(project_id, milestone_name)
        );

        CREATE TABLE IF NOT EXISTS activity_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER,
            action TEXT NOT NULL,
            detail TEXT DEFAULT '',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (project_id) REFERENCES projects(id) ON DELETE SET NULL
        );
    """)
    conn.commit()

    # Auto-seed from seed_data.json if DB is empty (first cloud deploy)
    count = conn.execute("SELECT COUNT(*) FROM projects").fetchone()[0]
    if count == 0:
        import json
        seed_file = os.path.join(os.path.dirname(__file__), "seed_data.json")
        if os.path.exists(seed_file):
            print("DATABASE EMPTY — seeding from seed_data.json ...")
            with open(seed_file) as f:
                seed = json.load(f)
            for table in ["projects", "documents", "milestones", "activity_log"]:
                rows = seed.get(table, [])
                if not rows:
                    continue
                cols = list(rows[0].keys())
                placeholders = ", ".join(["?"] * len(cols))
                col_names = ", ".join(cols)
                for row in rows:
                    vals = [row.get(c) for c in cols]
                    try:
                        conn.execute(f"INSERT OR IGNORE INTO {table} ({col_names}) VALUES ({placeholders})", vals)
                    except Exception as e:
                        print(f"  Seed warning ({table}): {e}")
            conn.commit()
            final = conn.execute("SELECT COUNT(*) FROM projects").fetchone()[0]
            print(f"  Seeded {final} projects successfully.")

    conn.close()


def safe_float(val, default=0):
    if val is None:
        return default
    try:
        return float(val)
    except (ValueError, TypeError):
        return default


def safe_date_str(val):
    """Convert various date formats to YYYY-MM-DD string or None."""
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, date):
        return val.strftime("%Y-%m-%d")
    s = str(val).strip()
    if s.upper() in ("", "TBD", "NONE"):
        return None
    # Try common formats
    for fmt in ["%Y-%m-%d", "%m/%d/%Y", "%B %d, %Y", "%b %d, %Y"]:
        try:
            return datetime.strptime(s[:10], fmt).strftime("%Y-%m-%d")
        except (ValueError, IndexError):
            continue
    return None


def get_project_summary(project_row):
    """Enrich a project row with computed fields."""
    p = dict(project_row)
    total_cost = sum([
        safe_float(p.get('labor')), safe_float(p.get('materials')),
        safe_float(p.get('equipment')), safe_float(p.get('subcontractors')),
        safe_float(p.get('overhead'))
    ])
    p['total_cost'] = total_cost
    p['bid_price'] = total_cost * (1 + safe_float(p.get('profit_pct'), 15) / 100) if total_cost > 0 else 0

    # Days until deadline
    if p.get('deadline'):
        try:
            dl = datetime.strptime(p['deadline'], "%Y-%m-%d")
            p['days_left'] = (dl - datetime.now()).days
        except (ValueError, TypeError):
            p['days_left'] = None
    else:
        p['days_left'] = None

    # Urgency
    if p.get('status') in ('NOT BIDDING', 'NOPE', 'LOST'):
        p['urgency'] = 'inactive'
    elif p.get('status') == 'WIN':
        p['urgency'] = 'ok'
    elif p['days_left'] is None:
        p['urgency'] = 'unknown'
    elif p['days_left'] < 0:
        p['urgency'] = 'expired'
    elif p['days_left'] <= 7:
        p['urgency'] = 'urgent'
    elif p['days_left'] <= 14:
        p['urgency'] = 'soon'
    else:
        p['urgency'] = 'ok'

    return p


def get_doc_readiness(project_id):
    """Calculate document readiness percentage for a project."""
    conn = get_db()
    docs = conn.execute("SELECT status FROM documents WHERE project_id = ?", (project_id,)).fetchall()
    conn.close()
    if not docs:
        return 0
    countable = [d for d in docs if d['status'] != 'N/A']
    if not countable:
        return 100
    done = sum(1 for d in countable if d['status'] == 'Done')
    return round(done / len(countable) * 100)
