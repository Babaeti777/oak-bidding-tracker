"""Folder scanner — checks project folders for documents and updates the DB."""
import os
from pathlib import Path
from datetime import datetime
from models import get_db, DOCUMENTS

COMPANY_DOCS_SUBDIR = "1111-Claude/COMPANY_DOCUMENTS"

DOC_PATTERNS = {
    "Bid Bond / Bid Security": {
        "folders": ["Bonding", "CLAUDE", "."],
        "keywords": ["bond", "bid bond", "surety", "bid security"],
        "company_folders": ["Bonding"],
    },
    "Performance Bond": {
        "folders": ["Bonding", "CLAUDE"],
        "keywords": ["performance bond"],
        "company_folders": ["Bonding"],
    },
    "Payment Bond": {
        "folders": ["Bonding", "CLAUDE"],
        "keywords": ["payment bond"],
        "company_folders": ["Bonding"],
    },
    "Certificate of Insurance": {
        "folders": ["Insurance", "CLAUDE", "."],
        "keywords": ["coi", "certificate insurance", "insurance cert", "insurance policy"],
        "company_folders": ["Insurance"],
    },
    "Bid Form / Proposal Form": {
        "folders": [".", "CLAUDE", "Forms"],
        "keywords": ["bid form", "proposal form", "pricing sheet", "bid tab"],
        "company_folders": [],
    },
    "Non-Collusion Affidavit": {
        "folders": [".", "CLAUDE", "Forms"],
        "keywords": ["affidavit", "non-collusion", "collusion"],
        "company_folders": [],
    },
    "MBE/WBE Compliance": {
        "folders": [".", "CLAUDE"],
        "keywords": ["mbe", "wbe", "dbe", "sbe", "compliance cert", "small business", "minority"],
        "company_folders": [],
    },
    "Safety Plan / OSHA Logs": {
        "folders": ["Safety", "CLAUDE"],
        "keywords": ["safety", "osha", "emr", "safety plan", "injury illness"],
        "company_folders": ["Safety"],
    },
    "Financial Statements": {
        "folders": ["Financial", "CLAUDE"],
        "keywords": ["financial statement", "balance sheet", "income statement", "p&l", "profit loss"],
        "company_folders": ["Financial"],
    },
    "References / Past Projects": {
        "folders": [".", "CLAUDE", "Personnel"],
        "keywords": ["reference", "past project", "experience", "client ref"],
        "company_folders": ["Personnel"],
    },
    "Subcontractor List": {
        "folders": [".", "CLAUDE"],
        "keywords": ["subcontractor", "sub list", "vendor list"],
        "company_folders": [],
    },
    "Project Schedule": {
        "folders": [".", "CLAUDE"],
        "keywords": ["schedule", "timeline", "gantt", "milestone schedule"],
        "company_folders": [],
    },
}

EXTENSIONS = {".pdf", ".docx", ".xlsx", ".doc", ".xls", ".jpg", ".png"}


def _search_folder(folder, keywords):
    """Search a folder for files matching keywords."""
    if not folder.exists():
        return None
    for f in folder.iterdir():
        if f.is_file() and f.suffix.lower() in EXTENSIONS:
            name_lower = f.name.lower()
            if any(kw in name_lower for kw in keywords):
                return f.name
    return None


def scan_project_docs(project_folder, company_docs_dir):
    """Scan a single project's folders for documents. Returns dict of doc_name -> (status, source)."""
    results = {}
    for doc_name, config in DOC_PATTERNS.items():
        found = None
        source = ""

        # Search project folders
        for subfolder in config["folders"]:
            search_dir = project_folder if subfolder == "." else project_folder / subfolder
            match = _search_folder(search_dir, config["keywords"])
            if match:
                found = "Done"
                source = f"[Project] {match}"
                break

        # Search company docs
        if not found and company_docs_dir.exists():
            for subfolder in config.get("company_folders", []):
                search_dir = company_docs_dir / subfolder
                match = _search_folder(search_dir, config["keywords"])
                if match:
                    found = "Done"
                    source = f"[Company] {match}"
                    break

        if found:
            results[doc_name] = (found, source)

    return results


def scan_all_projects(bidding_dir):
    """Scan all project folders and update the database."""
    bidding_dir = Path(bidding_dir)
    company_docs_dir = bidding_dir / COMPANY_DOCS_SUBDIR
    conn = get_db()
    changes = []

    # Skip archived/inactive projects — they don't need constant scanning
    ARCHIVED_STATUSES = ('LOST', 'NOT BIDDING', 'FOLLOWING UP', 'NOPE')
    placeholders = ','.join('?' for _ in ARCHIVED_STATUSES)
    projects = conn.execute(
        f"SELECT id, folder_name FROM projects WHERE status NOT IN ({placeholders})",
        ARCHIVED_STATUSES
    ).fetchall()

    exclude = {"COMPANY_DOCUMENTS", "Claude", "Bid submitted", "claude", ".skills", ".claude",
               "0000-Bid submitted", "1111-Claude", "__pycache__", "node_modules", ".git", "webapp"}

    for proj in projects:
        folder_name = proj['folder_name']
        if not folder_name or folder_name in exclude:
            continue

        # Try to find folder
        project_folder = bidding_dir / folder_name
        if not project_folder.exists():
            # Try submitted folder
            submitted = bidding_dir / "0000-Bid submitted" / folder_name
            if submitted.exists():
                project_folder = submitted
            else:
                continue

        scan_results = scan_project_docs(project_folder, company_docs_dir)

        for doc_name, (new_status, source) in scan_results.items():
            doc = conn.execute(
                "SELECT id, status FROM documents WHERE project_id = ? AND doc_name = ?",
                (proj['id'], doc_name)
            ).fetchone()

            if doc and doc['status'] not in ('Done', 'N/A'):
                if new_status != doc['status']:
                    conn.execute("UPDATE documents SET status = ?, source_file = ? WHERE id = ?",
                                 (new_status, source, doc['id']))
                    changes.append(f"{folder_name}: {doc_name} → {new_status}")

    if changes:
        conn.execute("INSERT INTO activity_log (action, detail) VALUES (?, ?)",
                     ("scan", f"Found {len(changes)} document updates"))
    conn.commit()
    conn.close()
    return changes
