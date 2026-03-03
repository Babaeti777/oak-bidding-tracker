# OAK Builders — Bidding Tracker

A lightweight, password-protected web app for tracking construction bids, documents, costs, and deadlines. Built for OAK Builders LLC.

## Features

- **Dashboard** — All active bids at a glance with urgency alerts, pipeline value, and status breakdown
- **Project Details** — Edit costs, deadlines, documents, milestones per project
- **Document Scanner** — Auto-detects documents in your project folders and updates status
- **Mobile-Friendly** — Responsive design; access from your phone on the same WiFi
- **Password Protected** — Set a password on first launch; sessions persist across restarts
- **SQLite Database** — No Excel lock-file issues; fast reads/writes
- **Excel Migration** — One-time import from an existing Bidding Tracker Pro Excel workbook

## Quick Start

### 1. Install Python (3.9+)
Download from [python.org](https://python.org). Make sure "Add to PATH" is checked.

### 2. Install dependencies
```
pip install -r requirements.txt
```

### 3. (Optional) Migrate from Excel
If you have an existing `Bidding_Tracker_Pro_v4.xlsx` in the parent directory:
```
python migrate_from_excel.py
```

### 4. Run the app
**Windows:** Double-click `START_TRACKER.bat`

**Manual:**
```
python app.py
```

Open `http://localhost:5000` in your browser. First visit will ask you to set a password.

### Phone Access
The `.bat` launcher shows your local IP. On your phone (same WiFi), open `http://<your-ip>:5000`.

## File Structure

```
webapp/
├── app.py                 # Flask web server
├── models.py              # Database schema and helpers
├── scanner.py             # Folder document scanner
├── migrate_from_excel.py  # One-time Excel → SQLite migration
├── requirements.txt       # Python dependencies
├── START_TRACKER.bat      # Windows launcher
├── templates/
│   ├── base.html          # Layout template
│   ├── dashboard.html     # Main dashboard
│   ├── project.html       # Project detail/edit
│   ├── login.html         # Login page
│   └── setup.html         # First-time setup
└── .gitignore
```

## Security Notes

- Password is hashed (SHA-256) and stored in `config.json` (not committed to git)
- Session secret key persists in `config.json` so logins survive restarts
- Database file (`bidding_tracker.db`) is excluded from git
- To reset your password, visit `/reset-password` in your browser

## Tech Stack

- Python 3 / Flask
- SQLite (via built-in `sqlite3`)
- Vanilla HTML/CSS/JS (no build step, no npm)
- openpyxl (for Excel migration only)
