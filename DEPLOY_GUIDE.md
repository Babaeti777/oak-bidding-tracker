# OAK Builders Bidding Tracker — Deployment Guide

## What Changed

Your app is now ready to run three ways:

1. **Locally on Windows** — double-click `START_TRACKER.bat` (same as before)
2. **Locally on Mac** — double-click `start_tracker.sh` (or run `bash start_tracker.sh` in Terminal)
3. **In the cloud** — deploy to Render.com for 24/7 access from your phone, any browser, anywhere

---

## Option 1: Keep Running Locally (Windows or Mac)

This is what you already have. It works great when you're at your office PC.

**Windows:** Double-click `START_TRACKER.bat`
**Mac:** Open Terminal, navigate to this folder, run `bash start_tracker.sh`

Both scripts show you the local URL and your network IP so you can access from your phone **when on the same Wi-Fi**.

---

## Option 2: Deploy to Render.com (Recommended for Phone Access)

This gives you a public URL like `oak-bidding-tracker.onrender.com` that works from anywhere — phone, tablet, another computer — even when your office PC is off.

### Step-by-Step Setup

#### 1. Create a GitHub account (if you don't have one)
- Go to https://github.com and sign up (free)

#### 2. Push your code to GitHub
Open a terminal/command prompt in this project folder and run:

```
git init
git add -A
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/oak-bidding-tracker.git
git push -u origin main
```

Replace `YOUR_USERNAME` with your GitHub username. You'll need to create the repository on GitHub first (click the "+" button → New repository).

#### 3. Sign up for Render.com
- Go to https://render.com and sign up with your GitHub account (free)
- Click **New → Web Service**
- Connect your GitHub repository (`oak-bidding-tracker`)

#### 4. Configure the service
Render will auto-detect the settings from the `render.yaml` file. Verify these:

| Setting | Value |
|---------|-------|
| Name | `oak-bidding-tracker` |
| Runtime | Python |
| Build Command | `pip install -r requirements.txt` |
| Start Command | `gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --timeout 120` |

#### 5. Set environment variables
In the Render dashboard, go to **Environment** and add:

| Key | Value |
|-----|-------|
| `CLOUD_MODE` | `1` |
| `SECRET_KEY` | *(click "Generate" for a random value)* |
| `DATABASE_PATH` | `/data/bidding_tracker.db` |

#### 6. Add a persistent disk
Under **Disks**, add one:
- **Name:** `data`
- **Mount Path:** `/data`
- **Size:** 1 GB

This keeps your database safe across deploys.

#### 7. Upload your existing data
After the first deploy, you'll need to migrate your data. The easiest way:
- Open the Render **Shell** tab
- Upload your `bidding_tracker.db` file to `/data/bidding_tracker.db`

#### 8. Done!
Your app will be live at `https://oak-bidding-tracker.onrender.com` (or whatever name you chose). Bookmark it on your phone.

---

## Render.com Free Tier Notes

- **Free tier** spins down after 15 minutes of inactivity. First visit after sleep takes ~30 seconds to wake up.
- **Paid tier ($7/mo)** keeps it always running — instant loads, no spin-down.
- Either way, your data is safe on the persistent disk.

---

## What's Different in Cloud Mode

When `CLOUD_MODE=1`:
- The "Scan Folders" button is hidden (no local folders to scan in the cloud)
- HTTPS proxy headers are handled correctly
- Session cookies are set to secure mode
- Everything else works the same — login, projects, documents, milestones

---

## Files Added/Changed

| File | Purpose |
|------|---------|
| `Procfile` | Tells cloud hosts how to run the app |
| `runtime.txt` | Specifies Python version |
| `render.yaml` | Auto-config for Render.com |
| `.env.example` | Template for environment variables |
| `start_tracker.sh` | Updated Mac/Linux startup script |
| `requirements.txt` | Added `gunicorn` for production server |
| `app.py` | Added cloud mode, proxy support, env var config |
| `models.py` | Database path now configurable via env var |
