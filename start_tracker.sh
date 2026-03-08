#!/bin/bash
# OAK BUILDERS - Bidding Tracker Web App (Mac/Linux)
echo ""
echo "  ============================================================"
echo "    OAK BUILDERS - Bidding Tracker (Web App)"
echo "  ============================================================"
echo ""

cd "$(dirname "$0")"

# Check Python
if ! command -v python3 &> /dev/null; then
    echo "  Python3 is not installed."
    echo "  Install from https://python.org or run: brew install python3"
    exit 1
fi

# Install dependencies if needed
python3 -c "import flask" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "  Installing required packages..."
    pip3 install -r requirements.txt -q
fi

# Get local IP for phone access
IP=$(ifconfig 2>/dev/null | grep "inet " | grep -v 127.0.0.1 | head -1 | awk '{print $2}')
if [ -z "$IP" ]; then
    IP=$(hostname -I 2>/dev/null | awk '{print $1}')
fi

echo "  Starting web server..."
echo ""
echo "  ────────────────────────────────────────"
echo "  Open in browser:"
echo ""
echo "    This Mac:  http://localhost:5000"
if [ -n "$IP" ]; then
echo "    Phone:     http://${IP}:5000"
fi
echo ""
echo "  ────────────────────────────────────────"
echo "  Press Ctrl+C to stop the server."
echo ""

python3 app.py
