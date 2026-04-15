#!/bin/bash

echo "========================================"
echo "   Excel to PPT Generator v6.0"
echo "========================================"
echo

cd "$(dirname "$0")"

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    echo "[ERROR] Python 3 is not installed"
    echo "Please install Python 3.8+ first"
    exit 1
fi

# Check if virtual environment exists
if [ ! -d "venv" ]; then
    echo "[INFO] Creating virtual environment..."
    python3 -m venv venv
    if [ $? -ne 0 ]; then
        echo "[ERROR] Failed to create virtual environment"
        exit 1
    fi
fi

# Activate virtual environment
echo "[INFO] Activating virtual environment..."
source venv/bin/activate

# Install dependencies
echo "[INFO] Installing dependencies..."
pip install -r requirements.txt --quiet

if [ $? -ne 0 ]; then
    echo "[ERROR] Failed to install dependencies"
    exit 1
fi

echo
echo "========================================"
echo "   Starting server..."
echo "   URL: http://127.0.0.1:8000/ (avoid localhost -> IPv6 mismatch)"
echo "   Press Ctrl+C to stop"
echo "========================================"
echo

# Foolproof: refuse to start if something already listens on 8000 (wrong app in browser)
if [ -z "${SKIP_PORT_CHECK:-}" ]; then
    if command -v ss >/dev/null 2>&1; then
        if ss -lnt 2>/dev/null | grep -qE ':(8000)\s'; then
            echo "[ERROR] Port 8000 is already in use. Run stop_server or free the port."
            echo "        (Advanced: export SKIP_PORT_CHECK=1 to skip this check)"
            exit 1
        fi
    elif command -v lsof >/dev/null 2>&1; then
        if lsof -iTCP:8000 -sTCP:LISTEN >/dev/null 2>&1; then
            echo "[ERROR] Port 8000 is already in use. Run stop_server or free the port."
            echo "        (Advanced: export SKIP_PORT_CHECK=1 to skip this check)"
            exit 1
        fi
    fi
fi

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
python "$SCRIPT_DIR/scripts/wait_and_open_browser.py" 127.0.0.1 8000 &

python -m uvicorn app.main:app --host 127.0.0.1 --port 8000
