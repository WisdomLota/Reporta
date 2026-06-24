#!/usr/bin/env bash
# Convenience launcher for macOS/Linux. Starts backend + frontend together.
set -e
cd "$(dirname "$0")"
( cd backend && [ -d venv ] || python3 -m venv venv; source venv/bin/activate; pip install -q -r requirements.txt; uvicorn main:app --port 8000 ) &
BACK=$!
( cd frontend && [ -d node_modules ] || npm install; npm run dev )
kill $BACK 2>/dev/null || true
