#!/usr/bin/env bash
set -e
ROOT="$(cd "$(dirname "$0")" && pwd)"
echo "==> Backend em :8000"
cd "$ROOT/backend" && uvicorn main:app --host 0.0.0.0 --port 8000 --reload &
echo "==> Frontend em :5173"
cd "$ROOT/frontend" && npm run dev &
echo -e "\n  Backend : http://localhost:8000\n  Frontend: http://localhost:5173\n"
trap "kill %1 %2 2>/dev/null; exit 0" SIGINT SIGTERM
wait
