# Thermal Fault Detection + Labeling (FastAPI + React)

React (Vite) frontend + FastAPI backend for scanning images for faults, reviewing detections, editing labels, and generating reports.

## Backend

From the workspace root:

1) Install deps
- `pip install -r backend/requirements.txt`
2) Run the API
- `uvicorn backend.main:app --reload --port 8000`

Faster local dev tips:
- Skip the heavy YOLO preload on startup: `PRELOAD_MODEL=0 uvicorn backend.main:app --reload --port 8000`
- Disable reload (often faster on Windows): `uvicorn backend.main:app --port 8000`

Notes:
- Backend CORS is configurable via `CORS_ALLOW_ORIGINS`.

## Frontend

From the workspace root:

1) Install deps
- `npm install`
2) Run the app
- `npm run dev`

Notes:
- Dev proxy target can be changed with `VITE_API_PROXY_TARGET` (defaults to `http://127.0.0.1:8000`).
- For separate frontend/backend deployments, build with `VITE_API_BASE_URL` (see `.env.example`).

Open http://localhost:5173
