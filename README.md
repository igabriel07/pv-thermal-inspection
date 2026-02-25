# Thermal Fault Detection + Labeling (FastAPI + React)

React (Vite) frontend + FastAPI backend for scanning images for faults, reviewing detections, editing labels, and generating reports.

This repo is structured as a small monorepo:
- `backend/` = FastAPI API (serves `/api/*`)
- `frontend/` = React (Vite)

## Backend

1) Create a virtual environment (optional)
2) Install deps:
- `pip install -r backend/requirements.txt`
3) Run:
- `uvicorn backend.main:app --reload --port 8000`

Faster local dev tips:
- Skip the heavy YOLO preload on startup: `PRELOAD_MODEL=0 uvicorn backend.main:app --reload --port 8000`
- Disable reload (often faster on Windows): `uvicorn backend.main:app --port 8000`

Optional environment variables (see `.env.example`):
- `CORS_ALLOW_ORIGINS` (comma-separated, or `*`)
- `FAULT_MODEL_PATH`, `FAULT_IMGSZ`, `FAULT_CONF`, `FAULT_IOU`

## Frontend

1) Install deps:
- `npm install` (run inside frontend/)
2) Run:
- `npm run dev` (run inside frontend/)

Open http://localhost:5173

## Production-ish

If you deploy frontend and backend under different origins, set `VITE_API_BASE_URL` at build time for the frontend.

Docker (builds both services):
- `docker compose up --build`
