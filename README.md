# FastAPI + React (Vite)

Simple one-page app with a left menu and a React frontend calling a FastAPI backend.

This repo is structured as a small monorepo:
- `backend/` = FastAPI API (serves `/api/*`)
- `frontend/` = React (Vite)

## Backend

1) Create a virtual environment (optional)
2) Install deps:
- `pip install -r backend/requirements.txt`
3) Run:
- `uvicorn backend.main:app --reload --port 8000`

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
