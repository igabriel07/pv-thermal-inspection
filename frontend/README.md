# FastAPI + React (Vite)

Simple one-page app with a left menu and a React frontend calling a FastAPI backend.

## Backend

From the workspace root:

1) Install deps
- `pip install -r backend/requirements.txt`
2) Run the API
- `uvicorn backend.main:app --reload --port 8000`

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
