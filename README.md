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
- `FAULT_MODEL_PATH`, `FAULT_TYPE_MODEL_PATH`
- `FAULT_IMGSZ`, `FAULT_CONF`, `FAULT_IOU`

Notes:
- YOLO weight files (`*.pt`, etc.) are intentionally ignored by git via `.gitignore`.
	Place them locally (default: `backend/models/`) or point to them with the env vars above.
	This repo allows tracking only `backend/models/best.pt` and `backend/models/best_8_class.pt` as explicit exceptions.

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

## Deploy online (Docker + HTTPS)

This repo includes a production compose file + reverse proxy setup:
- `docker-compose.prod.yml` (frontend + backend + Caddy)
- `Caddyfile` (routes `/api/*` to backend and everything else to frontend)

### Prereqs
- A Linux server/VPS with Docker + Compose installed
- A domain name pointing to your server IP (DNS A/AAAA record)
- YOLO weights available at `backend/models/best.pt` and `backend/models/best_8_class.pt` (either committed as repo exceptions or placed manually).

### Run
From the repo root on the server:
- `DOMAIN=app.example.com docker compose -f docker-compose.prod.yml up -d --build`

Then open:
- `https://app.example.com`

Notes:
- Ensure ports 80/443 are open in the firewall/security group.
- If you don't have a domain yet, you can run HTTP-only by setting `DOMAIN=:80` (no TLS).

## Deploy (free, no domain) on Fly.io

This repo includes a single-container Fly setup that:
- runs the FastAPI backend on `127.0.0.1:8000`
- serves the built frontend via Caddy on `:8080`
- proxies `/api/*` to the backend
- scales to zero (`min_machines_running = 0`) and auto-starts on request

Files:
- `fly.toml`
- `Dockerfile.fly`
- `Caddyfile.fly`

Steps:
1) Install `flyctl` and login: `fly auth login`
2) From repo root: `fly launch --no-deploy`
3) Deploy: `fly deploy`

Notes:
- First request after idle will be slower (cold start + optional YOLO preload).
- If cold starts are too slow, set `PRELOAD_MODEL=0` in `fly.toml`.
