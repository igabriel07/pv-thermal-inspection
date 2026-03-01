#!/usr/bin/env sh
set -eu

export PORT="${PORT:-8080}"

# Start backend (internal only)
uvicorn backend.main:app --host 127.0.0.1 --port 8000 &

# Serve frontend + proxy /api via Caddy (public)
exec /usr/local/bin/caddy run --config /etc/caddy/Caddyfile --adapter caddyfile
