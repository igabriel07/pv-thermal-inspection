"""Stable backend entrypoint.

Keep `uvicorn backend.main:app` working while the application code lives under
`backend/app/`.
"""

from backend.app.main import app  # noqa: F401
