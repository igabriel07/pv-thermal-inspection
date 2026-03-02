FROM python:3.11-slim

# Runtime deps
RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        libgl1 \
        libglib2.0-0 \
        libgomp1 \
        libreoffice-writer \
        fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY backend/requirements.txt backend/requirements.txt
# Install torch explicitly.
# - On amd64, force CPU-only wheels.
# - On arm64, fall back to PyPI wheels.
RUN arch="$(dpkg --print-architecture)" \
        && if [ "$arch" = "amd64" ]; then \
                pip install --no-cache-dir --index-url https://download.pytorch.org/whl/cpu torch; \
            else \
                pip install --no-cache-dir torch; \
            fi \
        && pip install --no-cache-dir -r backend/requirements.txt

COPY backend/ backend/

EXPOSE 8000

# Production-style command (no reload)
CMD ["uvicorn", "backend.main:app", "--host", "0.0.0.0", "--port", "8000"]
