FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    LANG=C.UTF-8 \
    LC_ALL=C.UTF-8

# Install LibreOffice for office document interaction
RUN apt-get update && \
    DEBIAN_FRONTEND=noninteractive \
    apt-get install -y --no-install-recommends \
        libreoffice-core libreoffice-writer libreoffice-calc libreoffice-impress \
        fonts-liberation fonts-noto-core fonts-noto-cjk && \
    fc-cache -f && \
    apt-get clean && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY pyproject.toml /app/pyproject.toml
COPY office2llm /app/office2llm
RUN pip install --no-cache-dir /app

ENTRYPOINT ["office2llm"]

