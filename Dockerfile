FROM python:3.11-slim

WORKDIR /app

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1

# Deps systeme pour reportlab/openpyxl/pillow (rendu PDF + images)
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    libjpeg-dev \
    zlib1g-dev \
    && rm -rf /var/lib/apt/lists/*

# Install Python deps en premier (cache Docker)
COPY requirements.txt ./
RUN pip install --upgrade pip && pip install -r requirements.txt

# Copie du code
COPY . .

EXPOSE 8000

# Lance le backend Flask (sert aussi le frontend HTML)
CMD ["python", "-m", "backend.main"]
