FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    QT_QPA_PLATFORM=offscreen

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        libdbus-1-3 libfontconfig1 libglib2.0-0 libxkbcommon-x11-0 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY gmaildownloader.py /app/gmaildownloader.py
COPY icon.png /app/icon.png
RUN python -m pip install --no-cache-dir PyQt6 anthropic cryptography reportlab pypdfium2 Pillow

ENTRYPOINT ["python", "/app/gmaildownloader.py", "--headless"]
