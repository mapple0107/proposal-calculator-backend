FROM ubuntu:24.04

ENV DEBIAN_FRONTEND=noninteractive

RUN apt-get update && apt-get install -y \
    libreoffice-calc \
    libreoffice-core \
    python3 \
    python3-pip \
    python3-uno \
    fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir --break-system-packages -r requirements.txt

COPY app.py uno_engine.py start.sh ./
COPY templates ./templates
RUN chmod +x start.sh

ENV UNO_HOST=localhost
ENV UNO_PORT=2002

CMD ["./start.sh"]
