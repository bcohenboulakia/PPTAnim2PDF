FROM python:3.12-slim

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        libreoffice-core \
        libreoffice-impress \
        fonts-dejavu \
        fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY pptanim2pdf.py /app/pptanim2pdf.py

RUN pip install --no-cache-dir lxml \
    && printf '#!/bin/sh\nexec python /app/pptanim2pdf.py "$@"\n' > /usr/local/bin/pptanim2pdf \
    && chmod +x /usr/local/bin/pptanim2pdf

ENTRYPOINT ["pptanim2pdf"]
