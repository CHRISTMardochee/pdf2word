FROM python:3.12-slim

LABEL maintainer="pdf2word"
LABEL description="PDF to DOCX converter with LibreOffice headless engine"

# Install LibreOffice Writer (headless) + fonts for fidelity
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
        libreoffice-writer \
        libreoffice-core \
        fonts-liberation \
        fonts-dejavu \
        fonts-noto \
        fonts-freefont-ttf \
        fontconfig \
    && fc-cache -f -v \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install Python dependencies
COPY pyproject.toml requirements.txt ./
RUN pip install --no-cache-dir -e .

# Copy application
COPY pdf2word/ pdf2word/

# Default: run as a module
ENTRYPOINT ["python", "-m", "pdf2word"]
CMD ["--help"]
