FROM python:3.12-slim

LABEL maintainer="pdf2word"
LABEL description="PDF to DOCX converter with LibreOffice headless engine"

# Install LibreOffice Writer (headless) + Microsoft fonts for best fidelity
RUN apt-get update && \
    # Accept EULA for Microsoft fonts
    echo "ttf-mscorefonts-installer msttcorefonts/accepted-mscorefonts-eula select true" | debconf-set-selections && \
    apt-get install -y --no-install-recommends \
        libreoffice-writer \
        libreoffice-core \
        fonts-liberation \
        fonts-dejavu \
        fonts-noto \
        fontconfig \
        ttf-mscorefonts-installer \
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
