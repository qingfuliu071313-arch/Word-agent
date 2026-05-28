# word-agent Docker image
# Multi-stage build: Python dependencies → MCP servers → runtime
#
# NOTE: Docker mode is file-level only. Live editing (COM/JXA) is not available
# because there is no native Word application inside the container.

FROM python:3.11-slim AS base

WORKDIR /app

# System dependencies for LibreOffice (doc→docx conversion)
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer \
    fonts-wqy-microhei \
    fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/*

# Python dependencies
COPY scripts/requirements.txt scripts/requirements.txt
RUN pip install --no-cache-dir -r scripts/requirements.txt

# Copy word-agent plugin files
COPY . /app/

# MCP server ports (configurable via environment)
ENV WORD_DOC_SERVER_PORT=3100
ENV WORD_MCP_LIVE_PORT=3101
ENV DOCX_MCP_PORT=3102
ENV MCP_AUTHOR=Claude

# Expose MCP server ports
EXPOSE 3100 3101 3102

# Default working directory for document mounting
VOLUME ["/documents"]

# Health check
HEALTHCHECK --interval=30s --timeout=5s --retries=3 \
    CMD python3 -c "import docx2python; print('ok')" || exit 1

CMD ["bash"]
