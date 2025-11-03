FROM python:3.10-slim

# Set working directory
WORKDIR /app

# Install system dependencies
RUN apt-get update && apt-get install -y \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

# Copy dependency files
COPY requirements.txt pyproject.toml setup.py ./

# Upgrade pip and install dependencies
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

# Copy source code
COPY src/ ./src/
COPY README.md ./

# Set environment variables
ENV PYTHONUNBUFFERED=1

# Expose port (Railway will set PORT env var)
EXPOSE 8000

# Start command - Railway will override with railway.toml if needed
CMD python -m uvicorn mcp_sharepoint.server_http:app --host 0.0.0.0 --port ${PORT:-8000}
