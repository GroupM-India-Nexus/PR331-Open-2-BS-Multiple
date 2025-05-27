FROM python:3.9-slim

# Set working directory
WORKDIR /app

# Install system-level build dependencies
RUN apt-get update && apt-get install -y --no-install-recommends \
    gcc python3-dev \
    && rm -rf /var/lib/apt/lists/*

# Upgrade pip, setuptools, and wheel
RUN pip install --upgrade pip setuptools wheel

# Copy and install only the requirements first (to leverage caching)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Create data directory with proper permissions
RUN mkdir -p /app/data && chmod 777 /app/data

# Copy the remaining application code
COPY . .

# Set environment variables
ENV PYTHONUNBUFFERED=1
ENV PORT=8080

# Expose port
EXPOSE 8080

# Define the container entrypoint
CMD exec gunicorn --bind 0.0.0.0:$PORT --workers 1 --threads 8 --timeout 600 --access-logfile - --error-logfile - app:app