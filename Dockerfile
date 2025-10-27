# Use a slim Python image
FROM python:3.10-slim

# Install system dependencies
RUN apt-get update && apt-get install -y \
  nginx \
  && rm -rf /var/lib/apt/lists/*

# Set the working directory
WORKDIR /app

# Install the required Python packages
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy the app code into the container
COPY app ./app

# Expose ports for Uvicorn (8000) and Nginx (80)
EXPOSE 8000 80

# Copy Nginx configuration
COPY nginx.conf /etc/nginx/nginx.conf

# Start Nginx and Uvicorn
CMD service nginx start && uvicorn app.main:app --host 0.0.0.0 --port 8000 --workers 4
