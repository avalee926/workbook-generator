# Use an official Python slim image
FROM python:3.11-slim

# Install system dependencies including LibreOffice and Java
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    default-jre && \  # Install Java Runtime Environment for LibreOffice
    apt-get clean && \
    rm -rf /var/lib/apt/lists/*

# Verify LibreOffice installation
RUN libreoffice --version

# Set a working directory
WORKDIR /app

# Copy the requirements file and install dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy your entire project into the container
COPY . .

# Expose the port (for documentation purposes)
EXPOSE 8000

# Start the app with Gunicorn
# Use a shell to expand the $PORT environment variable
CMD ["sh", "-c", "gunicorn app3:app --bind 0.0.0.0:$PORT --timeout 360"]
