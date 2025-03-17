# Use an official Python slim image
FROM python:3.11-slim

# Update apt and install system dependencies including LibreOffice
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    && rm -rf /var/lib/apt/lists/*

# Set a working directory
WORKDIR /app

# Copy the requirements file and install dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy your entire project into the container
COPY . .

# Expose the port (we'll use 8000)
EXPOSE 8000

# Start the app with Gunicorn
CMD ["gunicorn", "app3:app", "--bind", "0.0.0.0:8000"]
