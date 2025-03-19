# Use an official Python slim image
FROM python:3.11-slim

# Install system dependencies including LibreOffice
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    # Clean up to reduce image size
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

RUN libreoffice --version


# Set a working directory
WORKDIR /app

# Copy the requirements file and install dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy your entire project into the container
COPY . .

# Expose the port (Render uses $PORT, so we'll use that instead of 8000)
ENV PORT=8000
EXPOSE $PORT

# Start the app with Gunicorn
CMD ["gunicorn", "app3:app", "--bind", "0.0.0.0:$PORT"]
