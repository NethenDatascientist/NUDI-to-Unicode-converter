FROM python:3.11-slim

# Install LibreOffice Writer for .doc to .docx conversion
RUN apt-get update && \
    apt-get install -y --no-install-recommends libreoffice-writer default-jre-headless && \
    apt-get clean && \
    rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy all project files (app.py, kannada_converter.py, etc.)
COPY . .

# Start Streamlit on Render's dynamic port
CMD streamlit run app.py --server.port=${PORT:-10000} --server.address=0.0.0.0 --server.headless=true --server.enableWebsocketCompression=false --server.enableCORS=false --server.enableXsrfProtection=false
