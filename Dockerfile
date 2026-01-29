FROM python:3.11-slim-bookworm

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

WORKDIR /app

# Instala apenas o essencial
RUN apt-get update && apt-get install -y \
    build-essential \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Instala dependências
COPY requirements.txt .
RUN pip3 install --no-cache-dir -r requirements.txt

# Copia todo o código e o Excel.xlsm
COPY . .

# Expõe a porta
EXPOSE 8501

# Comando de arranque com flags de Headless (para não pedir email) 
# e de Segurança (CORS/XSRF) para funcionar no Easypanel
ENTRYPOINT ["streamlit", "run", "app.py", \
    "--server.port=8501", \
    "--server.address=0.0.0.0", \
    "--server.headless=true", \
    "--browser.gatherUsageStats=false", \
    "--server.enableCORS=false", \
    "--server.enableXsrfProtection=false"]