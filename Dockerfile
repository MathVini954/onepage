FROM python:3.11-slim-bookworm

# Otimização de logs e bytecode
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

WORKDIR /app

# Instalação de dependências do sistema
RUN apt-get update && apt-get install -y \
    build-essential \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Instalação das bibliotecas Python
COPY requirements.txt .
RUN pip3 install --no-cache-dir -r requirements.txt

# Copia os ficheiros (incluindo o Excel.xlsm)
COPY . .

# Altere para a porta 8502
EXPOSE 8502

# Comando de arranque atualizado para a porta 8502
ENTRYPOINT ["streamlit", "run", "app.py", \
    "--server.port=8502", \
    "--server.address=0.0.0.0", \
    "--server.headless=true", \
    "--browser.gatherUsageStats=false", \
    "--server.enableCORS=false", \
    "--server.enableXsrfProtection=false"]