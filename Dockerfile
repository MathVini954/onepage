FROM python:3.11-slim-bookworm

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

WORKDIR /app

# Instala curl para o Healthcheck e dependências básicas
RUN apt-get update && apt-get install -y \
    build-essential \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Instala dependências do Python
COPY requirements.txt .
RUN pip3 install --no-cache-dir -r requirements.txt

# Copia todo o código do GitHub (incluindo o Excel.xlsm)
COPY . .

# Porta do Streamlit
EXPOSE 8501

# Healthcheck para o Easypanel monitorar a aplicação
HEALTHCHECK CMD curl --fail http://localhost:8501/_stcore/health

# Inicia a aplicação
ENTRYPOINT ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]