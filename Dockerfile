FROM python:3.11-slim-bookworm

# Evita que o Python gere arquivos .pyc e permite logs em tempo real
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

WORKDIR /app

# Instala curl para o Healthcheck e dependências básicas do sistema
RUN apt-get update && apt-get install -y \
    build-essential \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Copia e instala as dependências do Python
# O arquivo requirements.txt deve estar na raiz do projeto
COPY requirements.txt .
RUN pip3 install --no-cache-dir -r requirements.txt

# Copia todo o código do repositório (incluindo app.py e Excel.xlsm)
COPY . .

# Porta padrão utilizada pelo Streamlit
EXPOSE 8501

# Healthcheck configurado para validar a saúde da aplicação através da rota interna do Streamlit
HEALTHCHECK CMD curl --fail http://localhost:8501/_stcore/health

# Inicia a aplicação com as flags necessárias para o modo não interativo
# --server.headless=true: suprime o prompt de e-mail e impede a tentativa de abrir o navegador
# --browser.gatherUsageStats=false: desativa a coleta de estatísticas de telemetria
ENTRYPOINT ["streamlit", "run", "app.py", \
    "--server.port=8501", \
    "--server.address=0.0.0.0", \
    "--server.headless=true", \
    "--browser.gatherUsageStats=false"]