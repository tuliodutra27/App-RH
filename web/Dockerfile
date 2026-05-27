# ── Indicadores RH — AliseoSA ────────────────────────────────────────────
FROM python:3.12-slim

LABEL maintainer="TI AliseoSA"
LABEL description="Indicadores RH — comparação de planilhas com interface web"

# Evita prompts e garante saída de log imediata
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

WORKDIR /app

# Instala dependências do sistema necessárias para ldap3 e openpyxl
RUN apt-get update && apt-get install -y --no-install-recommends \
    libldap2-dev \
    libsasl2-dev \
    && rm -rf /var/lib/apt/lists/*

# Instala dependências Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copia o código da aplicação
COPY app/ .

# Cria diretórios de dados (sobrescritos pelo volume)
RUN mkdir -p /data/snapshots /data/relatorios /data/uploads

EXPOSE 5000

# gunicorn: 2 workers, timeout de 120s para uploads grandes
CMD ["gunicorn", \
     "--bind", "0.0.0.0:5000", \
     "--workers", "2", \
     "--timeout", "120", \
     "--access-logfile", "-", \
     "--error-logfile", "-", \
     "main:app"]
