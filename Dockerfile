FROM python:3.11-slim

# Instalar LibreOffice + display virtual + dependencias del sistema
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    xvfb \
    fonts-liberation \
    fonts-dejavu \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

# Directorio de trabajo
WORKDIR /app

# Instalar dependencias Python primero (aprovecha cache de Docker)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copiar el proyecto
COPY . .

# Directorio para archivos generados
RUN mkdir -p /tmp/segurito_outputs

# Puerto que expone Railway
EXPOSE 8000

# Script de inicio
CMD ["bash", "start.sh"]
