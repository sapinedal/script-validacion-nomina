# Usa una imagen oficial de Python ligera
FROM python:3.11-slim

# Establece el directorio de trabajo en el contenedor
WORKDIR /app

# Instala dependencias del sistema necesarias para compilar ciertas librerías si es necesario (ej. psycopg2)
RUN apt-get update && apt-get install -y \
    libpq-dev gcc \
    && rm -rf /var/lib/apt/lists/*

# Copia el archivo de requerimientos
COPY requirements.txt .

# Instala las dependencias de Python
RUN pip install --no-cache-dir -r requirements.txt

# Copia el resto del código del proyecto
COPY . .

# Expone el puerto donde corre la aplicación FastAPI
EXPOSE 8000

# Comando para ejecutar la aplicación con Uvicorn
CMD ["uvicorn", "server:app", "--host", "0.0.0.0", "--port", "8000"]
