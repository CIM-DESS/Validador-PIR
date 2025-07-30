# Usa una imagen base de Python
FROM python:3.12-slim

# Establece el directorio de trabajo
WORKDIR /app

# Instala dependencias del sistema necesarias para compilar paquetes Python
RUN apt-get update && apt-get install -y \
    gcc \
    g++ \
    && rm -rf /var/lib/apt/lists/*

# Actualiza pip, setuptools y wheel
RUN pip install --upgrade pip setuptools wheel

# Copia los archivos de dependencias primero (para aprovechar cache de Docker)
COPY requirements.txt .

# Instala las dependencias de Python
RUN pip install --no-cache-dir -r requirements.txt

# Copia el código de la aplicación
COPY . .

# Expone el puerto 8000
EXPOSE 8000

# Comando por defecto para ejecutar la aplicación
CMD ["python", "manage.py", "runserver", "0.0.0.0:8000"]
