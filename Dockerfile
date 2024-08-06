# Imagen base de python
FROM python:3.12.0

# directorio de trabajo en el contenedor
WORKDIR /app

# Copiar archivos necesarios al contenedor
COPY app /app

# Instala las dependencias del proyecto
RUN pip install --upgrade pip && pip install -r requirements.txt

# Expone el puerto en el que Flask estará escuchando
EXPOSE 5000

# Comando para ejecutar la aplicación
# CMD ["gunicorn", "--bind", "0.0.0.0:5000", "app:app"]

CMD ["flask", "run", "--host", "0.0.0.0"]
