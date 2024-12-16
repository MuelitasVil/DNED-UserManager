# DNED-UserManager

**DNED-UserManager** es un proyecto diseñado para gestionar y procesar archivos Excel de manera eficiente. Utiliza Flask como framework principal para el backend y facilita el manejo de datos administrativos y de usuarios a través de una interfaz sencilla.

## Características del proyecto

- **Gestión automatizada de archivos Excel:** Procesa y organiza datos desde archivos `.xlsx` y `.csv`.
- **Creación y llenado de grupos:** Automatización mediante scripts `.bat` para generar y llenar grupos.
- **Interfaz de usuario sencilla:** Desarrollado en Flask, lo que permite una experiencia intuitiva para administrar los datos.
- **API robusta:** Posibilidad de extender funcionalidades para consumir datos a través de endpoints.
- **Dockerización:** Configuración para ejecutar el proyecto en un entorno de producción con Docker y Docker Compose.
- **Entorno virtualizado:** Uso de entornos virtuales para gestionar dependencias de manera eficiente.

## Tecnologías utilizadas

- **Backend:** Flask 3.0.0
- **Procesamiento de Excel:** OpenPyXL, Pandas
- **Servidor en producción:** Gunicorn
- **Autenticación JWT:** PyJWT
- **HTTP Requests:** HTTPX, Requests
- **Virtualización:** Docker, Docker Compose
- **Entorno virtual:** Python `venv`

## Ejecutar el proyecto

### Configurar ambiente local

1. **Clonar el repositorio en tu equipo personal:**
   ```bash
   git clone <URL_DEL_REPOSITORIO>
   ```

2. **Acceder a la carpeta del proyecto:**
   ```bash
   cd app
   ```

3. **Crear el entorno virtual:**
   ```bash
   py -m venv .venv
   ```

4. **Activar el entorno virtual:**
   - **Windows:**
     ```bash
     .venv\Scripts\activate
     ```
   - **macOS y Linux:**
     ```bash
     source .venv/bin/activate
     ```

5. **Verificar que el entorno virtual está activo:**  
   La consola debería mostrar `(.venv)` al principio.

6. **Instalar los paquetes necesarios:**
   ```bash
   pip install -r requirements.txt
   ```

7. **Ejecutar el proyecto con Flask:**
   ```bash
   flask run
   ```
   - **Modo Debug:**
     ```bash
     flask run --debug
     ```

8. **Desactivar el entorno virtual cuando termines:**
   ```bash
   deactivate
   ```

9. **Agregar nuevos paquetes al proyecto:**  
   Si necesitas agregar más paquetes, instálalos en el entorno virtual y actualiza `requirements.txt`:
   ```bash
   pip freeze > requirements.txt
   ```
   🚨 **Nota:** Este comando sobrescribirá el archivo `requirements.txt`, asegúrate de revisar los cambios antes de ejecutarlo.

### Configurar ambiente de producción

Si deseas probar el código en un ambiente de producción, utiliza Docker y Docker Compose para construir y ejecutar el contenedor.

1. **Construir y ejecutar el contenedor:**
   ```bash
   docker-compose up --build
   ```

2. **Detener el contenedor:**
   ```bash
   docker-compose down
   ```

## Estructura del proyecto

```
DNED-UserManager/
│
├── .venv/                  # Entorno virtual
├── archivos/               # Archivos generados xls
├── scripts/                # Archivos generados de  tipo .bat
├── routes/                 # Rutas de la aplicación Flask
├── services/               # Lógica de negocio y servicios
├── static/                 # Archivos estáticos (CSS, JS, imágenes)
├── templates/              # Plantillas HTML para la interfaz web
├── app.py                  # Archivo principal de la aplicación Flask
├── requirements.txt        # Dependencias del proyecto
└── __init__.py             # Archivo de inicialización del paquete
```

## Dependencias principales

Aquí se listan las dependencias clave del proyecto:

- **Flask 3.0.0:** Framework web ligero para Python.
- **Gunicorn 21.2.0:** Servidor WSGI para producción.
- **Pandas 2.1.4:** Análisis y manipulación de datos.
- **OpenPyXL 3.1.2:** Lectura y escritura de archivos Excel.
- **HTTPX 0.26.0 / Requests 2.31.0:** Para realizar peticiones HTTP.
- **PyJWT 2.8.0:** Manejo de tokens JWT para autenticación.

Para ver todas las dependencias, consulta el archivo `requirements.txt`.

## Contribuciones

Las contribuciones son bienvenidas. Por favor, abre un issue para reportar problemas o sugerir mejoras. Puedes enviar pull requests para agregar nuevas funcionalidades.
