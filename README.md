# DNED-UserManager

**DNED-UserManager** es un proyecto diseñado para gestionar y procesar archivos Excel de manera eficiente. Utiliza Flask como framework principal para el backend y facilita el manejo de datos administrativos y de usuarios a través de una interfaz sencilla.

## Características del proyecto

- **Gestión automatizada de archivos Excel:** Procesa y organiza datos desde archivos `.xlsx` y `.csv`.
- **Creación y llenado de grupos:** Automatización mediante scripts `.bat` para generar y llenar grupos.

## Tecnologías utilizadas

- **Framework:** Flask 3.0.0
- **Procesamiento de Excel:** OpenPyXL, Pandas
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

Para ver todas las dependencias, consulta el archivo `requirements.txt`.

## Contribuciones

Las contribuciones son bienvenidas. Por favor, abre un issue para reportar problemas o sugerir mejoras. Puedes enviar pull requests para agregar nuevas funcionalidades.
