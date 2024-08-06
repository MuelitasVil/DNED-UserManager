# DNED-UserManager
Proyecto destinado al manejo de formatos de excel. 

# Ejecutar el proyecto :  

## Configurar ambiente local 

1. Clonar el repositorio en equipo personal 

2. ingrese a la carpeta de la proyecto (app)

3. Cree el ambiente vitual usando el comando "py -m venv .venv"

4. Para activar el ambiente virtual ejecute el siguiente comando :
    - windouws : ".venv\Scripts\activate" 
    - macOS y Linux :  "source .venv/bin/activate"  

5. En su consola deberia aparecer (.venv) al principio si hizo este proceso correctamente. 

6. Descargue los paquetes pip asociados al proyecto con el comando "pip install -r .\requirements.txt"

7. Ejecute "flask run" para ejecutar el proyecto. Si desea ejecutar el proyecto en modo Debug ingrese el comando "flask run --debug".

8. Si desea salir del entorno virtual ejecute "deactivate"  

9. Llegado el caso que necesite agregar otro paquete al proyecto descaguelo dentro del ambiente virtual y ejecute le comando "pip freeze > requirements.txt" tenga cuidado con este comando pues sobreescribirta el archivo requeriments.txt 

## Configurar ambiente de produccion : 

Si desea probar su codigo en un ambiente de produccion, ejecute el siguiente comando para verificar si funcionan sus cambios en el docker.
- docker-compose up --build
