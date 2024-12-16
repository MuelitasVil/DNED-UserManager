import os

class CorreosDocentesAdministrativos:
    def getCreateGroup(self, group):
        # Eliminar el sufijo @unal.edu.co.csv si existe
        group_name = group.replace("@unal.edu.co.csv", "")
        return f"gam create group {group_name}"
    
    def fillGroupWithCsv(self, csv):
        return f'gam csv "{csv}" gam update group "~Group Email" add "~Member Role" "~Member Email"'


class GeneradorScriptGam:
    def __init__(self):
        self.estudiantes = "ESTUDIANTES"
        self.estudiantes_dir = "app/archivos/estudiantes"
        self.scripts_dir = "app/scripts"
        self.correos = CorreosDocentesAdministrativos()

    def generar_scripts(self, tipoCsv: str):
        if tipoCsv == self.estudiantes:
            path = self.estudiantes_dir

        # Crear carpeta principal de scripts si no existe
        if not os.path.exists(self.scripts_dir):
            os.makedirs(self.scripts_dir)

        # Recorre las sedes dentro de "estudiantes"
        for sede in sorted(os.listdir(path)):
            sede_path = os.path.join(path, sede)

            if os.path.isdir(sede_path):
                print(f"Generando scripts para la sede: {sede}")

                # Crear carpeta específica para la sede dentro de "scripts"
                sede_scripts_dir = os.path.join(self.scripts_dir, sede)
                os.makedirs(sede_scripts_dir, exist_ok=True)

                script_create_name = os.path.join(sede_scripts_dir, f"{sede}_CreateGroups.bat")
                script_fill_name = os.path.join(sede_scripts_dir, f"{sede}_FillGroups.bat")

                with open(script_create_name, "w") as script_create, open(script_fill_name, "w") as script_fill:
                    script_create.write(f"@echo off\n")
                    script_create.write(f"REM Script para la creación de grupos en la sede {sede}\n\n")

                    script_fill.write(f"@echo off\n")
                    script_fill.write(f"REM Script para asignar usuarios a los grupos en la sede {sede}\n\n")

                    # Generar comandos para FACULTAD y PLAN
                    self._procesar_csv_en_ruta(os.path.join(sede_path, "FACULTAD-csv"), script_create, script_fill, "FACULTAD")
                    self._procesar_csv_en_ruta(os.path.join(sede_path, "PLAN-csv"), script_create, script_fill, "PLAN")
                
                print(f"Scripts {script_create_name} y {script_fill_name} generados con éxito.\n")

    def _procesar_csv_en_ruta(self, ruta, script_create, script_fill, tipo):
        if os.path.exists(ruta):
            # Recorrer las carpetas de la ruta en la que estamos parados.
            for root, dirnames, files in os.walk(ruta):
                print(f"Procesando ruta: {root}")
                
                # Saltar carpetas si hay subcarpetas (condición temporal)
                if len(dirnames) > 0:
                    continue
                
                for file in sorted(files):
                    if file.endswith(".csv"):
                        file_path = os.path.join(root, file)

                        # Añadir comandos al archivo de creación de grupos
                        script_create.write(f"REM Grupo: {tipo}\n")
                        script_create.write(self.correos.getCreateGroup(file) + "\n\n")

                        # Añadir comandos al archivo de llenado de grupos
                        script_fill.write(f"REM Llenar grupo: {tipo}\n")
                        script_fill.write(self.correos.fillGroupWithCsv(file) + "\n\n")


# Ejecución del generador
generador = GeneradorScriptGam()
generador.generar_scripts("ESTUDIANTES")
