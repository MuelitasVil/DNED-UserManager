import os

class CorreosDocentesAdministrativos:
    def getCreateGroup(self, group):
        # Eliminar el sufijo @unal.edu.co si existe
        group_name = group.replace("@unal.edu.co.csv", "")
        return f"gam create group {group_name}"
    
    def fillGroupWithCsv(self, csv):
        return f'gam csv "{csv}" gam update group "~Group Email" add "~Member Role" "~Member Email"'

class GeneradorScriptGam:
    def __init__(self):
        self.estudiantes = "ESTUDIANTES"
        self.estudiantes_dir = "app/archivos/estudiantes"
        self.correos = CorreosDocentesAdministrativos()

    def generar_scripts(self, tipoCsv: str):
        if tipoCsv == self.estudiantes:
            path = self.estudiantes_dir

        # Recorre las sedes dentro de "estudiantes"
        for sede in sorted(os.listdir(path)):
            sede_path = os.path.join(path, sede)

            if os.path.isdir(sede_path):
                print(f"Generando scripts para la sede: {sede}")
                script_create_name = f"scriptCreateGroups_{sede}.bat"
                script_fill_name = f"scriptFillGroups_{sede}.bat"

                with open(script_create_name, "w") as script_create, open(script_fill_name, "w") as script_fill:
                    script_create.write(f"@echo off\n")
                    script_create.write(f"REM Script para la creación de grupos en la sede {sede}\n\n")

                    script_fill.write(f"@echo off\n")
                    script_fill.write(f"REM Script para asignar usuarios a los grupos en la sede {sede}\n\n")

                    self._procesar_csv_en_ruta(os.path.join(sede_path, "SEDE-csv"), script_create, script_fill, "SEDE")
                    self._procesar_csv_en_ruta(os.path.join(sede_path, "FACULTAD-csv"), script_create, script_fill, "FACULTAD", True)
                    self._procesar_csv_en_ruta(os.path.join(sede_path, "PLAN-csv"), script_create, script_fill, "PLAN")
                
                print(f"Scripts {script_create_name} y {script_fill_name} generados con éxito.\n")

    def _procesar_csv_en_ruta(self, ruta, script_create, script_fill, tipo, explorar_subcarpetas=False):
        if os.path.exists(ruta):
            for root, _, files in os.walk(ruta):
                for file in sorted(files):
                    if file.endswith(".csv"):
                        file_path = os.path.join(root, file)
                        group_name = f"{tipo}_{os.path.splitext(file)[0]}"

                        # Añadir comandos al archivo de creación de grupos
                        script_create.write(f"REM Grupo: {group_name}\n")
                        script_create.write(self.correos.getCreateGroup(file) + "\n\n")

                        # Añadir comandos al archivo de llenado de grupos
                        script_fill.write(f"REM Llenar grupo: {group_name}\n")
                        script_fill.write(self.correos.fillGroupWithCsv(file) + "\n\n")

# Ejecución del generador
generador = GeneradorScriptGam()
generador.generar_scripts("ESTUDIANTES")
