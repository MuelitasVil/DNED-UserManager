from app.services.excel.utils import getCantOfColumns
from app.services.excel.excel.archvivosExcel import ArchivosExcel
from io import StringIO
import time
from openpyxl.utils import get_column_letter
import openpyxl 
from openpyxl import Workbook, load_workbook

import csv

'''
La siguiente clase tendra la funcionalidad de recorrer el excel en el cual se
encuentran los datos de los inquilinos.

Ojo : Es necesario que entienda la estructura del excel para entender las
siguientes funciones.
'''

class ExcelManager:
    def __init__(self, file = None, files = []):
        

        # Lectura del excel 
        if file != None:
            self.excel = openpyxl.load_workbook(file)
            self.hojas = self.excel.get_sheet_names()

        self.UserFiles = files
        self.columnaInicial = 1
        self.filaInicial = 1    
    
    def MergueUsuarios(self):
        
        '''
        Diccionario = {
            CorreoUsuario : {
                Nombre : String
                Espacio Gastado : 
            }
        }
        '''
        inicio = time.time()
        print(" Inicio del mergue ")
        Usuarios = {}
        for file in self.UserFiles:
            print(file)
            if file.filename.endswith(".xlsx") or file.filename.endswith(".xlsm") :
                excelFile = openpyxl.load_workbook(file)
                self.getDataUsersExel(excelFile, file.filename, Usuarios)
            else:
                # Es csv LDAP
                self.getDataUsersCsv(file, Usuarios)

        print(" Estructura de datos generada ")
        print(" Inicio generar excel ")
        # Limberar espacio ram
        excelFile = None

        # Generar el excel con la informacion 
        ws = Workbook()
        hojaUsuario = ws.create_sheet("Informacion General")
        hojaEgresado = ws.create_sheet("Egresados")
        hojaLdap = ws.create_sheet("OnlyLdap")
        hojaWorkSpace = ws.create_sheet("OnlyWorkSpace")
        hojaPregrado = ws.create_sheet("Only Pregrado")
        hojaPostGrado = ws.create_sheet("Only Postgrado")
        hojaDocente = ws.create_sheet("Only Docente")
        hojaAministrativo = ws.create_sheet("Only Administrativo")
        hojaSinConexion = ws.create_sheet("No conexion")
        
        columnas = self.get_columnas_usuarios()
        cantidadColumnas = len(columnas)

        # Rellenar columnas de los archivos
        for c in range(cantidadColumnas):
            column = get_column_letter(c + 1)
            hojaUsuario[column + "1"].value = columnas[c]
            if c <= 4:
                hojaEgresado[column + "1"].value = columnas[c]
                hojaWorkSpace[column + "1"].value = columnas[c]
                hojaLdap[column + "1"].value = columnas[c]
                hojaPregrado[column + "1"].value = columnas[c]
                hojaPostGrado[column + "1"].value = columnas[c]
                hojaDocente[column + "1"].value = columnas[c]
                hojaAministrativo[column + "1"].value = columnas[c]
                hojaSinConexion[column + "1"].value = columnas[c]

        # Ingresar Columnas 
        rowUsuario = 2
        rowEgresado = 2
        rowLdap = 2
        rowWorkSpace = 2
        rowPregrado = 2
        rowPostgrado = 2
        rowDocente = 2 
        rowAdministrativo = 2
        rowNoConexion = 2
        for Correo in Usuarios:

            if rowUsuario % 50000 == 0:
                print("Generando ... ")

            hojaUsuario["A" + str(rowUsuario)] = Correo
            hojaUsuario["B" + str(rowUsuario)] = Usuarios[Correo]["Nombre"]
            hojaUsuario["C" + str(rowUsuario)] = Usuarios[Correo]["Apellido"] 
            hojaUsuario["D" + str(rowUsuario)] = Usuarios[Correo]["UltimaConexion"] 
            hojaUsuario["E" + str(rowUsuario)] = Usuarios[Correo]["Almacenamiento"]
            hojaUsuario["F" + str(rowUsuario)] = Usuarios[Correo][ArchivosExcel.Docentes.TipoArchivo]
            hojaUsuario["G" + str(rowUsuario)] = Usuarios[Correo][ArchivosExcel.Egresados.TipoArchivo] 
            hojaUsuario["H" + str(rowUsuario)] = Usuarios[Correo][ArchivosExcel.EstudiantesActivos.TipoArchivo]
            hojaUsuario["I" + str(rowUsuario)] = Usuarios[Correo][ArchivosExcel.WorkSpace.TipoArchivo] 
            hojaUsuario["J" + str(rowUsuario)] = Usuarios[Correo][ArchivosExcel.Ldap.TipoArchivo] 
            hojaUsuario["K" + str(rowUsuario)] = Usuarios[Correo]["IsEgresado"] 
            
            if Usuarios[Correo]["OnlyWorkSpace"]:
                # Esto se puede volver una funccion pero me dio pereza :'v
                hojaWorkSpace["A" + str(rowWorkSpace)] = Correo
                hojaWorkSpace["B" + str(rowWorkSpace)] = Usuarios[Correo]["Nombre"]
                hojaWorkSpace["C" + str(rowWorkSpace)] = Usuarios[Correo]["Apellido"] 
                hojaWorkSpace["D" + str(rowWorkSpace)] = Usuarios[Correo]["UltimaConexion"] 
                hojaWorkSpace["E" + str(rowWorkSpace)] = Usuarios[Correo]["Almacenamiento"]
                rowWorkSpace += 1
            elif Usuarios[Correo]["OnlyLdap"]:
                hojaLdap["A" + str(rowLdap)] = Correo
                hojaLdap["B" + str(rowLdap)] = Usuarios[Correo]["Nombre"]
                hojaLdap["C" + str(rowLdap)] = Usuarios[Correo]["Apellido"] 
                hojaLdap["D" + str(rowLdap)] = Usuarios[Correo]["UltimaConexion"] 
                hojaLdap["E" + str(rowLdap)] = Usuarios[Correo]["Almacenamiento"]
                rowLdap += 1
            elif Usuarios[Correo]["IsEgresado"]: # Con el elseIf se verifica que exista en usuarios y no solo en los espacios
                hojaEgresado["A" + str(rowEgresado)] = Correo
                hojaEgresado["B" + str(rowEgresado)] = Usuarios[Correo]["Nombre"]
                hojaEgresado["C" + str(rowEgresado)] = Usuarios[Correo]["Apellido"] 
                hojaEgresado["D" + str(rowEgresado)] = Usuarios[Correo]["UltimaConexion"] 
                hojaEgresado["E" + str(rowEgresado)] = Usuarios[Correo]["Almacenamiento"]
                rowEgresado += 1
            elif Usuarios[Correo]["IsPregrado"]:
                hojaPregrado["A" + str(rowPregrado)] = Correo
                hojaPregrado["B" + str(rowPregrado)] = Usuarios[Correo]["Nombre"]
                hojaPregrado["C" + str(rowPregrado)] = Usuarios[Correo]["Apellido"] 
                hojaPregrado["D" + str(rowPregrado)] = Usuarios[Correo]["UltimaConexion"] 
                hojaPregrado["E" + str(rowPregrado)] = Usuarios[Correo]["Almacenamiento"]
                rowPregrado += 1
            elif Usuarios[Correo]["IsPostgrado"]:
                hojaPostGrado["A" + str(rowPostgrado)] = Correo
                hojaPostGrado["B" + str(rowPostgrado)] = Usuarios[Correo]["Nombre"]
                hojaPostGrado["C" + str(rowPostgrado)] = Usuarios[Correo]["Apellido"] 
                hojaPostGrado["D" + str(rowPostgrado)] = Usuarios[Correo]["UltimaConexion"] 
                hojaPostGrado["E" + str(rowPostgrado)] = Usuarios[Correo]["Almacenamiento"]
                rowPostgrado += 1
            elif Usuarios[Correo]["IsDocente"]: 
                hojaDocente["A" + str(rowDocente)] = Correo
                hojaDocente["B" + str(rowDocente)] = Usuarios[Correo]["Nombre"]
                hojaDocente["C" + str(rowDocente)] = Usuarios[Correo]["Apellido"] 
                hojaDocente["D" + str(rowDocente)] = Usuarios[Correo]["UltimaConexion"] 
                hojaDocente["E" + str(rowDocente)] = Usuarios[Correo]["Almacenamiento"]
                rowDocente += 1
            elif Usuarios[Correo]["IsAdministrativo"]:
                hojaAministrativo["A" + str(rowAdministrativo)] = Correo
                hojaAministrativo["B" + str(rowAdministrativo)] = Usuarios[Correo]["Nombre"]
                hojaAministrativo["C" + str(rowAdministrativo)] = Usuarios[Correo]["Apellido"] 
                hojaAministrativo["D" + str(rowAdministrativo)] = Usuarios[Correo]["UltimaConexion"] 
                hojaAministrativo["E" + str(rowAdministrativo)] = Usuarios[Correo]["Almacenamiento"]
                rowAdministrativo += 1
            elif Usuarios[Correo]["NoConexion"]:
                hojaSinConexion["A" + str(rowNoConexion)] = Correo
                hojaSinConexion["B" + str(rowNoConexion)] = Usuarios[Correo]["Nombre"]
                hojaSinConexion["C" + str(rowNoConexion)] = Usuarios[Correo]["Apellido"] 
                hojaSinConexion["D" + str(rowNoConexion)] = Usuarios[Correo]["UltimaConexion"] 
                hojaSinConexion["E" + str(rowNoConexion)] = Usuarios[Correo]["Almacenamiento"]
                hojaSinConexion["F" + str(rowNoConexion)] = Usuarios[Correo][ArchivosExcel.Ldap.TipoArchivo]
                rowNoConexion += 1
                     
            rowUsuario += 1

        print(" Excel generado ")
        print(" Guardando Excel ")
        
        ws.save("sample.xlsx")
        fin = time.time()   
        
        print("Fin")
        print("Tiempo tomado : " + str(round(fin-inicio)))
        
        return False

    def getDataUsersCsv(self, csvFile, Usuarios):
        # Leer el contenido del archivo y decodificarlo a cadena de texto
        csv_content = csvFile.read().decode('utf-8')

        # Usar StringIO para crear un archivo en memoria a partir de la cadena de texto
        csv_file = StringIO(csv_content)

        # Crear el lector CSV
        csvReader = csv.reader(csv_file)

        ldapAchivo = ArchivosExcel.Ldap
        tipos = ArchivosExcel.Ldap.TipoUsuario
        for row in csvReader:
            Correo = row[ldapAchivo.Correo]
            gropu_tipo = row[ldapAchivo.Tipo]
            
            if not Correo.endswith("@unal.edu.co"):
                continue
            
            if Correo not in Usuarios:
                self.create_user_dict(Correo, Usuarios)

            TiposUsuario = ""
            array_tipos = gropu_tipo.split("|")
            for tipo in array_tipos: 

                if str(tipo) not in tipos:
                    continue

                tipoUsuario = tipos[str(tipo)]
        
                if tipoUsuario != "Egresado":
                    Usuarios[Correo]["IsEgresado"] = False
                
                Usuarios[Correo]["OnlyWorkSpace"] = False

                TiposUsuario = TiposUsuario + " | " + tipoUsuario

            Usuarios[Correo]["Ldap"] = TiposUsuario

    def getDataUsersExel(self, excelFile, filename, Usuarios):
        # Encontrar cual Archivo es 
        nombreHoja = ""
        if filename == ArchivosExcel.Docentes.NombreArchivo:
            Archivo = ArchivosExcel.Docentes
            nombreHoja = Archivo.NombreHoja
        if filename == ArchivosExcel.Egresados.NombreArchivo:
            Archivo = ArchivosExcel.Egresados
            nombreHoja = Archivo.NombreHoja
        if filename == ArchivosExcel.EstudiantesActivos.NombreArchivo:
            Archivo = ArchivosExcel.EstudiantesActivos
            nombreHoja = Archivo.NombreHoja
        if filename == ArchivosExcel.WorkSpace.NombreArchivo:
            Archivo = ArchivosExcel.WorkSpace
            nombreHoja = Archivo.NombreHoja

        information = excelFile.get_sheet_by_name(nombreHoja)
        cantOfRows = len(list(information.rows))
        cantOfColumns = getCantOfColumns(nombreHoja)
        maxColumns = self.columnaInicial + cantOfColumns
        filaInicial = self.filaInicial
        Datos = False

        for row in range(filaInicial ,cantOfRows):
            columnCorreo = get_column_letter(Archivo.Correo + 1)
            Correo = information[columnCorreo + str(row)].value
            
            if Correo not in Usuarios:
                self.create_user_dict(Correo, Usuarios)
                
            if ArchivosExcel.WorkSpace.TipoArchivo == Archivo.TipoArchivo:
                columnNombre = get_column_letter(Archivo.Nombre + 1)
                columnApellido = get_column_letter(Archivo.Apellidos + 1)
                columnUltimaConexion = get_column_letter(Archivo.UltimaConexion + 1)
                columnAlmacenamiento = get_column_letter(Archivo.Almacenamiento + 1)
                
                Nombre = information[columnNombre + str(row)].value
                Apellido = information[columnApellido + str(row)].value
                UltimaConexion = information[columnUltimaConexion + str(row)].value
                Almacenamiento = information[columnAlmacenamiento + str(row)].value

                Usuarios[Correo]["Nombre"] = Nombre
                Usuarios[Correo]["Apellido"] = Apellido
                Usuarios[Correo]["UltimaConexion"] = UltimaConexion
                Usuarios[Correo]["Almacenamiento"] = Almacenamiento

            if (Archivo.TipoArchivo != ArchivosExcel.WorkSpace.TipoArchivo): 
                Usuarios[Correo]["OnlyWorkSpace"] = False

            if (Archivo.TipoArchivo != ArchivosExcel.Ldap.TipoArchivo):
                Usuarios[Correo]["OnlyLdap"] = False

            if (Archivo.TipoArchivo != ArchivosExcel.Egresados.TipoArchivo
                and Archivo.TipoArchivo != ArchivosExcel.WorkSpace.TipoArchivo):
                 Usuarios[Correo]["IsEgresado"] = False

            if Archivo.TipoArchivo == ArchivosExcel.EstudiantesActivos.TipoArchivo:
                columnNivel = get_column_letter(Archivo.Nivel + 1)
                Nivel = str(information[columnNivel + str(row)].value).strip()
                if Nivel ==  "PREGRADO":
                    Usuarios[Correo]["IsPregrado"] = True
                elif Nivel == "POSGRADO":
                    Usuarios[Correo]["IsPostgrado"] = True

            if Archivo.TipoArchivo == ArchivosExcel.Docentes.TipoArchivo:
                columnVinculacion = get_column_letter(Archivo.Vinculacion + 1)
                Vinculacion = str(information[columnVinculacion + str(row)].value)
                if "DOCENTE" in Vinculacion:
                    Usuarios[Correo]["IsDocente"] = True
                else:
                    Usuarios[Correo]["IsAdministrativo"] = True

            if Archivo.TipoArchivo == ArchivosExcel.WorkSpace.TipoArchivo:
                columnConexion = get_column_letter(Archivo.UltimaConexion + 1)
                Conexion = str(information[columnConexion + str(row)].value)
                if Conexion.strip() == "Never logged in":
                    Usuarios[Correo]["NoConexion"] = True

            Usuarios[Correo][Archivo.TipoArchivo] = True
        
    def create_user_dict(self, Correo, Usuarios):
        Usuarios[Correo] = {}
        Usuarios[Correo][ArchivosExcel.Docentes.TipoArchivo] = False
        Usuarios[Correo][ArchivosExcel.Egresados.TipoArchivo] = False
        Usuarios[Correo][ArchivosExcel.EstudiantesActivos.TipoArchivo] = False
        Usuarios[Correo][ArchivosExcel.WorkSpace.TipoArchivo] = False
        Usuarios[Correo][ArchivosExcel.Ldap.TipoArchivo] = False
        Usuarios[Correo]["Ldap"] = "" 
        Usuarios[Correo]["Nombre"] = ""
        Usuarios[Correo]["Apellido"] = ""
        Usuarios[Correo]["UltimaConexion"] = ""
        Usuarios[Correo]["Almacenamiento"] = ""
        Usuarios[Correo]["IsEgresado"] = True
        Usuarios[Correo]["IsPregrado"] = False
        Usuarios[Correo]["IsPostgrado"] = False
        Usuarios[Correo]["IsDocente"] = False
        Usuarios[Correo]["IsAdministrativo"] = False
        Usuarios[Correo]["OnlyWorkSpace"] = True
        Usuarios[Correo]["OnlyLdap"] = True
        Usuarios[Correo]["NoConexion"] = False

    def get_columnas_usuarios(self):

        Columnas = [
            "Correo Usuario",
            "Nombre",
            "Apellido",
            "UltimaConexion",
            "Almacenamiento",
            ArchivosExcel.Docentes.TipoArchivo,
            ArchivosExcel.Egresados.TipoArchivo,
            ArchivosExcel.EstudiantesActivos.TipoArchivo,
            ArchivosExcel.WorkSpace.TipoArchivo,
            ArchivosExcel.Ldap.TipoArchivo,
            "IsEgresado"
            ]
        
        return Columnas
        
    def get_csv_estudiantes(self):
        hojas = ["ESTUDIANTES ACTIVOS 2024-1S"]
        for hoja in hojas:
            self.get_dataHojaEstudiantes(hoja)

    def get_dataHojaEstudiantes(self, nombreHoja):
        
        '''
        Se va a recorrer la hoja de exel extrayendo la cantidad de filas
        atravez de la libreria openpyxl, mietras que la cantidad de columnas la 
        extraemos dependiendo de hoja que se esta recorriendo. 
        '''

        # Datos : 

        columsCsv = ["Group Email [Required]","Member Email","Member Type","Member Role"]

        infoCsv = {
            "Group Email [Required]" : "acompmanconflicto@unal",
            "Member Email" : "",
            "Member Type" : "USER",
            "Member Role" : "MEMBER"
        }

        dictColumns = {
            'NOMBRES_APELLIDOS' : 1,
            "EMAIL" : 2,	
            "SEDE" : 3,
            "FACULTAD" : 4,	
            "COD_PLAN" : 5,	
            "PLAN" : 6,	
            "TIPO_NIVEL" : 7,
        }

        FiltroSede = ["SEDE BOGOTÁ"]
        FiltroFacultad = ["FACULTAD DE CIENCIAS HUMANAS"]
        FiltroPlan = []


        columns = list(dictColumns.keys())

        information = self.excel.get_sheet_by_name(nombreHoja)
        cantOfRows = len(list(information.rows))
        cantOfColumns = getCantOfColumns(nombreHoja)
        maxColumns = self.columnaInicial + cantOfColumns
        filaInicial = self.filaInicial
        Datos = False

        dataUsuarios = []
        datacsv = []

        # CSV
        file_users = 'usuarios.csv'
        file_csv = 'archivoCsv.csv'

        # Abre los archivos CSV en modo de escritura
        with open(file_users, mode='w', newline='') as file1, open(file_csv, mode='w', newline='') as file2:
            writerUsers = csv.writer(file1)
            writerCsv = csv.writer(file2)
    
            # Escribe los encabezados de las columnas en ambos archivos
            writerUsers.writerow(columns)
            writerCsv.writerow(columsCsv)

            # Itera sobre las filas de datos
            for row in range(filaInicial, cantOfRows):
                indexEmail = get_column_letter(2)
                email = information[indexEmail + str(row)].value

                indexSede = get_column_letter(3)
                sede = information[indexSede + str(row)].value

                indexFacultad = get_column_letter(4)
                facultad = information[indexFacultad + str(row)].value

                indexCod_plan = get_column_letter(5)
                cod_plan = information[indexCod_plan + str(row)].value

                indexPlan = get_column_letter(6)
                plan = information[indexPlan + str(row)].value

                indexNivel = get_column_letter(6)
                nivel = information[indexNivel + str(row)].value

                # Filtros

                if sede not in ["SEDE BOGOTÁ"]:
                    continue
                
                if facultad not in ['FACULTAD DE INGENIERÍA']:
                    continue
                
                # Escribe la fila en ambos archivos CSV
                row_user = [email, sede, facultad, cod_plan, plan, nivel]
                row_csv = list(infoCsv.values())
                row_csv[1] = email 

                writerUsers.writerow(row_user)
                writerCsv.writerow(row_csv)
    
    # ------------- Manejo de excel ----------------------------------

    def print_data(self, nombreHoja):
        
        '''
        Se va a recorrer la hoja de exel extrayendo la cantidad de filas 
        atravez de la libreria, mietras que la cantidad de columnas la 
        extraemos dependiendo de hoja que se esta recorriendo. 
        '''

        information = self.excel.get_sheet_by_name(nombreHoja)
        cantOfRows = len(list(information.rows))
        cantOfColumns = getCantOfColumns(nombreHoja)
        maxColumns = self.columnaInicial + cantOfColumns
        filaInicial = self.filaInicial

        Datos = False
        for row in range(filaInicial ,cantOfRows):
            for column in range(self.columnaInicial, maxColumns):
                columnChar = get_column_letter(column)
                value = information[columnChar + str(row)].value
                print("cell "+columnChar + str(row), end= " : ")
                print(str(value) + " | ", end= " ")
            print()
            print("--------")
            print()

        print("Cantidad de filas : ")
        print(cantOfRows)
        print("Cantidad de columnas : ")
        print(maxColumns)

    def getHeaders(self, information, maxColumns, nombreHoja):
                
        '''
        Lee la hoja de excel seleccionada, ingresa el nombre de las columnas
        en un arreglo y va indicando en un arreglo de boleanos si ese dato es
        obligatorio.
        '''

        headers = []
        columnasObligatorias = []
        
        filaInicial = self.filaInicial
        if nombreHoja == "GANADO":
            filaInicial = self.filaInicialGanado

        for column in range(self.columnaInicial, maxColumns):
            columnChar = get_column_letter(column)
            value = information[columnChar + str(filaInicial - 1)].value
            value = str(value).strip()
            
            if value.startswith("*"):
                columnasObligatorias.append(True)
            else:
                columnasObligatorias.append(False)
            
            headers.append(value)
        
        return (headers, columnasObligatorias)
    
    # ------------------------ Utilidades --------------------
                          
    def IsemptyRow(self, rowData):
        cantOfCells = len(rowData) - 1
        cantEmptyCell = 0

        for cell in rowData:
            if cell == None:
                cantEmptyCell += 1

        if cantEmptyCell >= cantOfCells:
            return True
        
        return False
    