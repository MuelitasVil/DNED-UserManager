from app.services.excel.utils import getCantOfColumns
from app.services.excel.excel.archvivosExcel import ArchivosExcel
from openpyxl.utils import get_column_letter
from openpyxl import Workbook, load_workbook
import openpyxl 
import os
import csv

'''
La siguiente clase tendra la funcionalidad de recorrer el excel en el cual se
encuentran los datos de los inquilinos.

Ojo : Es necesario que entienda la estructura del excel para entender las
siguientes funciones.
'''

class CorreosEstudiantes:
    def __init__(self, file = None, files = []):
        # Lectura del excel 
        if file != None:
            self.excel = openpyxl.load_workbook(file)
            self.hojas = self.excel.get_sheet_names()

        self.folder_pathDocentes = ""
        if not os.path.exists(self.folder_path):
            os.makedirs(self.folder_path)

        self.UserFiles = files
        self.columnaInicial = 1
        self.filaInicial = 1   

        self.bogota = "SEDE BOGOTÁ"
        self.amazona = "SEDE AMAZONÍA"
        self.caribe = "SEDE CARIBE"
        self.paz = "SEDE DE LA PAZ" # 
        self.manizales = "SEDE MANIZALES" #
        self.medellin = "SEDE MEDELLÍN" # 
        self.orinoquia = "SEDE ORINOQUÍA" #
        self.palmira = "SEDE PALMIRA" #
        self.tumaco = "SEDE TUMACO"

    def FilterEstudiantes(self):
        
        '''
        Nota : Esta solucion se podria hacer con un arbol binario.
        Sin embargo lo implemente con diccionarios por facilidad

        Estructura : 
        Sedes -> facultades -> unidadres -> profesores

        Dict-Profesores:
        dict = {
            "sede" : {
                "Unidad" : {
                    
                    }
                }
            }

        Estructura : 
        Sedes  -> administrativos

        Dict-Administrativos:
        dict = {
            "sede" : {
                }
            }


        '''

        # Datos : 

        ArchivoDocentesAdministrativos = ArchivosExcel.Docentes 
        information = self.excel["Hoja1"]
        cantOfRows = len(list(information.rows))
        filaInicial = self.filaInicial
 
        #woorkbookOfplans.create_sheet("Informacion General")
        
        # Obtener informacion : 
        dict_Of_Docentes = {}
        dict_Of_Administrativos = {}

        print("OBTENIENDO INFORMACION : ")
        for row in range(filaInicial, cantOfRows):
            
            columnNombreVinculacion = get_column_letter(ArchivoDocentesAdministrativos.Vinculacion + 1)
            nombreVinculacion = str(information[columnNombreVinculacion + str(row)].value)
            
            if "DOCENTE" in nombreVinculacion:
                self.fillDictDocentes(row, dict_Of_Docentes, information)
            else: 
                self.fillDictAdministrativos(row, dict_Of_Administrativos, information)

        for sede in dict_Of_Sedes:
            
            if sede == "SEDE":
                continue
            
            print("Rellenar excel " + sede)
            woorkbookSEDE = Workbook()
            woorkbookPLAN = Workbook()
            hojaSede = woorkbookSEDE.create_sheet(sede)

            dict_sede = dict_Of_Sedes[sede]
            usuariosSede = list(dict_sede.keys())
            
            self.fillListaCorreos(hojaSede, sede, usuariosSede, "SEDE", "FACULTAD", sede)
            
            for facultad in dict_sede:
                hojaFacultad = woorkbookSEDE.create_sheet(facultad)
                dict_facultad = dict_sede[facultad]
                
                usuariosFacultad = list(dict_facultad.keys())
                self.fillListaCorreos(hojaFacultad, facultad, usuariosFacultad, "FACULTAD", "PLAN", sede)

                for plan in dict_facultad:
                    hojaPlan = woorkbookPLAN.create_sheet(plan)
                    usuariosEstudiantes = dict_facultad[plan]
                    self.fillListaCorreos(hojaPlan, plan, usuariosEstudiantes, "PLAN", "ESTUDIANTE", sede, facultad)
            
            woorkbookSEDE.save("docentes/" + sede + ".xlsx")
            woorkbookPLAN.save("docentes/" + "PLANES " + sede + ".xlsx")
    
    def fillDictDocentes(self, row, dict_Of_Docentes, information):
            
            ArchivoDocentesAdministrativos = ArchivosExcel.Docentes

            columnSede = get_column_letter(ArchivoDocentesAdministrativos.Sede + 1)
            sede = str(information[columnSede + str(row)].value)
            
            if sede not in dict_Of_Docentes:
                dict_Of_Docentes[sede] = {}
            
            dict_Of_Facultades = dict_Of_Docentes[sede]
            
            columFacultad = get_column_letter(ArchivoDocentesAdministrativos.Facultad + 1)
            facultad = str(information[columFacultad + str(row)].value)

            if facultad not in dict_Of_Facultades:
                dict_Of_Docentes[facultad] = {}
            
            columUnidad = get_column_letter(ArchivoDocentesAdministrativos.Unidad + 1)
            unidad = str(information[columUnidad + str(row)].value)

            dict_Of_Unidades = dict_Of_Facultades[facultad]
            
            columUnidad = get_column_letter(ArchivoDocentesAdministrativos.Unidad + 1)
            unidad = str(information[columUnidad + str(row)].value)

            if unidad not in dict_Of_Unidades:
                dict_Of_Unidades[unidad] = []

            columnCorreo = get_column_letter(ArchivoDocentesAdministrativos.Correo + 1)
            correo = str(information[columnCorreo + str(row)].value)

            dict_Of_Unidades[unidad].append(correo)

    def fillDictAdministrativos(self, row, dict_Of_Administrativos, information):
        
        ArchivoDocentesAdministrativos = ArchivosExcel.Docentes

        columnSede = get_column_letter(ArchivoDocentesAdministrativos.Sede + 1)
        sede = str(information[columnSede + str(row)].value)
            
        if sede not in dict_Of_Administrativos:
            dict_Of_Administrativos[sede] = []

        columnCorreo = get_column_letter(ArchivoDocentesAdministrativos.Correo + 1)
        correo = str(information[columnCorreo + str(row)].value)
        
        dict_Of_Administrativos[sede].append(correo)
           
    def fillListaCorreos(self, hoja, GroupMember, users, tipoGroup, tipoUser, sede, facultad = None):
        hoja["A1"] = "Group Email"
        hoja["B1"] = "Member Email"
        hoja["C1"] = "Member Type"
        hoja["D1"] = "Member Role"
        hoja["G1"] = "Member NAME"

        row = 2
        userGroupMember = self.get_EmailMember(GroupMember, tipoGroup, sede)
        row = self.PropietariosAllListas(hoja, row, userGroupMember)
        row = self.PropietariosSede(hoja, row, userGroupMember, sede)
        
        if tipoGroup == "FACULTAD" or tipoGroup == "PLAN":
            row = self.PropietariosFacultad(hoja,userGroupMember, GroupMember, tipoGroup, sede, row, facultad)

        for user in users: 
            hoja["A" + str(row)] = userGroupMember
            hoja["B" + str(row)] = self.get_EmailMember(user, tipoUser, sede)
            hoja["C" + str(row)] = "USER" 
            hoja["D" + str(row)] = "MEMBER"
            hoja["G" + str(row)] = user
            row += 1 
            
    def get_EmailMember(self, user : str, tipoUser : str, sede):
        if tipoUser == "ESTUDIANTE":
            return user
        
        if tipoUser == "SEDE":
            # "SEDE BOGOTA"
            sede = user.split(" ")
            # "[SEDE, BOGOTA]"
            sede = sede[1][:3].lower()
            # "bog"
            return "estudiante_" + sede + "@unal.edu.co"

         # "SEDE BOGOTA"
        sede = sede.split(" ")
        # "[SEDE, BOGOTA]"
        sede = sede[1][:3].lower()
        # "bog"  

        user = user.split(" ")
        acronimo = ""

        if tipoUser == "FACULTAD":
        
            if (sede == "ama" or sede == "car" 
                or sede == "ori" or sede == "tum"):
                return "estf_" + sede + "@unal.edu.co"

            for palabra in user:
                if len(palabra) > 2:
                    acronimo += palabra.lower()[0]    
            
            return "estf" + acronimo + "_" + sede + "@unal.edu.co"
        
        if tipoUser == "PLAN":
            for palabra in user:
                if len(palabra) > 2:
                    acronimo += palabra.capitalize()[:3]
            
            return acronimo + "_" + sede + "@unal.edu.co"
    
    def PropietariosAllListas(self, hoja, row, userGroupMember):
    
        listaNacional = [
        "boletin_un@unal.edu.co",
        "comdninfoa_nal@unal.edu.co",
        "enviosvri_nal@unal.edu.co",
        "rectorinforma@unal.edu.co",
        "comunicado_csu_bog@unal.edu.co",
        "reconsejobu_nal@unal.edu.co",
        "dninfoacad_nal@unal.edu.co",
        "dgt_dned@unal.edu.co",
        "gruposeguridad_nal@unal.edu.co",
        "sisii_nal@unal.edu.co",
        "postmaster_unal@unal.edu.co",
        "postmasterdnia_nal@unal.edu.co",
        "protecdatos_na@unal.edu.co",
        "infraestructurati_dned@unal.edu.co",
        "dre@unal.edu.co",
        "dned@unal.edu.co",

        # Representacion estudiantil 
        "estudiantilcsu@unal.edu.co",
        "estudiantilca@unal.edu.co"
        ]
        
        for owner in listaNacional:
            hoja["A" + str(row)] = userGroupMember
            hoja["B" + str(row)] = owner
            hoja["C" + str(row)] = "USER" 
            hoja["D" + str(row)] = "OWNER"
            hoja["G" + str(row)] = "OWNER COLOMBIA"
            row += 1 
        
        return row
    
    def PropietariosSede(self, hoja, row, userGroupMember, sede):
        lista_sede = []
        
        if sede == self.medellin:
            lista_sede = [
                "alertas_med@unal.edu.co",
                "informa_biblioteca@unal.edu.co",
                "informa_comunicaciones@unal.edu.co",
                "informa_direccion_administrativa@unal.edu.co",
                "informa_direccion_laboratorios@unal.edu.co",
                "informa_fac_ciencias_humanas_y_economicas@unal.edu.co",
                "informa_juridica@unal.edu.co",
                "inf_aplicaciones_med@unal.edu.co",
                "informa_vicerrectoria@unal.edu.co",
                "informa_bienestar_universitario@unal.edu.co",
                "infservcomp_med@unal.edu.co",
                "inflogistica_med@unal.edu.co",
                "informa_fac_ciencias@unal.edu.co",
                "informa_fac_minas@unal.edu.co",
                "informa_fac_ciencias_agrarias@unal.edu.co",
                'info_aplica_med@unal.edu.co',
                "informa_secretaria_sede@unal.edu.co",
                "innovaacad_med@unal.edu.co",
                "unalternativac_nal@unal.edu.co",
                "pcm@unal.edu.co",
                "postmaster_med@unal.edu.co",
                "infeducontinua@unal.edu.co",
                "informa_direccion_academica@unal.edu.co",
                "informa_direccion_de_investigacion_y_extension@unal.edu.co",
                "informa_direccion_ordenamiento_y_desarrollo_fisico@unal.edu.co",
                "informa_fac_arquitectura@unal.edu.co",
                "informa_registro_y_matricula@unal.edu.co",
                "informa_unimedios@unal.edu.co",
                "infpersonal_med@unal.edu.co",
                # Repesentacion estudiantil
                "reestudia_med@unal.edu.co"
            ]

        if sede == self.manizales:
            lista_sede = [
                "ventanilla_man@unal.edu.co",
                "bienestar_man@unal.edu.co",
                "planea_man@unal.edu.co",
                "postmaster_man@unal.edu.co",
                "vicsede_man@unal.edu.co",
                "personaladm_man@unal.edu.co",
                "personaldoc_man@unal.edu.co",
                "saludocup_man@unal.edu.co",
                # Repesentacion manizales
                "estudiantilcs_man@unal.edu.co"
            ]
        
        if sede == self.palmira:
            lista_sede = [
                "unnoticias_pal@unal.edu.co",
                "postmaster_pal@unal.edu.co",
                # Representacion 
                "estudiantilcs_pal@unal.edu.co"
            ]

        if sede == self.orinoquia:
            lista_sede = [
                "divcultural_ori@unal.edu.co"
            ]

        if sede == self.paz:
            lista_sede = [
                "secsedelapaz@unal.edu.co",
                "sedelapaz@unal.edu.co",
                "tics_paz@unal.edu.co",
                "vicesedelapaz@unal.edu.co"
            ]

        if sede == self.bogota:
            lista_sede = [
                "divulgaciondrm_bog@unal.edu.co",
                "talenhumano_bog@unal.edu.co",
                "reprecarrera_bog@unal.edu.co",
                "comunicaciones_bog@unal.edu.co",
                "diracasede_bog@unal.edu.co",
                "dircultural_bog@unal.edu.co",
                "notificass_bog@unal.edu.co",
                "personaladm_bog@unal.edu.co",
                "postmaster_bog@unal.edu.co",
                "salarialp_bog@unal.edu.co"
            ]
        
        for owner in lista_sede:
            hoja["A" + str(row)] = userGroupMember
            hoja["B" + str(row)] = owner
            hoja["C" + str(row)] = "USER" 
            hoja["D" + str(row)] = "OWNER"
            hoja["G" + str(row)] = "OWNER SEDE"
            row += 1 
        
        return row

    def PropietariosFacultad(self, hoja, userGroupMember, GroupMember, tipoGroup, sede, row, facultad):
        
        if sede != self.bogota:
            return row
        
        if tipoGroup == "PLAN":
            facultad = facultad
        
        if tipoGroup == "FACULTAD":
            facultad = GroupMember

        FacultadBogota = {
            "FACULTAD DE CIENCIAS HUMANAS" : "correo_fchbog@unal.edu.co",
            "FACULTAD DE INGENIERÍA" : "correo_fibog@unal.edu.co",
            "FACULTAD DE CIENCIAS" : "correo_fcbog@unal.edu.co",
            "FACULTAD DE ARTES" : "correo_farbog@unal.edu.co",
            "FACULTAD DE CIENCIAS ECONÓMICAS" : "correo_fcebog@unal.edu.co",
            "FACULTAD DE MEDICINA" : "correo_fmbog@unal.edu.co ",
            "FACULTAD DE DERECHO, CIENCIAS POLÍTICAS Y SOCIALES" : "correo_fdbog@unal.edu.co",
            "FACULTAD DE MEDICINA VETERINARIA Y DE ZOOTECNIA" : "correo_fmvbog@unal.edu.co",
            "FACULTAD DE CIENCIAS AGRARIAS" : "correo_fcabog@unal.edu.co",
            "FACULTAD DE ENFERMERÍA" : "correo_febog@unal.edu.co",
            "FACULTAD DE ODONTOLOGÍA" : "correo_fobog@unal.edu.co"
        }

        hoja["A" + str(row)] = userGroupMember
        hoja["B" + str(row)] = FacultadBogota[facultad]
        hoja["C" + str(row)] = "USER" 
        hoja["D" + str(row)] = "OWNER"
        hoja["G" + str(row)] = "OWNER SEDE"
        row += 1

        return row 
        

    # ------------- Manejo de excel ----------------------------------

    def print_data(self, nombreHoja):
        
        '''
        Se va a recorrer la hoja de exel extrayendo la cantidad de filas 
        atravez de la libreria, mietras que la cantidad de columnas la 
        extraemos dependiendo de hoja que se esta recorriendo. 
        '''

        information = self.excel[nombreHoja]
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

