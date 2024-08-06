class EgresadosIndex:    
    TipoArchivo = "Egresados"
    NombreArchivo = "Precenso Egresados - Rector 2024-2027 (1).xlsx"
    NombreHoja = "Hoja1"
    DocumentoIdentidad = 0
    Nombres = 1
    Apellidos = 2
    Genero = 4
    Correo = 5
    CorreoAternativa = 6
    Facultad = 7
    Sede = 8
    Plan = 9
    CodigoSede = 10
    CodigoFacultad = 11
    CodigoPlan = 12

class EstudiantesActivos:
    TipoArchivo = "EstudiantesActivos"
    NombreArchivo = "2024_04_02 Estudiantes activos.xlsx"
    NombreHoja = "ESTUDIANTES ACTIVOS 2024-1S"
    NombresXApellidos = 0
    Correo = 1
    Sede = 2
    Facultad = 3
    CodigoPlan = 4
    Plan = 5
    Nivel = 6

class Docentes:
    TipoArchivo = "Docentes"
    NombreArchivo = "Listado Docentes y Administrativos 25072024.xlsx"
    NombreHoja = "Export Worksheet"
    Sede = 0
    ApellidosXNombre = 1
    Correo = 2
    Cargo = 3
    Unidad = 4
    Funcion = 5
    UnidadFuncion = 6
    Vinculacion = 7

class WorkSpace:
    TipoArchivo = "WorkSpace"
    NombreArchivo = "usuarios workspace almacenamiento 01082024.xlsx"
    NombreHoja = "User_Download_01082024_091650"
    Nombre = 0
    Apellidos = 1
    Correo = 2
    Estado = 3
    UltimaConexion = 4
    Almacenamiento = 5
    LimiteAlmacenamiento = 6
    
class Ldap:
    # Ejemplo entrada csv
    # ['uid=rleonc,ou=People,o=unal.edu.co', 'rleonc','Rosalba Leon Castro','52361091', '1', 'rleonc@unal.edu.co', 'bogota']
    TipoArchivo = "Ldap"
    NombreArchivo = "Export full LDAP.csv"
    TipoUsuario = {
    "1" : "Estudiante",
    "2" : "Docente",
    "3" : "Administrativo",
    "4" : "Dependencia",
    "5" : "Contratista",
    "6" : "Dependencia",  
    "7" : "Pensionado",
    "8" : "Egresado"
    }

    # Ojo tener cuidado que no todas las filas csv no son iguales.
    uui = 0
    usuario = 1
    NombreXApellidos = 2 
    Dependencia = 3
    Tipo = 4
    Correo = 5
    Sede = 6

    TamañoArreglo = 9

class ArchivosExcel:
    Egresados = EgresadosIndex
    EstudiantesActivos = EstudiantesActivos
    Docentes = Docentes
    WorkSpace = WorkSpace
    Ldap = Ldap