from flask import Blueprint, render_template, request, abort, redirect, url_for, send_file, flash, session
from openpyxl import Workbook, load_workbook
from app.services.excel.excel.excelManager import ExcelManager
from openpyxl.drawing.image import Image
from werkzeug.utils import secure_filename
from datetime import datetime

carga = Blueprint("carga", __name__, static_folder="static", template_folder="templates")

@carga.route('/', methods = ["POST", "GET"])
def subir():
    return render_template("DatosInformacion/cargas.html")

'''
@carga.route("/downloadCargasEjemplo", methods = ["POST", "GET"])
def Download_FileDatos():
    PATH = "static/excel/Plantilla de cargas inventarios - Datos.xlsm"
    return send_file(PATH, as_attachment=True)
'''

@carga.route('/Create-csv', methods = ["POST"])
def upload():
    file = request.files['uploadFile']

    if not file:
        flash('Por favor ingrese su archivo')
        return render_template("DatosInformacion/cargas.html")
     
    fileName = secure_filename(file.filename)

    if not VerifyExel(fileName):
        flash('Por favor ingrese su archivo excel para cargas masivas')
        return render_template("DatosInformacion/cargas.html")
    
    excel = ExcelManager(file = file)
    respuesta = excel.get_csv()
    
    if respuesta != True:
        flash("Se han creado los csv")    
        return render_template("DatosInformacion/cargas.html")    
    
    session['sync'] = False
    return redirect(url_for("carga.subir"))

@carga.route('/Create-Merge', methods = ["POST"])
def Merge():

    files = {
        "file1" : request.files['uploadFile1'],
        "file2" : request.files['uploadFile2'],
        "file3" : request.files['uploadFile3'],
        "file4" : request.files['uploadFile4'],
        "file5" : request.files['uploadFile5']
    }

    for file in files:
        archivo = files[file]
        if not archivo:
            flash('Por favor ingrese su archivo en ' + file)
            return render_template("DatosInformacion/cargas.html")
            
        fileName = secure_filename(archivo.filename)
        if not VerifyExel(fileName):
            flash('Por favor ingrese su archivo excel para cargas masivas')
            return render_template("DatosInformacion/cargas.html")
    
    excel = ExcelManager(files = list(files.values()))
    respuesta = excel.MergueUsuarios()
    respuesta = False
    if respuesta != True:
        flash("Se han creado los csv")    
        return render_template("DatosInformacion/cargas.html")    
    
    session['sync'] = False
    return redirect(url_for("carga.subir"))


def VerifyExel(fileName):
    if fileName.endswith(".xlsm") or fileName.endswith(".xlsb") or fileName.endswith(".xlsx") or fileName.endswith("csv"):
        return True
    return False
