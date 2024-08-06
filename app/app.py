from flask import Flask, render_template, request, redirect, url_for
from app.routes.DatosInformacion.cargaExcel import carga

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = "static/uploads"
app.secret_key = 'th3r4n0ms1ng'

app.register_blueprint(carga)

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=5000)
