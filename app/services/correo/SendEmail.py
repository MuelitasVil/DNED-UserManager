import tempfile
import smtplib
from email.message import EmailMessage

def sendEmailSoporte(fileName, archivoExcel, respuesta, tipo):
    
    fileName = fileName.split(".")
    fileName = fileName[0] + ".xlsx"

    mensaje = (
        ""
    )

    if respuesta == True:
        estado = "Exitoso"
        mensaje += "Estado : Exitoso \n"
    else:
        estado = "Fallido"
        mensaje += "Estado : Fallido. \n"
        mensaje += "Error : " + str(respuesta)

    remitente = ""
    destinatarios = [""]
    
    subject = (
        "{} (carga masiva : {} | estado : {})".format(nombre, tipo, estado)
    )
    
    for destinatario in destinatarios:
        email = EmailMessage()
        email['From'] = remitente
        email["Subject"] = subject
        email['To'] = destinatario
        email.set_content(mensaje)
    
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_file:
            archivoExcel.save(temp_file.name)
            email.add_attachment(open(temp_file.name, 'rb').read(), maintype='application', subtype='octet-stream', filename=fileName)

        try:
            smtp = smtplib.SMTP("", port=587)
            smtp.starttls()
            smtp.login(remitente, "")
            smtp.sendmail(remitente, destinatario, email.as_string())
            smtp.quit()
        except Exception as e:
            print("Error al enviar el correo electrónico:", e)
        finally:
            return
