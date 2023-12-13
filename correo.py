import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

def enviarCorreo(fecha_formateada_inicio, fecha_formateada_final, errores):
    # Configuración del servidor SMTP
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587    # El puerto puede variar según el servidor

    # Datos de la cuenta de correo
    correo_emisor = 'mpolanco@timix.cl'
    contrasena = 'bmxp vyaq mrpe qcwq'

    # Destinatario
    correo_destinatario = ['martinpolancosaavedra@gmail.com','fvillalobos@timix.cl','vherrera@imagina.cl','juanfrancisco@timix.cl','kbenavideso@imagina.cl','lchacon@imagina.cl']
    # Crear el mensaje
    asunto = 'Bot Contribuciones Terrenos'
    if errores==True:
        cuerpo_mensaje = """Estimados/as
        El Bot de Contribuciones terrenos se ha ejecutado con fecha inicio: """+str(fecha_formateada_inicio)+""" y fecha termino: """+str(fecha_formateada_final)+""" y a presentado errores, se adjunta la lista de errores
        Un saludo
        Bot Imagina"""
    
    else:
        cuerpo_mensaje="""Estimados/as
        El Bot de Contribuciones terrenos se ha ejecutado correctamente con fecha inicio: """+str(fecha_formateada_inicio)+""" y fecha termino: """+str(fecha_formateada_final)+"""
        Un saludo
        Bot Imagina"""
    

    mensaje = MIMEMultipart()
    mensaje['From'] = correo_emisor
    mensaje['To'] = ', '.join(correo_destinatario)
    mensaje['Subject'] = asunto

    # Adjuntar un archivo Excel
    archivo_excel = "Correo/Formato Reporte.xlsx"
    nombre_adjunto = "Reporte Bot Contribuciones terrenos "+str(fecha_formateada_final)+".xlsx"

    adjunto = open(archivo_excel, 'rb')

    parte_adjunta = MIMEBase('application', 'octet-stream')
    parte_adjunta.set_payload(adjunto.read())
    encoders.encode_base64(parte_adjunta)
    parte_adjunta.add_header('Content-Disposition', f'attachment; filename= {nombre_adjunto}')

    mensaje.attach(MIMEText(cuerpo_mensaje, 'plain'))
    mensaje.attach(parte_adjunta)

    # Configurar el servidor SMTP y enviar el correo
    try:
        servidor_smtp = smtplib.SMTP(smtp_server, smtp_port)
        servidor_smtp.starttls()  # Usar TLS para seguridad
        servidor_smtp.login(correo_emisor, contrasena)
        servidor_smtp.sendmail(correo_emisor, correo_destinatario, mensaje.as_string())
        servidor_smtp.quit()
        print("Correo enviado exitosamente.")
    except Exception as e:
        print("Error al enviar el correo:", str(e))
