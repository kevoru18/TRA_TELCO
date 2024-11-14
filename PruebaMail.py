import smtplib
from email.mime.text import MIMEText

def test_email():
    destinatario = "federico_dargallo@yahoo.es"
    remitente = "federico.dargallo@gmail.com"
    password = "nruuvmqjnpucpdie"  # Asegúrate de usar la contraseña correcta o la contraseña de aplicación


    msg = MIMEText("Este es un mensaje de prueba.")
    msg['Subject'] = "Prueba de correo"
    msg['From'] = remitente
    msg['To'] = destinatario

    try:
        # Usar SMTP_SSL para el puerto 465
        with smtplib.SMTP_SSL('smtp.mail.yahoo.com', 465) as server:
            server.login(remitente, password)
            server.send_message(msg)
        print("Correo enviado correctamente.")
    except smtplib.SMTPAuthenticationError:
        print("Error de autenticación: verifica tu correo y contraseña.")
    except smtplib.SMTPException as e:
        print(f"Error al enviar el correo: {e}")
    except Exception as e:
        print(f"Error inesperado: {e}")

test_email()