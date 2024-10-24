import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Configura tu dirección de correo de Outlook y la del destinatario
sender_email = "lopezr@uif.gob.ar"
receiver_email = "juanignacio_lopez@hotmail.com"
password = "Miranda2024"

# Crear el mensaje
subject = "Correo desde Outlook"
body = "Este correo fue enviado automáticamente usando Python y Outlook."

msg = MIMEMultipart()
msg['From'] = sender_email
msg['To'] = receiver_email
msg['Subject'] = subject

# Adjuntar el cuerpo del mensaje
msg.attach(MIMEText(body, 'plain'))

# Conexión con el servidor SMTP de Outlook
try:
    server = smtplib.SMTP('smtp-mail.outlook.com', 587)
    server.starttls()  # Iniciar la encriptación TLS
    server.login(sender_email, password)  # Iniciar sesión en tu cuenta de Outlook

    # Enviar el correo
    text = msg.as_string()
    server.sendmail(sender_email, receiver_email, text)

    print("Correo enviado con éxito desde Outlook")
except Exception as e:
    print(f"Error al enviar el correo: {e}")
finally:
    server.quit()  # Cerrar la conexión con el servidor
