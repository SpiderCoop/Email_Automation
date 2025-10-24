import os
from email_automation import send_email
from dotenv import load_dotenv

# Se toma variables del archivo .env
load_dotenv()
cuenta = os.getenv('Cuenta')
password = os.getenv('password')
destinatarios = os.getenv("Destinatarios").split(',')

asunto = "Prueba del correo"
cuerpo_correo = """
<html>
<head></head>
<body style='font-family: Arial, sans-serif; font-size: 15px;'>
<p><b>Prueba</b></p>
</body>
</html>"""

send_email(cuenta, password, destinatarios, asunto, cuerpo_correo)