# -*- coding: utf-8 -*-
"""
Created on Fri Mar 22 21:37:35 2024

@author: DJIMENEZ
"""

# Librerias necesarias -------------------------------------------------------------------------

import os
import uuid
import ssl
import smtplib
import mimetypes
from email import encoders
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart

# Clase para enviar correo electrónico
class EmailManager:
    def __init__(self, account:str, password:str, smtp_server:str='smtp.office365.com', smtp_port:str=587):
        self.account = account
        self.__password = password
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
            
    def send(self, subject:str, body:str, direct_recipients:list[str], copied_recipients:list[str] = [], blind_recipients:list[str] = [], files:list[str] | None = None, inline_images:list[str] | None = None, signature_file: str | None = None) -> bool:
        """
        Send an email with HTML body, embedded images, and properly typed files.

        Params:
            subject (str): Email subject.
            body (str): HTML content of the message.
            direct_recipients (list): List of direct recipients.
            copied_recipients (list): List of copied recipients.
            blind_recipients (list): List of blind copied recipients.
            files (list): List of file paths to attach.
            inline_images (list): List of image paths to embed in the email body.
            signature_file (str): Path to HTML signature file.
        Returns:
            bool: True if email was sent successfully, False otherwise.
        """

        # Crear mensaje base (outer multipart/mixed)
        msg_mixed = MIMEMultipart('mixed')
        msg_mixed['From'] = self.account
        msg_mixed['To'] = "; ".join(direct_recipients)
        msg_mixed['Cc'] = "; ".join(copied_recipients)
        msg_mixed['Bcc'] = "; ".join(blind_recipients)
        msg_mixed['Subject'] = subject
        success = True

        # Procesar imágenes inline y añadirlas al related
        msg_related = MIMEMultipart('related')
        msg_alternative = MIMEMultipart('alternative')


        # Procesar inline_images proporcionadas: no las adjuntamos todavía, solo las registramos y modificamos el body
        images_to_attach: list[tuple[str, str, str]] = []
        if inline_images:
            for image in inline_images:
                if not os.path.isfile(image):
                    print(f"❌ Image not found: {image}")
                    success = False
                    continue

                file_name = os.path.basename(image)
                base_file_name = os.path.splitext(file_name)[0]
                content_id = str(uuid.uuid4())
                images_to_attach.append((image, content_id, file_name))
                body = body.replace(f'cid:{file_name}', f'cid:{content_id}')
                body = body.replace(f'cid:{base_file_name}', f'cid:{content_id}')

        
        # Si se proporcionó un archivo de firma, se lee y adjunta
        if signature_file:
            if os.path.isfile(signature_file):
                with open(signature_file, 'r', encoding='utf-8') as sf:
                    signature_html = sf.read()
            else:
                print(f"❌ Signature file not found: {signature_file}")
                signature_html = ''
        else:
            signature_html = ''

        # Añadir la firma al cuerpo
        if "</body>" in body.lower():
            body = body.replace("</body>", f"{signature_html}</body>")
        else:
            body += signature_html

        # Cuerpo HTML
        msg_alternative.attach(MIMEText(body, 'html', 'utf-8'))
        msg_related.attach(msg_alternative)
        msg_mixed.attach(msg_related)


        # Adjuntar las imágenes encontradas al related
        for path, cid, filename in images_to_attach:
            try:
                with open(path, 'rb') as imgf:
                    img_part = MIMEImage(imgf.read(), name=filename)
                    img_part.add_header('Content-ID', f'<{cid}>')
                    img_part.add_header('Content-Disposition', 'inline', filename=filename)
                    msg_related.attach(img_part)

            except Exception as e:
                print(f"❌ Error attaching inline image {path}: {e}")
                success = False

        # Adjuntar archivos
        if files:
            for file in files:
                if not os.path.isfile(file):
                    print(f"❌ File not found: {file}")
                    success = False
                    continue

                ctype, encoding = mimetypes.guess_type(file)
                if ctype is None or encoding is not None:
                    ctype = 'application/octet-stream'
                maintype, subtype = ctype.split('/', 1)

                with open(file, 'rb') as f:
                    mime_part = MIMEBase(maintype, subtype)
                    mime_part.set_payload(f.read())

                encoders.encode_base64(mime_part)
                file_name = os.path.basename(file)
                mime_part.add_header('Content-Disposition', f'attachment; filename="{file_name}"')
                msg_mixed.attach(mime_part)


        # Enviar correo con conexión segura
        # Obtener todos los destinatarios (to + cc + bcc)
        all_recipients = direct_recipients + copied_recipients + blind_recipients
        
        if success:
            context = ssl.create_default_context()
            try:
                with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
                    server.starttls(context=context)
                    server.login(self.account, self.__password)
                    server.sendmail(self.account, all_recipients, msg_mixed.as_string())

                print("✅ Email sent successfully.")

            except smtplib.SMTPException as e:
                print(f"❌ Error sending email: {e}")
                success = False
        else:
            print("❌ Not sending email due to previous errors.")

        return success
