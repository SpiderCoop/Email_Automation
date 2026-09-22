"""
Description:   Manages sending emails with HTML content, inline images,attachments, and IMAP drafts.
Author:        David Jiménez Cooper - SpiderCoop
Date:          2024-03-22
"""

from __future__ import annotations

import imaplib
import mimetypes
import os
import re
import smtplib
import ssl
import time
import uuid
from email.message import EmailMessage
from email.utils import formataddr
from pathlib import Path
from typing import Sequence


class EmailManager:
    """
    Gestor de correo electrónico para Microsoft 365 / SMTP + IMAP.

    Permite:
        - Destinatarios To, Cc y Bcc.
        - Contenido HTML.
        - Firma desde archivo HTML o string HTML.
        - Imágenes inline mediante CID.
        - Archivos adjuntos.
        - Envío mediante SMTP + STARTTLS.
        - Guardado de borradores mediante IMAP.

    NOTA:
        La autenticación actualmente utiliza usuario/contraseña.
        Para Microsoft 365 puede ser necesario implementar OAuth 2.0
        / Modern Authentication.
    """

    def __init__(
        self,
        account: str,
        password: str,
        smtp_server: str = "smtp.office365.com",
        smtp_port: int = 587,
        imap_server: str = "outlook.office365.com",
        imap_port: int = 993,
    ) -> None:

        self.account = account
        self._password = password

        self.smtp_server = smtp_server
        self.smtp_port = smtp_port

        self.imap_server = imap_server
        self.imap_port = imap_port

        # Mensaje generado por build_message()
        self.message: EmailMessage | None = None

        # Archivo de firma HTML
        self.signature_file: str | None = None

        # Destinatarios
        self.direct_recipients: list[str] = []
        self.copied_recipients: list[str] = []
        self.blind_recipients: list[str] = []
        self.all_recipients: list[str] = []

    # ==================================================================
    # DESTINATARIOS
    # ==================================================================

    def set_recipients(
        self,
        direct_recipients: Sequence[str] | None = None,
        copied_recipients: Sequence[str] | None = None,
        blind_recipients: Sequence[str] | None = None,
    ) -> None:
        """
        Define los destinatarios del correo.

        Args:
            direct_recipients:
                Destinatarios principales (To).

            copied_recipients:
                Destinatarios en copia (Cc).

            blind_recipients:
                Destinatarios en copia oculta (Bcc).

        Raises:
            ValueError:
                Si no se especifica ningún destinatario.
        """

        self.direct_recipients = self._clean_recipients(
            direct_recipients
        )

        self.copied_recipients = self._clean_recipients(
            copied_recipients
        )

        self.blind_recipients = self._clean_recipients(
            blind_recipients
        )

        self.all_recipients = (
            self.direct_recipients
            + self.copied_recipients
            + self.blind_recipients
        )

        if not self.all_recipients:
            raise ValueError(
                "❌ At least one recipient must be specified."
            )

    @staticmethod
    def _clean_recipients(
        recipients: Sequence[str] | None,
    ) -> list[str]:
        """
        Limpia la lista de destinatarios.

        Elimina:
            - espacios al inicio/final
            - valores vacíos
        """

        if not recipients:
            return []

        return [
            recipient.strip()
            for recipient in recipients
            if recipient and recipient.strip()
        ]
    

    def set_signature(
        self,
        signature_file: str | None = None,
    ) -> None:
        """
        Define el archivo de firma que se agregará al final del correo.

        Args:
            signature_file: Ruta al archivo HTML que contiene la firma.
        """
        self.signature_file = signature_file

    # ==================================================================
    # CONSTRUCCIÓN DEL MENSAJE
    # ==================================================================

    def build_message(
        self,
        subject: str,
        body: str,
        files: Sequence[str | os.PathLike[str]] | None = None,
        inline_images: Sequence[str | os.PathLike[str]] | None = None,
    ) -> EmailMessage:
        """
        Construye el mensaje de correo.

        Args:
            subject: Asunto del correo.
            body: Contenido HTML.
            files: Archivos que serán enviados como attachments.
            inline_images: Imágenes que serán incrustadas dentro del HTML mediante referencias CID.

        Returns:
            EmailMessage: Mensaje completo listo para enviar o guardar como borrador.
        """


        # Verificamos destinatarios antes de construir el mensaje.
        self._validate_recipients()


        body, prepared_images = self._prepare_inline_images(
            body,
            inline_images,
        )

        # --------------------------------------------------------------
        # Procesamos la firma.
        # --------------------------------------------------------------

        signature_html = self._load_signature()

        body = self._append_signature(
            body,
            signature_html,
        )


        message = EmailMessage()

        message["From"] = self.account

        # EmailMessage acepta directamente una lista de destinatarios.
        if self.direct_recipients:
            message["To"] = ", ".join(self.direct_recipients)
            
        if self.copied_recipients:
            message["Cc"] = ", ".join(self.copied_recipients)

        if self.blind_recipients:
            message["Bcc"] = ", ".join(self.blind_recipients)

        message["Subject"] = subject


        message.set_content(
            body,
            subtype="html",
        )

        # --------------------------------------------------------------
        # Agregamos imágenes inline.
        # --------------------------------------------------------------

        for image_data in prepared_images:

            self._attach_inline_image(
                message=message,
                image_data=image_data,
            )

        # --------------------------------------------------------------
        # Agregamos attachments.
        # --------------------------------------------------------------

        if files:

            for file_path in files:

                self._attach_file(
                    message=message,
                    file_path=file_path,
                )

        self.message = message

        return self.message

    # ==================================================================
    # IMÁGENES INLINE
    # ==================================================================

    def _prepare_inline_images(
        self,
        body: str,
        inline_images: Sequence[str | os.PathLike[str]] | None,
    ) -> tuple[str, list[dict[str, object]]]:
        """
        Prepara las imágenes inline y reemplaza sus referencias CID.

        Soporta referencias como:

            cid:logo.png
            cid:logo

        Returns:
            tuple:
                - HTML modificado.
                - Lista de imágenes preparadas.
        """

        if not inline_images:
            return body, []

        prepared_images: list[dict[str, object]] = []

        for image in inline_images:

            path = Path(image)

            # ----------------------------------------------------------
            # CORRECCIÓN:
            # Validamos que la imagen exista y sea un archivo.
            # ----------------------------------------------------------

            if not path.is_file():

                raise FileNotFoundError(
                    f"❌ Image not found: {path}"
                )

            filename = path.name
            base_filename = path.stem


            content_id = uuid.uuid4().hex

            # ----------------------------------------------------------
            # Reemplazamos ambas variantes:
            #
            # cid:logo.png
            # cid:logo
            # ----------------------------------------------------------

            body = body.replace(
                f"cid:{filename}",
                f"cid:{content_id}",
            )

            body = body.replace(
                f"cid:{base_filename}",
                f"cid:{content_id}",
            )

            prepared_images.append(
                {
                    "path": path,
                    "filename": filename,
                    "content_id": content_id,
                }
            )

        return body, prepared_images


    def _attach_inline_image(
        self,
        message: EmailMessage,
        image_data: dict[str, object],
    ) -> None:
        """
        Adjunta una imagen como contenido inline relacionado con el HTML.
        """

        path = image_data["path"]
        filename = image_data["filename"]
        content_id = image_data["content_id"]

        # --------------------------------------------------------------
        # Validación de tipos.
        # --------------------------------------------------------------

        if not isinstance(path, Path):
            path = Path(path)

        if not isinstance(filename, str):
            raise TypeError(
                "Invalid inline image filename."
            )

        if not isinstance(content_id, str):
            raise TypeError(
                "Invalid inline image Content-ID."
            )

        # --------------------------------------------------------------
        # Determinamos el MIME type.
        # --------------------------------------------------------------

        mime_type, _ = mimetypes.guess_type(
            str(path)
        )

        if not mime_type or not mime_type.startswith("image/"):

            raise ValueError(
                f"❌ Unsupported image MIME type: {path}"
            )

        maintype, subtype = mime_type.split(
            "/",
            1,
        )

        # --------------------------------------------------------------
        # Leemos la imagen.
        # --------------------------------------------------------------

        try:

            image_bytes = path.read_bytes()

        except OSError as e:

            raise FileNotFoundError(
                f"❌ Error reading inline image {path}: {e}"
            ) from e


        message.add_related(
            image_bytes,
            maintype=maintype,
            subtype=subtype,
            cid=content_id,
            filename=filename,
        )

    # ==================================================================
    # ATTACHMENTS
    # ==================================================================

    def _attach_file(
        self,
        message: EmailMessage,
        file_path: str | os.PathLike[str],
    ) -> None:
        """
        Agrega un archivo como attachment.
        """

        path = Path(file_path)

        if not path.is_file():

            raise FileNotFoundError(
                f"❌ File not found: {path}"
            )

        mime_type, encoding = mimetypes.guess_type(
            str(path)
        )

        # --------------------------------------------------------------
        # CORRECCIÓN:
        # Si no podemos determinar el MIME type, utilizamos
        # application/octet-stream.
        # --------------------------------------------------------------

        if mime_type is None or encoding is not None:

            mime_type = "application/octet-stream"

        maintype, subtype = mime_type.split(
            "/",
            1,
        )

        try:

            file_data = path.read_bytes()

        except OSError as e:

            raise FileNotFoundError(
                f"❌ Error reading attachment {path}: {e}"
            ) from e

        # --------------------------------------------------------------
        # CORRECCIÓN:
        # EmailMessage genera automáticamente:
        #
        # Content-Type
        # Content-Disposition
        # Content-Transfer-Encoding
        #
        # y realiza el encoding apropiado.
        # --------------------------------------------------------------

        message.add_attachment(
            file_data,
            maintype=maintype,
            subtype=subtype,
            filename=path.name,
        )

    # ==================================================================
    # FIRMA
    # ==================================================================

    def _load_signature(self) -> str:
        """
        Carga la firma desde un archivo HTML o directamente desde
        un string que contenga HTML.
        """

        if not self.signature_file:
            return ""

        signature = self.signature_file

        # --------------------------------------------------------------
        # Si es un archivo existente.
        # --------------------------------------------------------------

        if os.path.isfile(signature):

            try:

                with open(
                    signature,
                    "r",
                    encoding="utf-8",
                ) as signature_file:

                    return signature_file.read()

            except OSError as e:

                raise RuntimeError(
                    f"❌ Error reading signature file "
                    f"{signature}: {e}"
                ) from e

        # --------------------------------------------------------------
        # Si parece ser HTML directamente.
        # --------------------------------------------------------------

        if "<" in signature and ">" in signature:

            return signature

        # --------------------------------------------------------------
        # Mantenemos el comportamiento original:
        # avisamos y no agregamos firma.
        # --------------------------------------------------------------

        print(
            "⚠️ The signature content is not a valid file "
            "nor does it appear to be HTML: "
            f"{signature[:30]}..."
        )

        return ""

    @staticmethod
    def _append_signature(
        body: str,
        signature_html: str,
    ) -> str:
        """
        Inserta la firma antes de </body>.

        Si el HTML no contiene </body>, agrega la firma al final.
        """

        if not signature_html:
            return body

        # --------------------------------------------------------------
        # CORRECCIÓN:
        # Utilizamos una búsqueda case-insensitive.
        #
        # Esto permite detectar:
        #
        # </body>
        # </BODY>
        # </Body>
        # --------------------------------------------------------------

        match = re.search(
            r"</body\s*>",
            body,
            flags=re.IGNORECASE,
        )

        if match:

            return (
                body[:match.start()]
                + signature_html
                + body[match.start():]
            )

        return body + signature_html

    # ==================================================================
    # VALIDACIÓN
    # ==================================================================

    def _validate_recipients(self) -> None:
        """
        Verifica que exista al menos un destinatario.
        """

        if not self.all_recipients:

            raise ValueError(
                "❌ At least one recipient must be specified. "
                "Use set_recipients() first."
            )

    def _validate_message(self) -> None:
        """
        Verifica que exista un mensaje construido.
        """

        if self.message is None:

            raise ValueError(
                "❌ No email message has been built. "
                "Call build_message() first."
            )

    # ==================================================================
    # ENVÍO SMTP
    # ==================================================================

    def send(self) -> bool:
        """
        Envía el correo mediante SMTP + STARTTLS.

        Returns:
            True si el correo fue enviado correctamente.

        Raises:
            RuntimeError:
                Si ocurre un error durante el envío.
        """

        self._validate_message()
        self._validate_recipients()


        message = self.message

        if message is None:
            raise ValueError(
                "❌ Email message is not available."
            )

        context = ssl.create_default_context()

        try:

            with smtplib.SMTP(
                self.smtp_server,
                self.smtp_port,
                timeout=30,
            ) as server:

                server.ehlo()

                # ------------------------------------------------------
                # STARTTLS
                # ------------------------------------------------------

                server.starttls(
                    context=context
                )

                server.ehlo()

                # ------------------------------------------------------
                # AUTENTICACIÓN
                #
                # NOTA:
                # Microsoft 365 puede requerir OAuth 2.0.
                # ------------------------------------------------------

                server.login(
                    self.account,
                    self._password,
                )

                # ------------------------------------------------------
                # CORRECCIÓN:
                #
                # send_message() puede utilizar directamente el
                # EmailMessage.
                #
                # Sin embargo, pasamos recipients explícitamente para
                # asegurarnos de incluir BCC.
                # ------------------------------------------------------

                server.send_message(
                    message,
                    from_addr=self.account,
                    to_addrs=self.all_recipients,
                )

            print(
                "✅ Email sent successfully."
            )

            return True

        except smtplib.SMTPAuthenticationError as e:

            raise RuntimeError(
                "❌ SMTP authentication failed. "
                "Verify the credentials and whether SMTP AUTH "
                "is enabled for the mailbox. "
                "For Microsoft 365, OAuth 2.0 may be required."
            ) from e

        except smtplib.SMTPException as e:

            raise RuntimeError(
                f"❌ SMTP error occurred: {e}"
            ) from e

        except OSError as e:

            raise RuntimeError(
                f"❌ Network/connection error while sending email: {e}"
            ) from e

        except Exception as e:

            raise RuntimeError(
                f"❌ An unexpected error occurred while sending email: {e}"
            ) from e

    # ==================================================================
    # IMAP - BORRADORES
    # ==================================================================

    def _get_drafts_folder(
        self,
        imap: imaplib.IMAP4_SSL,
    ) -> str:
        """
        Busca dinámicamente el nombre de la carpeta de borradores.

        Intenta detectar:
            - \\Drafts
            - Drafts
            - Borradores

        Si no encuentra ninguna, devuelve "Drafts".
        """

        status, folder_list = imap.list()

        if status != "OK" or not folder_list:

            return "Drafts"

        for folder_bytes in folder_list:

            if not folder_bytes:
                continue

            folder_str = folder_bytes.decode(
                "utf-8",
                errors="ignore",
            )

            folder_lower = folder_str.lower()

            # ----------------------------------------------------------
            # CORRECCIÓN:
            # Detectamos el flag estándar IMAP \Drafts.
            # ----------------------------------------------------------

            if "\\drafts" in folder_lower:

                folder_name = self._extract_imap_folder_name(
                    folder_str
                )

                if folder_name:
                    return folder_name

            # ----------------------------------------------------------
            # También soportamos nombres comunes.
            # ----------------------------------------------------------

            if (
                "drafts" in folder_lower
                or "borradores" in folder_lower
            ):

                folder_name = self._extract_imap_folder_name(
                    folder_str
                )

                if folder_name:
                    return folder_name

        return "Drafts"

    @staticmethod
    def _extract_imap_folder_name(
        folder_str: str,
    ) -> str | None:
        """
        Intenta extraer el nombre de la carpeta de una respuesta
        IMAP LIST.

        Ejemplo típico:

            (\\Drafts) "/" Drafts

        """

        # --------------------------------------------------------------
        # Intentamos obtener la parte posterior al delimitador.
        # --------------------------------------------------------------

        match = re.search(
            r'\)\s+"[^"]*"\s+(.+)$',
            folder_str,
        )

        if match:

            folder_name = match.group(1).strip()

            # Eliminamos comillas exteriores.
            if (
                len(folder_name) >= 2
                and folder_name.startswith('"')
                and folder_name.endswith('"')
            ):

                folder_name = folder_name[1:-1]

            if folder_name:
                return folder_name

        # --------------------------------------------------------------
        # Fallback para formatos diferentes.
        # --------------------------------------------------------------

        parts = folder_str.split('"')

        if parts:

            candidate = parts[-1].strip()

            if candidate:
                return candidate

        return None

    def save_draft(self) -> bool:
        """
        Guarda el mensaje como borrador mediante IMAP.

        Returns:
            True si el borrador se guardó correctamente.
        """

        self._validate_message()

        message = self.message

        if message is None:
            raise ValueError(
                "❌ Email message is not available."
            )

        try:

            with imaplib.IMAP4_SSL(
                self.imap_server,
                self.imap_port,
                timeout=30,
            ) as imap:

                # ------------------------------------------------------
                # LOGIN
                # ------------------------------------------------------

                status, response = imap.login(
                    self.account,
                    self._password,
                )

                if status != "OK":

                    raise RuntimeError(
                        "❌ IMAP login failed: "
                        f"{response}"
                    )

                # ------------------------------------------------------
                # Detectamos Drafts.
                # ------------------------------------------------------

                drafts_folder = self._get_drafts_folder(
                    imap
                )

                # ------------------------------------------------------
                # Fecha en formato IMAP.
                # ------------------------------------------------------

                date_time = imaplib.Time2Internaldate(
                    time.time()
                )

                # ------------------------------------------------------
                # CORRECCIÓN:
                # EmailMessage proporciona as_bytes(), por lo que no
                # necesitamos reconstruir manualmente el MIME.
                # ------------------------------------------------------

                status, response = imap.append(
                    f'"{drafts_folder}"', 
                    "(\\Draft)",
                    date_time,
                    message.as_bytes(),
                )

                if status != "OK":

                    raise RuntimeError(
                        "❌ IMAP APPEND failed: "
                        f"{response}"
                    )

            print(
                "📝 Draft saved successfully."
            )

            return True

        except imaplib.IMAP4.error as e:

            raise RuntimeError(
                f"❌ IMAP error while saving draft: {e}"
            ) from e

        except OSError as e:

            raise RuntimeError(
                f"❌ Network/connection error while saving draft: {e}"
            ) from e

        except Exception as e:

            raise RuntimeError(
                f"❌ Error saving draft via IMAP: {e}"
            ) from e


if __name__ == "__main__":

    from dotenv import load_dotenv
    load_dotenv()

    email_manager = EmailManager(
        account=os.getenv("Cuenta"),
        password=os.getenv("password"),
    )

    # Enviamos un correo de prueba.
    email_manager.set_recipients(
        os.getenv("Destinatarios").split(",")
    )
    email_manager.build_message(
        subject="Correo de prueba",
        body="Este es un correo de prueba."
    )
    email_manager.save_draft()
    email_manager.send(
    )