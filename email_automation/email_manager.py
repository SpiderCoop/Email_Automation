"""
Description:   Manages sending emails with HTML content, inline images, attachments, and IMAP drafts.
Author:        David Jiménez Cooper - SpiderCoop
Date:          2024-03-22
"""

from __future__ import annotations

import base64
import imaplib
import mimetypes
import os
import re
import smtplib
import ssl
import time
import uuid
from collections.abc import Sequence
from email import policy
from email.message import EmailMessage
from pathlib import Path
from typing import Literal

PathLike = str | os.PathLike[str]


class EmailManager:
    """
    Gestor de correo para Microsoft 365 (SMTP + IMAP).

    - Destinatarios To, Cc y Bcc.
    - Cuerpo HTML con firma (archivo HTML o HTML directo).
    - Imágenes inline (CID) y adjuntos.
    - Envío por SMTP + STARTTLS y borradores por IMAP.

    Autenticación (`auth_method`):
        "password": usuario/contraseña.
        "token":    access tokens ya obtenidos (XOAUTH2), uno por protocolo.
        "oauth2":   token obtenido con MSAL (XOAUTH2). Con `client_secret`
                    usa client credentials (requiere `tenant_id`); sin él,
                    flujo delegado con Device Code.

    Los endpoints de Gmail están tipados como opciones de servidor, pero
    `oauth2` usa scopes de Microsoft; para Gmail se requieren credenciales
    o tokens obtenidos por separado.

    `token_cache_path` persiste la caché MSAL del flujo delegado en texto
    plano: protege el archivo y no lo incluyas en el control de versiones.
    """

    SMTP_SCOPE = "https://outlook.office.com/SMTP.Send"
    IMAP_SCOPE = "https://outlook.office.com/IMAP.AccessAsUser.All"
    AUTH_METHODS = ("password", "token", "oauth2")
    CLIENT_CREDENTIALS_SCOPE = "https://outlook.office365.com/.default"

    def __init__(
        self,
        account: str,
        password: str | None = None,
        smtp_server: Literal[
            "smtp.office365.com", "smtp.gmail.com"
        ] = "smtp.office365.com",
        smtp_port: int = 587,
        imap_server: Literal[
            "outlook.office365.com", "imap.gmail.com"
        ] = "outlook.office365.com",
        imap_port: int = 993,
        auth_method: Literal["password", "token", "oauth2"] = "password",
        smtp_access_token: str | None = None,
        imap_access_token: str | None = None,
        client_id: str | None = None,
        tenant_id: str = "common",
        client_secret: str | None = None,
        token_cache_path: PathLike | None = None,
    ) -> None:
        if auth_method not in self.AUTH_METHODS:
            raise ValueError(
                f"Invalid auth_method '{auth_method}'. "
                f"Use one of: {', '.join(self.AUTH_METHODS)}."
            )

        if auth_method == "password" and not password:
            raise ValueError("Password is required.")

        if auth_method == "token" and not (smtp_access_token or imap_access_token):
            raise ValueError("At least one protocol access token is required.")

        if auth_method == "oauth2":
            if not client_id:
                raise ValueError("client_id is required.")

            if client_secret and tenant_id == "common":
                raise ValueError(
                    "A specific tenant_id is required for client credentials."
                )

        self.account = account
        self.auth_method = auth_method
        self._password = password
        self._access_tokens = {
            self.SMTP_SCOPE: smtp_access_token,
            self.IMAP_SCOPE: imap_access_token,
        }

        self.client_id = client_id
        self.tenant_id = tenant_id
        self._client_secret = client_secret
        self.token_cache_path = Path(token_cache_path) if token_cache_path else None
        self._msal_app = None
        self._token_cache = None

        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.imap_server = imap_server
        self.imap_port = imap_port

        self.message: EmailMessage | None = None
        self.signature_source: PathLike | None = None

        self.direct_recipients: list[str] = []
        self.copied_recipients: list[str] = []
        self.blind_recipients: list[str] = []
        self.all_recipients: list[str] = []

    # ==================================================================
    # DESTINATARIOS Y FIRMA
    # ==================================================================

    def set_recipients(
        self,
        direct_recipients: Sequence[str] | str | None = None,
        copied_recipients: Sequence[str] | str | None = None,
        blind_recipients: Sequence[str] | str | None = None,
    ) -> None:
        """
        Define los destinatarios (To, Cc, Bcc).

        Reemplaza por completo los destinatarios anteriores, y no modifica
        el estado si no hay ninguno válido.

        Raises:
            ValueError: Si no se especifica ningún destinatario.
        """

        direct = self._clean_recipients(direct_recipients)
        copied = self._clean_recipients(copied_recipients)
        blind = self._clean_recipients(blind_recipients)

        if not (direct or copied or blind):
            raise ValueError("At least one recipient must be specified.")

        self.direct_recipients = direct
        self.copied_recipients = copied
        self.blind_recipients = blind
        self.all_recipients = direct + copied + blind

        # Mantiene coherente un mensaje ya construido con los nuevos
        # destinatarios (el borrador y el envío usan estos encabezados).
        if self.message is not None:
            self._write_recipient_headers(self.message)

    def _write_recipient_headers(self, message: EmailMessage) -> None:
        """Reescribe los encabezados To, Cc y Bcc en `message`."""

        for header in ("To", "Cc", "Bcc"):
            del message[header]

        if self.direct_recipients:
            message["To"] = ", ".join(self.direct_recipients)

        if self.copied_recipients:
            message["Cc"] = ", ".join(self.copied_recipients)

        # Bcc se omite de los encabezados enviados por SMTP, pero se
        # conserva para que el borrador IMAP lo incluya.
        if self.blind_recipients:
            message["Bcc"] = ", ".join(self.blind_recipients)

    @staticmethod
    def _as_path_list(
        values: Sequence[PathLike] | PathLike | None, name: str
    ) -> list[PathLike]:
        """Convierte a lista; una ruta suelta se envuelve como `[ruta]`."""

        if isinstance(values, (str, os.PathLike)):
            return [values]

        return list(values or [])

    @staticmethod
    def _clean_recipients(recipients: Sequence[str] | str | None) -> list[str]:
        """Elimina espacios y valores vacíos o no textuales."""

        values = [recipients] if isinstance(recipients, str) else recipients or []

        return [r.strip() for r in values if isinstance(r, str) and r.strip()]

    def set_signature(self, signature_source: PathLike | None = None) -> None:
        """
        Define la firma que se agrega al final del correo.

        Solo modifica la configuración: no lee archivos ni toca el mensaje
        ya construido. Para aplicar la nueva firma hay que volver a llamar
        a `build_message()`.

        Args:
            signature_source: Ruta (str o Path) a un archivo HTML, o un
                string con el HTML de la firma. None la desactiva.
        """

        self.signature_source = signature_source

    # ==================================================================
    # CONSTRUCCIÓN DEL MENSAJE
    # ==================================================================

    def build_message(
        self,
        subject: str,
        body: str,
        files: Sequence[PathLike] | None = None,
        inline_images: Sequence[PathLike] | None = None,
    ) -> EmailMessage:
        """
        Construye el mensaje.

        Args:
            subject: Asunto.
            body: Contenido HTML.
            files: Archivos adjuntos.
            inline_images: Imágenes referenciadas en el HTML como
                `cid:<nombre.ext>` o `cid:<nombre>`.

        Raises:
            ValueError: Si no hay destinatarios.
            FileNotFoundError: Si falta una imagen, adjunto o firma.
        """

        # Si la construcción falla no debe quedar un mensaje anterior
        # disponible para enviar o guardar como borrador.
        self.message = None

        self._validate_recipients()

        files = self._as_path_list(files, "files")
        inline_images = self._as_path_list(inline_images, "inline_images")

        body_with_images, images = self._prepare_inline_images(body, inline_images)
        full_body = self._append_signature(body_with_images, self._load_signature())

        message = EmailMessage()
        message["From"] = self.account
        self._write_recipient_headers(message)
        message["Subject"] = subject
        message.set_content(full_body, subtype="html")

        for path, content_id in images:
            self._attach_inline_image(message, path, content_id)

        for file_path in files:
            self._attach_file(message, Path(file_path))

        self.message = message

        return message

    # ==================================================================
    # IMÁGENES INLINE
    # ==================================================================

    def _prepare_inline_images(
        self,
        body: str,
        inline_images: list[PathLike],
    ) -> tuple[str, list[tuple[Path, str]]]:
        """
        Valida las imágenes y reemplaza sus referencias CID en una pasada.

        Cada imagen responde a `cid:<nombre.ext>` y `cid:<nombre>`. Los
        nombres completos tienen prioridad sobre los nombres sin extensión.
        El lookahead evita coincidencias parciales (`cid:logo` dentro de
        `cid:logo2.png`).

        Raises:
            FileNotFoundError: Si una imagen no existe.
            ValueError: Si dos imágenes comparten el mismo nombre base
                (`banner.png` y `banner.jpg`), pues `cid:banner` sería
                ambiguo.
        """

        images: list[tuple[Path, str]] = []
        seen_stems: dict[str, Path] = {}

        for image in inline_images:
            path = Path(image)

            if not path.is_file():
                raise FileNotFoundError(f"Image not found: {path}")

            previous = seen_stems.setdefault(path.stem.casefold(), path)

            if previous is not path:
                raise ValueError(
                    "Inline images must have distinct names, even with "
                    f"different extensions: '{previous.name}' and '{path.name}'."
                )

            images.append((path, uuid.uuid4().hex))

        if not images:
            return body, images

        references: dict[str, str] = {}

        for path, content_id in images:
            references[path.name] = content_id

        for path, content_id in images:
            references.setdefault(path.stem, content_id)

        # Más largos primero para que `logo.png` no se resuelva como `logo`.
        names = sorted(references, key=len, reverse=True)
        pattern = re.compile(
            r"cid:(" + "|".join(map(re.escape, names)) + r")(?![\w.\-])"
        )

        body = pattern.sub(lambda m: f"cid:{references[m.group(1)]}", body)

        return body, images

    @staticmethod
    def _attach_inline_image(
        message: EmailMessage,
        path: Path,
        content_id: str,
    ) -> None:
        """Adjunta una imagen relacionada con el HTML."""

        mime_type, _ = mimetypes.guess_type(path)

        if not mime_type or not mime_type.startswith("image/"):
            raise ValueError(f"Unsupported image MIME type: {path}")

        maintype, subtype = mime_type.split("/", 1)

        message.add_related(
            path.read_bytes(),
            maintype=maintype,
            subtype=subtype,
            cid=f"<{content_id}>",
            filename=path.name,
        )

    @staticmethod
    def _attach_file(message: EmailMessage, path: Path) -> None:
        """Adjunta un archivo; usa application/octet-stream si es desconocido."""

        if not path.is_file():
            raise FileNotFoundError(f"File not found: {path}")

        mime_type, encoding = mimetypes.guess_type(path)

        if mime_type is None or encoding is not None:
            mime_type = "application/octet-stream"

        maintype, subtype = mime_type.split("/", 1)

        message.add_attachment(
            path.read_bytes(),
            maintype=maintype,
            subtype=subtype,
            filename=path.name,
        )

    # ==================================================================
    # FIRMA
    # ==================================================================

    def _load_signature(self) -> str:
        """
        Devuelve el HTML de la firma.

        Un `Path` es siempre un archivo. Un `str` se considera HTML si
        contiene etiquetas; en caso contrario se trata como ruta.

        Raises:
            FileNotFoundError: Si la firma no es HTML ni un archivo existente.
        """

        source = self.signature_source

        if not source:
            return ""

        if isinstance(source, str) and "<" in source and ">" in source:
            return source

        path = Path(source)

        if not path.is_file():
            raise FileNotFoundError(
                f"Signature is neither HTML nor an existing file: {str(source)[:50]}"
            )

        return path.read_text(encoding="utf-8")

    @staticmethod
    def _append_signature(body: str, signature_html: str) -> str:
        """Inserta la firma antes de </body>, o al final si no existe."""

        if not signature_html:
            return body

        match = re.search(r"</body\s*>", body, flags=re.IGNORECASE)

        if match:
            return body[: match.start()] + signature_html + body[match.start() :]

        return body + signature_html

    # ==================================================================
    # VALIDACIÓN
    # ==================================================================

    def _validate_recipients(self) -> None:
        if not self.all_recipients:
            raise ValueError(
                "At least one recipient must be specified. Use set_recipients() first."
            )

    def _require_message(self) -> EmailMessage:
        if self.message is None:
            raise ValueError(
                "No email message has been built. Call build_message() first."
            )

        return self.message

    # ==================================================================
    # AUTENTICACIÓN
    # ==================================================================

    def _load_token_cache(self, msal):
        """Crea la caché MSAL y la hidrata desde `token_cache_path`."""

        self._token_cache = msal.SerializableTokenCache()

        path = self.token_cache_path

        if path and path.is_file():
            try:
                self._token_cache.deserialize(path.read_text(encoding="utf-8"))
            except (OSError, ValueError):
                # Caché ilegible o corrupta: se vuelve a autenticar.
                pass

        return self._token_cache

    def _save_token_cache(self) -> None:
        cache = self._token_cache

        if cache is not None and self.token_cache_path and cache.has_state_changed:
            self.token_cache_path.parent.mkdir(parents=True, exist_ok=True)
            self.token_cache_path.write_text(cache.serialize(), encoding="utf-8")

    def _get_access_token(self, scope: str) -> str:
        """Devuelve un access token para el scope (token directo o MSAL)."""

        if self.auth_method == "token":
            token = self._access_tokens[scope]

            if not token:
                raise ValueError(f"No access token was provided for scope {scope}.")

            return token

        try:
            import msal
        except ImportError as e:
            raise RuntimeError(
                "The 'msal' package is required for OAuth 2.0 (pip install msal)."
            ) from e

        authority = f"https://login.microsoftonline.com/{self.tenant_id}"

        if self._client_secret:
            if self._msal_app is None:
                self._msal_app = msal.ConfidentialClientApplication(
                    self.client_id,
                    authority=authority,
                    client_credential=self._client_secret,
                )

            # Client credentials solo admite ".default" del recurso de
            # Exchange Online, común a SMTP e IMAP.
            result = self._msal_app.acquire_token_for_client(
                scopes=[self.CLIENT_CREDENTIALS_SCOPE]
            )
        else:
            if self._msal_app is None:
                self._msal_app = msal.PublicClientApplication(
                    self.client_id,
                    authority=authority,
                    token_cache=self._load_token_cache(msal),
                )

            result = None
            accounts = self._msal_app.get_accounts(username=self.account)

            if accounts:
                result = self._msal_app.acquire_token_silent(
                    [scope], account=accounts[0]
                )

            if not result:
                flow = self._msal_app.initiate_device_flow(scopes=[scope])

                if "user_code" not in flow:
                    error = flow.get("error", "unknown_error")
                    description = flow.get(
                        "error_description", "Could not initiate device flow."
                    )
                    raise RuntimeError(f"Device flow failed: {error} - {description}")

                print(flow["message"])
                result = self._msal_app.acquire_token_by_device_flow(flow)

            self._save_token_cache()

        if not isinstance(result, dict):
            raise TypeError("MSAL returned an invalid token response.")

        token = result.get("access_token")

        if not token:
            error = result.get("error", "unknown_error")
            description = result.get(
                "error_description", "Access token was not returned."
            )
            raise RuntimeError(
                f"Could not obtain access token: {error} - {description}"
            )

        return token

    def _xoauth2_string(self, token: str) -> str:
        return f"user={self.account}\x01auth=Bearer {token}\x01\x01"

    def _authenticate_smtp(self, server: smtplib.SMTP) -> None:
        """
        Autentica SMTP (tras STARTTLS).

        Un fallo lanza smtplib.SMTPAuthenticationError.
        """

        if self.auth_method == "password":
            server.login(self.account, self._password)
            return

        auth_string = self._xoauth2_string(self._get_access_token(self.SMTP_SCOPE))

        # smtplib codifica en base64 y exige que el callback retorne str.
        # Llama sin argumentos para la respuesta inicial; ante un desafío
        # (error del servidor) se responde vacío para cerrar el intercambio.
        def auth_callback(challenge=None) -> str:
            if challenge is not None:
                return ""

            return auth_string

        server.auth("XOAUTH2", auth_callback, initial_response_ok=True)

    def _authenticate_imap(self, imap: imaplib.IMAP4) -> None:
        """Autentica IMAP; el callback de imaplib debe retornar bytes."""

        if self.auth_method == "password":
            status, response = imap.login(self.account, self._password)
        else:
            auth_bytes = self._xoauth2_string(
                self._get_access_token(self.IMAP_SCOPE)
            ).encode()

            # imaplib invoca el callback primero con el "+" vacío del
            # servidor; un desafío con contenido es un error y se responde
            # vacío en lugar de reenviar el token.
            def auth_callback(challenge: bytes | None) -> bytes:
                if challenge:
                    return b""

                return auth_bytes

            status, response = imap.authenticate("XOAUTH2", auth_callback)

        if status != "OK":
            raise RuntimeError(f"IMAP authentication failed: {response}")

    # ==================================================================
    # ENVÍO SMTP
    # ==================================================================

    def send(self) -> bool:
        """
        Envía el correo por SMTP + STARTTLS.

        Raises:
            RuntimeError: Si falla la autenticación, el protocolo o la red.
        """

        message = self._require_message()

        try:
            with smtplib.SMTP(self.smtp_server, self.smtp_port, timeout=30) as server:
                server.ehlo()
                server.starttls(context=ssl.create_default_context())
                server.ehlo()

                self._authenticate_smtp(server)

                # Destinatarios explícitos para incluir Bcc. Devuelve un
                # dict con los rechazados si la entrega fue parcial.
                refused = server.send_message(
                    message,
                    from_addr=self.account,
                    to_addrs=self.all_recipients,
                )

            if refused:
                rejected = ", ".join(refused)
                raise RuntimeError(
                    f"Email was partially delivered. Rejected recipients: {rejected}"
                )

        except smtplib.SMTPAuthenticationError as e:
            raise RuntimeError(
                "SMTP authentication failed. Verify the credentials and that "
                "SMTP AUTH is enabled for the mailbox (Microsoft 365 may "
                "require OAuth 2.0)."
            ) from e

        except smtplib.SMTPException as e:
            raise RuntimeError(f"SMTP error occurred: {e}") from e

        except OSError as e:
            raise RuntimeError(f"Network error while sending email: {e}") from e

        print("✅ Email sent successfully.")

        return True

    # ==================================================================
    # IMAP - BORRADORES
    # ==================================================================

    _LIST_RESPONSE = re.compile(
        r"""^\((?P<attrs>[^)]*)\)\s+                 # atributos
            (?:"(?P<delim>(?:[^"\\]|\\.)*)"|NIL)\s+  # delimitador o NIL
            (?P<name>.+)$                            # nombre del buzón
        """,
        re.VERBOSE | re.DOTALL,
    )

    @staticmethod
    def _decode_imap_utf7(name: str) -> str:
        """Decodifica un nombre de buzón en Modified UTF-7 (RFC 3501)."""

        def replace(match: re.Match[str]) -> str:
            encoded = match.group(1)

            if not encoded:
                return "&"

            encoded = encoded.replace(",", "/")
            encoded += "=" * (-len(encoded) % 4)

            return base64.b64decode(encoded).decode("utf-16-be")

        try:
            return re.sub(r"&([^-]*)-", replace, name)
        except ValueError:
            return name

    @classmethod
    def _parse_list_response(
        cls, item: bytes | tuple | None
    ) -> tuple[set[str], str] | None:
        """
        Interpreta una línea de respuesta LIST de imaplib.

        Admite delimitador entrecomillado o NIL, nombre entrecomillado con
        escapes, atom y literal (imaplib lo entrega como tupla).

        Returns:
            (atributos en minúsculas, nombre tal como viaja en el protocolo,
            es decir, en Modified UTF-7), o None si no se pudo interpretar.
        """

        literal: bytes | None = None

        if isinstance(item, tuple):
            item, literal = item[0], item[1]

        if not item:
            return None

        match = cls._LIST_RESPONSE.match(item.decode("utf-8", errors="replace"))

        if not match:
            return None

        attributes = {a.lower() for a in match.group("attrs").split()}
        name = match.group("name").strip()

        if literal is not None:
            name = literal.decode("utf-8", errors="replace")
        elif name.startswith('"') and name.endswith('"') and len(name) >= 2:
            name = re.sub(r"\\(.)", r"\1", name[1:-1])

        return (attributes, name) if name else None

    def _get_drafts_folder(self, imap: imaplib.IMAP4) -> str:
        """
        Devuelve el nombre (en formato de protocolo) de la carpeta de
        borradores.

        Usa primero el atributo especial `\\Drafts`. Solo si ningún buzón
        lo tiene, busca un buzón llamado exactamente "Drafts" o
        "Borradores" (sin distinguir mayúsculas); si tampoco existe,
        devuelve "Drafts".
        """

        status, items = imap.list()

        if status != "OK" or not items:
            return "Drafts"

        by_name: str | None = None

        for item in items:
            parsed = self._parse_list_response(item)

            if parsed is None:
                continue

            attributes, name = parsed

            if "\\noselect" in attributes:
                continue

            if "\\drafts" in attributes:
                return name

            decoded = self._decode_imap_utf7(name).casefold()

            if by_name is None and decoded in ("drafts", "borradores"):
                by_name = name

        return by_name or "Drafts"

    @staticmethod
    def _quote_imap_mailbox(mailbox: str) -> str:
        """Entrecomilla y escapa un nombre de buzón IMAP."""

        escaped = mailbox.replace("\\", "\\\\").replace('"', '\\"')

        return f'"{escaped}"'

    def save_draft(self) -> bool:
        """
        Guarda el mensaje como borrador por IMAP.

        Raises:
            RuntimeError: Si falla la autenticación, el protocolo o la red.
        """

        message = self._require_message()

        try:
            with imaplib.IMAP4_SSL(
                self.imap_server, self.imap_port, timeout=30
            ) as imap:
                self._authenticate_imap(imap)

                mailbox = self._quote_imap_mailbox(self._get_drafts_folder(imap))

                status, response = imap.append(
                    mailbox,
                    "(\\Draft)",
                    imaplib.Time2Internaldate(time.time()),
                    message.as_bytes(policy=policy.SMTP),
                )

                if status != "OK":
                    raise RuntimeError(f"IMAP APPEND failed: {response}")

        except imaplib.IMAP4.error as e:
            raise RuntimeError(f"IMAP error while saving draft: {e}") from e

        except OSError as e:
            raise RuntimeError(f"Network error while saving draft: {e}") from e

        print("📝 Draft saved successfully.")

        return True


if __name__ == "__main__":
    from dotenv import load_dotenv

    load_dotenv()

    def _require_env(name: str) -> str:
        value = os.getenv(name, "").strip()

        if not value:
            raise SystemExit(f"Missing required environment variable: {name}")

        return value

    email_manager = EmailManager(
        account=_require_env("Cuenta"),
        password=_require_env("password"),
    )

    email_manager.set_recipients(_require_env("Destinatarios").split(","))
    email_manager.build_message(
        subject="Correo de prueba",
        body="Este es un correo de prueba.",
    )
    # email_manager.save_draft()
    email_manager.send()
