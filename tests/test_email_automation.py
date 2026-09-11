"""
Description:   Unit tests for the email automation module.
Author:        David Jiménez Cooper - SpiderCoop
Date:          2026-09-11
"""

import smtplib
from unittest.mock import MagicMock, patch
import pytest
import smtplib
from email_automation.email_manager import EmailManager


def test_send_email_success(email_manager, temp_signature, temp_file, temp_image):
    email_manager.signature_file = temp_signature
    with patch("smtplib.SMTP") as mock_smtp:
        mock_server = MagicMock()
        mock_smtp.return_value.__enter__.return_value = mock_server

        result = email_manager.send(
            subject="Test Subject",
            body="<html><body>Hola cid:logo.png</body></html>",
            direct_recipients=["to@example.com"],
            copied_recipients=["cc@example.com"],
            files=[temp_file],
            inline_images=[temp_image],
        )

        assert result is True
        mock_server.starttls.assert_called_once()
        mock_server.login.assert_called_once_with("test@example.com", "secret_password")
        mock_server.sendmail.assert_called_once()


def test_missing_recipients_raises_value_error(email_manager):
    with pytest.raises(ValueError, match="At least one recipient"):
        email_manager.send(subject="Test", body="Body")


def test_file_not_found_raises(email_manager, temp_file):
    with pytest.raises(FileNotFoundError):
        email_manager.send(
            subject="Test",
            body="Body",
            direct_recipients=["to@example.com"],
            files=[temp_file, "nonexistent.txt"],
        )


def test_inline_image_not_found_raises(email_manager, temp_image):
    with pytest.raises(FileNotFoundError):
        email_manager.send(
            subject="Test",
            body="Body",
            direct_recipients=["to@example.com"],
            inline_images=[temp_image, "nonexistent.jpg"],
        )


def test_signature_as_html_string(email_manager):
    email_manager.signature_file = "<p>Firma directa HTML</p>"
    with patch("smtplib.SMTP") as mock_smtp:
        mock_server = MagicMock()
        mock_smtp.return_value.__enter__.return_value = mock_server

        result = email_manager.send(
            subject="Test",
            body="<html><body>Body</body></html>",
            direct_recipients=["to@example.com"],
        )

        assert result is True


def test_invalid_signature_string(email_manager):
    email_manager.signature_file = "texto plano sin formato ni archivo"
    with patch("smtplib.SMTP") as mock_smtp:
        mock_server = MagicMock()
        mock_smtp.return_value.__enter__.return_value = mock_server

        result = email_manager.send(
            subject="Test Invalid Signature",
            body="<html><body>Body</body></html>",
            direct_recipients=["to@example.com"],
        )

        assert result is True


def test_send_email_only_cc_and_bcc(email_manager):
    with patch("smtplib.SMTP") as mock_smtp:
        mock_server = MagicMock()
        mock_smtp.return_value.__enter__.return_value = mock_server

        result = email_manager.send(
            subject="Test CC/BCC",
            body="<html><body>Test Body</body></html>",
            copied_recipients=["cc@example.com"],
            blind_recipients=["bcc@example.com"],
        )

        assert result is True
        mock_server.sendmail.assert_called_once()


def test_multiple_files_attachment(email_manager, temp_file, tmp_path):
    file2 = tmp_path / "document2.pdf"
    file2.write_bytes(b"%PDF-1.4")

    with patch("smtplib.SMTP") as mock_smtp:
        mock_server = MagicMock()
        mock_smtp.return_value.__enter__.return_value = mock_server

        result = email_manager.send(
            subject="Test Multiple Files",
            body="<html><body>Body</body></html>",
            direct_recipients=["to@example.com"],
            files=[temp_file, str(file2)],
        )

        assert result is True


def test_multiple_inline_images(email_manager, temp_image, tmp_path):
    img2 = tmp_path / "logo2.png"
    img2.write_bytes(b"\x89PNG\r\n\x1a\n")

    with patch("smtplib.SMTP") as mock_smtp:
        mock_server = MagicMock()
        mock_smtp.return_value.__enter__.return_value = mock_server

        result = email_manager.send(
            subject="Test Multiple Inline Images",
            body='<html><body><img src="cid:logo.png"><img src="cid:logo2.png"></body></html>',
            direct_recipients=["to@example.com"],
            inline_images=[temp_image, str(img2)],
        )

        assert result is True


def test_custom_smtp_server_and_port():
    manager = EmailManager(
        account="custom@example.com",
        password="password",
        smtp_server="smtp.customserver.com",
        smtp_port=465,
    )
    assert manager.smtp_server == "smtp.customserver.com"
    assert manager.smtp_port == 465


def test_smtp_error_raises_runtime_error(email_manager):
    with patch("smtplib.SMTP") as mock_smtp:
        mock_server = MagicMock()
        mock_server.sendmail.side_effect = smtplib.SMTPException("Connection failed")
        mock_smtp.return_value.__enter__.return_value = mock_server

        with pytest.raises(RuntimeError, match="SMTP error occurred"):
            email_manager.send(
                subject="Test", body="Body", direct_recipients=["to@example.com"]
            )


def test_unexpected_exception_raises_runtime_error(email_manager):
    with patch("smtplib.SMTP") as mock_smtp:
        mock_smtp.side_effect = Exception("Unexpected network drop")

        with pytest.raises(RuntimeError, match="An unexpected error occurred"):
            email_manager.send(
                subject="Test", body="Body", direct_recipients=["to@example.com"]
            )
