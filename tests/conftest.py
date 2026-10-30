"""
Description:   Provides pytest fixtures for testing the EmailManager class.
Author:        David Jiménez Cooper - SpiderCoop
Date:          2026-09-11
"""

import pytest

from email_automation import EmailManager

ACCOUNT = "test@example.com"
PASSWORD = "secret_password"


@pytest.fixture
def manager():
    """Manager con contraseña y un destinatario ya definido."""
    m = EmailManager(account=ACCOUNT, password=PASSWORD)
    m.set_recipients(["to@example.com"])
    return m


@pytest.fixture
def built_manager(manager):
    manager.build_message(subject="Asunto", body="<html><body>Hola</body></html>")
    return manager


@pytest.fixture
def signature_file(tmp_path):
    f = tmp_path / "signature.html"
    f.write_text("<p>Atentamente, Soporte</p>", encoding="utf-8")
    return f


@pytest.fixture
def attachment(tmp_path):
    f = tmp_path / "document.txt"
    f.write_text("Contenido de prueba", encoding="utf-8")
    return f


@pytest.fixture
def image(tmp_path):
    img = tmp_path / "logo.png"
    img.write_bytes(b"\x89PNG\r\n\x1a\n")
    return img
