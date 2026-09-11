"""
Description:   Provides pytest fixtures for testing the EmailManager class.
Author:        David Jiménez Cooper - SpiderCoop
Date:          2026-09-11
"""

import pytest
from email_automation import EmailManager

@pytest.fixture
def email_manager():
    return EmailManager(
        account="test@example.com",
        password="secret_password"
    )

@pytest.fixture
def temp_signature(tmp_path):
    sig_file = tmp_path / "signature.html"
    sig_file.write_text("<p>Atentamente, Soporte</p>", encoding="utf-8")
    return str(sig_file)

@pytest.fixture
def temp_file(tmp_path):
    f = tmp_path / "document.txt"
    f.write_text("Contenido de prueba", encoding="utf-8")
    return str(f)

@pytest.fixture
def temp_image(tmp_path):
    img = tmp_path / "logo.png"
    img.write_bytes(b"\x89PNG\r\n\x1a\n")
    return str(img)