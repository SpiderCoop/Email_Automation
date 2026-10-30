"""
Description:   Unit tests for the EmailManager class.
Author:        David Jiménez Cooper - SpiderCoop
Date:          2026-09-11
"""

import imaplib
import re
import smtplib
import sys
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

import pytest

from email_automation import EmailManager
from email_automation import email_manager as em

ACCOUNT = "test@example.com"
PASSWORD = "secret_password"

# ======================================================================
# Constructor
# ======================================================================


class TestConstructor:
    def test_defaults(self):
        m = EmailManager(ACCOUNT, PASSWORD)
        assert m.smtp_server == "smtp.office365.com"
        assert m.smtp_port == 587
        assert m.imap_server == "outlook.office365.com"
        assert m.imap_port == 993

    def test_invalid_auth_method(self):
        with pytest.raises(ValueError, match="Invalid auth_method"):
            EmailManager(ACCOUNT, PASSWORD, auth_method="kerberos")

    def test_password_required(self):
        with pytest.raises(ValueError, match="Password is required"):
            EmailManager(ACCOUNT)

    def test_token_requires_at_least_one(self):
        with pytest.raises(ValueError, match="access token"):
            EmailManager(ACCOUNT, auth_method="token")

    def test_token_with_single_protocol_is_valid(self):
        EmailManager(ACCOUNT, auth_method="token", smtp_access_token="t")

    def test_oauth2_requires_client_id(self):
        with pytest.raises(ValueError, match="client_id"):
            EmailManager(ACCOUNT, auth_method="oauth2")

    def test_client_credentials_require_specific_tenant(self):
        with pytest.raises(ValueError, match="tenant_id"):
            EmailManager(
                ACCOUNT, auth_method="oauth2", client_id="c", client_secret="s"
            )

    def test_client_credentials_with_tenant(self):
        EmailManager(
            ACCOUNT,
            auth_method="oauth2",
            client_id="c",
            client_secret="s",
            tenant_id="tenant",
        )

    def test_token_cache_path_is_path(self, tmp_path):
        m = EmailManager(
            ACCOUNT,
            auth_method="oauth2",
            client_id="c",
            token_cache_path=str(tmp_path / "cache.json"),
        )
        assert isinstance(m.token_cache_path, Path)


# ======================================================================
# Destinatarios
# ======================================================================


class TestRecipients:
    def test_cleans_values(self):
        m = EmailManager(ACCOUNT, PASSWORD)
        m.set_recipients([" a@x.com ", "", "  "], ["b@x.com"], None)
        assert m.direct_recipients == ["a@x.com"]
        assert m.copied_recipients == ["b@x.com"]
        assert m.blind_recipients == []
        assert m.all_recipients == ["a@x.com", "b@x.com"]

    def test_only_cc_and_bcc(self):
        m = EmailManager(ACCOUNT, PASSWORD)
        m.set_recipients(copied_recipients=["c@x.com"], blind_recipients=["b@x.com"])
        assert m.all_recipients == ["c@x.com", "b@x.com"]

    def test_replaces_instead_of_accumulating(self, manager):
        manager.set_recipients(["new@x.com"])
        assert manager.all_recipients == ["new@x.com"]
        assert manager.copied_recipients == []

    def test_single_string_is_one_recipient(self):
        m = EmailManager(ACCOUNT, PASSWORD)
        m.set_recipients("user@x.com", "cc@x.com", "bcc@x.com")
        assert m.direct_recipients == ["user@x.com"]
        assert m.copied_recipients == ["cc@x.com"]
        assert m.blind_recipients == ["bcc@x.com"]

    def test_non_string_items_are_ignored(self):
        assert EmailManager._clean_recipients(["a@x.com", None, 5]) == ["a@x.com"]

    def test_clean_recipients_none(self):
        assert EmailManager._clean_recipients(None) == []

    def test_headers_of_built_message_are_synced(self, built_manager):
        msg = built_manager.message
        built_manager.set_recipients(["new@x.com"], blind_recipients=["bcc@x.com"])

        assert msg["To"] == "new@x.com"
        assert msg["Cc"] is None
        assert msg["Bcc"] == "bcc@x.com"
        assert len(msg.get_all("To")) == 1

    def test_removed_headers_are_deleted_from_built_message(self, manager):
        manager.set_recipients(["to@x.com"], ["cc@x.com"], ["bcc@x.com"])
        msg = manager.build_message("s", "b")
        manager.set_recipients(["to@x.com"])

        assert msg["Cc"] is None
        assert msg["Bcc"] is None

    def test_failed_update_leaves_built_message_untouched(self, built_manager):
        msg = built_manager.message
        with pytest.raises(ValueError):
            built_manager.set_recipients([])
        assert msg["To"] == "to@example.com"

    def test_empty_raises_and_keeps_state(self, manager):
        with pytest.raises(ValueError, match="At least one recipient"):
            manager.set_recipients([" "])
        assert manager.all_recipients == ["to@example.com"]

    def test_build_message_requires_recipients(self):
        m = EmailManager(ACCOUNT, PASSWORD)
        with pytest.raises(ValueError, match="At least one recipient"):
            m.build_message("s", "b")


# ======================================================================
# Construcción del mensaje
# ======================================================================


class TestBuildMessage:
    def test_headers(self, manager):
        manager.set_recipients(["to@x.com"], ["cc@x.com"], ["bcc@x.com"])
        msg = manager.build_message("Asunto", "<p>Hola</p>")
        assert msg["From"] == ACCOUNT
        assert msg["To"] == "to@x.com"
        assert msg["Cc"] == "cc@x.com"
        assert msg["Bcc"] == "bcc@x.com"
        assert msg["Subject"] == "Asunto"
        assert manager.message is msg

    def test_omits_empty_headers(self, built_manager):
        assert built_manager.message["Cc"] is None
        assert built_manager.message["Bcc"] is None

    def test_attachment(self, manager, attachment):
        msg = manager.build_message("s", "<p>b</p>", files=[attachment])
        names = [p.get_filename() for p in msg.iter_attachments()]
        assert names == ["document.txt"]

    def test_attachment_unknown_mime_uses_octet_stream(self, manager, tmp_path):
        f = tmp_path / "data.unknownext"
        f.write_bytes(b"123")
        msg = manager.build_message("s", "b", files=[f])
        assert next(msg.iter_attachments()).get_content_type() == (
            "application/octet-stream"
        )

    def test_attachment_compressed_uses_octet_stream(self, manager, tmp_path):
        f = tmp_path / "data.tar.gz"
        f.write_bytes(b"123")
        msg = manager.build_message("s", "b", files=[f])
        assert next(msg.iter_attachments()).get_content_type() == (
            "application/octet-stream"
        )

    @pytest.mark.parametrize("as_str", [True, False])
    def test_single_attachment_path_is_wrapped(self, manager, attachment, as_str):
        single = str(attachment) if as_str else attachment
        msg = manager.build_message("s", "b", files=single)
        assert len(list(msg.iter_attachments())) == 1

    def test_attachment_not_found(self, manager):
        with pytest.raises(FileNotFoundError, match="File not found"):
            manager.build_message("s", "b", files=["nonexistent.txt"])


# ======================================================================
# Imágenes inline
# ======================================================================


class TestInlineImages:
    def test_cid_replaced_and_attached(self, manager, image):
        msg = manager.build_message(
            "s", '<body><img src="cid:logo.png"></body>', inline_images=[image]
        )
        html = msg.get_body(("html",)).get_content()
        cid = re.search(r"cid:(\w+)", html).group(1)
        assert "logo" not in html
        related = [p for p in msg.walk() if p.get("Content-ID")]
        assert related[0]["Content-ID"] == f"<{cid}>"

    @pytest.mark.parametrize("as_str", [True, False])
    def test_single_image_path_is_wrapped(self, manager, image, as_str):
        single = str(image) if as_str else image
        msg = manager.build_message("s", "<img src='cid:logo.png'>", inline_images=single)
        assert [p for p in msg.walk() if p.get("Content-ID")]

    def test_base_name_reference(self, manager, image):
        msg = manager.build_message("s", "<img src='cid:logo'>", inline_images=[image])
        assert "cid:logo" not in msg.get_body(("html",)).get_content()

    def test_same_base_name_different_extension_raises(self, manager, tmp_path):
        png = tmp_path / "banner.png"
        jpg = tmp_path / "banner.jpg"
        png.write_bytes(b"\x89PNG")
        jpg.write_bytes(b"\xff\xd8")

        with pytest.raises(ValueError, match="distinct names.*banner.png.*banner.jpg"):
            manager.build_message("s", "cid:banner.png", inline_images=[png, jpg])

    def test_same_base_name_is_case_insensitive(self, manager, tmp_path):
        a = tmp_path / "Logo.png"
        b = tmp_path / "logo.jpg"
        a.write_bytes(b"\x89PNG")
        b.write_bytes(b"\xff\xd8")

        with pytest.raises(ValueError, match="distinct names"):
            manager._prepare_inline_images("", [a, b])

    def test_same_name_in_different_folders_raises(self, manager, tmp_path):
        (tmp_path / "a").mkdir()
        (tmp_path / "b").mkdir()
        one = tmp_path / "a" / "logo.png"
        two = tmp_path / "b" / "logo.png"
        one.write_bytes(b"\x89PNG")
        two.write_bytes(b"\x89PNG")

        with pytest.raises(ValueError, match="distinct names"):
            manager._prepare_inline_images("", [one, two])

    def test_distinct_names_with_shared_prefix_are_allowed(self, manager, tmp_path):
        a = tmp_path / "logo.png"
        b = tmp_path / "logo2.jpg"
        a.write_bytes(b"\x89PNG")
        b.write_bytes(b"\xff\xd8")

        body, images = manager._prepare_inline_images(
            "cid:logo.png cid:logo2.jpg cid:logo cid:logo2", [a, b]
        )
        a_id, b_id = (cid for _, cid in images)

        assert body == f"cid:{a_id} cid:{b_id} cid:{a_id} cid:{b_id}"

    def test_similar_names_do_not_corrupt_each_other(self, manager, tmp_path):
        png = tmp_path / "logo.png"
        jpg = tmp_path / "logoo.jpg"
        png.write_bytes(b"\x89PNG")
        jpg.write_bytes(b"\xff\xd8")

        body, images = manager._prepare_inline_images(
            "cid:logo.png cid:logoo.jpg cid:logo", [png, jpg]
        )
        png_id, jpg_id = (cid for _, cid in images)

        assert body == f"cid:{png_id} cid:{jpg_id} cid:{png_id}"

    def test_prefix_name_not_matched_inside_longer_name(self, manager, tmp_path):
        short = tmp_path / "logo.png"
        short.write_bytes(b"\x89PNG")

        body, _ = manager._prepare_inline_images("cid:logo2.png", [short])

        assert body == "cid:logo2.png"

    def test_no_images_returns_body_unchanged(self, manager):
        assert manager._prepare_inline_images("cid:x", []) == ("cid:x", [])

    def test_image_not_found(self, manager):
        with pytest.raises(FileNotFoundError, match="Image not found"):
            manager.build_message("s", "b", inline_images=["missing.png"])

    def test_unsupported_mime(self, manager, tmp_path):
        f = tmp_path / "file.txt"
        f.write_text("x")
        with pytest.raises(ValueError, match="Unsupported image"):
            manager.build_message("s", "b", inline_images=[f])


# ======================================================================
# Firma
# ======================================================================


class TestSignature:
    def test_none_means_no_signature(self, manager):
        assert manager._load_signature() == ""

    @pytest.mark.parametrize("as_str", [True, False])
    def test_from_file(self, manager, signature_file, as_str):
        manager.set_signature(str(signature_file) if as_str else signature_file)
        assert "Atentamente" in manager._load_signature()

    def test_from_html_string(self, manager):
        manager.set_signature("<p>Firma</p>")
        assert manager._load_signature() == "<p>Firma</p>"

    def test_invalid_string_raises(self, manager):
        manager.set_signature("texto plano")
        with pytest.raises(FileNotFoundError, match="Signature"):
            manager._load_signature()

    def test_missing_path_object_raises(self, manager, tmp_path):
        manager.set_signature(tmp_path / "nope.html")
        with pytest.raises(FileNotFoundError):
            manager._load_signature()

    def test_inserted_before_closing_body_case_insensitive(self, manager):
        manager.set_signature("<p>sig</p>")
        msg = manager.build_message("s", "<HTML><BODY>x</BODY></HTML>")
        assert "x<p>sig</p></BODY>" in msg.get_content()

    def test_appended_when_no_body_tag(self, manager):
        manager.set_signature("<p>sig</p>")
        msg = manager.build_message("s", "<p>x</p>")
        assert msg.get_content().strip().endswith("<p>x</p><p>sig</p>")

    def test_append_without_signature(self):
        assert EmailManager._append_signature("body", "") == "body"

    def test_set_signature_does_not_touch_existing_message(self, manager):
        manager.set_signature("<p>old</p>")
        msg = manager.build_message("s", "<body>x</body>")

        manager.set_signature("<p>new</p>")

        assert manager.message is msg
        assert "<p>old</p>" in msg.get_content()

    def test_new_signature_applies_on_next_build(self, manager):
        manager.set_signature("<p>old</p>")
        manager.build_message("s", "<body>x</body>")

        manager.set_signature("<p>new</p>")
        html = manager.build_message("s", "<body>x</body>").get_content()

        assert "<p>new</p>" in html
        assert "<p>old</p>" not in html

    def test_set_signature_does_not_read_files(self, manager, tmp_path):
        manager.set_signature(tmp_path / "nope.html")
        assert manager.message is None

    def test_invalid_signature_fails_on_build(self, manager):
        manager.set_signature("not html nor file")
        with pytest.raises(FileNotFoundError, match="Signature"):
            manager.build_message("s", "<body>x</body>")
        assert manager.message is None

    def test_failed_build_discards_previous_message(self, built_manager):
        with pytest.raises(FileNotFoundError):
            built_manager.build_message("s", "b", files=["missing.txt"])

        assert built_manager.message is None


# ======================================================================
# Tokens y MSAL
# ======================================================================


@pytest.fixture
def fake_msal():
    """Sustituye el módulo msal por un doble controlable."""
    app = MagicMock()
    cache = MagicMock()
    cache.has_state_changed = True
    cache.serialize.return_value = "SERIALIZED"

    module = SimpleNamespace(
        PublicClientApplication=MagicMock(return_value=app),
        ConfidentialClientApplication=MagicMock(return_value=app),
        SerializableTokenCache=MagicMock(return_value=cache),
    )
    with patch.dict(sys.modules, {"msal": module}):
        yield SimpleNamespace(module=module, app=app, cache=cache)


def delegated(tmp_path=None, **kwargs):
    return EmailManager(
        ACCOUNT,
        auth_method="oauth2",
        client_id="client",
        token_cache_path=(tmp_path / "cache.json") if tmp_path else None,
        **kwargs,
    )


class TestAccessToken:
    def test_direct_tokens_per_protocol(self):
        m = EmailManager(
            ACCOUNT,
            auth_method="token",
            smtp_access_token="smtp-t",
            imap_access_token="imap-t",
        )
        assert m._get_access_token(m.SMTP_SCOPE) == "smtp-t"
        assert m._get_access_token(m.IMAP_SCOPE) == "imap-t"

    def test_missing_protocol_token(self):
        m = EmailManager(ACCOUNT, auth_method="token", smtp_access_token="t")
        with pytest.raises(ValueError, match="No access token"):
            m._get_access_token(m.IMAP_SCOPE)

    def test_msal_not_installed(self):
        with (
            patch.dict(sys.modules, {"msal": None}),
            pytest.raises(RuntimeError, match="msal"),
        ):
            delegated()._get_access_token(EmailManager.SMTP_SCOPE)

    def test_client_credentials_uses_exchange_default_scope(self, fake_msal):
        fake_msal.app.acquire_token_for_client.return_value = {"access_token": "tok"}
        m = delegated(client_secret="secret", tenant_id="tenant")

        assert m._get_access_token(m.SMTP_SCOPE) == "tok"
        assert m._get_access_token(m.IMAP_SCOPE) == "tok"
        for call in fake_msal.app.acquire_token_for_client.call_args_list:
            assert call.kwargs == {"scopes": ["https://outlook.office365.com/.default"]}

    def test_delegated_silent_token(self, fake_msal):
        fake_msal.app.get_accounts.return_value = [{"a": 1}]
        fake_msal.app.acquire_token_silent.return_value = {"access_token": "tok"}
        m = delegated()

        assert m._get_access_token(m.IMAP_SCOPE) == "tok"
        fake_msal.app.initiate_device_flow.assert_not_called()

    def test_delegated_device_flow(self, fake_msal, capsys):
        fake_msal.app.get_accounts.return_value = []
        fake_msal.app.initiate_device_flow.return_value = {
            "user_code": "X",
            "message": "Go to the page",
        }
        fake_msal.app.acquire_token_by_device_flow.return_value = {
            "access_token": "tok"
        }

        assert delegated()._get_access_token(EmailManager.SMTP_SCOPE) == "tok"
        assert "Go to the page" in capsys.readouterr().out

    def test_device_flow_failure(self, fake_msal):
        fake_msal.app.get_accounts.return_value = []
        fake_msal.app.initiate_device_flow.return_value = {
            "error": "invalid_client",
            "error_description": "Bad client",
            "device_code": "SENSITIVE",
        }

        with pytest.raises(RuntimeError, match="invalid_client - Bad client") as exc:
            delegated()._get_access_token(EmailManager.SMTP_SCOPE)
        assert "SENSITIVE" not in str(exc.value)

    def test_device_flow_failure_defaults(self, fake_msal):
        fake_msal.app.get_accounts.return_value = []
        fake_msal.app.initiate_device_flow.return_value = {}

        with pytest.raises(
            RuntimeError, match="unknown_error - Could not initiate device flow"
        ):
            delegated()._get_access_token(EmailManager.SMTP_SCOPE)

    def test_no_token_in_result(self, fake_msal):
        fake_msal.app.get_accounts.return_value = []
        fake_msal.app.initiate_device_flow.return_value = {
            "user_code": "X",
            "message": "",
        }
        fake_msal.app.acquire_token_by_device_flow.return_value = {
            "error": "access_denied",
            "error_description": "denied",
        }

        with pytest.raises(RuntimeError, match="access_denied - denied"):
            delegated()._get_access_token(EmailManager.SMTP_SCOPE)

    def test_token_missing_without_details(self, fake_msal):
        fake_msal.app.acquire_token_for_client.return_value = {}
        m = delegated(client_secret="s", tenant_id="t")

        with pytest.raises(
            RuntimeError, match="unknown_error - Access token was not returned"
        ):
            m._get_access_token(m.SMTP_SCOPE)

    def test_invalid_msal_response(self, fake_msal):
        fake_msal.app.acquire_token_for_client.return_value = None
        m = delegated(client_secret="s", tenant_id="t")

        with pytest.raises(TypeError, match="invalid token response"):
            m._get_access_token(m.SMTP_SCOPE)

    def test_app_is_reused(self, fake_msal):
        fake_msal.app.get_accounts.return_value = [{"a": 1}]
        fake_msal.app.acquire_token_silent.return_value = {"access_token": "tok"}
        m = delegated()
        m._get_access_token(m.SMTP_SCOPE)
        m._get_access_token(m.IMAP_SCOPE)

        fake_msal.module.PublicClientApplication.assert_called_once()


class TestTokenCache:
    def test_saved_when_changed(self, fake_msal, tmp_path):
        fake_msal.app.get_accounts.return_value = [{"a": 1}]
        fake_msal.app.acquire_token_silent.return_value = {"access_token": "tok"}
        m = delegated(tmp_path)

        m._get_access_token(m.SMTP_SCOPE)

        assert (tmp_path / "cache.json").read_text() == "SERIALIZED"

    def test_not_saved_when_unchanged(self, fake_msal, tmp_path):
        fake_msal.cache.has_state_changed = False
        fake_msal.app.get_accounts.return_value = [{"a": 1}]
        fake_msal.app.acquire_token_silent.return_value = {"access_token": "tok"}
        m = delegated(tmp_path)

        m._get_access_token(m.SMTP_SCOPE)

        assert not (tmp_path / "cache.json").exists()

    def test_not_saved_without_path(self, fake_msal):
        m = delegated()
        m._load_token_cache(fake_msal.module)
        m._save_token_cache()
        fake_msal.cache.serialize.assert_not_called()

    def test_loaded_from_disk(self, fake_msal, tmp_path):
        (tmp_path / "cache.json").write_text("STORED")
        m = delegated(tmp_path)

        m._load_token_cache(fake_msal.module)

        fake_msal.cache.deserialize.assert_called_once_with("STORED")

    def test_corrupt_cache_is_ignored(self, fake_msal, tmp_path):
        (tmp_path / "cache.json").write_text("garbage")
        fake_msal.cache.deserialize.side_effect = ValueError
        m = delegated(tmp_path)

        assert m._load_token_cache(fake_msal.module) is fake_msal.cache

    def test_creates_parent_directories(self, fake_msal, tmp_path):
        m = EmailManager(
            ACCOUNT,
            auth_method="oauth2",
            client_id="c",
            token_cache_path=tmp_path / "sub" / "dir" / "cache.json",
        )
        m._load_token_cache(fake_msal.module)
        m._save_token_cache()
        assert (tmp_path / "sub" / "dir" / "cache.json").is_file()


# ======================================================================
# Autenticación SMTP / IMAP
# ======================================================================


def token_manager(**kwargs):
    return EmailManager(
        ACCOUNT,
        auth_method="token",
        smtp_access_token="smtp-t",
        imap_access_token="imap-t",
        **kwargs,
    )


class TestAuthentication:
    def test_xoauth2_string(self):
        assert token_manager()._xoauth2_string("T") == (
            f"user={ACCOUNT}\x01auth=Bearer T\x01\x01"
        )

    def test_smtp_password(self, manager):
        server = MagicMock(spec=smtplib.SMTP)
        manager._authenticate_smtp(server)
        server.login.assert_called_once_with(ACCOUNT, PASSWORD)

    def test_smtp_xoauth2_callback_returns_str(self):
        server = MagicMock(spec=smtplib.SMTP)
        token_manager()._authenticate_smtp(server)

        mechanism, callback = server.auth.call_args.args
        assert mechanism == "XOAUTH2"
        assert callback() == f"user={ACCOUNT}\x01auth=Bearer smtp-t\x01\x01"
        assert callback(b'{"status":"401"}') == ""
        assert server.auth.call_args.kwargs == {"initial_response_ok": True}

    def test_imap_password(self, manager):
        imap = MagicMock(spec=imaplib.IMAP4)
        imap.login.return_value = ("OK", [b"done"])
        manager._authenticate_imap(imap)
        imap.login.assert_called_once_with(ACCOUNT, PASSWORD)

    def test_imap_xoauth2_callback_returns_bytes(self):
        imap = MagicMock(spec=imaplib.IMAP4)
        imap.authenticate.return_value = ("OK", [b"done"])
        token_manager()._authenticate_imap(imap)

        mechanism, callback = imap.authenticate.call_args.args
        assert mechanism == "XOAUTH2"
        assert callback(b"") == f"user={ACCOUNT}\x01auth=Bearer imap-t\x01\x01".encode()
        assert callback(None) == callback(b"")
        assert callback(b'{"status":"401"}') == b""

    @pytest.mark.parametrize("method", ["login", "authenticate"])
    def test_imap_non_ok_status_raises(self, manager, method):
        imap = MagicMock(spec=imaplib.IMAP4)
        getattr(imap, method).return_value = ("NO", [b"denied"])
        target = manager if method == "login" else token_manager()

        with pytest.raises(RuntimeError, match="IMAP authentication failed"):
            target._authenticate_imap(imap)


# ======================================================================
# Envío SMTP
# ======================================================================


class TestSend:
    def test_requires_message(self, manager):
        with pytest.raises(ValueError, match="build_message"):
            manager.send()

    def test_success_flow(self, built_manager):
        server = MagicMock(spec=smtplib.SMTP)
        server.send_message.return_value = {}
        with patch.object(em.smtplib, "SMTP") as cls:
            cls.return_value.__enter__.return_value = server
            with patch.object(em.EmailManager, "_authenticate_smtp") as auth:
                assert built_manager.send() is True

        cls.assert_called_once_with("smtp.office365.com", 587, timeout=30)
        server.starttls.assert_called_once()
        auth.assert_called_once_with(server)
        message = server.send_message.call_args.args[0]
        assert message is built_manager.message
        assert server.send_message.call_args.kwargs == {
            "from_addr": ACCOUNT,
            "to_addrs": ["to@example.com"],
        }

    def test_bcc_included_in_envelope(self, manager):
        manager.set_recipients(["to@x.com"], blind_recipients=["bcc@x.com"])
        manager.build_message("s", "b")
        server = MagicMock(spec=smtplib.SMTP)
        server.send_message.return_value = {}
        with patch.object(em.smtplib, "SMTP") as cls:
            cls.return_value.__enter__.return_value = server
            with patch.object(em.EmailManager, "_authenticate_smtp"):
                manager.send()

        assert server.send_message.call_args.kwargs["to_addrs"] == [
            "to@x.com",
            "bcc@x.com",
        ]

    def test_partial_delivery_raises(self, manager, capsys):
        manager.set_recipients(["a@x.com", "b@x.com", "c@x.com"])
        manager.build_message("s", "b")
        server = MagicMock(spec=smtplib.SMTP)
        server.send_message.return_value = {
            "c@x.com": (550, b"No such user"),
            "b@x.com": (552, b"Mailbox full"),
        }
        with patch.object(em.smtplib, "SMTP") as cls:
            cls.return_value.__enter__.return_value = server
            with (
                patch.object(em.EmailManager, "_authenticate_smtp"),
                pytest.raises(
                    RuntimeError,
                    match="partially delivered. Rejected recipients: c@x.com, b@x.com",
                ),
            ):
                manager.send()

        assert "sent successfully" not in capsys.readouterr().out

    def test_all_recipients_refused_is_wrapped(self, built_manager):
        server = MagicMock(spec=smtplib.SMTP)
        server.send_message.side_effect = smtplib.SMTPRecipientsRefused(
            {"to@example.com": (550, b"no")}
        )
        with patch.object(em.smtplib, "SMTP") as cls:
            cls.return_value.__enter__.return_value = server
            with (
                patch.object(em.EmailManager, "_authenticate_smtp"),
                pytest.raises(RuntimeError, match="SMTP error occurred"),
            ):
                built_manager.send()

    @pytest.mark.parametrize(
        ("error", "match"),
        [
            (smtplib.SMTPAuthenticationError(535, b"bad"), "authentication failed"),
            (smtplib.SMTPException("boom"), "SMTP error occurred"),
            (ConnectionError("down"), "Network error"),
        ],
    )
    def test_errors_are_wrapped(self, built_manager, error, match):
        with (
            patch.object(em.smtplib, "SMTP", side_effect=error),
            pytest.raises(RuntimeError, match=match),
        ):
            built_manager.send()

    def test_authentication_error_is_wrapped(self, built_manager):
        server = MagicMock(spec=smtplib.SMTP)
        error = smtplib.SMTPAuthenticationError(535, b"bad")
        with patch.object(em.smtplib, "SMTP") as cls:
            cls.return_value.__enter__.return_value = server
            with (
                patch.object(em.EmailManager, "_authenticate_smtp", side_effect=error),
                pytest.raises(RuntimeError, match="authentication failed"),
            ):
                built_manager.send()


# ======================================================================
# IMAP / borradores
# ======================================================================


class TestDraftsFolder:
    @pytest.mark.parametrize(
        ("listing", "expected"),
        [
            ([b'(\\HasNoChildren \\Drafts) "/" "Drafts"'], "Drafts"),
            ([b'(\\HasNoChildren \\DRAFTS) "/" "Mis Borradores"'], "Mis Borradores"),
            ([b"(\\Drafts) NIL Drafts"], "Drafts"),
            ([b'(\\Drafts) "." INBOX.Drafts'], "INBOX.Drafts"),
            ([b'(\\Drafts) "/" "Dr\\"afts"'], 'Dr"afts'),
            ([(b'(\\Drafts) "/" {10}', b"My Drafts!")], "My Drafts!"),
            ([b'(\\HasNoChildren) "/" "Borradores"'], "Borradores"),
            ([b'(\\HasNoChildren) "/" "drafts"'], "drafts"),
            ([b'(\\HasNoChildren) "/" "Bor&APE-radores"'], "Drafts"),
            ([None, b'(\\HasNoChildren) "/" "INBOX"'], "Drafts"),
            ([b"unparseable drafts"], "Drafts"),
        ],
    )
    def test_detection(self, manager, listing, expected):
        imap = MagicMock()
        imap.list.return_value = ("OK", listing)
        assert manager._get_drafts_folder(imap) == expected

    def test_special_use_attribute_beats_name(self, manager):
        imap = MagicMock()
        imap.list.return_value = (
            "OK",
            [
                b'(\\HasNoChildren) "/" "Drafts"',
                b'(\\HasNoChildren \\Drafts) "/" "Entwuerfe"',
            ],
        )
        assert manager._get_drafts_folder(imap) == "Entwuerfe"

    @pytest.mark.parametrize(
        "name", ["Archive/DraftsBackup", "Drafts Old", "MyDrafts", "Archive/Drafts"]
    )
    def test_partial_name_matches_are_ignored(self, manager, name):
        imap = MagicMock()
        imap.list.return_value = ("OK", [f'(\\HasNoChildren) "/" "{name}"'.encode()])
        assert manager._get_drafts_folder(imap) == "Drafts"

    def test_noselect_mailboxes_are_skipped(self, manager):
        imap = MagicMock()
        imap.list.return_value = (
            "OK",
            [
                b'(\\Noselect \\Drafts) "/" "Container"',
                b'(\\HasNoChildren \\Drafts) "/" "Real"',
            ],
        )
        assert manager._get_drafts_folder(imap) == "Real"

    def test_encoded_name_is_returned_in_wire_format(self, manager):
        imap = MagicMock()
        imap.list.return_value = ("OK", [b'(\\Drafts) "/" "Entw&APw-rfe"'])
        assert manager._get_drafts_folder(imap) == "Entw&APw-rfe"

    @pytest.mark.parametrize("response", [("NO", []), ("OK", [])])
    def test_default_when_list_fails(self, manager, response):
        imap = MagicMock()
        imap.list.return_value = response
        assert manager._get_drafts_folder(imap) == "Drafts"


class TestImapNameHelpers:
    @pytest.mark.parametrize(
        ("encoded", "decoded"),
        [
            ("Drafts", "Drafts"),
            ("Entw&APw-rfe", "Entwürfe"),
            ("Bor&APE-radores", "Borñradores"),
            ("A &- B", "A & B"),
            ("&ZeVnLIqe-", "日本語"),
            ("&AOk-/&AOk-", "é/é"),
            ("&bad!-", "&bad!-"),
        ],
    )
    def test_decode_modified_utf7(self, encoded, decoded):
        assert EmailManager._decode_imap_utf7(encoded) == decoded

    def test_parse_list_response_details(self):
        attrs, name = EmailManager._parse_list_response(
            b'(\\HasNoChildren \\Drafts) "/" "Drafts"'
        )
        assert attrs == {"\\hasnochildren", "\\drafts"}
        assert name == "Drafts"

    def test_parse_list_response_invalid(self):
        assert EmailManager._parse_list_response(b"garbage") is None
        assert EmailManager._parse_list_response(None) is None
        assert EmailManager._parse_list_response(b"") is None

    @pytest.mark.parametrize(
        ("mailbox", "quoted"),
        [
            ("Drafts", '"Drafts"'),
            ("My Drafts", '"My Drafts"'),
            ('Say "hi"', '"Say \\"hi\\""'),
            ("back\\slash", '"back\\\\slash"'),
            ('a\\"b', '"a\\\\\\"b"'),
        ],
    )
    def test_quote_imap_mailbox(self, mailbox, quoted):
        assert EmailManager._quote_imap_mailbox(mailbox) == quoted


@pytest.fixture
def imap():
    server = MagicMock(spec=imaplib.IMAP4_SSL)
    server.list.return_value = ("OK", [b'(\\Drafts) "/" "Drafts"'])
    server.append.return_value = ("OK", [b"done"])
    with patch.object(em.imaplib, "IMAP4_SSL") as cls:
        cls.return_value.__enter__.return_value = server
        yield SimpleNamespace(cls=cls, server=server)


class TestSaveDraft:
    def test_requires_message(self, manager):
        with pytest.raises(ValueError, match="build_message"):
            manager.save_draft()

    def test_success(self, built_manager, imap):
        with patch.object(em.EmailManager, "_authenticate_imap") as auth:
            assert built_manager.save_draft() is True

        imap.cls.assert_called_once_with("outlook.office365.com", 993, timeout=30)
        auth.assert_called_once_with(imap.server)
        folder, flags, _, payload = imap.server.append.call_args.args
        assert folder == '"Drafts"'
        assert flags == "(\\Draft)"
        assert b"Subject: Asunto" in payload
        assert b"\r\n" in payload
        assert b"\n" not in payload.replace(b"\r\n", b"")

    def test_mailbox_name_is_quoted_and_escaped(self, built_manager, imap):
        imap.server.list.return_value = ("OK", [b'(\\Drafts) "/" "Dr\\"af\\\\ts"'])
        with patch.object(em.EmailManager, "_authenticate_imap"):
            built_manager.save_draft()

        assert imap.server.append.call_args.args[0] == '"Dr\\"af\\\\ts"'

    def test_append_failure(self, built_manager, imap):
        imap.server.append.return_value = ("NO", [b"quota"])
        with (
            patch.object(em.EmailManager, "_authenticate_imap"),
            pytest.raises(RuntimeError, match="APPEND failed"),
        ):
            built_manager.save_draft()

    def test_imap_error_is_wrapped(self, built_manager, imap):
        with (
            patch.object(
                em.EmailManager,
                "_authenticate_imap",
                side_effect=imaplib.IMAP4.error("x"),
            ),
            pytest.raises(RuntimeError, match="IMAP error"),
        ):
            built_manager.save_draft()

    def test_network_error_is_wrapped(self, built_manager):
        with (
            patch.object(em.imaplib, "IMAP4_SSL", side_effect=OSError("down")),
            pytest.raises(RuntimeError, match="Network error"),
        ):
            built_manager.save_draft()
