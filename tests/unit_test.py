import unittest
from unittest.mock import patch, mock_open, MagicMock
import os
import tempfile
import smtplib
import json
from dotenv import load_dotenv
from email_automation import send_email

class TestSendEmail(unittest.TestCase):

    def setUp(self):
        # Cargar variables desde el .env para emular el flujo de secrets en producción
        dotted = os.path.join(os.path.dirname(__file__), '..', '.env')
        load_dotenv(dotenv_path=dotted)

        self.account = os.getenv('Cuenta', 'test@example.com')
        self.password = os.getenv('password', 'password')

        # Si en .env existe Destinatarios, lo parseamos de JSON (o fallback sigue estático)
        raw_dest = os.getenv('Destinatarios')
        
        if raw_dest:
            try:
                parsed = json.loads(raw_dest)
                # Asegurar listas
                recipients = {
                    'to': parsed.get('to', '').split(',') if isinstance(parsed.get('to', ''), str) else parsed.get('to', []),
                    'cc': parsed.get('cc', '').split(',') if isinstance(parsed.get('cc', ''), str) else parsed.get('cc', []),
                    'bcc': parsed.get('bcc', '').split(',') if isinstance(parsed.get('bcc', ''), str) else parsed.get('bcc', []),
                }
            except Exception:
                recipients = {'to': ['to@example.com'], 'cc': ['cc@example.com'], 'bcc': ['bcc@example.com']}
        else:
            recipients = {'to': ['to@example.com'], 'cc': ['cc@example.com'], 'bcc': ['bcc@example.com']}

        self.direct_recipients = recipients['to']
        self.copied_recipients = recipients['cc']
        self.blind_recipients = recipients['bcc']

        self.subject = 'Test Subject'
        self.body = '<html><body>Test Body</body></html>'

    def test_invalid_direct_recipients_not_list(self):
        result = send_email(self.account, self.password, self.subject, self.body, 'invalid', self.copied_recipients, self.blind_recipients)
        self.assertFalse(result)

    @patch('email_automation.send_email.smtplib.SMTP')
    @patch('email_automation.send_email.os.path.isfile')
    def test_inline_image_not_found(self, mock_isfile, mock_smtp):
        mock_isfile.return_value = False
        mock_server = MagicMock()
        mock_smtp.return_value.__enter__.return_value = mock_server
        result = send_email(self.account, self.password, self.subject, self.body, self.direct_recipients, self.copied_recipients, self.blind_recipients, inline_images=['nonexistent.jpg'])
        self.assertFalse(result)

    @patch('email_automation.send_email.smtplib.SMTP')
    @patch('email_automation.send_email.os.path.isfile')
    def test_attachment_not_found(self, mock_isfile, mock_smtp):
        mock_isfile.return_value = False
        mock_server = MagicMock()
        mock_smtp.return_value.__enter__.return_value = mock_server
        result = send_email(self.account, self.password, self.subject, self.body, self.direct_recipients, self.copied_recipients, self.blind_recipients, files=['nonexistent.txt'])
        self.assertFalse(result)

    @patch('email_automation.send_email.smtplib.SMTP')
    @patch('email_automation.send_email.os.path.isfile')
    def test_signature_file_not_found(self, mock_isfile, mock_smtp):
        mock_isfile.return_value = False
        mock_server = MagicMock()
        mock_smtp.return_value.__enter__.return_value = mock_server
        result = send_email(self.account, self.password, self.subject, self.body, self.direct_recipients, self.copied_recipients, self.blind_recipients, signature_file='nonexistent.html')
        self.assertTrue(result)  # Should still succeed, just no signature

    @patch('email_automation.send_email.os.path.isfile')
    @patch('email_automation.send_email.smtplib.SMTP')
    def test_successful_send(self, mock_smtp, mock_isfile):
        mock_isfile.return_value = True
        mock_server = MagicMock()
        mock_smtp.return_value.__enter__.return_value = mock_server

        with patch('builtins.open', mock_open(read_data='signature')):
            result = send_email(self.account, self.password, self.subject, self.body, self.direct_recipients, self.copied_recipients, self.blind_recipients, signature_file='sig.html')
        self.assertTrue(result)
        mock_server.starttls.assert_called_once()
        mock_server.login.assert_called_once_with(self.account, self.password)
        mock_server.sendmail.assert_called_once()

    @patch('email_automation.send_email.smtplib.SMTP')
    def test_smtp_error(self, mock_smtp):
        mock_server = MagicMock()
        mock_server.sendmail.side_effect = smtplib.SMTPException('SMTP Error')
        mock_smtp.return_value.__enter__.return_value = mock_server

        result = send_email(self.account, self.password, self.subject, self.body, self.direct_recipients, self.copied_recipients, self.blind_recipients)
        self.assertFalse(result)

    @patch('email_automation.send_email.smtplib.SMTP')
    def test_recipients_as_lists(self, mock_smtp):
        direct_recipients = ['to@example.com']
        copied_recipients = ['cc@example.com']
        blind_recipients = ['bcc@example.com']
        mock_server = MagicMock()
        mock_smtp.return_value.__enter__.return_value = mock_server
        result = send_email(self.account, self.password, self.subject, self.body, direct_recipients, copied_recipients, blind_recipients)
        self.assertTrue(result)

if __name__ == '__main__':
    unittest.main()