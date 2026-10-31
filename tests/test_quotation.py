import base64
import copy
import io
import json
import os
import smtplib
import tempfile
import unittest
from decimal import Decimal
from email import message_from_bytes
from email.policy import default
from pathlib import Path
from unittest.mock import patch
from urllib.error import HTTPError

from quotation import ROOT, calculate, load_json, main, render_pdf, send_email


class QuotationTests(unittest.TestCase):
    def setUp(self):
        self.request = load_json(ROOT / "examples/requirements.json")
        self.catalogue = load_json(ROOT / "data/catalogue.json")
        self.charges = load_json(ROOT / "data/charges.json")

    def calculate(self):
        return calculate(self.request, self.catalogue, self.charges)

    def test_sample_totals(self):
        result = self.calculate()
        expected = {"Light Fixtures": "17300.00", "Wires and Consumables": "12300.00",
                    "Labour Charges": "11000.00", "Other Charges": "1800.00"}
        for category, amount in expected.items():
            self.assertEqual(sum(line["amount"] for line in result["lines"] if line["category"] == category), Decimal(amount))
        self.assertEqual(result["subtotal"], Decimal("42400.00"))
        self.assertEqual(result["tax"], Decimal("0.00"))
        self.assertEqual(result["total"], Decimal("42400.00"))

    def test_decimal_rounding_and_line_tax(self):
        self.request["materials"] = [{"code": "LED-LINEAR-25", "quantity": "0.5"}]
        self.request["charges"] = []
        self.catalogue[0].update(rate="0.09", tax_percent="10")
        result = self.calculate()
        self.assertEqual(result["subtotal"], Decimal("0.05"))
        self.assertEqual(result["tax"], Decimal("0.01"))
        self.assertEqual(result["total"], Decimal("0.06"))

    def test_invalid_quantities(self):
        for value in ("0", "-1", "NaN", "Infinity", "abc", True, None, "1000000001"):
            with self.subTest(value=value), self.assertRaises(ValueError):
                self.request["materials"][0]["quantity"] = value
                self.calculate()

    def test_unknown_or_missing_inputs_fail(self):
        for field, value in (("materials", []), ("charges", None), ("customer", ""),
                             ("site", None), ("validity_days", True), ("quotation_id", "../escape")):
            with self.subTest(field=field), self.assertRaises(ValueError):
                request = {**self.request, field: value}
                calculate(request, self.catalogue, self.charges)
        self.request["materials"][0]["code"] = "UNKNOWN"
        with self.assertRaisesRegex(ValueError, "Unknown materials code"):
            self.calculate()

    def test_invalid_catalogue_fails(self):
        for field, value in (("rate", "-1"), ("rate", "1.234"), ("rate", "NaN"),
                             ("tax_percent", "101"), ("category", "Unknown")):
            catalogue = copy.deepcopy(self.catalogue)
            catalogue[0][field] = value
            with self.subTest(field=field, value=value), self.assertRaises(ValueError):
                calculate(self.request, catalogue, self.charges)
        with self.assertRaisesRegex(ValueError, "Duplicate catalogue code"):
            calculate(self.request, self.catalogue + [self.catalogue[0]], self.charges)

    def test_pdf_generation_escapes_customer_markup(self):
        self.request["customer"] = "Demo <Centre> & Partners"
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "quote.pdf"
            render_pdf(path, self.request, self.calculate(), load_json(ROOT / "templates/quotation.json"))
            self.assertTrue(path.read_bytes().startswith(b"%PDF-"))
            self.assertGreater(path.stat().st_size, 1000)

    def test_pipeline_json_overrides_sample_and_sends_only_when_requested(self):
        request = {**self.request, "quotation_id": "POC-PIPELINE"}
        with patch.dict(os.environ, {"QUOTATION_REQUIREMENTS_JSON": json.dumps(request)}), \
                patch("quotation.render_pdf") as render, patch("quotation.send_email") as send, \
                patch("builtins.print"):
            with patch("sys.argv", ["quotation.py"]):
                main()
            self.assertEqual(render.call_args.args[1]["quotation_id"], "POC-PIPELINE")
            send.assert_not_called()
            with patch("sys.argv", ["quotation.py", "--send-email"]):
                main()
            send.assert_called_once_with(ROOT / "output/pdf/POC-PIPELINE.pdf", "POC-PIPELINE")

    def test_email_tls_attachment_and_recipient(self):
        settings = {"SMTP_HOST": "smtp.example.invalid", "SMTP_PORT": "587",
                    "SMTP_USERNAME": "smtp-user", "SMTP_PASSWORD": "mock-password",
                    "SMTP_FROM": "sender@example.invalid", "QUOTE_EMAIL_TO": "owner@another.invalid"}
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "POC-001.pdf"
            path.write_bytes(b"%PDF-mock")
            for port in ("587", "465"):
                with self.subTest(port=port), patch.dict(os.environ, {**settings, "SMTP_PORT": port}, clear=True), \
                        patch("quotation.smtplib.SMTP") as plain, patch("quotation.smtplib.SMTP_SSL") as secure:
                    client = secure if port == "465" else plain
                    smtp = client.return_value.__enter__.return_value
                    smtp.send_message.return_value = {}
                    send_email(path, "POC-001")
                    smtp.login.assert_called_once_with("smtp-user", "mock-password")
                    if port == "587":
                        smtp.starttls.assert_called_once()
                    else:
                        smtp.starttls.assert_not_called()
                    message = smtp.send_message.call_args.args[0]
                    self.assertEqual(message["To"], "owner@another.invalid")
                    self.assertEqual(message["From"], "sender@example.invalid")
                    attachment = next(message.iter_attachments())
                    self.assertEqual(attachment.get_content_type(), "application/pdf")
                    self.assertEqual(attachment.get_filename(), path.name)
                    self.assertEqual(attachment.get_payload(decode=True), b"%PDF-mock")

    def test_missing_credentials_and_recipient_refusal_fail(self):
        with patch.dict(os.environ, {}, clear=True), self.assertRaisesRegex(ValueError, "Missing email settings"):
            send_email(Path("unused.pdf"), "POC-001")
        settings = {"SMTP_HOST": "smtp.example.invalid", "SMTP_USERNAME": "sender@example.invalid",
                    "SMTP_PASSWORD": "mock-password", "SMTP_FROM": "sender@example.invalid",
                    "QUOTE_EMAIL_TO": "owner@another.invalid"}
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "quote.pdf"
            path.write_bytes(b"%PDF-mock")
            with patch.dict(os.environ, settings, clear=True), patch("quotation.smtplib.SMTP") as client:
                client.return_value.__enter__.return_value.send_message.return_value = {"owner@another.invalid": (550, b"Rejected")}
                with self.assertRaises(smtplib.SMTPException):
                    send_email(path, "POC-001")

    def test_multiple_recipients_in_graph_and_smtp(self):
        settings = {"SMTP_FROM": "sender@example.invalid", "QUOTE_EMAIL_TO": " owner@another.invalid, accounts@example.invalid ",
                    "SMTP_HOST": "smtp.example.invalid", "SMTP_USERNAME": "smtp-user", "SMTP_PASSWORD": "mock-password"}
        expected = ["owner@another.invalid", "accounts@example.invalid"]
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "quote.pdf"
            path.write_bytes(b"%PDF-mock")
            with patch.dict(os.environ, settings, clear=True), patch("quotation.smtplib.SMTP") as client:
                smtp = client.return_value.__enter__.return_value
                smtp.send_message.return_value = {}
                send_email(path, "POC-001")
                message = smtp.send_message.call_args.args[0]
                self.assertEqual([address.addr_spec for address in message["To"].addresses], expected)
                smtp.send_message.return_value = {expected[1]: (550, b"Rejected")}
                with self.assertRaises(smtplib.SMTPException):
                    send_email(path, "POC-001")
            oauth_settings = {**settings, "MICROSOFT_CLIENT_ID": "mock-client", "MICROSOFT_REFRESH_TOKEN": "mock-refresh"}
            with patch.dict(os.environ, oauth_settings, clear=True), patch("quotation.urlopen") as open_url, \
                    patch("quotation.refresh_access_token", return_value="mock-access"):
                open_url.return_value.__enter__.return_value.status = 202
                send_email(path, "POC-001")
                message = message_from_bytes(base64.b64decode(open_url.call_args.args[0].data), policy=default)
                self.assertEqual([address.addr_spec for address in message["To"].addresses], expected)
                self.assertEqual(next(message.iter_attachments()).get_payload(decode=True), b"%PDF-mock")

    def test_invalid_recipient_lists_fail_before_authentication_or_sending(self):
        settings = {"SMTP_FROM": "sender@example.invalid", "MICROSOFT_CLIENT_ID": "mock-client",
                    "MICROSOFT_REFRESH_TOKEN": "mock-refresh"}
        invalid = [" ", "not-an-email", "owner@", "@example.invalid", "a@example.invalid,",
                   "a@example.invalid,,b@example.invalid", "a@example.invalid;b@example.invalid",
                   "a@example.invalid, private-invalid-entry", "a@example.invalid\r\nBcc: b@example.invalid"]
        for recipients in invalid:
            with self.subTest(recipients=recipients), patch.dict(os.environ, {**settings, "QUOTE_EMAIL_TO": recipients}, clear=True), \
                    patch("quotation.refresh_access_token") as refresh, patch("quotation.urlopen") as open_url, \
                    patch("quotation.smtplib.SMTP") as smtp:
                with self.assertRaisesRegex(ValueError, "valid email addresses separated by commas") as error:
                    send_email(Path("unused.pdf"), "POC-001")
                self.assertNotIn("private-invalid-entry", str(error.exception))
                refresh.assert_not_called()
                open_url.assert_not_called()
                smtp.assert_not_called()

    def test_outlook_uses_graph_and_attaches_pdf_without_smtp(self):
        settings = {"SMTP_PORT": "unused-invalid-port", "SMTP_FROM": " sender@outlook.com\n",
                    "QUOTE_EMAIL_TO": " owner@another.invalid\n",
                    "MICROSOFT_CLIENT_ID": "mock-client", "MICROSOFT_REFRESH_TOKEN": "mock-refresh"}
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "quote.pdf"
            path.write_bytes(b"%PDF-mock")
            with patch.dict(os.environ, settings, clear=True), patch("quotation.smtplib.SMTP") as client, \
                    patch("quotation.smtplib.SMTP_SSL") as secure, patch("quotation.urlopen") as open_url, \
                    patch("quotation.refresh_access_token", return_value="mock-access") as refresh:
                open_url.return_value.__enter__.return_value.status = 202
                send_email(path, "POC-001")
                refresh.assert_called_once_with("mock-client", "mock-refresh")
                client.assert_not_called()
                secure.assert_not_called()
                open_url.assert_called_once()
                request = open_url.call_args.args[0]
                self.assertEqual(request.full_url, "https://graph.microsoft.com/v1.0/me/sendMail")
                self.assertEqual(request.get_method(), "POST")
                self.assertEqual(request.get_header("Authorization"), "Bearer mock-access")
                self.assertEqual(request.get_header("Content-type"), "text/plain")
                message = message_from_bytes(base64.b64decode(request.data, validate=True), policy=default)
                self.assertEqual(message["Subject"], "Mock electrical quotation POC-001")
                self.assertEqual(message["From"], "sender@outlook.com")
                self.assertEqual(message["To"], "owner@another.invalid")
                attachment = next(message.iter_attachments())
                self.assertEqual(attachment.get_filename(), "quote.pdf")
                self.assertEqual(attachment.get_content_type(), "application/pdf")
                self.assertEqual(attachment.get_payload(decode=True), b"%PDF-mock")

    def test_graph_failures_are_private_and_do_not_retry_or_fall_back(self):
        settings = {"SMTP_FROM": "sender@outlook.com", "QUOTE_EMAIL_TO": "owner@another.invalid",
                    "MICROSOFT_CLIENT_ID": "mock-client", "MICROSOFT_REFRESH_TOKEN": "mock-refresh"}
        for status in (401, 403, 429, 500):
            with self.subTest(status=status), tempfile.TemporaryDirectory() as directory:
                path = Path(directory) / "quote.pdf"
                path.write_bytes(b"%PDF-mock")
                error = HTTPError("https://graph.microsoft.com/v1.0/me/sendMail", status,
                                  "PRIVATE_PROVIDER_DETAILS", {}, io.BytesIO(b"private@example.invalid PRIVATE_ACCESS_TOKEN"))
                with patch.dict(os.environ, settings, clear=True), patch("quotation.urlopen", side_effect=error) as open_url, \
                        patch("quotation.refresh_access_token", return_value="mock-access"), \
                        patch("quotation.smtplib.SMTP") as client:
                    with self.assertRaisesRegex(ValueError, f"HTTP {status}") as raised:
                        send_email(path, "POC-001")
                    open_url.assert_called_once()
                    client.assert_not_called()
                    self.assertNotIn("PRIVATE", str(raised.exception))
                    self.assertNotIn("private@example.invalid", str(raised.exception))

    def test_graph_requires_confirmed_submission_and_a_valid_token(self):
        settings = {"SMTP_FROM": "sender@outlook.com", "QUOTE_EMAIL_TO": "owner@another.invalid",
                    "MICROSOFT_CLIENT_ID": "mock-client", "MICROSOFT_REFRESH_TOKEN": "mock-refresh"}
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "quote.pdf"
            path.write_bytes(b"%PDF-mock")
            with patch.dict(os.environ, settings, clear=True), patch("quotation.urlopen") as open_url, \
                    patch("quotation.refresh_access_token", return_value="mock-access") as refresh:
                open_url.return_value.__enter__.return_value.status = 200
                with self.assertRaisesRegex(ValueError, "did not confirm email submission"):
                    send_email(path, "POC-001")
                open_url.reset_mock()
                refresh.side_effect = ValueError("Microsoft authorisation failed")
                with self.assertRaisesRegex(ValueError, "authorisation failed"):
                    send_email(path, "POC-001")
                open_url.assert_not_called()

    def test_authentication_failure_reports_only_safe_codes_and_guidance(self):
        cases = [
            (b"5.7.139 Authentication unsuccessful, SmtpClientAuthentication is disabled for the Mailbox", "disabled for the mailbox or tenant", "5.7.139"),
            (b"5.7.139 Authentication unsuccessful, basic authentication is disabled", "Password-based SMTP is disabled", "5.7.139"),
            (b"5.7.139 user is locked by security policy", "account or sign-in restriction", "5.7.139"),
            (b"5.7.3 Authentication unsuccessful", "SMTP_USERNAME matches the Outlook account", "5.7.3"),
            ("Authentication unsuccessful", "SMTP_USERNAME matches the Outlook account", None),
        ]
        for response, guidance, status in cases:
            private = " private@example.invalid PRIVATE_ACCESS_TOKEN"
            response = response + (private.encode() if isinstance(response, bytes) else private)
            stderr = io.StringIO()
            with self.subTest(response=response), patch("sys.argv", ["quotation.py", "--send-email"]), \
                    patch.dict(os.environ, {"QUOTATION_REQUIREMENTS_JSON": ""}), \
                    patch("quotation.render_pdf"), patch("builtins.print"), patch("sys.stderr", stderr), \
                    patch("quotation.send_email", side_effect=smtplib.SMTPAuthenticationError(535, response)):
                with self.assertRaises(SystemExit) as error:
                    main()
                self.assertEqual(error.exception.code, 1)
                self.assertIn("SMTP 535", stderr.getvalue())
                if status:
                    self.assertIn(status, stderr.getvalue())
                self.assertIn(guidance, stderr.getvalue())
                self.assertNotIn("private@example.invalid", stderr.getvalue())
                self.assertNotIn("PRIVATE_ACCESS_TOKEN", stderr.getvalue())

    def test_incomplete_outlook_credentials_do_not_fall_back_to_password(self):
        with patch.dict(os.environ, {"MICROSOFT_CLIENT_ID": "mock-client", "SMTP_PASSWORD": "unused"}, clear=True), \
                patch("quotation.smtplib.SMTP") as client:
            with self.assertRaisesRegex(ValueError, "MICROSOFT_REFRESH_TOKEN"):
                send_email(Path("unused.pdf"), "POC-001")
            client.assert_not_called()


if __name__ == "__main__":
    unittest.main()
