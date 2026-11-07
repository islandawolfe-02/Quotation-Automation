import io
import os
import unittest
from unittest.mock import patch
from urllib.error import HTTPError
from urllib.request import HTTPSHandler, build_opener
from urllib.response import addinfourl

from quotation import NoRedirect, download_onedrive_workbook, main


class OneDriveInputTests(unittest.TestCase):
    def redirect(self, url="https://files.1drv.com/private-download"):
        return HTTPError("https://graph.microsoft.com/content", 302, "Found",
                         {"Location": url}, io.BytesIO(b""))

    def test_encoded_path_and_download_without_bearer_token(self):
        with patch("quotation.build_opener") as build, \
                patch("quotation.urlopen", return_value=io.BytesIO(b"workbook")) as open_url:
            build.return_value.open.side_effect = self.redirect()
            with download_onedrive_workbook("Quotations/Client #1 & café.xlsx", "private-access") as workbook:
                self.assertEqual(workbook.read(), b"workbook")
            request = build.return_value.open.call_args.args[0]
            self.assertEqual(request.full_url, "https://graph.microsoft.com/v1.0/me/drive/root:/"
                             "Quotations/Client%20%231%20%26%20caf%C3%A9.xlsx:/content")
            self.assertEqual(request.get_header("Authorization"), "Bearer private-access")
            self.assertIsInstance(build.call_args.args[0], NoRedirect)
            open_url.assert_called_once_with("https://files.1drv.com/private-download", timeout=30)

    def test_content_can_be_returned_directly_without_metadata(self):
        with patch("quotation.build_opener") as build, patch("quotation.urlopen") as open_url:
            build.return_value.open.return_value = io.BytesIO(b"workbook")
            with download_onedrive_workbook("Quotations/client.xlsx", "access") as workbook:
                self.assertEqual(workbook.read(), b"workbook")
            open_url.assert_not_called()

    def test_standard_library_redirect_handler_does_not_forward_authorisation(self):
        requests = []

        class GraphResponse(HTTPSHandler):
            def https_open(self, request):
                requests.append(request)
                response = addinfourl(io.BytesIO(), {"Location": "https://files.1drv.com/download"},
                                      request.full_url, 302)
                response.msg = "Found"
                return response

        opener = build_opener(NoRedirect(), GraphResponse())
        with patch("quotation.build_opener", return_value=opener), \
                patch("quotation.urlopen", return_value=io.BytesIO(b"workbook")) as open_url:
            with download_onedrive_workbook("Quotations/client.xlsx", "private-access") as workbook:
                self.assertEqual(workbook.read(), b"workbook")
            self.assertEqual(len(requests), 1)
            self.assertEqual(requests[0].get_header("Authorization"), "Bearer private-access")
            open_url.assert_called_once_with("https://files.1drv.com/download", timeout=30)

    def test_invalid_paths_never_reach_network(self):
        for path in ("", "/client.xlsx", "../client.xlsx", "Quotes/../client.xlsx", "Quotes//client.xlsx",
                     "https://example.com/client.xlsx", "Quotes\\client.xlsx", "Quotes/client.json", "Quotes/\nclient.xlsx"):
            with self.subTest(path=path), patch("quotation.urlopen") as open_url, patch("quotation.build_opener") as build:
                with self.assertRaises(ValueError):
                    download_onedrive_workbook(path, "private-access")
                open_url.assert_not_called()
                build.assert_not_called()

    def test_download_errors_hide_provider_details_and_do_not_retry(self):
        for status in (401, 403, 404, 429, 500):
            for download_started in (False, True):
                error = HTTPError("https://files.1drv.com/private-url", status, "PRIVATE_DETAILS", {},
                                  io.BytesIO(b"private@example.com PRIVATE_TOKEN"))
                with self.subTest(status=status, download=download_started), \
                        patch("quotation.build_opener") as build, patch("quotation.urlopen", side_effect=error) as open_url:
                    build.return_value.open.side_effect = self.redirect() if download_started else error
                    with self.assertRaisesRegex(ValueError, f"HTTP {status}") as raised:
                        download_onedrive_workbook("Quotations/client.xlsx", "private-access")
                    build.return_value.open.assert_called_once()
                    self.assertEqual(open_url.call_count, int(download_started))
                    self.assertNotIn("PRIVATE", str(raised.exception))
                    self.assertNotIn("private", str(raised.exception))

    def test_missing_or_unsafe_redirect_urls_are_rejected(self):
        for url in ("", "http://files.1drv.com/file", "file:///tmp/client.xlsx",
                    "https://user:password@files.1drv.com/file"):
            with self.subTest(url=url), patch("quotation.build_opener") as build, patch("quotation.urlopen") as open_url:
                build.return_value.open.side_effect = self.redirect(url)
                with self.assertRaisesRegex(ValueError, "secure workbook download URL"):
                    download_onedrive_workbook("Quotations/client.xlsx", "private-access")
                open_url.assert_not_called()
        with patch("quotation.build_opener") as build, patch("quotation.urlopen") as open_url:
            build.return_value.open.side_effect = HTTPError("https://graph.microsoft.com/content", 302, "Found", {}, io.BytesIO())
            with self.assertRaisesRegex(ValueError, "secure workbook download URL"):
                download_onedrive_workbook("Quotations/client.xlsx", "private-access")
            open_url.assert_not_called()

    def test_empty_or_oversized_workbooks_are_rejected(self):
        for content in (b"", b"12345"):
            for redirected in (False, True):
                with self.subTest(content=content, redirected=redirected), patch("quotation.MAX_WORKBOOK_BYTES", 4), \
                        patch("quotation.build_opener") as build, patch("quotation.urlopen", return_value=io.BytesIO(content)):
                    if redirected:
                        build.return_value.open.side_effect = self.redirect()
                    else:
                        build.return_value.open.return_value = io.BytesIO(content)
                    with self.assertRaisesRegex(ValueError, "non-empty and at most"):
                        download_onedrive_workbook("Quotations/client.xlsx", "private-access")

    def test_cli_downloads_latest_workbooks_and_closes_them_on_failure(self):
        for fail_client in (False, True):
            master, client = io.BytesIO(b"master"), io.BytesIO(b"client")
            downloads = [master, ValueError("Download failed")] if fail_client else [master, client]
            with self.subTest(fail_client=fail_client), \
                    patch.dict(os.environ, {"MICROSOFT_CLIENT_ID": "client", "MICROSOFT_REFRESH_TOKEN": "refresh"}, clear=True), \
                    patch("sys.argv", ["quotation.py", "--onedrive", "--send-email"]), \
                    patch("quotation.recipient_addresses"), \
                    patch("quotation.refresh_access_token", return_value="access") as refresh, \
                    patch("quotation.download_onedrive_workbook", side_effect=downloads) as download, \
                    patch("quotation.load_excel_inputs", side_effect=ValueError("Stop after parsing")) as load, \
                    patch("quotation.render_pdf") as render, patch("quotation.send_email") as send, \
                    patch("sys.stderr", io.StringIO()):
                with self.assertRaises(SystemExit):
                    main()
                refresh.assert_called_once_with("client", "refresh", read_files=True)
                self.assertEqual([call.args for call in download.call_args_list],
                                 [("Quotations/master_catalogue.xlsx", "access"), ("Quotations/client_requirements.xlsx", "access")])
                if fail_client:
                    load.assert_not_called()
                else:
                    load.assert_called_once_with(master, client)
                    self.assertTrue(client.closed)
                self.assertTrue(master.closed)
                render.assert_not_called()
                send.assert_not_called()

    def test_zero_recipients_fail_before_authentication_or_download(self):
        flags = {f"SEND_PARTNER_{partner}": "false" for partner in (1, 2, 3)}
        with patch.dict(os.environ, flags, clear=True), \
                patch("sys.argv", ["quotation.py", "--onedrive", "--send-email"]), \
                patch("quotation.refresh_access_token") as refresh, patch("quotation.urlopen") as open_url, \
                patch("quotation.render_pdf") as render, patch("quotation.send_email") as send, \
                patch("sys.stderr", io.StringIO()):
            with self.assertRaises(SystemExit):
                main()
            refresh.assert_not_called()
            open_url.assert_not_called()
            render.assert_not_called()
            send.assert_not_called()


if __name__ == "__main__":
    unittest.main()
