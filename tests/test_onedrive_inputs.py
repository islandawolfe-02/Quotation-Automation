import io
import json
import os
import unittest
from unittest.mock import patch
from urllib.error import HTTPError

from quotation import download_onedrive_workbook, main


class OneDriveInputTests(unittest.TestCase):
    def metadata(self, url="https://files.1drv.com/private-download"):
        return io.BytesIO(json.dumps({"@microsoft.graph.downloadUrl": url}).encode())

    def test_encoded_path_and_download_without_bearer_token(self):
        with patch("quotation.urlopen", side_effect=[self.metadata(), io.BytesIO(b"workbook")]) as open_url:
            with download_onedrive_workbook("Quotations/Client #1 & café.xlsx", "private-access") as workbook:
                self.assertEqual(workbook.read(), b"workbook")
            request = open_url.call_args_list[0].args[0]
            self.assertEqual(request.full_url, "https://graph.microsoft.com/v1.0/me/drive/root:/"
                             "Quotations/Client%20%231%20%26%20caf%C3%A9.xlsx?$select=id,@microsoft.graph.downloadUrl")
            self.assertEqual(request.get_header("Authorization"), "Bearer private-access")
            self.assertEqual(open_url.call_args_list[1].args, ("https://files.1drv.com/private-download",))
            self.assertEqual(open_url.call_args_list[1].kwargs, {"timeout": 30})

    def test_invalid_paths_never_reach_network(self):
        for path in ("", "/client.xlsx", "../client.xlsx", "Quotes/../client.xlsx", "Quotes//client.xlsx",
                     "https://example.com/client.xlsx", "Quotes\\client.xlsx", "Quotes/client.json", "Quotes/\nclient.xlsx"):
            with self.subTest(path=path), patch("quotation.urlopen") as open_url:
                with self.assertRaises(ValueError):
                    download_onedrive_workbook(path, "private-access")
                open_url.assert_not_called()

    def test_download_errors_hide_provider_details_and_do_not_retry(self):
        for status in (401, 403, 404, 429, 500):
            for download_started in (False, True):
                error = HTTPError("https://files.1drv.com/private-url", status, "PRIVATE_DETAILS", {},
                                  io.BytesIO(b"private@example.com PRIVATE_TOKEN"))
                responses = [self.metadata(), error] if download_started else [error]
                with self.subTest(status=status, download=download_started), \
                        patch("quotation.urlopen", side_effect=responses) as open_url:
                    with self.assertRaisesRegex(ValueError, f"HTTP {status}") as raised:
                        download_onedrive_workbook("Quotations/client.xlsx", "private-access")
                    self.assertEqual(open_url.call_count, len(responses))
                    self.assertNotIn("PRIVATE", str(raised.exception))
                    self.assertNotIn("private", str(raised.exception))

    def test_missing_or_unsafe_download_urls_are_rejected(self):
        for metadata in ({}, [], {"@microsoft.graph.downloadUrl": 123},
                         {"@microsoft.graph.downloadUrl": "http://files.1drv.com/file"},
                         {"@microsoft.graph.downloadUrl": "https://user:password@files.1drv.com/file"}):
            with self.subTest(metadata=metadata), \
                    patch("quotation.urlopen", return_value=io.BytesIO(json.dumps(metadata).encode())) as open_url:
                with self.assertRaisesRegex(ValueError, "secure workbook download URL"):
                    download_onedrive_workbook("Quotations/client.xlsx", "private-access")
                open_url.assert_called_once()

    def test_empty_or_oversized_workbooks_are_rejected(self):
        for content in (b"", b"12345"):
            with self.subTest(content=content), patch("quotation.MAX_WORKBOOK_BYTES", 4), \
                    patch("quotation.urlopen", side_effect=[self.metadata(), io.BytesIO(content)]):
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
