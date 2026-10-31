import io
import json
import os
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from unittest.mock import patch
from urllib.error import HTTPError
from urllib.parse import parse_qs

from outlook_auth import AUTHORITY, SCOPE, authorise, post_form, refresh_access_token

CLIENT_ID = "12345678-1234-1234-1234-123456789012"
DEVICE = {"device_code": "private-device-code", "user_code": "DEMO-CODE",
          "verification_uri": "https://microsoft.com/devicelogin", "interval": 1, "expires_in": 900}


class OutlookAuthTests(unittest.TestCase):
    def test_scope_requests_graph_mail_send_and_offline_access(self):
        self.assertEqual(set(SCOPE.split()), {"https://graph.microsoft.com/Mail.Send", "offline_access"})

    def test_form_request_and_http_error_response(self):
        with patch("outlook_auth.urlopen") as open_url:
            open_url.return_value.__enter__.return_value = io.StringIO('{"access_token":"mock-access"}')
            self.assertEqual(post_form("token", {"refresh_token": "mock+refresh"})["access_token"], "mock-access")
            request = open_url.call_args.args[0]
            self.assertEqual(request.full_url, f"{AUTHORITY}/token")
            self.assertEqual(parse_qs(request.data.decode()), {"refresh_token": ["mock+refresh"]})
            open_url.side_effect = HTTPError(request.full_url, 400, "Bad Request", {},
                                             io.BytesIO(json.dumps({"error": "invalid_grant"}).encode()))
            self.assertEqual(post_form("token", {})["error"], "invalid_grant")

    def test_refresh_token_request_and_failed_authorisation(self):
        with patch("outlook_auth.post_form", return_value={"access_token": "mock-access"}) as post:
            self.assertEqual(refresh_access_token(CLIENT_ID, "mock-refresh"), "mock-access")
            post.assert_called_once_with("token", {"client_id": CLIENT_ID, "refresh_token": "mock-refresh",
                                                    "grant_type": "refresh_token", "scope": SCOPE})
            post.return_value = {"error": "invalid_grant", "error_description": "private-provider-details"}
            with self.assertRaisesRegex(ValueError, "Run outlook_auth.py again") as error:
                refresh_access_token(CLIENT_ID, "mock-refresh")
            self.assertNotIn("private-provider-details", str(error.exception))

    def test_device_flow_pending_slowdown_and_private_token_file(self):
        responses = [DEVICE, {"error": "authorization_pending"}, {"error": "slow_down"},
                     {"refresh_token": "private-refresh-token", "access_token": "private-access-token"}]
        output = io.StringIO()
        with tempfile.TemporaryDirectory() as directory, patch("outlook_auth.post_form", side_effect=responses) as post, \
                patch("outlook_auth.time.sleep") as sleep, redirect_stdout(output):
            path = Path(directory) / "token"
            authorise(CLIENT_ID, path)
            self.assertEqual(path.read_text(), "private-refresh-token")
            if os.name != "nt":
                self.assertEqual(path.stat().st_mode & 0o777, 0o600)
            self.assertEqual([call.args[0] for call in sleep.call_args_list], [1, 1, 6])
            self.assertEqual(post.call_args_list[0].args, ("devicecode", {"client_id": CLIENT_ID, "scope": SCOPE}))
            for secret in ("private-refresh-token", "private-access-token", "private-device-code"):
                self.assertNotIn(secret, output.getvalue())

    def test_existing_file_is_not_overwritten(self):
        with tempfile.TemporaryDirectory() as directory, patch("outlook_auth.post_form") as post:
            path = Path(directory) / "token"
            path.write_text("existing-token")
            with self.assertRaisesRegex(ValueError, "already exists"):
                authorise(CLIENT_ID, path)
            self.assertEqual(path.read_text(), "existing-token")
            post.assert_not_called()

    def test_declined_or_expired_device_flow_does_not_write_token(self):
        with tempfile.TemporaryDirectory() as directory, patch("outlook_auth.time.sleep"), \
                patch("builtins.print"):
            path = Path(directory) / "token"
            with patch("outlook_auth.post_form", side_effect=[DEVICE, {"error": "authorization_declined"}]), \
                    self.assertRaisesRegex(ValueError, "failed or was declined"):
                authorise(CLIENT_ID, path)
            self.assertFalse(path.exists())
            with patch("outlook_auth.post_form", return_value=DEVICE), \
                    patch("outlook_auth.time.monotonic", side_effect=[0, 901]), \
                    self.assertRaisesRegex(ValueError, "timed out"):
                authorise(CLIENT_ID, path)
            self.assertFalse(path.exists())


if __name__ == "__main__":
    unittest.main()
