"""One-time Outlook.com authorisation and access-token renewal using OAuth."""

import argparse
import json
import os
import time
from pathlib import Path
from urllib.error import HTTPError
from urllib.parse import urlencode
from urllib.request import Request, urlopen
from uuid import UUID

AUTHORITY = "https://login.microsoftonline.com/consumers/oauth2/v2.0"
SCOPE = "https://graph.microsoft.com/Mail.Send offline_access"
TOKEN_FILE = Path(__file__).resolve().parent / ".outlook-refresh-token"


def post_form(endpoint, fields):
    request = Request(f"{AUTHORITY}/{endpoint}", data=urlencode(fields).encode(),
                      headers={"Content-Type": "application/x-www-form-urlencoded"})
    try:
        with urlopen(request, timeout=30) as response:
            return json.load(response)
    except HTTPError as exc:
        # Only callers' fixed error messages reach logs, never tokens or provider responses.
        with exc:
            return json.load(exc)


def refresh_access_token(client_id, refresh_token):
    result = post_form("token", {"client_id": client_id, "refresh_token": refresh_token,
                                 "grant_type": "refresh_token", "scope": SCOPE})
    if not result.get("access_token"):
        raise ValueError("Microsoft authorisation failed. Run outlook_auth.py again and update MICROSOFT_REFRESH_TOKEN.")
    # shortcut: the POC keeps the configured refresh token; reauthorise when it expires or is revoked.
    return result["access_token"]


def authorise(client_id, output=TOKEN_FILE):
    if output.exists():
        raise ValueError("Local token file already exists. Remove it before authorising again.")
    UUID(client_id)
    device = post_form("devicecode", {"client_id": client_id, "scope": SCOPE})
    if not device.get("device_code"):
        raise ValueError("Microsoft sign-in could not start. Check the client ID and allow public client flows.")
    print(f'Open {device["verification_uri"]} and enter code {device["user_code"]}.')
    print("Sign in as dlg-electrical-solutions@outlook.com and approve sending email.")
    interval = int(device.get("interval", 5))
    deadline = time.monotonic() + int(device["expires_in"])
    while time.monotonic() < deadline:
        time.sleep(interval)
        result = post_form("token", {"client_id": client_id, "device_code": device["device_code"],
                                     "grant_type": "urn:ietf:params:oauth:grant-type:device_code"})
        if result.get("refresh_token"):
            descriptor = os.open(output, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
            with os.fdopen(descriptor, "w", encoding="utf-8") as stream:
                stream.write(result["refresh_token"])
            print(f"Authorised. Refresh token saved privately to {output.name}; it is Git-ignored.")
            print("Add that file's contents to MICROSOFT_REFRESH_TOKEN in GitHub Secrets, then delete the local file.")
            return
        error = result.get("error")
        if error == "slow_down":
            interval += 5
        elif error != "authorization_pending":
            raise ValueError("Microsoft sign-in failed or was declined. Run the helper again to retry.")
    raise ValueError("Microsoft sign-in timed out. Run the helper again to retry.")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--client-id", required=True, help="Application (client) ID from your Microsoft app registration")
    args = parser.parse_args()
    try:
        authorise(args.client_id)
    except (ValueError, KeyError, OSError):
        parser.exit(1, "Outlook setup failed. Check the client ID, public-client setting, connectivity and local token file.\n")


if __name__ == "__main__":
    main()
