# Electrical Quotation Automation POC

Enter materials and quantities, generate a mock quotation PDF, and email it to
your configured address through GitHub Actions. The recipient can use any email
provider. Outlook.com sends through Microsoft Graph with OAuth; other senders
can use authenticated SMTP with TLS.

The sample follows the supplied Tripura Bhavan quotation's section structure:
light fixtures, wires and consumables, labour, summary and terms. Other charges
have a separate section. The client document is not copied into this repository.
All business details, rates, tax settings and terms in this POC are fictional.

## Editable templates

| File | What to edit |
| --- | --- |
| `templates/quotation.json` | Mock business details and quotation terms |
| `data/catalogue.json` | Material codes, descriptions, specifications, units, selling rates and tax percentages |
| `data/charges.json` | Labour rates, transport, testing and other charges |
| `examples/requirements.json` | Quotation ID, customer, site, validity, material quantities/areas and charge quantities |
| `quotation.py` | PDF layout, if its visual arrangement needs changing |

The sample totals **INR 42,400.00**, with zero tax explicitly selected for the
demonstration. These are not market prices or tax advice. A wire coil is priced
per **90 metre coil**, not per metre; PVC conduit is per **3 metre length**.
Quantities always use the unit defined in the catalogue. Include only the
labour/other charge codes that apply; an empty `charges` list is allowed.

Rates exclude tax. Amounts are rounded to two decimals using half-up rounding
per line, then tax is calculated and rounded per line. The summary adds these
rounded amounts. Negative amounts, missing/unknown codes, invalid quantities,
duplicate catalogue codes and invalid tax percentages fail the run.

## Set up email once

For the selected sender, **dlg-electrical-solutions@outlook.com**, follow the
Outlook.com instructions below. Outlook sends through the Microsoft Graph mail
API with delegated `Mail.Send` permission. This supports personal Outlook
accounts and does not require SMTP authentication to be enabled on the mailbox.
[Microsoft Graph sendMail documentation](https://learn.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0).

### Outlook.com sender

Open this private repository on GitHub, then **Settings → Secrets and variables
→ Actions → Secrets**. No repository variables are required for Outlook.

Under **Secrets → New repository secret**, enter the **Name**, paste its value
into **Secret**, then click **Add secret**. Repeat for these four secrets:

| Secret | Value |
| --- | --- |
| `SMTP_FROM` | `dlg-electrical-solutions@outlook.com` |
| `QUOTE_EMAIL_TO` | Your receiving email address; use the same Outlook address for a self-test if desired |
| `MICROSOFT_CLIENT_ID` | Application (client) ID from the app registration below |
| `MICROSOFT_REFRESH_TOKEN` | Token created by the one-time local sign-in below |

`SMTP_FROM` keeps its existing name and supplies the sender header for either
mail service. Outlook does not use `SMTP_HOST`, `SMTP_PORT`, `SMTP_USERNAME` or
`SMTP_PASSWORD`; existing values can remain configured. If either Microsoft
OAuth secret is present, both are required and Graph is used. Failures do not
fall back to SMTP or retry the mail submission automatically.

**Already set up the earlier SMTP version?** Keep your existing app registration
and `MICROSOFT_CLIENT_ID`. Use the updated helper to approve `Mail.Send` and
replace the `MICROSOFT_REFRESH_TOKEN` secret:

```sh
rm -f .outlook-refresh-token
python3 outlook_auth.py --client-id YOUR_APPLICATION_CLIENT_ID
pbcopy < .outlook-refresh-token
```

Complete the Microsoft sign-in before the copy command. On GitHub, edit the
existing `MICROSOFT_REFRESH_TOKEN` secret and paste the replacement token. Delete
the local token file after saving it in GitHub. Push the updated code and start
a new manual run on `main`. The old `SMTP.Send` consent does not grant the new
`Mail.Send` permission. No new application registration is necessary.

Register and authorise the application once:

1. Open [Microsoft Entra app registrations](https://entra.microsoft.com/) and
   choose **App registrations → New registration**. Name it `Quotation POC`.
   Registration requires an Entra tenant where you have permission to register
   an app; a personal Outlook mailbox alone does not provide that directory.
   See [Microsoft's registration prerequisites](https://learn.microsoft.com/en-us/graph/auth-register-app-v2).
2. Select **Personal Microsoft accounts only**, or **Accounts in any
   organizational directory and personal Microsoft accounts**. Register it
   and copy **Application (client) ID** into `MICROSOFT_CLIENT_ID` above.
3. In the app's **Authentication** settings, set **Allow public client flows**
   to **Yes** and save. No client secret or redirect URI is needed for this
   device-code flow. The helper requests delegated `Mail.Send` and
   `offline_access` permissions when you sign in; do not choose application
   permissions or the client-credentials flow.
4. From this repository on your own computer, run:

   ```sh
   python3 outlook_auth.py --client-id YOUR_APPLICATION_CLIENT_ID
   ```

5. Open the Microsoft verification page shown by the helper, enter its code,
   and sign in as **dlg-electrical-solutions@outlook.com**. Approve permission
   to send mail. Only authorise a device code you just requested yourself.
   This helper signs in only; it does not send email.
6. The helper creates `.outlook-refresh-token` locally with private file
   permissions. Copy its contents into the `MICROSOFT_REFRESH_TOKEN` repository
   secret. On macOS, `pbcopy < .outlook-refresh-token` copies the token without
   printing it. Once saved in GitHub, delete the local token file. It is also
   Git-ignored. Do not paste it into chat, commit it, or put it in workflow inputs.

The workflow renews the short-lived access token on every run. This POC does
not write rotated refresh tokens back to GitHub Secrets. Re-run the local
sign-in helper and replace `MICROSOFT_REFRESH_TOKEN` when authorisation expires
or is revoked; remove any existing local token file first. Refresh tokens
generally last 90 days for this flow but can be revoked earlier.
[Microsoft token lifetime guidance](https://learn.microsoft.com/en-us/entra/identity-platform/refresh-tokens).

[Microsoft Graph sendMail](https://learn.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0),
[device-code sign-in](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-device-code).

### Other SMTP senders

For other providers, set `SMTP_HOST` and `SMTP_PORT` to their SMTP settings
(`587` for STARTTLS or `465` for implicit TLS). Keep `SMTP_USERNAME`, `SMTP_FROM`
and `QUOTE_EMAIL_TO` as secrets, add a `SMTP_PASSWORD` secret, and leave the two
`MICROSOFT_*` secrets unset. The sender address must be authorised by the provider.

For example, a Gmail sender uses `smtp.gmail.com`, port `587`, your Gmail
address as username/from, and an app password. Google requires 2-Step
Verification for app passwords, and availability depends on the account:
[Google app-password instructions](https://support.google.com/accounts/answer/185833).
You can send to an Outlook or business inbox from that Gmail sender.

Other SMTP services work with their own credentials and authorised sender.
OAuth support here is specifically for personal Outlook.com accounts. Other
OAuth providers need their own integration. An ordinary account password does
not necessarily enable SMTP. Get SMTP settings from your provider.

## Run the pipeline

1. Push these files to the default branch. Pushes and pull requests run tests
   without sending email.
2. Open **Actions → Quotation POC → Run workflow**.
3. Leave `requirements_json` blank to use `examples/requirements.json`, or paste
   a complete JSON object following that file's format to supply a new job
   directly at run time.
4. Run the workflow. It validates the inputs, generates the PDF, and emails it
   to `QUOTE_EMAIL_TO`. No PDF download is required.

The recipient is deliberately configured in a secret, so customer input cannot
redirect the email. A failed email submission fails the job. Successful Graph
submission (`202 Accepted`) or SMTP submission means the service accepted the
request; check the inbox and spam folder to confirm delivery. Graph also saves
a copy in the sending account's Sent Items. There are no automatic send retries.
Rerunning a manual job can send another copy with the same quotation ID.

### If Outlook authentication fails

If you still see **SMTP 535 / 5.7.139**, that run is using the earlier SMTP code,
or both Microsoft OAuth secrets are absent. The updated Outlook path makes no
SMTP connection. Check that the updated code is pushed and both secrets exist.

- **Microsoft authorisation failed**: run the updated helper, approve `Mail.Send`,
  and replace `MICROSOFT_REFRESH_TOKEN`.
- **Microsoft Graph rejected email (HTTP 401 or 403)**: reauthorise for
  `Mail.Send` and check that `SMTP_FROM` is the Outlook address you signed in as.
- **HTTP 429**: wait before starting a new run; Microsoft rate-limited the request.
- Other failed or unconfirmed submissions: check the recipient inbox and the
  sender's Sent Items before retrying, to avoid duplicate messages.

The script reports HTTP status and fixed guidance without logging the Graph
response body, email addresses, or tokens.

Share the safe error code and fixed explanation when asking for help, not tokens
or passwords. After pushing a code fix, start a **new Run workflow** on `main`;
re-running an older failed run uses the code from that older run.

No PDF artifacts or quotation history are uploaded, and generated PDFs are
removed at the end of the job. The emailed copy is your record. GitHub still
retains workflow logs/metadata, and JSON entered as a workflow input is not a
secret. Committed requirements remain in Git history. Use mock information for
this POC; it does not offer confidential or zero-retention input handling.

## Run locally

Requires Python 3.12 or later:

```sh
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python -m unittest discover -s tests -v
python quotation.py
```

The sample is written to `output/pdf/POC-001.pdf`; `output/` is Git-ignored.
For another input file, run `python quotation.py --requirements path/to/job.json`.
For local email, supply the same SMTP environment variables and run
`python quotation.py --send-email`. The script does not load `.env` files.

The PDF always says **MOCK QUOTATION - POC ONLY**. Before using it commercially,
replace the demo prices/details/terms, agree the applicable tax treatment and
rounding rules, and review the generated quotation.
