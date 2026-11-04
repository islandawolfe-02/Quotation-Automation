# Electrical Quotation Automation POC

Edit the master Excel catalogue and a client requirements workbook, then run
GitHub Actions to generate and email a mock quotation PDF. Select one, two or
all three recipient aliases when starting the run. Outlook.com sends through
Microsoft Graph with OAuth; other senders can use authenticated SMTP with TLS.
The user confirmed that Graph delivery is working on 2026-11-04.

The sample follows the supplied Tripura Bhavan quotation's section structure:
light fixtures, wires and consumables, labour, summary and terms. Other charges
have a separate section. The client document is not copied into this repository.
All business details, rates, tax settings and terms in this POC are fictional.

## Editable templates

| File | What to edit |
| --- | --- |
| `templates/quotation.json` | Mock business details and quotation terms |
| `data/master_catalogue.xlsx` | Editable master with 33 material variants and 6 labour/other charge rows |
| `examples/client_requirements.xlsx` | Client details and requested component quantities |
| `config/recipients.json` | Private-repo mapping from three aliases to receiving addresses |
| `quotation.py` | PDF layout, if its visual arrangement needs changing |

## Edit the Excel workbooks

Open `data/master_catalogue.xlsx` in Excel. The **Catalogue** sheet has a header
in row 5 and one component variant per row. Edit the mock specs, brands and
rates to match your own supplier information. Each combination of product,
specification, poles and brand needs its own stable, unique Item Code.

Both workbooks have the same columns:

| Column | Meaning |
| --- | --- |
| Item Code | Unique identifier for a specific component variant |
| Item Name | MCB, RCCB, SPD, wire, fixture, labour item, etc. |
| Quantity | Master: always 1 pricing unit. Client: requested quantity |
| Specs | For example `6A B-curve`, `16A C-curve`, `40A 30mA`, `63A 300mA` |
| Poles | For example `1P`, `2P`, `3P`, `4P`, `1P+N`; use `N/A` for other items |
| Unit | Piece, coil, length, point, trip, etc. |
| Rate INR | Price for one Unit, excluding tax |
| Tax % | Numeric percentage such as 0 or 18, not 0.18 for 18% |
| Category | Light Fixtures, Wires and Consumables, Labour Charges or Other Charges |
| Brand | Manufacturer/brand; sample rows use Demo |
| Area | Room or work area; client-specific values may differ from the master |

In `examples/client_requirements.xlsx`, fill the yellow fields **Client Name,
Address, Contact, Site, Quotation ID and Validity Days** in B4:B9. Contact is
stored as text to preserve phone prefixes and leading zeroes. All these details
are picked up for the PDF, including the client's address and contact.

The client component table starts at row 12. Edit **Quantity** and **Area**;
zero or blank Quantity omits that row. The other component details must match
the master. Copy whole master component rows into the client table to add new
variants, then set their requested quantities. Rows appended below the current
table are also read; keep the header and client detail labels in place.

**The current master owns the prices and tax settings.** Rate and tax columns
in the client workbook are previews and cannot override the master. Changing a
master rate updates the next quotation; changing specs, poles, units or brand
requires refreshing that component's row in the client workbook. Mismatches,
unknown/duplicate codes, missing details, negative quantities, formulas and
Excel error cells fail the run. Enter plain values in these input workbooks.

The Excel sample totals **INR 50,940.00**, with zero tax selected for the demo.
All sample rates, brands and specifications are illustrative, not verified
market prices or procurement recommendations. Replace them with your supplier's
data. A wire coil is priced per **90 metre coil**, not per metre; PVC conduit is
per **3 metre length**. For a pack price, use a pack as the Unit and retain
Quantity = 1 in the master. The pipeline does not estimate material quantities.

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
into **Secret**, then click **Add secret**. Outlook needs these three secrets:

| Secret | Value |
| --- | --- |
| `SMTP_FROM` | `dlg-electrical-solutions@outlook.com` |
| `MICROSOFT_CLIENT_ID` | Application (client) ID from the app registration below |
| `MICROSOFT_REFRESH_TOKEN` | Token created by the one-time local sign-in below |

### Configure the three receivers

Edit `config/recipients.json` in this private repo. Replace the three demo
addresses with your actual receiving addresses, keeping the alias keys:

```json
{
  "office": "office@your-domain.com",
  "accounts": "accounts@your-domain.com",
  "owner": "owner@your-domain.com"
}
```

The workflow's **recipients** choice offers all seven non-empty combinations:
`office`, `accounts`, `owner`, `office,accounts`, `office,owner`, `accounts,owner`,
or `office,accounts,owner`. Only these aliases appear in the run form. The
selected addresses are delivered using **BCC**, so recipients do not see one
another's addresses. Invalid or unconfigured selections fail before sending;
the shipped `.invalid` demo addresses cannot be used for email delivery.

Receiver addresses are no longer a required GitHub secret and the workflow does
not pass `QUOTE_EMAIL_TO`. For compatibility, local runs can still use that
environment variable when no alias selection is supplied. Alias selections
always use the configuration file.

This is the practical option for a private repo with trusted collaborators.
Addresses are not printed in ordinary logs, but collaborators with repo access
can read the file and its Git history. It is not encrypted. BCC hides addresses
from other recipients, not from the sender or Microsoft. If repo readers must
not know the addresses, use encrypted storage with a separately protected key.
GitHub variables do not add confidentiality because their values are unmasked
by default: [GitHub variables](https://docs.github.com/en/actions/concepts/workflows-and-actions/variables).

If an SMTP server accepts some recipients and rejects others, the job fails
after partial delivery. Check who received it before retrying.

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
(`587` for STARTTLS or `465` for implicit TLS). Keep `SMTP_USERNAME` and `SMTP_FROM`
as secrets, add a `SMTP_PASSWORD` secret, configure `config/recipients.json`,
and leave the two `MICROSOFT_*` secrets unset. The sender address must be
authorised by the provider.

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

1. Save the edited Excel workbooks, commit them and the recipient configuration,
   and push your branch. The runner reads the committed workbook versions;
   unsaved or unpushed edits on your computer are not available to GitHub.
   Pushes and pull requests run tests without sending email.
2. Open **Actions → Quotation POC → Run workflow**.
3. Select your branch. Set **requirements_file** to the client workbook path
   (default `examples/client_requirements.xlsx`) and **catalogue_file** to
   `data/master_catalogue.xlsx` or another master workbook in that branch.
4. Choose the **recipients** alias combination and run. The pipeline validates
   and prices the workbook, generates the PDF, and emails the selected receivers.
   A PDF download is not required. GitHub's run form accepts file paths, not
   direct Excel file uploads.

While this change is on the feature branch, the existing default-branch form
may still show the older input fields. After pushing, GitHub CLI can dispatch
the updated workflow explicitly on the feature branch:

```sh
gh workflow run quotation.yml --ref feature/input-details-with-excel-sheets \
  -f requirements_file=examples/client_requirements.xlsx \
  -f catalogue_file=data/master_catalogue.xlsx \
  -f recipients=office,accounts
```

Use the actual branch name if it changes. The workflow already exists on the
default branch; `--ref` selects the branch version to execute.
[GitHub manual runs](https://docs.github.com/en/actions/how-tos/manage-workflow-runs/manually-run-a-workflow).

Recipient aliases resolve through the configuration file; the client workbook
does not control delivery addresses. A failed email submission fails the job. Successful Graph
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
retains workflow logs/metadata. Committed client workbooks, including names,
addresses and contacts, remain in private Git history and are readable by repo
collaborators. Use mock client information while evaluating this POC; it does
not offer zero-retention input handling.

## Run locally

Requires Python 3.12 or later:

```sh
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python -m unittest discover -s tests -v
python quotation.py
```

The Excel sample is written to `output/pdf/POC-EXCEL-001.pdf`; `output/` is
Git-ignored. To generate another job locally:

```sh
python quotation.py --requirements path/to/client.xlsx --catalogue data/master_catalogue.xlsx
```

For local email, configure the sender environment variables, set
`QUOTE_RECIPIENT_ALIASES=office,owner`, and run `python quotation.py --send-email`.
The script does not load `.env` files. With no selection or legacy email
environment variable, the local default alias is `office`.

The original `examples/requirements.json`, `data/catalogue.json` and
`data/charges.json` remain available for the legacy JSON mode:
`python quotation.py --requirements examples/requirements.json`. That original
sample totals INR 42,400.00. `QUOTATION_REQUIREMENTS_JSON` also remains supported
locally. JSON mode uses those JSON catalogues; the workbook workflow uses the
Excel master and does not expose a JSON job input.

The PDF always says **MOCK QUOTATION - POC ONLY**. Before using it commercially,
replace the demo prices/details/terms, agree the applicable tax treatment and
rounding rules, and review the generated quotation.
