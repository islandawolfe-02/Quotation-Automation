# Electrical Quotation Automation POC

Save the master Excel catalogue and a client requirements workbook to OneDrive,
then run GitHub Actions to generate and email a mock quotation PDF. After one-time
setup, Excel changes do not need a Git commit or push. Select one, two or
all three partners using checkboxes when starting the run. Outlook.com sends through
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
| `data/master_catalogue.xlsx` | Starter master to copy to OneDrive; 33 material variants and 6 labour/other charge rows |
| `examples/client_requirements.xlsx` | Starter client workbook to copy to OneDrive |
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

### Configure partner receiver secrets

In **Settings → Secrets and variables → Actions → Secrets**, create these
repository secrets. Each value must be one plain email address:

| Secret | Value |
| --- | --- |
| `PARTNER_1_EMAIL` | Partner 1's receiving email address |
| `PARTNER_2_EMAIL` | Partner 2's receiving email address |
| `PARTNER_3_EMAIL` | Partner 3's receiving email address |

The run form shows **Send to Partner 1**, **Send to Partner 2** and **Send to
Partner 3** as checkboxes. Partner 1 starts checked; the other two start
unchecked. Choose any one, any two, or all three. If all are unchecked, the
workflow fails with **Select at least one partner** before checkout or dependency
installation. The local CLI also checks before reading workbooks or generating
a PDF. GitHub's form cannot express a cross-field checkbox requirement, so this
is enforced by the workflow and application.

Only selected partners need populated, valid secrets. All selected recipients
appear in **To** and can see who else received the email. The sender stays in
**From**. Addresses are not workflow inputs or configuration files and are not printed by the application. The
previous recipient file and comma-separated receiver environment option have
been removed.

If addresses were committed in the old recipient file, deleting it on this
branch does not erase those earlier Git commits. Treat past repository readers
as having had access to those addresses. No history rewrite is performed here.

If an SMTP server accepts some recipients and rejects others, the job fails
after partial delivery. Check who received it before retrying.

`SMTP_FROM` keeps its existing name and supplies the sender header for either
mail service. Outlook does not use `SMTP_HOST`, `SMTP_PORT`, `SMTP_USERNAME` or
`SMTP_PASSWORD`; existing values can remain configured. If either Microsoft
OAuth secret is present, both are required and Graph is used. Failures do not
fall back to SMTP or retry the mail submission automatically.

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
   `offline_access` permissions, plus `Files.Read` when `--read-files` is
   supplied; do not choose application permissions or the client-credentials flow.
4. From this repository on your own computer, run:

   ```sh
   python3 outlook_auth.py --client-id YOUR_APPLICATION_CLIENT_ID --read-files
   ```

5. Open the Microsoft verification page shown by the helper, enter its code,
   and sign in as **dlg-electrical-solutions@outlook.com**. Approve permission
   to send mail and read OneDrive files. Only authorise a device code you just
   requested yourself.
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
as secrets, add a `SMTP_PASSWORD` secret and the partner receiver secrets,
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

## Set up OneDrive once

The workflow reads private files from the OneDrive belonging to the same account
used for Microsoft Graph: **dlg-electrical-solutions@outlook.com**. No public
sharing links or additional service credentials are needed. The starter Excel
files in this repository remain available for copying and local use.

1. Open [OneDrive](https://onedrive.live.com/) using that account. Create a folder
   called **Quotations** at the top level of **My files**.
2. Upload `data/master_catalogue.xlsx` and `examples/client_requirements.xlsx`
   into that folder. Their OneDrive paths will be:

   ```text
   Quotations/master_catalogue.xlsx
   Quotations/client_requirements.xlsx
   ```

3. Grant read access using the existing app registration. From this repository,
   run the updated helper (remove an old local `.outlook-refresh-token` first
   if it still exists):

   ```sh
   python3 outlook_auth.py --client-id YOUR_APPLICATION_CLIENT_ID --read-files
   ```

   Sign in as the sender and approve sending mail and reading files. Replace
   `MICROSOFT_REFRESH_TOKEN` in GitHub Secrets with the newly saved token, then
   delete the local token file. Keep `MICROSOFT_CLIENT_ID`, `SMTP_FROM` and the
   three partner secrets. See the sign-in instructions above for copying the
   token without printing it. **An old mail-only consent is insufficient for
   OneDrive reads.** Omitting `--read-files` keeps mail-only authorisation for
   local-file use.
4. Commit and push these **code/workflow changes once**. Ordinary Excel changes
   afterward happen in OneDrive and require no commits. To use the updated web
   run form reliably, the updated workflow must also be on the default branch;
   the CLI below can select this feature branch in the meantime.

The app requests delegated **Files.Read**, the least-privileged download
permission for a personal Microsoft account. This permits reading that account's
files, not only the Quotations folder, so use the business account for these
inputs. It does not grant file write or delete permission.
[Microsoft download permissions](https://learn.microsoft.com/en-us/graph/api/driveitem-get-content?view=graph-rest-1.0),
[Files.Read permission](https://learn.microsoft.com/en-us/graph/permissions-reference#filesread).

## Save Excel, then run the pipeline

1. Open the **OneDrive copy** of `master_catalogue.xlsx` in Excel online or from
   your synced OneDrive folder. Update rates and specs as needed, keeping a
   unique Item Code per variant and master Quantity = 1.
2. Open the **OneDrive copy** of `client_requirements.xlsx`. Fill client details
   in B4:B9 and set quantities and areas. Copy updated component rows from the
   master when their specifications change. Use a fresh Quotation ID per job.
3. **Save and wait for both files to finish syncing to OneDrive.** Check them
   in OneDrive online if unsure. Saving only the repository copies on your Mac
   does not update these cloud files. Avoid editing either workbook during a run.
4. Open **GitHub → Actions → Quotation POC → Run workflow**, select the branch
   containing the updated code, and use:

   | Field | Value |
   | --- | --- |
   | `requirements_file` | `Quotations/client_requirements.xlsx` |
   | `catalogue_file` | `Quotations/master_catalogue.xlsx` |
   | `send_partner_1` | Check to include Partner 1 |
   | `send_partner_2` | Check to include Partner 2 |
   | `send_partner_3` | Check to include Partner 3 |

5. Select at least one partner and click **Run workflow**. The workflow downloads
   the current saved workbooks, validates them, generates the PDF, and emails
   it to the selected partners in **To**. It never falls back to the repository
   samples if OneDrive fails. No commit, push or PDF download is needed per job.

You can save client workbooks under different filenames and enter the desired
OneDrive-relative path, such as `Quotations/Jobs/job-002.xlsx`. File fields
accept paths within this account's OneDrive, **not sharing URLs, Mac paths or
file uploads**. Paths and filenames appear in workflow metadata, so avoid
putting client names or other confidential details in them.

If the web form still shows older inputs from the default branch, use GitHub
CLI after the one-time code push. This example selects Partners 1 and 3:

```sh
gh workflow run quotation.yml --ref feature/input-details-with-excel-sheets \
  -f requirements_file=Quotations/client_requirements.xlsx \
  -f catalogue_file=Quotations/master_catalogue.xlsx \
  -f send_partner_1=true \
  -f send_partner_2=false \
  -f send_partner_3=true
```

Use the actual branch name if it changes. The workflow must exist on the default
branch for manual dispatch; `--ref` selects the branch version to execute.
[GitHub manual runs](https://docs.github.com/en/actions/how-tos/manage-workflow-runs/manually-run-a-workflow).

Pushes and pull requests run tests without reading OneDrive or sending email.
Only manual runs read the cloud workbooks and send the quotation. A successful
Graph submission (`202 Accepted`) means Microsoft accepted the request; check
the inbox and spam folder to confirm delivery. Graph also saves a copy in Sent
Items. There are no automatic mail retries; starting another run reads the latest
saved files and can send another quotation even with the same Quotation ID.

### If OneDrive downloads fail

- **Microsoft authorisation failed**, **HTTP 401** or **403**: run the helper
  with `--read-files`, approve file access and replace the refresh token secret.
- **HTTP 404**: check the path relative to **My files**, filename, sign-in
  account and completed sync. Open OneDrive online once if it has not been set up.
- **HTTP 429**: wait before starting another run.
- **OneDrive did not provide a secure workbook download URL**: check that each
  path points to a saved Excel file, not a folder or shortcut. The updated
  downloader uses Microsoft's `/content` endpoint and follows its HTTPS
  redirect without forwarding the Graph token; it no longer requires the
  download-URL field in file metadata. Push the fix and start a new workflow
  run on that branch, rather than rerunning a job using older code.
  [Microsoft file downloads](https://learn.microsoft.com/en-us/graph/api/driveitem-get-content?view=graph-rest-1.0).
- A workbook must be a non-empty `.xlsx` file no larger than **10 MiB**. It must
  retain the supplied sheet layout and plain-value cells. Invalid inputs stop
  the run before generating or sending a quotation.

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
or passwords. After pushing a code fix, start a **new Run workflow** on the updated branch;
re-running an older failed run uses the code from that older run.

No PDF artifacts or quotation history are uploaded, and generated PDFs are
removed at the end of the job. The emailed copy is your record. GitHub still
retains workflow logs/metadata. Downloaded Excel copies stay in runner memory
and are not uploaded as artifacts; the source workbooks remain in OneDrive
under its own retention/version-history settings. Previously committed client
data remains in Git history. Use mock client information while evaluating
this POC; this is not a zero-retention system.

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

To use the same cloud workbooks locally, configure `MICROSOFT_CLIENT_ID` and
`MICROSOFT_REFRESH_TOKEN` (with file-read consent), then run:

```sh
python quotation.py --onedrive
```

This generates the PDF without sending email. Add `--requirements` and
`--catalogue` to select other OneDrive-relative paths. Add `--send-email` only
when you want delivery.

For local email, configure the sender and partner secret environment variables,
then select partners explicitly. This example selects Partners 1 and 3:

```sh
SEND_PARTNER_1=true SEND_PARTNER_2=false SEND_PARTNER_3=true python quotation.py --send-email
```

The script does not load `.env` files. With no selection flags, local email
selects Partner 1 and requires `PARTNER_1_EMAIL`. Quote inputs are Excel-only;
the old JSON catalogue, charges and job files have been removed. Shared mock
data for automated tests now lives in `tests/fixtures/quotation.json` and is
not an operator input.

The PDF always says **MOCK QUOTATION - POC ONLY**. Before using it commercially,
replace the demo prices/details/terms, agree the applicable tax treatment and
rounding rules, and review the generated quotation.
