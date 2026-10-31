# Electrical Quotation Automation POC

Enter materials and quantities, generate a mock quotation PDF, and email it to
your configured address through GitHub Actions. The recipient can use any email
provider. The sender must offer authenticated SMTP with TLS.

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

Open this private repository on GitHub, then **Settings → Secrets and variables
→ Actions**. Add these **repository variables**:

| Variable | Value |
| --- | --- |
| `SMTP_HOST` | Your sending provider's SMTP hostname |
| `SMTP_PORT` | `587` for STARTTLS (default), or `465` for implicit TLS |

Add these **repository secrets**, not source files or workflow inputs:

| Secret | Value |
| --- | --- |
| `SMTP_USERNAME` | SMTP login supplied by the sending provider |
| `SMTP_PASSWORD` | SMTP credential or app password supplied by the sending provider |
| `SMTP_FROM` | Sender address authorised by that provider |
| `QUOTE_EMAIL_TO` | Your receiving email address, with any provider |

For example, a Gmail sender uses `smtp.gmail.com`, port `587`, your Gmail
address as username/from, and an app password. Google requires 2-Step
Verification for app passwords, and availability depends on the account:
[Google app-password instructions](https://support.google.com/accounts/answer/185833).
You can send to an Outlook or business inbox from that Gmail sender.

Other SMTP services work with their own credentials and authorised sender.
This POC supports username/password or app-password authentication. Providers
or organisations requiring OAuth-only authentication need a later integration;
an ordinary account password does not necessarily enable SMTP. Get SMTP
settings from your provider. Do not paste credentials into chat or commit them.

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
redirect the email. A failed email submission fails the job. Successful SMTP
submission means the sending server accepted it, not that inbox delivery is
guaranteed; check spam folders too. There are no automatic send retries.
Rerunning a manual job can send another copy with the same quotation ID.

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
