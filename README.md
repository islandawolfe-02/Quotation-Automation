# Quotation Automation

A proof of concept for generating PDF quotations from structured inputs and
emailing them to selected recipients.

Sample inputs and rates are included for demonstration only.

## Local use

Requires Python 3.12 or later.

```sh
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python -m unittest discover -s tests -v
python quotation.py
```

Local generation saves a PDF under `output/`. Email delivery requires separate
configuration and explicit execution.

## GitHub Actions

Pushes and pull requests run automated checks. A manually triggered workflow
can generate and email a quotation after deployment configuration is complete.
