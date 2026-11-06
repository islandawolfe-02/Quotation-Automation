"""Generate a mock electrical quotation and optionally email its PDF."""

import argparse
import base64
import json
import os
import re
import smtplib
import ssl
from contextlib import ExitStack
from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from email.errors import HeaderParseError
from email.headerregistry import Address
from email.message import EmailMessage
from email.policy import SMTP as SMTP_POLICY
from io import BytesIO
from pathlib import Path
from urllib.error import HTTPError
from urllib.parse import quote, urlsplit
from urllib.request import Request, urlopen
from xml.sax.saxutils import escape
from zipfile import BadZipFile
from zoneinfo import ZoneInfo

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import KeepTogether, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from outlook_auth import refresh_access_token

ROOT = Path(__file__).resolve().parent
CENT = Decimal("0.01")
GROUPS = ("Light Fixtures", "Wires and Consumables", "Labour Charges", "Other Charges")
EXCEL_COLUMNS = ("Item Code", "Item Name", "Quantity", "Specs", "Poles", "Unit", "Rate INR", "Tax %", "Category", "Brand", "Area")
MAX_WORKBOOK_BYTES = 10 * 1024 * 1024


def download_onedrive_workbook(path, token):
    path = text(path, "OneDrive workbook path")
    if (not path.lower().endswith(".xlsx") or any(part in ("", ".", "..") for part in path.split("/"))
            or any(character in path for character in '\\:*<>?|' + '\r\n')):
        raise ValueError("Use a OneDrive-relative .xlsx path, for example Quotations/client_requirements.xlsx")
    request = Request(f"https://graph.microsoft.com/v1.0/me/drive/root:/{quote(path, safe='/')}"
                      "?$select=id,@microsoft.graph.downloadUrl", headers={"Authorization": f"Bearer {token}"})
    try:
        with urlopen(request, timeout=30) as response:
            metadata = json.loads(response.read(64 * 1024))
        download_url = metadata.get("@microsoft.graph.downloadUrl", "") if isinstance(metadata, dict) else ""
        if not isinstance(download_url, str):
            raise ValueError("OneDrive did not provide a secure workbook download URL")
        parsed = urlsplit(download_url)
        if parsed.scheme != "https" or not parsed.hostname or parsed.username or parsed.password:
            raise ValueError("OneDrive did not provide a secure workbook download URL")
        # Download URLs are preauthenticated; never forward the Graph bearer token.
        with urlopen(download_url, timeout=30) as response:
            content = response.read(MAX_WORKBOOK_BYTES + 1)
        if not content or len(content) > MAX_WORKBOOK_BYTES:
            raise ValueError("OneDrive workbook must be non-empty and at most 10 MiB")
        return BytesIO(content)
    except HTTPError as exc:
        with exc:
            status = exc.code
        if status in (401, 403):
            reason = "Run outlook_auth.py with --read-files and update MICROSOFT_REFRESH_TOKEN."
        elif status == 404:
            reason = "Check the workbook path and that it has finished syncing to the authorised account's OneDrive."
        elif status == 429:
            reason = "Microsoft rate-limited the request. Wait before starting a new run."
        else:
            reason = "Could not download the saved workbook. Try a new run after checking OneDrive."
        raise ValueError(f"OneDrive workbook download failed (HTTP {status}). {reason}") from None


def excel_value(cell):
    if cell.data_type in ("f", "e"):
        raise ValueError(f"{cell.parent.title}!{cell.coordinate} must contain a plain value, not a formula or Excel error")
    return cell.value


def excel_rows(sheet, header_row):
    headers = [excel_value(cell) for cell in sheet[header_row]]
    if any(headers.count(name) != 1 for name in EXCEL_COLUMNS):
        raise ValueError(f"{sheet.title} row {header_row} must contain each template column exactly once")
    for cells in sheet.iter_rows(min_row=header_row + 1):
        values = {name: excel_value(cells[headers.index(name)]) for name in EXCEL_COLUMNS}
        if any(value is not None for value in values.values()):
            yield cells[0].row, values


def load_excel_inputs(master_path, client_path):
    try:
        master = load_workbook(master_path, read_only=True, data_only=False, keep_links=False)
        try:
            if "Catalogue" not in master.sheetnames:
                raise ValueError("Master workbook needs a Catalogue worksheet")
            entries = {}
            for row_number, row in excel_rows(master["Catalogue"], 5):
                location = f"Catalogue row {row_number}"
                code = text(row["Item Code"], f"{location} Item Code")
                if code in entries:
                    raise ValueError(f"Duplicate Item Code in {location}")
                if number(row["Quantity"], f"{location} Quantity", positive=True) != 1:
                    raise ValueError(f"{location} Quantity must be 1; Rate INR is the price per Unit")
                for field in ("Item Name", "Specs", "Poles", "Unit", "Category", "Brand"):
                    text(row[field], f"{location} {field}")
                if row["Item Name"].strip().upper() in ("MCB", "RCCB", "RCBO", "SPD", "ISOLATOR") and row["Poles"].strip().upper() == "N/A":
                    raise ValueError(f"{location} needs Poles for its protective device")
                specification = row["Specs"].strip()
                if row["Poles"].strip().upper() != "N/A":
                    specification += f'; {row["Poles"].strip()}'
                entry = {"code": code, "description": row["Item Name"].strip(),
                         "category": row["Category"].strip(), "unit": row["Unit"].strip(),
                         "specification": f'{row["Brand"].strip()} | {specification}',
                         "rate": row["Rate INR"], "tax_percent": row["Tax %"]}
                entries[code] = (row, entry)
        finally:
            master.close()
        client = load_workbook(client_path, read_only=True, data_only=False, keep_links=False)
        try:
            if "Requirements" not in client.sheetnames:
                raise ValueError("Client workbook needs a Requirements worksheet")
            sheet = client["Requirements"]
            request = {"materials": [], "charges": []}
            fields = (("Client Name", "customer"), ("Address", "customer_address"),
                      ("Contact", "customer_contact"), ("Site", "site"),
                      ("Quotation ID", "quotation_id"), ("Validity Days", "validity_days"))
            for position, (label, field) in enumerate(fields, 4):
                if excel_value(sheet.cell(position, 1)) != label:
                    raise ValueError(f"Requirements!A{position} must be labelled {label}")
                value = excel_value(sheet.cell(position, 2))
                request[field] = value if field == "validity_days" else text(value, f"Requirements!B{position}")
            for row_number, row in excel_rows(sheet, 12):
                location = f"Requirements row {row_number}"
                quantity = number(row["Quantity"] if row["Quantity"] is not None else 0, f"{location} Quantity")
                if quantity == 0:
                    continue
                code = text(row["Item Code"], f"{location} Item Code")
                if code not in entries:
                    raise ValueError(f"Unknown Item Code in {location}; copy the component from the master")
                original, entry = entries[code]
                for field in ("Item Name", "Specs", "Poles", "Unit", "Category", "Brand"):
                    if text(row[field], f"{location} {field}").casefold() != original[field].strip().casefold():
                        raise ValueError(f"{location} {field} does not match the master; copy the current component row")
                kind = "materials" if entry["category"] in GROUPS[:2] else "charges"
                request[kind].append({"code": code, "quantity": quantity, "area": row["Area"] or "General"})
        finally:
            client.close()
        catalogue = [entry for _, entry in entries.values() if entry["category"] in GROUPS[:2]]
        charges = [entry for _, entry in entries.values() if entry["category"] not in GROUPS[:2]]
        return request, catalogue, charges
    except (BadZipFile, InvalidFileException, KeyError):
        raise ValueError("Could not read the Excel templates. Use valid .xlsx workbooks with the supplied sheet names and columns.") from None


def recipient_addresses():
    selected = []
    for partner in (1, 2, 3):
        flag = os.environ.get(f"SEND_PARTNER_{partner}", "true" if partner == 1 else "false").strip().lower()
        if flag not in ("true", "false"):
            raise ValueError(f"SEND_PARTNER_{partner} must be true or false")
        if flag == "true":
            selected.append(f"PARTNER_{partner}_EMAIL")
    if not selected:
        raise ValueError("Select at least one partner to receive the quotation")
    addresses = []
    for name in selected:
        if not os.environ.get(name, "").strip():
            raise ValueError(f"Missing selected recipient secret: {name}")
        try:
            address = Address(addr_spec=os.environ[name].strip())
            if not address.username or not address.domain or address.domain.lower().endswith(".invalid"):
                raise ValueError
        except (ValueError, HeaderParseError):
            raise ValueError(f"{name} must contain exactly one valid email address") from None
        addresses.append(address)
    return addresses


def load_json(path):
    return json.loads(Path(path).read_text(encoding="utf-8"), parse_float=Decimal)


def text(value, field):
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f"{field} must be non-empty text")
    return value.strip()


def number(value, field, *, positive=False, maximum=Decimal("1000000000")):
    try:
        result = Decimal(str(value))
    except InvalidOperation:
        raise ValueError(f"{field} must be a decimal number") from None
    if not result.is_finite() or result < 0 or result > maximum or (positive and result == 0):
        raise ValueError(f"{field} is outside the allowed range")
    return result


def calculate(request, catalogue, charges):
    if not isinstance(request, dict):
        raise ValueError("Requirements must be a JSON object")
    quote_id = text(request.get("quotation_id"), "quotation_id")
    if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_-]{0,59}", quote_id):
        raise ValueError("quotation_id must contain 1-60 letters, numbers, underscores or hyphens")
    for field in ("customer", "site"):
        text(request.get(field), field)
    for field in ("customer_address", "customer_contact"):
        if field in request:
            text(request[field], field)
    validity = request.get("validity_days")
    if type(validity) is not int or not 1 <= validity <= 365:
        raise ValueError("validity_days must be a whole number from 1 to 365")

    lines = []
    seen_codes = set()
    for kind, entries, allowed in (("materials", catalogue, GROUPS[:2]), ("charges", charges, GROUPS[2:])):
        if not isinstance(entries, list):
            raise ValueError(f"{kind} catalogue must be a list")
        lookup = {}
        for entry in entries:
            if not isinstance(entry, dict):
                raise ValueError(f"{kind} catalogue entries must be objects")
            code = text(entry.get("code"), "catalogue code")
            if code in seen_codes:
                raise ValueError(f"Duplicate catalogue code: {code}")
            seen_codes.add(code)
            if entry.get("category") not in allowed:
                raise ValueError(f"Invalid category for {code}")
            for field in ("description", "specification", "unit"):
                text(entry.get(field), f"{code} {field}")
            rate = number(entry.get("rate"), f"{code} rate")
            if rate != rate.quantize(CENT):
                raise ValueError(f"{code} rate must have at most two decimal places")
            tax = number(entry.get("tax_percent"), f"{code} tax_percent", maximum=Decimal(100))
            lookup[code] = {**entry, "rate": rate, "tax_percent": tax}
        selections = request.get(kind)
        if not isinstance(selections, list) or (kind == "materials" and not selections):
            raise ValueError(f"{kind} must be a list; at least one material is required")
        for selection in selections:
            if not isinstance(selection, dict):
                raise ValueError(f"{kind} entries must be objects")
            code = text(selection.get("code"), "item code")
            if code not in lookup:
                raise ValueError(f"Unknown {kind} code: {code}")
            item = lookup[code]
            quantity = number(selection.get("quantity"), f"{code} quantity", positive=True)
            area = text(selection.get("area", "General"), f"{code} area")
            amount = (quantity * item["rate"]).quantize(CENT, rounding=ROUND_HALF_UP)
            tax = (amount * item["tax_percent"] / 100).quantize(CENT, rounding=ROUND_HALF_UP)
            lines.append({**item, "quantity": quantity, "area": area, "amount": amount, "tax": tax})
    subtotal = sum((line["amount"] for line in lines), Decimal("0.00"))
    tax = sum((line["tax"] for line in lines), Decimal("0.00"))
    return {"lines": lines, "subtotal": subtotal, "tax": tax, "total": subtotal + tax}


def render_pdf(path, request, result, template):
    styles = getSampleStyleSheet()
    styles["Title"].textColor = colors.black
    styles["Heading2"].textColor = colors.black
    styles["BodyText"].fontSize = 9
    styles["BodyText"].leading = 12
    styles["BodyText"].spaceAfter = 5

    def paragraph(value, style="BodyText"):
        return Paragraph(escape(str(value)).replace("\n", "<br/>"), styles[style])

    def table(rows, widths, numeric=()):
        cells = [[paragraph(cell) for cell in row] for row in rows]
        tab = Table(cells, colWidths=widths, repeatRows=1, hAlign="LEFT")
        commands = [("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E8EDF2")),
                    ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D9D9D9")),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("TOPPADDING", (0, 0), (-1, -1), 5),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 5)]
        for column in numeric:
            for row in cells[1:]:
                row[column].style = styles["BodyText"].clone("number", alignment=2)
        tab.setStyle(TableStyle(commands))
        return tab

    business = template["business"]
    customer_details = f'Prepared for: {request["customer"]}'
    if request.get("customer_address"):
        customer_details += f'\nAddress: {request["customer_address"]}'
    if request.get("customer_contact"):
        customer_details += f'\nContact: {request["customer_contact"]}'
    customer_details += f'\nSite: {request["site"]}'
    today = datetime.now(ZoneInfo("Asia/Kolkata")).date().isoformat()
    story = [paragraph(business["name"], "Title"),
             paragraph("Electrical Works Quotation", "Heading2"),
             paragraph("MOCK QUOTATION - POC ONLY - NOT FOR COMMERCIAL USE"),
             paragraph(f'{business["address"]}\n{business["phone"]} | {business["email"]}\n{business["registration"]}'),
             Spacer(1, 3 * mm),
             paragraph(f'Quotation: {request["quotation_id"]} | Date: {today} | Validity: {request["validity_days"]} days'),
             paragraph(customer_details),
             paragraph("All amounts are in INR. Line amounts exclude tax; tax is shown separately in the summary.")]
    for category in GROUPS:
        lines = [line for line in result["lines"] if line["category"] == category]
        if not lines:
            continue
        story.append(paragraph(category, "Heading2"))
        rows = [["Area / Description", "Specification", "Qty / Unit", "Rate INR", "Amount INR"]]
        for line in lines:
            description = f'{line["area"]}\n{line["description"]}\n{line["code"]}' if category in GROUPS[:2] else line["description"]
            rows.append([description, f'{line["specification"]}\nTax: {line["tax_percent"]}%',
                         f'{line["quantity"]:g} {line["unit"]}', f'{line["rate"]:,.2f}', f'{line["amount"]:,.2f}'])
        story.append(table(rows, [59 * mm, 42 * mm, 27 * mm, 24 * mm, 28 * mm], (2, 3, 4)))
        total = sum((line["amount"] for line in lines), Decimal("0.00"))
        story.append(paragraph(f"{category} subtotal: INR {total:,.2f}"))

    rows = [["Scope / Category", "Amount INR"]]
    for category in GROUPS:
        amount = sum((line["amount"] for line in result["lines"] if line["category"] == category), Decimal("0.00"))
        rows.append([category, f"{amount:,.2f}"])
    rows.extend([["Subtotal before tax", f'{result["subtotal"]:,.2f}'],
                 ["Tax (illustrative settings)", f'{result["tax"]:,.2f}'],
                 ["Grand Total", f'{result["total"]:,.2f}']])
    story.append(KeepTogether([paragraph("Quotation Summary", "Heading2"), table(rows, [135 * mm, 45 * mm], (1,))]))
    story.append(paragraph("Terms and Conditions", "Heading2"))
    for index, term in enumerate(template["terms"], 1):
        story.append(paragraph(f"{index}. {term}"))
    story.extend([Spacer(1, 5 * mm), paragraph("Contractor signature: ____________________    Client acceptance: ____________________")])

    def footer(canvas, doc):
        canvas.setFont("Helvetica", 8)
        canvas.drawString(15 * mm, 10 * mm, "Mock electrical quotation | All prices are illustrative")
        canvas.drawRightString(195 * mm, 10 * mm, f"Page {doc.page}")

    path.parent.mkdir(parents=True, exist_ok=True)
    SimpleDocTemplate(str(path), pagesize=(210 * mm, 297 * mm), leftMargin=15 * mm,
                      rightMargin=15 * mm, topMargin=15 * mm, bottomMargin=18 * mm,
                      title="Mock Electrical Works Quotation", author=business["name"]).build(
                          story, onFirstPage=footer, onLaterPages=footer)


def send_email(path, quote_id):
    oauth = bool(os.environ.get("MICROSOFT_CLIENT_ID") or os.environ.get("MICROSOFT_REFRESH_TOKEN"))
    credentials = ("MICROSOFT_CLIENT_ID", "MICROSOFT_REFRESH_TOKEN") if oauth else ("SMTP_HOST", "SMTP_USERNAME", "SMTP_PASSWORD")
    names = ("SMTP_FROM",) + credentials
    missing = [name for name in names if not os.environ.get(name)]
    if missing:
        raise ValueError("Missing email settings: " + ", ".join(missing))
    sender = str(Address(addr_spec=text(os.environ["SMTP_FROM"], "SMTP_FROM")))
    recipients = recipient_addresses()
    message = EmailMessage()
    message["Subject"] = f"Mock electrical quotation {quote_id}"
    message["From"] = sender
    message["To"] = recipients
    message.set_content("Attached is your mock electrical quotation. All prices and tax settings are illustrative.\n")
    message.add_attachment(path.read_bytes(), maintype="application", subtype="pdf", filename=path.name)
    if oauth:
        token = refresh_access_token(text(os.environ["MICROSOFT_CLIENT_ID"], "MICROSOFT_CLIENT_ID"),
                                     text(os.environ["MICROSOFT_REFRESH_TOKEN"], "MICROSOFT_REFRESH_TOKEN"))
        request = Request("https://graph.microsoft.com/v1.0/me/sendMail",
                          data=base64.b64encode(message.as_bytes(policy=SMTP_POLICY)),
                          headers={"Authorization": f"Bearer {token}", "Content-Type": "text/plain"})
        try:
            with urlopen(request, timeout=30) as response:
                if response.status != 202:
                    raise ValueError("Microsoft Graph did not confirm email submission; check the mailbox before retrying.")
        except HTTPError as exc:
            # Provider bodies may contain addresses or tokens; report only fixed guidance and status.
            with exc:
                status = exc.code
            if status in (401, 403):
                reason = "Run outlook_auth.py again, approve Mail.Send and update MICROSOFT_REFRESH_TOKEN. SMTP_FROM must be the authorised Outlook address."
            elif status == 429:
                reason = "Microsoft rate-limited the request. Wait before starting a new run."
            else:
                reason = "Microsoft did not confirm email submission. Check the mailbox before retrying."
            raise ValueError(f"Microsoft Graph rejected email (HTTP {status}). {reason}") from None
        return

    username = text(os.environ["SMTP_USERNAME"], "SMTP_USERNAME")
    port = int(os.environ.get("SMTP_PORT", "587"))
    if not 1 <= port <= 65535:
        raise ValueError("SMTP_PORT must be between 1 and 65535")
    context = ssl.create_default_context()
    client = smtplib.SMTP_SSL if port == 465 else smtplib.SMTP
    options = {"context": context} if port == 465 else {}
    with client(os.environ["SMTP_HOST"], port, timeout=30, **options) as smtp:
        if port != 465:
            smtp.starttls(context=context)
        smtp.login(username, os.environ["SMTP_PASSWORD"])
        refused = smtp.send_message(message)
        if refused:
            raise smtplib.SMTPException("Recipient rejected")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--requirements", help="Client .xlsx path; relative to OneDrive root with --onedrive")
    parser.add_argument("--catalogue", help="Master .xlsx path; relative to OneDrive root with --onedrive")
    parser.add_argument("--onedrive", action="store_true", help="Read the latest saved Excel workbooks from OneDrive")
    parser.add_argument("--send-email", action="store_true")
    args = parser.parse_args()
    try:
        if args.send_email:
            recipient_addresses()
        requirements = args.requirements or ("Quotations/client_requirements.xlsx" if args.onedrive else ROOT / "examples/client_requirements.xlsx")
        catalogue = args.catalogue or ("Quotations/master_catalogue.xlsx" if args.onedrive else ROOT / "data/master_catalogue.xlsx")
        if Path(requirements).suffix.lower() != ".xlsx" or Path(catalogue).suffix.lower() != ".xlsx":
            raise ValueError("Requirements and master catalogue must be .xlsx workbooks")
        with ExitStack() as inputs:
            if args.onedrive:
                token = refresh_access_token(text(os.environ.get("MICROSOFT_CLIENT_ID"), "MICROSOFT_CLIENT_ID"),
                                             text(os.environ.get("MICROSOFT_REFRESH_TOKEN"), "MICROSOFT_REFRESH_TOKEN"),
                                             read_files=True)
                catalogue = inputs.enter_context(download_onedrive_workbook(str(catalogue), token))
                requirements = inputs.enter_context(download_onedrive_workbook(str(requirements), token))
            else:
                catalogue, requirements = Path(catalogue), Path(requirements)
            request, catalogue, charges = load_excel_inputs(catalogue, requirements)
        result = calculate(request, catalogue, charges)
        output = ROOT / "output/pdf" / f'{request["quotation_id"]}.pdf'
        render_pdf(output, request, result, load_json(ROOT / "templates/quotation.json"))
        print("Mock quotation PDF generated successfully.", flush=True)
        if args.send_email:
            send_email(output, request["quotation_id"])
            print("Email accepted by the sending server. Check the recipient inbox and spam folder.")
    except smtplib.SMTPAuthenticationError as exc:
        detail = exc.smtp_error.decode(errors="replace") if isinstance(exc.smtp_error, bytes) else str(exc.smtp_error)
        status = re.search(r"\b[245]\.\d{1,3}\.\d{1,3}\b", detail)
        code = f"SMTP {exc.smtp_code}" + (f" / {status.group()}" if status else "")
        if "smtpclientauthentication is disabled" in detail.lower():
            reason = "SMTP authentication is disabled for the mailbox or tenant. The mail provider must resolve this restriction."
        elif "basic authentication is disabled" in detail.lower():
            reason = "Password-based SMTP is disabled. Configure both Microsoft OAuth secrets for Outlook."
        elif "locked" in detail.lower() or "blocked" in detail.lower():
            reason = "The provider reports an account or sign-in restriction. Check the account's security alerts."
        else:
            reason = "Authentication was rejected. Verify SMTP_USERNAME matches the Outlook account used during authorisation and SMTP_FROM."
        parser.exit(1, f"Quotation failed ({code}). {reason} Provider response text was omitted to protect private data.\n")
    except (ValueError, KeyError, TypeError, OSError, smtplib.SMTPException) as exc:
        # SMTP responses can contain private addresses; keep provider responses out of CI logs.
        if isinstance(exc, (smtplib.SMTPException, OSError)):
            parser.exit(1, f"Quotation failed ({type(exc).__name__}). Check files, connectivity and email settings.\n")
        parser.exit(1, f"Quotation failed: {exc}\n")


if __name__ == "__main__":
    main()
