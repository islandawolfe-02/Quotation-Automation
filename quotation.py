"""Generate a mock electrical quotation and optionally email its PDF."""

import argparse
import json
import os
import re
import smtplib
import ssl
from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from email.headerregistry import Address
from email.message import EmailMessage
from pathlib import Path
from xml.sax.saxutils import escape
from zoneinfo import ZoneInfo

from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from outlook_auth import refresh_access_token

ROOT = Path(__file__).resolve().parent
CENT = Decimal("0.01")
GROUPS = ("Light Fixtures", "Wires and Consumables", "Labour Charges", "Other Charges")


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
    today = datetime.now(ZoneInfo("Asia/Kolkata")).date().isoformat()
    story = [paragraph(business["name"], "Title"),
             paragraph("Electrical Works Quotation", "Heading2"),
             paragraph("MOCK QUOTATION - POC ONLY - NOT FOR COMMERCIAL USE"),
             paragraph(f'{business["address"]}\n{business["phone"]} | {business["email"]}\n{business["registration"]}'),
             Spacer(1, 3 * mm),
             paragraph(f'Quotation: {request["quotation_id"]} | Date: {today} | Validity: {request["validity_days"]} days'),
             paragraph(f'Prepared for: {request["customer"]}\nSite: {request["site"]}'),
             paragraph("All amounts are in INR. Line amounts exclude tax; tax is shown separately in the summary.")]
    if any(line["category"] in GROUPS[2:] for line in result["lines"]):
        second_page_category = next(group for group in GROUPS[2:] if any(line["category"] == group for line in result["lines"]))
    else:
        second_page_category = None
    for category in GROUPS:
        lines = [line for line in result["lines"] if line["category"] == category]
        if not lines:
            continue
        if category == second_page_category:
            story.append(PageBreak())
        story.append(paragraph(category, "Heading2"))
        rows = [["Area / Description", "Specification", "Qty / Unit", "Rate INR", "Amount INR"]]
        for line in lines:
            description = f'{line["area"]}\n{line["description"]}\n{line["code"]}' if category in GROUPS[:2] else line["description"]
            rows.append([description, f'{line["specification"]}\nTax: {line["tax_percent"]}%',
                         f'{line["quantity"]:g} {line["unit"]}', f'{line["rate"]:,.2f}', f'{line["amount"]:,.2f}'])
        story.append(table(rows, [59 * mm, 42 * mm, 27 * mm, 24 * mm, 28 * mm], (2, 3, 4)))
        total = sum((line["amount"] for line in lines), Decimal("0.00"))
        story.append(paragraph(f"{category} subtotal: INR {total:,.2f}"))

    story.append(paragraph("Quotation Summary", "Heading2"))
    rows = [["Scope / Category", "Amount INR"]]
    for category in GROUPS:
        amount = sum((line["amount"] for line in result["lines"] if line["category"] == category), Decimal("0.00"))
        rows.append([category, f"{amount:,.2f}"])
    rows.extend([["Subtotal before tax", f'{result["subtotal"]:,.2f}'],
                 ["Tax (illustrative settings)", f'{result["tax"]:,.2f}'],
                 ["Grand Total", f'{result["total"]:,.2f}']])
    story.append(table(rows, [135 * mm, 45 * mm], (1,)))
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
    credentials = ("MICROSOFT_CLIENT_ID", "MICROSOFT_REFRESH_TOKEN") if oauth else ("SMTP_PASSWORD",)
    names = ("SMTP_HOST", "SMTP_USERNAME", "SMTP_FROM", "QUOTE_EMAIL_TO") + credentials
    missing = [name for name in names if not os.environ.get(name)]
    if missing:
        raise ValueError("Missing email settings: " + ", ".join(missing))
    sender = str(Address(addr_spec=os.environ["SMTP_FROM"]))
    recipient = str(Address(addr_spec=os.environ["QUOTE_EMAIL_TO"]))
    port = int(os.environ.get("SMTP_PORT", "587"))
    if not 1 <= port <= 65535:
        raise ValueError("SMTP_PORT must be between 1 and 65535")
    message = EmailMessage()
    message["Subject"] = f"Mock electrical quotation {quote_id}"
    message["From"] = sender
    message["To"] = recipient
    message.set_content("Attached is your mock electrical quotation. All prices and tax settings are illustrative.\n")
    message.add_attachment(path.read_bytes(), maintype="application", subtype="pdf", filename=path.name)
    context = ssl.create_default_context()
    token = refresh_access_token(os.environ["MICROSOFT_CLIENT_ID"], os.environ["MICROSOFT_REFRESH_TOKEN"]) if oauth else None
    client = smtplib.SMTP_SSL if port == 465 else smtplib.SMTP
    options = {"context": context} if port == 465 else {}
    with client(os.environ["SMTP_HOST"], port, timeout=30, **options) as smtp:
        if port != 465:
            smtp.starttls(context=context)
        if oauth:
            smtp.ehlo()
            auth = f'user={os.environ["SMTP_USERNAME"]}\x01auth=Bearer {token}\x01\x01'
            smtp.auth("XOAUTH2", lambda challenge=None: auth if challenge is None else "")
        else:
            smtp.login(os.environ["SMTP_USERNAME"], os.environ["SMTP_PASSWORD"])
        refused = smtp.send_message(message)
        if refused:
            raise smtplib.SMTPException("Recipient rejected")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--requirements", type=Path, default=ROOT / "examples/requirements.json")
    parser.add_argument("--send-email", action="store_true")
    args = parser.parse_args()
    try:
        raw = os.environ.get("QUOTATION_REQUIREMENTS_JSON", "").strip()
        request = json.loads(raw, parse_float=Decimal) if raw else load_json(args.requirements)
        result = calculate(request, load_json(ROOT / "data/catalogue.json"), load_json(ROOT / "data/charges.json"))
        output = ROOT / "output/pdf" / f'{request["quotation_id"]}.pdf'
        render_pdf(output, request, result, load_json(ROOT / "templates/quotation.json"))
        print("Mock quotation PDF generated successfully.")
        if args.send_email:
            send_email(output, request["quotation_id"])
            print("Email accepted by the sending server. Check the recipient inbox and spam folder.")
    except (ValueError, KeyError, TypeError, OSError, smtplib.SMTPException) as exc:
        # SMTP responses can contain private addresses; keep provider responses out of CI logs.
        if isinstance(exc, (smtplib.SMTPException, OSError)):
            parser.exit(1, f"Quotation failed ({type(exc).__name__}). Check files, connectivity and SMTP settings.\n")
        parser.exit(1, f"Quotation failed: {exc}\n")


if __name__ == "__main__":
    main()
