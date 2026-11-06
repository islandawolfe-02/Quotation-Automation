import io
import json
import os
import shutil
import tempfile
import unittest
from decimal import Decimal
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from quotation import EXCEL_COLUMNS, ROOT, calculate, load_excel_inputs, load_json, main


class ExcelInputTests(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.master = Path(directory.name) / "master.xlsx"
        self.client = Path(directory.name) / "client.xlsx"
        self.master_fixture = Path(directory.name) / "master_fixture.xlsx"
        self.client_fixture = Path(directory.name) / "client_fixture.xlsx"
        fixture = load_json(ROOT / "tests/fixtures/quotation.json")
        materials = fixture["catalogue"]
        charges = fixture["charges"]
        quantities = {item["code"]: item["quantity"] for kind in ("materials", "charges")
                      for item in fixture["request"][kind]}
        rows = []
        for item in materials + charges:
            code = "MCB-16A-C-1P" if item["code"] == "MCB-16A" else item["code"]
            rows.append([code, "MCB" if item["code"] == "MCB-16A" else item["description"], 1,
                         "16A C-curve" if item["code"] == "MCB-16A" else item["specification"],
                         "1P" if item["code"] == "MCB-16A" else "N/A", item["unit"],
                         float(item["rate"]), float(item["tax_percent"]), item["category"], "Demo", "General"])
        variants = [("MCB-6A-B-1P", "MCB", "6A B-curve", "1P", 220, 2),
                    ("MCB-6A-B-2P", "MCB", "6A B-curve", "2P", 480, 0),
                    ("RCCB-40A-30MA-2P", "RCCB", "40A 30mA", "2P", 1700, 1),
                    ("RCCB-63A-300MA-4P", "RCCB", "63A 300mA", "4P", 3600, 1),
                    ("SPD-T2-4P", "SPD", "Type 2, 275V AC, 20kA", "4P", 2800, 1)]
        for code, name, specs, poles, rate, quantity in variants:
            rows.append([code, name, 1, specs, poles, "piece", rate, 0, "Wires and Consumables", "Demo", "General"])
            quantities[code] = quantity
        quantities["MCB-16A-C-1P"] = quantities.pop("MCB-16A")
        for path, title, header in ((self.master_fixture, "Catalogue", 5), (self.client_fixture, "Requirements", 12)):
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = title
            if title == "Requirements":
                metadata = [("Client Name", "Demo Community Centre"), ("Address", "Sample client address, Shillong"),
                            ("Contact", "+91 00000 00000"), ("Site", "Sample project site, Shillong"),
                            ("Quotation ID", "POC-EXCEL-001"), ("Validity Days", 15)]
                for position, (label, value) in enumerate(metadata, 4):
                    sheet.cell(position, 1, label)
                    sheet.cell(position, 2, value)
            for column, value in enumerate(EXCEL_COLUMNS, 1):
                sheet.cell(header, column, value)
            for position, source in enumerate(rows, header + 1):
                row = list(source)
                if title == "Requirements":
                    row[2] = float(quantities.get(row[0], 0))
                for column, value in enumerate(row, 1):
                    sheet.cell(position, column, value)
            workbook.save(path)
            workbook.close()
        shutil.copyfile(self.master_fixture, self.master)
        shutil.copyfile(self.client_fixture, self.client)

    def edit(self, path, sheet, cell, value):
        workbook = load_workbook(path)
        try:
            workbook[sheet][cell] = value
            workbook.save(path)
        finally:
            workbook.close()

    def load(self):
        return load_excel_inputs(self.master, self.client)

    def total(self):
        return calculate(*self.load())["total"]

    def test_sample_client_details_and_specification_variants(self):
        request, materials, charges = self.load()
        self.assertEqual(request["customer"], "Demo Community Centre")
        self.assertEqual(request["customer_address"], "Sample client address, Shillong")
        self.assertEqual(request["customer_contact"], "+91 00000 00000")
        self.assertEqual(len(materials), 13)
        self.assertEqual(len(charges), 6)
        self.assertEqual(calculate(request, materials, charges)["total"], Decimal("50940.00"))
        lookup = {item["code"]: item for item in materials}
        self.assertIn("6A B-curve; 1P", lookup["MCB-6A-B-1P"]["specification"])
        self.assertIn("6A B-curve; 2P", lookup["MCB-6A-B-2P"]["specification"])
        self.assertIn("63A 300mA; 4P", lookup["RCCB-63A-300MA-4P"]["specification"])

    def test_master_rate_is_authoritative(self):
        self.edit(self.master, "Catalogue", "G6", 1300)
        self.edit(self.client, "Requirements", "G13", 999999)
        self.edit(self.client, "Requirements", "H13", 99)
        self.assertEqual(self.total(), Decimal("51740.00"))

    def test_onedrive_cli_uses_saved_client_and_latest_master_rates(self):
        for iteration, total in enumerate(("50940.00", "51740.00")):
            if iteration:
                self.edit(self.master, "Catalogue", "G6", 1300)
            self.edit(self.client, "Requirements", "B4", f"Saved client {iteration}")
            metadata = json.dumps({"@microsoft.graph.downloadUrl": "https://files.1drv.com/download"}).encode()
            responses = [io.BytesIO(metadata), io.BytesIO(self.master.read_bytes()),
                         io.BytesIO(metadata), io.BytesIO(self.client.read_bytes())]
            with self.subTest(iteration=iteration), \
                    patch.dict(os.environ, {"MICROSOFT_CLIENT_ID": "client", "MICROSOFT_REFRESH_TOKEN": "refresh"}, clear=True), \
                    patch("sys.argv", ["quotation.py", "--onedrive", "--send-email",
                                       "--requirements", "Quotations/saved-client.xlsx",
                                       "--catalogue", "Quotations/master.xlsx"]), \
                    patch("quotation.recipient_addresses"), patch("quotation.refresh_access_token", return_value="access"), \
                    patch("quotation.urlopen", side_effect=responses) as open_url, \
                    patch("quotation.render_pdf") as render, patch("quotation.send_email") as send, patch("builtins.print"):
                main()
                self.assertEqual(render.call_args.args[1]["customer"], f"Saved client {iteration}")
                self.assertEqual(render.call_args.args[2]["total"], Decimal(total))
                self.assertIn("Quotations/master.xlsx", open_url.call_args_list[0].args[0].full_url)
                self.assertIn("Quotations/saved-client.xlsx", open_url.call_args_list[2].args[0].full_url)
                send.assert_called_once_with(ROOT / "output/pdf/POC-EXCEL-001.pdf", "POC-EXCEL-001")

    def test_blank_and_zero_quantities_are_omitted(self):
        self.edit(self.client, "Requirements", "C13", None)
        self.edit(self.client, "Requirements", "C14", 0)
        self.assertEqual(self.total(), Decimal("36240.00"))

    def test_wrong_specs_poles_or_brand_are_rejected(self):
        for cell, value in (("D19", "6A B-curve"), ("E19", "4P"), ("J19", "Different brand")):
            shutil.copyfile(self.client_fixture, self.client)
            self.edit(self.client, "Requirements", cell, value)
            with self.subTest(cell=cell), self.assertRaisesRegex(ValueError, "does not match the master"):
                self.load()

    def test_unknown_and_duplicate_codes_are_rejected(self):
        self.edit(self.client, "Requirements", "A13", "UNKNOWN")
        with self.assertRaisesRegex(ValueError, "Unknown Item Code"):
            self.load()
        self.edit(self.master, "Catalogue", "A7", "LED-LINEAR-25")
        with self.assertRaisesRegex(ValueError, "Duplicate Item Code"):
            self.load()

    def test_missing_contact_and_invalid_quantities_are_rejected(self):
        self.edit(self.client, "Requirements", "B6", None)
        with self.assertRaisesRegex(ValueError, "Requirements!B6"):
            self.load()
        for quantity in (-1, True, "NaN"):
            shutil.copyfile(self.client_fixture, self.client)
            self.edit(self.client, "Requirements", "C13", quantity)
            with self.subTest(quantity=quantity), self.assertRaises(ValueError):
                self.load()

    def test_formula_error_and_missing_headers_are_rejected(self):
        for cell, value in (("G6", "=1200+50"), ("G6", "#REF!"), ("A5", "Wrong Header"), ("C6", 2), ("E12", "N/A")):
            shutil.copyfile(self.master_fixture, self.master)
            self.edit(self.master, "Catalogue", cell, value)
            with self.subTest(cell=cell, value=value), self.assertRaises(ValueError):
                self.total()

    def test_new_component_rows_are_read_without_table_range_updates(self):
        master = load_workbook(self.master)
        client = load_workbook(self.client)
        try:
            row = [cell.value for cell in master["Catalogue"][6]]
            row[0:2] = ["LED-NEW", "New demo light"]
            row[6] = 1600
            master["Catalogue"].append(row)
            row[2] = 2
            client["Requirements"].append(row)
            master.save(self.master)
            client.save(self.client)
        finally:
            master.close()
            client.close()
        self.assertEqual(self.total(), Decimal("54140.00"))


if __name__ == "__main__":
    unittest.main()
