"""Regression checks for alternate headings and real PDF column layouts."""
import tempfile
import unittest
from pathlib import Path

from reportlab.pdfgen import canvas
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle

from convert import _read_mapped_rows, _statement_number, parse_statement


class AlternativeStatementTests(unittest.TestCase):
    def test_alternative_headings_and_dates(self):
        for date_label in ("Txn Date", "Posting Date", "Value Date", "Transaction Date"):
            for description_label in ("Narration", "Particulars", "Remarks", "Transaction Details"):
                rows, warnings = _read_mapped_rows([
                    [date_label, description_label, "Transaction Amount (INR)", "Running Balance"],
                    ["2025-01-02", "Shop", "(10.00)", "90.00 CR"],
                    ["03-Jan-2025", "Payroll", "INR 50.00 CR", "140.00"],
                ])
                self.assertEqual(rows, [["2025-01-02", "Shop", -10.0, 90.0],
                                        ["03-Jan-2025", "Payroll", 50.0, 140.0]])
                self.assertFalse(warnings)

    def test_empty_columns_and_continuations(self):
        rows, warnings = _read_mapped_rows([
            ["Date", "Narration", "Withdrawal", "Deposit", "Closing Balance"],
            ["01/02/2025", "Card purchase", "10.00", "", "90.00"],
            ["", "reference 001234", "", "", ""],
            ["02/02/2025", "Payroll", "-", "50.00", "140.00"],
            ["", "Total", "10.00", "50.00", "140.00"],
        ])
        self.assertEqual(rows[0], ["01/02/2025", "Card purchase reference 001234", -10.0, 90.0])
        self.assertEqual(rows[1][2:], [50.0, 140.0])
        self.assertFalse(warnings)

    def test_ambiguous_or_missing_values_are_flagged(self):
        rows, warnings = _read_mapped_rows([
            ["Date", "Details", "Money Out", "Money In", "Balance"],
            ["01/02", "Ambiguous", "10.00", "20.00", "110.00"],
            ["01/03", "Missing balance", "10.00", "", ""],
        ])
        self.assertFalse(rows)
        self.assertEqual(len(warnings), 2)

    def test_reference_columns_not_mistaken_for_amounts(self):
        rows, warnings = _read_mapped_rows([
            ["Posting Date", "Ref No", "Particulars", "Paid Out", "Paid In", "Ledger Balance"],
            ["03.01.2025", "123456789", "Salary", "", "2,000.00", "2,100.00"],
        ])
        self.assertEqual(rows, [["03.01.2025", "Salary", 2000.0, 2100.0]])
        self.assertFalse(warnings)

    def test_currency_signs(self):
        for value in ("INR 1,234.56 Dr", "(1,234.56)", "-1,234.56", "Rs 1,234.56 DR"):
            self.assertEqual(_statement_number(value), -1234.56)
        self.assertEqual(_statement_number("0.00"), 0)

    def test_bordered_pdf(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "statement.pdf"
            table = Table([
                ["Value Date", "Particulars", "Debit Amount", "Credit Amount", "Running Balance"],
                ["01-01-2025", "Shop", "10.00", "", "90.00"],
                ["02-01-2025", "Payroll", "", "50.00", "140.00"],
            ])
            table.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), .5, "black")]))
            SimpleDocTemplate(str(path)).build([table])
            statement = parse_statement(path)
            self.assertEqual(len(statement["transactions"]), 2)
            self.assertEqual(statement["transactions"][0][2], -10.0)
            self.assertFalse(statement["mismatches"])

    def test_legacy_heading_with_split_debit_credit(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "statement.pdf"
            pdf = canvas.Canvas(str(path), pagesize=(700, 850))
            pdf.setFont("Helvetica", 10)
            for y, row in [(750, ["DATE", "DESCRIPTION", "DEBIT", "CREDIT", "BALANCE"]),
                           (730, ["01/01", "Shop", "10.00", "", "90.00"])]:
                for x, text in zip([40, 160, 350, 450, 550], row):
                    pdf.drawString(x, y, text)
            pdf.save()
            statement = parse_statement(path)
            self.assertEqual(statement["transactions"], [["01/01", "Shop", -10.0, 90.0]])

    def test_borderless_pdf(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "statement.pdf"
            pdf = canvas.Canvas(str(path), pagesize=(700, 850))
            pdf.setFont("Helvetica", 10)
            pdf.drawString(40, 810, "Opening Balance 100.00")
            pdf.drawString(40, 790, "Closing Balance 140.00")
            columns = [40, 160, 350, 450, 550]
            for y, row in [(750, ["Txn Date", "Narration", "Money Out", "Money In", "Balance"]),
                           (730, ["01/01/2025", "Shop", "10.00", "", "90.00"]),
                           (710, ["02/01/2025", "Salary", "", "50.00", "140.00"])]:
                for x, text in zip(columns, row):
                    pdf.drawString(x, y, text)
            pdf.save()
            statement = parse_statement(path)
            self.assertEqual(statement["transactions"], [
                ["01/01/2025", "Shop", -10.0, 90.0],
                ["02/01/2025", "Salary", 50.0, 140.0]])
            self.assertEqual(statement["summary"]["Beginning Balance"], 100.0)
            self.assertFalse(statement["mismatches"])


if __name__ == "__main__":
    unittest.main()
