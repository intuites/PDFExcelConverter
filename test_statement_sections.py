"""Regression coverage for section-based statements without row balances."""
import tempfile
import unittest
from pathlib import Path

from reportlab.pdfgen import canvas
from convert import parse_statement


LINES = [
    "January 31, 2026throughFebruary 27, 2026",
    "Account Number:", "000000123456789",
    "CHECKING SUMMARY", "INSTANCES AMOUNT",
    "Beginning Balance $100.00",
    "Deposits and Additions 1 50.00",
    "Checks Paid 1 -20.00",
    "Electronic Withdrawals 1 -10.00",
    "Fees 1 -5.00",
    "Ending Balance 4 $115.00",
    "DEPOSITS AND ADDITIONS", "DATE DESCRIPTION AMOUNT",
    "02/02 Payroll $50.00", "Reference 001234",
    "Total Deposits and Additions $50.00",
    "CHECKS PAID", "DATE", "CHECK NO. DESCRIPTION PAID AMOUNT",
    "10 ^ 02/01 $20.00", "Total Checks Paid $20.00",
    "ELECTRONIC WITHDRAWALS", "DATE DESCRIPTION AMOUNT",
    "02/03 Card purchase $10.00", "Reference 005678",
    "Total Electronic Withdrawals $10.00",
    "FEES", "DATE DESCRIPTION AMOUNT",
    "02/04 Monthly fee $5.00", "Total Fees $5.00",
    "DAILY ENDING BALANCE", "DATE AMOUNT",
    "02/01 $80.00", "02/02 130.00", "02/03 120.00", "02/04 115.00",
]


def parse_lines(lines):
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "sections.pdf"
        pdf = canvas.Canvas(str(path), pagesize=(620, 1000))
        pdf.setFont("Helvetica", 10)
        for index, line in enumerate(lines):
            pdf.drawString(30, 970 - index * 18, line)
        pdf.save()
        return parse_statement(path)


class SectionStatementTests(unittest.TestCase):
    def test_transactions_chronology_signs_and_daily_validation(self):
        statement = parse_lines(LINES)
        self.assertEqual(statement["transactions"], [
            ["02/01", "Check #10", -20.0, 80.0],
            ["02/02", "Payroll Reference 001234", 50.0, 130.0],
            ["02/03", "Card purchase Reference 005678", -10.0, 120.0],
            ["02/04", "Monthly fee", -5.0, 115.0],
        ])
        self.assertFalse(statement["mismatches"])
        self.assertEqual(statement["balance_source"], "calculated")
        self.assertIn("intraday", statement["balance_note"])
        self.assertEqual(statement["info"]["Account Number"], "000000123456789")

    def test_daily_balance_disagreement_is_not_silently_accepted(self):
        statement = parse_lines([line.replace("02/03 120.00", "02/03 121.00") for line in LINES])
        self.assertTrue(any("Daily balance 02/03" in error for error in statement["mismatches"]))

    def test_summary_count_disagreement_is_flagged(self):
        statement = parse_lines([line.replace("Deposits and Additions 1 50.00", "Deposits and Additions 2 50.00") for line in LINES])
        self.assertTrue(any("transaction count" in error for error in statement["mismatches"]))


if __name__ == "__main__":
    unittest.main()
