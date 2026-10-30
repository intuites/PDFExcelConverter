"""PDF <-> Excel converter.

    python convert.py statement.pdf out.xlsx   # bank statement -> Transactions + Summary sheets
    python convert.py statement.pdf out.csv    # bank statement -> Date,Description,Amount,Balance
    python convert.py statement.pdf out.json   # what the web preview shows
    python convert.py data.xlsx data.pdf       # each sheet -> a table in the PDF
PDFs that are not statements fall back to raw table extraction.
"""
import csv
import json
import re
import sys
from io import BytesIO
from pathlib import Path

import pdfplumber
import xlrd
from fontTools.ttLib import TTFont, TTLibError
from pdfminer.pdftypes import resolve1
from pdfminer.psparser import literal_name
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import LongTable, PageBreak, Paragraph, SimpleDocTemplate, TableStyle
from statement_sections import parse_sections

HEADER = ["Date", "Description", "Amount", "Balance"]
EXCEL_HEADER = ["Date", "Withdrawals", "Deposits", "Description"]
MONEY = r"-?\$?[\d,]+\.\d\d"
TXN = re.compile(rf"^(\d\d/\d\d)\s*(.*?)\s+({MONEY})\s+({MONEY})$")
LABELLED = re.compile(rf"^([A-Za-z][A-Za-z &]*?)\s+({MONEY})$")  # "Fees -4.95"
DETAIL_END = re.compile(r"^(\*end\*|Ending Balance|Page \d+ of \d+)")
CID = re.compile(r"\(cid:(\d+)\)")

# Match explicit statement terminology, not arbitrary words in transaction text.
COLUMN_ALIASES = {
    "date": {"date", "transaction date", "txn date", "tran date", "posting date", "posted date", "value date", "booking date", "trans date"},
    "description": {"description", "transaction description", "particulars", "narration", "details", "transaction details", "remarks", "transaction remarks", "memo", "reference", "merchant"},
    "amount": {"amount", "transaction amount", "txn amount", "net amount", "value", "amt"},
    "debit": {"debit", "debits", "debit amount", "withdrawal", "withdrawals", "withdrawal amount", "money out", "paid out", "dr"},
    "credit": {"credit", "credits", "credit amount", "deposit", "deposits", "deposit amount", "money in", "paid in", "cr"},
    "balance": {"balance", "running balance", "closing balance", "available balance", "account balance", "ledger balance", "bal"},
}
DATE_VALUE = re.compile(r"^(?:\d{4}[-/]\d{1,2}[-/]\d{1,2}|\d{1,2}[-/.]\d{1,2}(?:[-/.]\d{2,4})?|\d{1,2}[- /][A-Za-z]{3,9}(?:[- /]\d{2,4})?)$")


def _column_role(label):
    label = re.sub(r"\b(?:inr|usd|gbp|eur|rs)\b|[^a-z ]", " ", str(label or "").lower())
    label = " ".join(label.split())
    return next((role for role, aliases in COLUMN_ALIASES.items() if label in aliases), None)


def _header_columns(cells):
    columns = {}
    for index, cell in enumerate(cells):
        role = _column_role(cell)
        if role:
            columns.setdefault(role, index)
    if {"date", "description", "balance"} <= columns.keys() and (
            "amount" in columns or {"debit", "credit"} <= columns.keys()):
        return columns
    return None


def _statement_number(value):
    """Read currency, grouping, parentheses, and explicit CR/DR signs."""
    text = str(value or "").strip()
    if text in {"", "-", "--", "—"}:
        return None
    negative = bool(re.search(r"\bdr\b", text, re.I)) or text.startswith("(")
    text = re.sub(r"\b(?:inr|usd|gbp|eur|rs|cr|dr)\b|[$£€₹,\s()]", "", text, flags=re.I)
    if not re.fullmatch(r"[+-]?\d+(?:\.\d{1,2})?", text):
        raise ValueError("Not a statement amount")
    number = float(text)
    return -abs(number) if negative else number


def _mapped_transaction(cells, columns):
    def cell(role):
        index = columns.get(role)
        return str(cells[index] or "").strip() if index is not None and index < len(cells) else ""
    date = cell("date")
    if not DATE_VALUE.fullmatch(date):
        return None
    balance = _statement_number(cell("balance"))
    if "debit" in columns and "credit" in columns:
        debit, credit = _statement_number(cell("debit")), _statement_number(cell("credit"))
        if debit is None and credit is None:
            return None
        if debit and credit:
            raise ValueError("Both debit and credit contain values; review this row")
        amount = abs(credit or 0) - abs(debit or 0)
    else:
        amount = _statement_number(cell("amount"))
    if amount is None or balance is None:
        return None
    return [date, " ".join(cell("description").split()), amount, balance]


def _read_mapped_rows(records):
    rows, warnings, columns = [], [], None
    for cells in records:
        header = _header_columns(cells)
        if header:
            columns = header
            continue
        if columns is None:
            continue
        date = str(cells[columns["date"]] or "").strip() if len(cells) > columns["date"] else ""
        if DATE_VALUE.fullmatch(date):
            try:
                transaction = _mapped_transaction(cells, columns)
            except ValueError:
                transaction = None
            if transaction:
                rows.append(transaction)
            else:
                warnings.append(f"Transaction dated {date} could not be mapped safely; review the original PDF.")
        elif rows and not date:
            description = str(cells[columns["description"]] or "").strip() if len(cells) > columns["description"] else ""
            # A continuation has only description text; totals/footer values
            # must not be appended to the preceding transaction.
            if description and all(not str(value or "").strip() for i, value in enumerate(cells) if i != columns["description"]):
                rows[-1][1] += " " + " ".join(description.split())
    return rows, warnings


def _positioned_records(page):
    """Use word coordinates to preserve empty debit/credit columns without borders."""
    lines = []
    for word in sorted(page.extract_words(), key=lambda w: (w["top"], w["x0"])):
        if not lines or abs(lines[-1][0]["top"] - word["top"]) > 3:
            lines.append([word])
        else:
            lines[-1].append(word)
    headings = None
    for line in lines:
        words = sorted(line, key=lambda w: w["x0"])
        candidate = []
        index = 0
        while index < len(words):
            found = False
            for count in (3, 2, 1):
                group = words[index:index + count]
                if len(group) == count and _column_role(" ".join(w["text"] for w in group)):
                    candidate.append((" ".join(w["text"] for w in group), group[0]["x0"], group[-1]["x1"]))
                    index += count
                    found = True
                    break
            if not found:
                index += 1
        if _header_columns([item[0] for item in candidate]):
            headings = candidate
            yield [item[0] for item in headings]
            continue
        if headings:
            # Boundaries use the gap between headings, so a right-aligned
            # number can begin to the left of its own heading.
            boundaries = [(a[2] + b[1]) / 2 for a, b in zip(headings, headings[1:])]
            cells = [[] for _ in headings]
            for word in words:
                center = (word["x0"] + word["x1"]) / 2
                index = sum(center >= boundary for boundary in boundaries)
                cells[index].append(word["text"])
            yield [" ".join(cell) for cell in cells]


def _alternative_transactions(src):
    rows, warnings = [], []
    with pdfplumber.open(src) as pdf:
        for page_number, page in enumerate(pdf.pages, 1):
            _page_text(page)  # repair embedded glyphs before either strategy
            tables = page.extract_tables()
            page_rows, page_warnings = [], []
            for table in tables:
                extracted, problems = _read_mapped_rows(table)
                page_rows.extend(extracted)
                page_warnings.extend(problems)
            if not page_rows:
                page_rows, page_warnings = _read_mapped_rows(_positioned_records(page))
            rows.extend(page_rows)
            warnings.extend(f"Page {page_number}: {warning}" for warning in page_warnings)
            page.flush_cache()
    return rows, warnings


def _embedded_glyph_map(font):
    """Recover Unicode from an embedded TrueType cmap when PDF ToUnicode is incomplete.

    Only Identity-H/Identity CID mappings are supported here: other encodings
    require a different CID-to-glyph translation and must not be guessed.
    """
    if literal_name(font.get("Encoding")) != "Identity-H":
        return {}
    descendants = resolve1(font.get("DescendantFonts", []))
    if not descendants:
        return {}
    descendant = resolve1(descendants[0])
    if literal_name(descendant.get("CIDToGIDMap")) != "Identity":
        return {}
    descriptor = resolve1(descendant.get("FontDescriptor", {}))
    stream = resolve1(descriptor.get("FontFile2"))
    if stream is None:
        return {}
    with TTFont(BytesIO(stream.get_data())) as embedded:
        # Some fonts map several codepoints to the same glyph. Prefer the
        # lowest codepoint, e.g. a normal space rather than a nonbreaking space.
        result = {}
        for codepoint, name in sorted((embedded.getBestCmap() or {}).items()):
            if codepoint >= 32:
                result.setdefault(embedded.getGlyphID(name), chr(codepoint))
        return result


def _page_text(page):
    """Extract text after repairing only glyphs with a verified font mapping."""
    missing = {c["fontname"] for c in page.chars if CID.fullmatch(c["text"])}
    if missing:
        resources = resolve1(page.page_obj.resources)
        fonts = resolve1(resources.get("Font", {}))
        mappings = {}
        for reference in fonts.values():
            font = resolve1(reference)
            name = literal_name(font.get("BaseFont"))
            if name in missing:
                try:
                    mappings[name] = _embedded_glyph_map(font)
                except (TTLibError, ValueError, KeyError, OSError):
                    mappings[name] = {}
        for char in page.chars:
            match = CID.fullmatch(char["text"])
            if match:
                char["text"] = mappings.get(char["fontname"], {}).get(
                    int(match[1]), char["text"])
    # Chase uses tiny white indexing text such as *end* inside visible rows.
    # It is not statement content and can split a date during extraction.
    def visible(obj):
        return not (obj.get("object_type") == "char" and obj.get("size", 10) <= 1
                    and obj.get("non_stroking_color") in ((1,), (1, 1, 1)))
    if any(not visible(char) for char in page.chars):
        return page.filter(visible).extract_text() or ""
    return page.extract_text() or ""


def _money(s):
    return float(s.replace("$", "").replace(",", ""))


def parse_statement(src):
    """Read a text-based bank statement into info, summary, transactions and a balance check."""
    info, summary, rows, warnings = {}, {}, [], []
    in_detail, prefix = False, ""
    page_texts = []
    balance_source, daily_balances, section_issues = "printed", {}, []
    balance_note = ""
    with pdfplumber.open(src) as pdf:
        for page_number, page in enumerate(pdf.pages, 1):
            text = _page_text(page)
            page_texts.append(text)
            if CID.search(text):
                warnings.append(f"Page {page_number}: some characters could not be decoded; review descriptions containing (cid:...).")
            for line in text.splitlines():
                line = line.strip()
                if m := re.search(r"(?:account\s*(?:number|no\.?|#)|a/c\s*(?:number|no\.?))\s*:?\s*([\dXx*]+)", line, re.I):
                    info.setdefault("Account Number", m[1])
                if m := re.search(r"(\w+ \d\d?, \d{4}) through (\w+ \d\d?, \d{4})", line):
                    info.setdefault("Statement Period", f"{m[1]} - {m[2]}")
                if re.fullmatch(r"DATE\s+DESCRIPTION\s+AMOUNT\s+BALANCE", line):
                    in_detail = True
                elif not in_detail:
                    if not rows and (m := LABELLED.match(line)):
                        label = {"opening balance": "Beginning Balance", "beginning balance": "Beginning Balance", "closing balance": "Ending Balance", "ending balance": "Ending Balance"}.get(m[1].lower(), m[1])
                        summary.setdefault(label, _money(m[2]))
                elif DETAIL_END.match(line):
                    in_detail = False
                elif m := TXN.match(line):
                    rows.append([m[1], f"{prefix} {m[2]}".strip(), _money(m[3]), _money(m[4])])
                    prefix = ""
                elif m := re.match(r"^\d\d/\d\d\s*(.*)$", line):
                    prefix = m[1]  # description wrapped above its amounts, e.g. "02/28Online"
                elif rows and not LABELLED.match(line):
                    rows[-1][1] += " " + line  # wrapped tail, e.g. card number "1425"
            page.flush_cache()

    if not rows:
        rows, alternative_warnings = _alternative_transactions(src)
        warnings.extend(alternative_warnings)
    if not rows:
        sectioned = parse_sections(page_texts)
        if sectioned:
            info.update(sectioned["info"])
            summary = sectioned["summary"]
            rows = sectioned["transactions"]
            warnings.extend(sectioned["warnings"])
            daily_balances = sectioned["daily_balances"]
            section_issues = sectioned["issues"]
            balance_source = "calculated"
            balance_note = sectioned["balance_note"]

    # each balance must equal previous balance + amount; flags misread rows
    bal, bad = summary.get("Beginning Balance"), []
    for r in rows:
        if bal is not None and round(bal + r[2], 2) != r[3]:
            bad.append(f"{r[0]} {r[1]}")
        bal = r[3]
    ending = summary.get("Ending Balance")
    if rows and ending is not None and rows[-1][3] != ending:
        bad.append(f"Last balance {rows[-1][3]} != Ending Balance {ending}")
    last_by_date = {row[0]: row[3] for row in rows}
    for date, expected in daily_balances.items():
        if date not in last_by_date or last_by_date[date] != expected:
            bad.append(f"Daily balance {date}: calculated {last_by_date.get(date)} != printed {expected}")
    bad.extend(section_issues)
    return {"info": info, "summary": summary, "header": HEADER, "transactions": rows, "mismatches": bad, "warnings": warnings, "balance_source": balance_source, "balance_note": balance_note}


def pdf_to_json(src, dst):
    Path(dst).write_text(json.dumps(parse_statement(src)), encoding="utf-8")


def pdf_to_csv(src, dst):
    st = parse_statement(src)
    with open(dst, "w", newline="", encoding="utf-8-sig") as f:  # BOM so Excel reads "…" correctly
        w = csv.writer(f)
        if st["transactions"]:
            w.writerows([HEADER] + st["transactions"])
        else:
            for _, _, table in _raw_tables(src):
                w.writerows([(c or "").strip() for c in row] for row in table)


def pdf_to_excel(src, dst):
    st = parse_statement(src)
    if not st["transactions"]:
        return _tables_to_excel(src, dst)
    wb = Workbook()
    ws = wb.active
    ws.title = "Transactions"
    ws.append(EXCEL_HEADER)
    # Keep the export aligned with the supplied workbook template. These rows
    # act as the visible category guide beneath the two amount columns.
    ws.append(["", "checks paid", "Deposits and additions", ""])
    ws.append(["", "electronic withdrawal", "", ""])
    ws.append(["", "fees", "", ""])
    for date, description, amount, _balance in st["transactions"]:
        ws.append([date, abs(amount) if amount < 0 else "", amount if amount > 0 else "", description])
    for c in ws[1]:
        c.font = Font(bold=True)
    for col in "BC":
        for c in ws[col][1:]:
            c.number_format = "#,##0.00"
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = ws.dimensions
    for col, w in zip("ABCD", (12, 20, 22, 70)):
        ws.column_dimensions[col].width = w

    sm = wb.create_sheet("Summary")
    for k, v in [*st["info"].items(), *st["summary"].items()]:
        sm.append([k, v])
    sm.append(["Balances reconcile", "; ".join(st["mismatches"]) or "Yes"])
    sm.append(["Balance source", st["balance_source"]])
    if st["balance_note"]:
        sm.append(["Balance calculation", st["balance_note"]])
    for warning in st["warnings"]:
        sm.append(["Text extraction warning", warning])
    for c in sm["A"]:
        c.font = Font(bold=True)
    for c in sm["B"]:
        c.number_format = "#,##0.00"
    sm.column_dimensions["A"].width, sm.column_dimensions["B"].width = 28, 40
    wb.save(dst)


def _raw_tables(src):
    with pdfplumber.open(src) as pdf:
        for pno, page in enumerate(pdf.pages, 1):
            tables = page.extract_tables()
            # shortcut: pages without ruled tables become one text line per row; use table_settings={"vertical_strategy": "text"} for borderless tables
            if not tables:
                text = _page_text(page)
                tables = [[[line] for line in text.splitlines()]] if text.strip() else []
            for tno, table in enumerate(tables, 1):
                yield pno, tno, table
            page.flush_cache()  # free per-page objects so long PDFs don't grow memory


def _tables_to_excel(src, dst):
    wb = Workbook(write_only=True)  # write_only streams rows: low memory on big files
    found = False
    for pno, tno, table in _raw_tables(src):
        ws = wb.create_sheet(f"P{pno}_T{tno}"[:31])
        for i, row in enumerate(table):
            cells = [(c or "").strip() for c in row]
            ws.append([_num(c) for c in cells] if i else cells)
        found = True
    if not found:
        wb.create_sheet("Empty")
    wb.save(dst)


def _num(s):
    try:
        return float(s.replace(",", "")) if "." in s or "," in s else int(s)
    except ValueError:
        return s


def excel_to_pdf(src, dst):
    if Path(src).suffix.lower() == ".xls":
        workbook = xlrd.open_workbook(src, on_demand=True)
        sheets = [(sheet.name, sheet.get_rows()) for sheet in workbook.sheets()]
    else:
        wb = load_workbook(src, read_only=True, data_only=True)  # data_only: formula results, not formulas
        sheets = [(ws.title, ws.iter_rows(values_only=True)) for ws in wb.worksheets]
    style = getSampleStyleSheet()
    cell_style = style["BodyText"]
    cell_style.fontSize = 8
    cell_style.leading = 10
    story = []
    for sheet_name, sheet_rows in sheets:
        rows = []
        for row in sheet_rows:
            values = [cell.value for cell in row] if Path(src).suffix.lower() == ".xls" else row
            rows.append(["" if v is None else str(v) for v in values])
        rows = [r for r in rows if any(r)]
        if not rows:
            continue
        width = max(map(len, rows))
        rows = [[Paragraph(_esc(c), cell_style) for c in r + [""] * (width - len(r))] for r in rows]
        if story:
            story.append(PageBreak())
        story.append(Paragraph(_esc(sheet_name), style["Heading2"]))
        t = LongTable(rows, repeatRows=1)  # header row repeats on every page
        t.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#DDE4EE")),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]))
        story.append(t)
    if Path(src).suffix.lower() == ".xls":
        workbook.release_resources()
    else:
        wb.close()
    if not story:
        story = [Paragraph("Workbook is empty", style["BodyText"])]
    SimpleDocTemplate(dst, pagesize=landscape(A4), leftMargin=20, rightMargin=20,
                      topMargin=20, bottomMargin=20).build(story)


def _esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


# input extension -> {output format: converter}; first format is the default
CONVERTERS = {
    ".pdf": {"xlsx": pdf_to_excel, "csv": pdf_to_csv, "json": pdf_to_json},
    ".xlsx": {"pdf": excel_to_pdf},
    ".xlsm": {"pdf": excel_to_pdf},
    ".xls": {"pdf": excel_to_pdf},
}


def main():
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    src, dst = map(Path, sys.argv[1:])
    if not src.is_file():
        sys.exit(f"Not found: {src}")
    fn = CONVERTERS.get(src.suffix.lower(), {}).get(dst.suffix.lower().lstrip("."))
    if not fn:
        sys.exit("Use .pdf -> .xlsx/.csv/.json or .xlsx/.xls/.xlsm -> .pdf")
    fn(str(src), str(dst))
    print(f"Wrote {dst}")


if __name__ == "__main__":
    main()
