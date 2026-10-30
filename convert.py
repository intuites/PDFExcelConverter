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
from pathlib import Path

import pdfplumber
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import LongTable, PageBreak, Paragraph, SimpleDocTemplate, TableStyle

HEADER = ["Date", "Description", "Amount", "Balance"]
MONEY = r"-?\$?[\d,]+\.\d\d"
TXN = re.compile(rf"^(\d\d/\d\d)\s*(.*?)\s+({MONEY})\s+({MONEY})$")
LABELLED = re.compile(rf"^([A-Za-z][A-Za-z &]*?)\s+({MONEY})$")  # "Fees -4.95"
DETAIL_END = re.compile(r"^(\*end\*|Ending Balance|Page \d+ of \d+)")


def _money(s):
    return float(s.replace("$", "").replace(",", ""))


def parse_statement(src):
    """Read a text-based bank statement into info, summary, transactions and a balance check."""
    info, summary, rows = {}, {}, []
    in_detail, prefix = False, ""
    with pdfplumber.open(src) as pdf:
        for page in pdf.pages:
            for line in (page.extract_text() or "").splitlines():
                line = re.sub(r"\(cid:\d+\)", "", line).strip()  # glyphs with no unicode mapping
                if m := re.search(r"Account Number:\s*(\d+)", line):
                    info.setdefault("Account Number", m[1])
                if m := re.search(r"(\w+ \d\d?, \d{4}) through (\w+ \d\d?, \d{4})", line):
                    info.setdefault("Statement Period", f"{m[1]} - {m[2]}")
                if line.startswith("DATE DESCRIPTION"):
                    in_detail = True
                elif not in_detail:
                    if not rows and (m := LABELLED.match(line)):
                        summary.setdefault(m[1], _money(m[2]))
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

    # each balance must equal previous balance + amount; flags misread rows
    bal, bad = summary.get("Beginning Balance"), []
    for r in rows:
        if bal is not None and round(bal + r[2], 2) != r[3]:
            bad.append(f"{r[0]} {r[1]}")
        bal = r[3]
    ending = summary.get("Ending Balance")
    if rows and ending is not None and rows[-1][3] != ending:
        bad.append(f"Last balance {rows[-1][3]} != Ending Balance {ending}")
    return {"info": info, "summary": summary, "header": HEADER, "transactions": rows, "mismatches": bad}


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
    for r in [HEADER] + st["transactions"]:
        ws.append(r)
    for c in ws[1]:
        c.font = Font(bold=True)
    for col in "CD":
        for c in ws[col][1:]:
            c.number_format = "#,##0.00"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for col, w in zip("ABCD", (8, 70, 13, 13)):
        ws.column_dimensions[col].width = w

    sm = wb.create_sheet("Summary")
    for k, v in [*st["info"].items(), *st["summary"].items()]:
        sm.append([k, v])
    sm.append(["Balances reconcile", "; ".join(st["mismatches"]) or "Yes"])
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
                text = page.extract_text() or ""
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
    wb = load_workbook(src, read_only=True, data_only=True)  # data_only: formula results, not formulas
    style = getSampleStyleSheet()
    cell_style = style["BodyText"]
    cell_style.fontSize = 8
    cell_style.leading = 10
    story = []
    for ws in wb.worksheets:
        rows = [["" if v is None else str(v) for v in r] for r in ws.iter_rows(values_only=True)]
        rows = [r for r in rows if any(r)]
        if not rows:
            continue
        width = max(map(len, rows))
        rows = [[Paragraph(_esc(c), cell_style) for c in r + [""] * (width - len(r))] for r in rows]
        if story:
            story.append(PageBreak())
        story.append(Paragraph(_esc(ws.title), style["Heading2"]))
        t = LongTable(rows, repeatRows=1)  # header row repeats on every page
        t.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#DDE4EE")),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]))
        story.append(t)
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
}


def main():
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    src, dst = map(Path, sys.argv[1:])
    if not src.is_file():
        sys.exit(f"Not found: {src}")
    fn = CONVERTERS.get(src.suffix.lower(), {}).get(dst.suffix.lower().lstrip("."))
    if not fn:
        sys.exit("Use .pdf -> .xlsx/.csv/.json or .xlsx -> .pdf")
    fn(str(src), str(dst))
    print(f"Wrote {dst}")


if __name__ == "__main__":
    main()
