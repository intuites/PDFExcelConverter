"""Read statements that list credits/debits separately from daily balances."""
import re
from decimal import Decimal

MONEY = r"-?\$?[\d,]+\.\d{2}"
SECTIONS = {
    "DEPOSITS AND ADDITIONS": ("Deposits and Additions", 1),
    "CHECKS PAID": ("Checks Paid", -1),
    "ELECTRONIC WITHDRAWALS": ("Electronic Withdrawals", -1),
    "FEES": ("Fees", -1),
}


def number(text):
    return Decimal(text.replace("$", "").replace(",", ""))


def parse_sections(pages):
    text = "\n".join(pages)
    if not all(heading in text for heading in ("CHECKING SUMMARY", "DEPOSITS AND ADDITIONS")):
        return None
    summary, info, counts, daily = {}, {}, {}, {}
    period = re.search(r"(\w+\s+\d{1,2},\s*\d{4})\s*through\s*(\w+\s+\d{1,2},\s*\d{4})", text)
    start_month = 1
    if period:
        from datetime import datetime
        start, end = (" ".join(value.split()) for value in period.groups())
        info["Statement Period"] = f"{start} - {end}"
        start_month = datetime.strptime(start, "%B %d, %Y").month
    # Account values are sometimes on a separate line, or above the label.
    account = re.search(r"Account Number:\s*\n(\d{8,})\b", text)
    if account:
        info["Account Number"] = account[1]

    labels = ["Beginning Balance", *[value[0] for value in SECTIONS.values()], "Ending Balance"]
    for label in labels:
        match = re.search(rf"^{re.escape(label)}\s+(?:(\d+)\s+)?({MONEY})$", text, re.M)
        if match:
            summary[label] = number(match[2])
            if match[1]:
                counts[label] = int(match[1])
    if "Beginning Balance" not in summary:
        return None

    records, warnings, issues = [], [], []
    section, reading, last = None, False, None
    totals = {}
    for page in pages:
        # Do not attach page headers to the last description on the prior page.
        section, reading, last = None, False, None
        for line in page.splitlines():
            line = " ".join(line.split())
            if line in SECTIONS or line == "DAILY ENDING BALANCE":
                section, reading, last = line, False, None
                continue
            if not section:
                continue
            if line in {"DATE DESCRIPTION AMOUNT", "CHECK NO. DESCRIPTION PAID AMOUNT", "DATE AMOUNT"}:
                reading = True
                continue
            if not reading:
                continue
            if line.startswith("Total "):
                total = re.search(rf"({MONEY})$", line)
                if total:
                    totals[section] = number(total[1])
                section, reading, last = None, False, None
                continue
            if section == "DAILY ENDING BALANCE":
                match = re.fullmatch(rf"(\d{{2}}/\d{{2}})\s+({MONEY})", line)
                if match:
                    daily[match[1]] = number(match[2])
                else:
                    section, reading = None, False
                continue
            if section == "CHECKS PAID":
                match = re.fullmatch(rf"(\d+)\s*\^?\s+(.*?)\b(\d{{2}}/\d{{2}})\s+({MONEY})", line)
                if match:
                    check, description, date, amount = match.groups()
                    last = [date, f"Check #{check}" + (f" {description.strip()}" if description.strip() else ""), -abs(number(amount)), section]
                    records.append(last)
                continue
            match = re.fullmatch(rf"(\d{{2}}/\d{{2}})\s+(.*?)\s+({MONEY})", line)
            if match:
                date, description, amount = match.groups()
                last = [date, description, abs(number(amount)) * SECTIONS[section][1], section]
                records.append(last)
            elif re.match(r"^\d{2}/\d{2}\b", line):
                warnings.append(f"A transaction in {section.lower()} could not be read: {line[:5]}.")
                last = None
            elif line.startswith(("Page of", "Page ", "Account Number:")):
                reading, last = False, None
            elif last:
                last[1] += " " + line

    if not records:
        return None
    for heading, (label, _) in SECTIONS.items():
        actual = sum((abs(row[2]) for row in records if row[3] == heading), Decimal(0))
        expected = summary.get(label)
        if expected is not None and actual != abs(expected):
            issues.append(f"{label}: extracted total {actual} != printed total {abs(expected)}")
        if heading in totals and actual != abs(totals[heading]):
            issues.append(f"{label}: extracted total {actual} != section total {abs(totals[heading])}")
        if label in counts and sum(row[3] == heading for row in records) != counts[label]:
            issues.append(f"{label}: transaction count differs from statement summary")

    # Preserve source order for transactions on the same date. The statement
    # does not supply intraday ordering, so these balances must be labelled.
    records.sort(key=lambda row: ((int(row[0][:2]) - start_month) % 12, int(row[0][3:])))
    balance = summary["Beginning Balance"]
    rows = []
    for date, description, amount, _ in records:
        balance += amount
        rows.append([date, description, float(amount), float(balance)])
    note = "Running balances are calculated from the opening balance and dated transactions. Same-day order follows source sections; intraday balances are not printed in this PDF."
    return {"info": info, "summary": {key: float(value) for key, value in summary.items()},
            "transactions": rows, "warnings": warnings,
            "daily_balances": {key: float(value) for key, value in daily.items()}, "issues": issues,
            "balance_note": note}
