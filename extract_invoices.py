"""Extract invoice data from a folder of PDF invoices into a clean Excel workbook.

Usage:
    python extract_invoices.py --input sample_invoices --output output/invoices.xlsx

For every PDF it reads the vendor, invoice number, dates, line items and totals,
then writes three sheets: Invoices, Line Items and By Vendor. Each invoice is
checked with live Excel formulas (line items add up to the subtotal, subtotal +
tax equals the total) and anything that does not match is flagged for review.
"""

from __future__ import annotations

import argparse
import re
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

import pdfplumber
from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

AMOUNT = r"\$?\s?([\d,]+\.\d{2})"
DATE_FORMATS = ["%Y-%m-%d", "%b %d, %Y", "%B %d, %Y", "%d.%m.%Y", "%m/%d/%Y"]
DATE_PATTERN = re.compile(
    r"\d{4}-\d{2}-\d{2}|[A-Z][a-z]+ \d{1,2}, \d{4}|\d{2}\.\d{2}\.\d{4}|\d{1,2}/\d{1,2}/\d{4}"
)
LINE_ITEM = re.compile(
    rf"^(?P<desc>.+?)\s+(?P<qty>\d+(?:\.\d+)?)\s+{AMOUNT}\s+{AMOUNT}$"
)
INVOICE_NO = re.compile(r"Invoice\s*(?:No\.?|Number|#)\s*:?\s*([A-Z0-9][A-Z0-9/\-]*)", re.I)
SUBTOTAL = re.compile(rf"^Sub-?total\s*:?\s*{AMOUNT}", re.I | re.M)
TAX = re.compile(rf"^(?:Tax|VAT|Sales Tax)\b.*?{AMOUNT}\s*$", re.I | re.M)
TOTAL = re.compile(rf"^(?:Total|Total Due|Amount Due|Balance Due)\s*:?\s*{AMOUNT}", re.I | re.M)


@dataclass
class LineItem:
    description: str
    qty: float
    unit_price: float
    amount: float


@dataclass
class Invoice:
    source_file: str
    vendor: str = ""
    invoice_no: str = ""
    invoice_date: date | None = None
    due_date: date | None = None
    subtotal: float | None = None
    tax: float | None = None
    total: float | None = None
    items: list[LineItem] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


def to_number(text: str) -> float:
    return float(text.replace(",", ""))


def parse_date(text: str) -> date | None:
    match = DATE_PATTERN.search(text)
    if not match:
        return None
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(match.group(0), fmt).date()
        except ValueError:
            continue
    return None


def find_vendor(page) -> str:
    """The vendor name is usually the largest text on the page (ignoring 'INVOICE')."""
    words = page.extract_words(extra_attrs=["size"])
    lines: dict[tuple[int, float], list[str]] = {}
    for w in words:
        key = (round(w["top"]), round(w["size"], 1))
        lines.setdefault(key, []).append(w["text"])
    candidates = [
        (size, " ".join(parts)) for (_, size), parts in lines.items()
        if "invoice" not in " ".join(parts).lower()
    ]
    return max(candidates)[1] if candidates else ""


def parse_invoice(path: Path) -> Invoice:
    inv = Invoice(source_file=path.name)
    with pdfplumber.open(path) as pdf:
        inv.vendor = find_vendor(pdf.pages[0])
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)

    if m := INVOICE_NO.search(text):
        inv.invoice_no = m.group(1)

    for line in text.splitlines():
        if re.search(r"\bDue(?:\s+Date)?\s*:", line, re.I):
            inv.due_date = inv.due_date or parse_date(line.split(":", 1)[1])
        elif re.search(r"\b(?:Invoice\s+Date|Date|Issued)\s*:", line, re.I):
            inv.invoice_date = inv.invoice_date or parse_date(line.split(":", 1)[1])

        if m := LINE_ITEM.match(line.strip()):
            inv.items.append(LineItem(
                description=m.group("desc").strip(),
                qty=float(m.group("qty")),
                unit_price=to_number(m.group(3)),
                amount=to_number(m.group(4)),
            ))

    for name, pattern in [("subtotal", SUBTOTAL), ("tax", TAX), ("total", TOTAL)]:
        if m := pattern.search(text):
            setattr(inv, name, to_number(m.group(1)))
        else:
            inv.warnings.append(f"{name} not found")

    for attr in ["vendor", "invoice_no", "invoice_date", "due_date"]:
        if not getattr(inv, attr):
            inv.warnings.append(f"{attr} not found")
    if not inv.items:
        inv.warnings.append("no line items found")
    return inv


# ---------------------------------------------------------------- Excel output

FONT = "Arial"
HEADER_FILL = PatternFill("solid", fgColor="1F3A5F")
HEADER_FONT = Font(name=FONT, bold=True, color="FFFFFF")
BODY_FONT = Font(name=FONT)
BOLD_FONT = Font(name=FONT, bold=True)
REVIEW_FILL = PatternFill("solid", fgColor="F8D7DA")
THIN = Side(style="thin", color="BFC7D1")
MONEY = '$#,##0.00'
DATE_FMT = "yyyy-mm-dd"


def write_table(ws, headers: list[str], rows: list[list], widths: list[int]) -> None:
    ws.append(headers)
    for row in rows:
        ws.append(row)
    for col, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = width
    for cell in ws[1]:
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font = BODY_FONT
            cell.border = Border(bottom=THIN)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].height = 22


def build_workbook(invoices: list[Invoice], output: Path) -> None:
    wb = Workbook()

    # Line Items first, because the Invoices sheet references it.
    ws_inv = wb.active
    ws_inv.title = "Invoices"
    ws_items = wb.create_sheet("Line Items")
    ws_vendor = wb.create_sheet("By Vendor")

    item_rows = [
        [inv.invoice_no, inv.vendor, it.description, it.qty, it.unit_price, it.amount]
        for inv in invoices for it in inv.items
    ]
    write_table(ws_items, ["Invoice No", "Vendor", "Description", "Qty", "Unit Price", "Amount"],
                item_rows, [16, 26, 40, 8, 13, 13])
    for row in ws_items.iter_rows(min_row=2):
        row[4].number_format = MONEY
        row[5].number_format = MONEY

    inv_rows = []
    for i, inv in enumerate(invoices, start=2):
        inv_rows.append([
            inv.source_file, inv.vendor, inv.invoice_no, inv.invoice_date, inv.due_date,
            inv.subtotal, inv.tax, inv.total,
            f"=SUMIFS('Line Items'!$F:$F,'Line Items'!$A:$A,C{i})",
            f'=IF(AND(ABS(I{i}-F{i})<0.005,ABS(F{i}+G{i}-H{i})<0.005),"OK","Review")',
            "; ".join(inv.warnings),
        ])
    write_table(ws_inv, ["Source File", "Vendor", "Invoice No", "Invoice Date", "Due Date",
                         "Subtotal", "Tax", "Total", "Line Items Sum", "Check", "Parser Notes"],
                inv_rows, [22, 26, 16, 13, 13, 13, 12, 13, 15, 10, 24])
    last = len(invoices) + 1
    for row in ws_inv.iter_rows(min_row=2, max_row=last):
        row[3].number_format = DATE_FMT
        row[4].number_format = DATE_FMT
        for cell in row[5:9]:
            cell.number_format = MONEY
        row[9].alignment = Alignment(horizontal="center")
    ws_inv.conditional_formatting.add(
        f"A2:K{last}", FormulaRule(formula=[f'$J2="Review"'], fill=REVIEW_FILL)
    )

    vendors = sorted({inv.vendor for inv in invoices})
    vendor_rows = []
    for i, vendor in enumerate(vendors, start=2):
        vendor_rows.append([
            vendor,
            f"=COUNTIF(Invoices!$B:$B,A{i})",
            f"=SUMIFS(Invoices!$H:$H,Invoices!$B:$B,A{i})",
            f'=COUNTIFS(Invoices!$B:$B,A{i},Invoices!$J:$J,"Review")',
        ])
    write_table(ws_vendor, ["Vendor", "Invoices", "Total Billed", "Needs Review"],
                vendor_rows, [28, 11, 15, 14])
    total_row = len(vendors) + 2
    ws_vendor.append(["Total", f"=SUM(B2:B{total_row - 1})", f"=SUM(C2:C{total_row - 1})",
                      f"=SUM(D2:D{total_row - 1})"])
    for cell in ws_vendor[total_row]:
        cell.font = BOLD_FONT
        cell.border = Border(top=Side(style="thin", color="1F3A5F"))
    for row in ws_vendor.iter_rows(min_row=2, max_row=total_row):
        row[2].number_format = MONEY
    ws_vendor.auto_filter.ref = f"A1:D{total_row - 1}"

    ws_vendor[f"A{total_row + 2}"] = (
        f"Generated by extract_invoices.py on {date.today():%Y-%m-%d} "
        f"from {len(invoices)} PDF files."
    )
    ws_vendor[f"A{total_row + 2}"].font = Font(name=FONT, italic=True, color="666666")

    wb.calculation.fullCalcOnLoad = True
    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--input", type=Path, default=Path("sample_invoices"),
                        help="folder containing PDF invoices")
    parser.add_argument("--output", type=Path, default=Path("output/invoices.xlsx"),
                        help="Excel file to create")
    args = parser.parse_args()

    pdfs = sorted(args.input.glob("*.pdf"))
    if not pdfs:
        raise SystemExit(f"No PDF files found in {args.input}")

    invoices = [parse_invoice(p) for p in pdfs]
    invoices.sort(key=lambda inv: (inv.invoice_date or date.max, inv.invoice_no))
    build_workbook(invoices, args.output)

    for inv in invoices:
        status = "OK"
        if inv.subtotal is None or inv.total is None or inv.tax is None:
            status = "INCOMPLETE"
        elif (abs(sum(it.amount for it in inv.items) - inv.subtotal) > 0.005
              or abs(inv.subtotal + inv.tax - inv.total) > 0.005):
            status = "REVIEW"
        total = f"{inv.total:,.2f}" if inv.total is not None else "-"
        print(f"{inv.source_file:<22} {inv.vendor:<24} {inv.invoice_no:<14} "
              f"{len(inv.items)} items  total {total:>10}  {status}")
    print(f"\nSaved {len(invoices)} invoices to {args.output}")


if __name__ == "__main__":
    main()
