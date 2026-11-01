"""Generate fictional sample invoices (PDF) for the extractor demo.

All companies, numbers and addresses are made up. Two different layouts are
used on purpose, so the extractor has to cope with varying labels, date
formats and table styles - just like real supplier invoices.
"""

from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas

OUT_DIR = Path(__file__).parent / "sample_invoices"

VENDORS = {
    "bluepeak": {
        "name": "Bluepeak Office Supply",
        "address": ["48 Harbor Lane", "Portland, OR 97201", "billing@bluepeak.example"],
        "layout": "classic",
        "tax_label": "Tax (8%)",
        "tax_rate": 0.08,
    },
    "brightline": {
        "name": "Brightline Web Studio",
        "address": ["Unit 7, Canal Works", "Manchester M1 5AN", "hello@brightline.example"],
        "layout": "modern",
        "tax_label": "VAT 20%",
        "tax_rate": 0.20,
    },
    "atlas": {
        "name": "Atlas Freight Co.",
        "address": ["900 Industrial Pkwy", "Columbus, OH 43215", "ar@atlasfreight.example"],
        "layout": "classic",
        "tax_label": "Tax (6%)",
        "tax_rate": 0.06,
    },
}

BILL_TO = ["Acme Retail Inc.", "215 Main Street", "Springfield, IL 62701"]

# (vendor, invoice no, issue date, due date, line items, total override)
# A total override simulates a supplier's typo, so the checks have something to catch.
INVOICES = [
    ("bluepeak", "BP-1043", "2026-09-02", "2026-10-02",
     [("A4 copy paper, 5-ream box", 12, 38.50), ("Black toner cartridge", 4, 72.00),
      ("Desk organizer set", 6, 14.25)], None),
    ("brightline", "BWS/2026/077", "Sep 3, 2026", "Oct 3, 2026",
     [("Landing page design", 1, 850.00), ("Copywriting (3 pages)", 3, 120.00),
      ("Hosting setup", 1, 95.00)], None),
    ("atlas", "AF-55120", "2026-09-08", "2026-09-22",
     [("Pallet shipment Columbus to Chicago", 3, 410.00), ("Fuel surcharge", 1, 62.40),
      ("Liftgate service", 2, 45.00)], None),
    ("bluepeak", "BP-1058", "2026-09-16", "2026-10-16",
     [("Ergonomic office chair", 2, 289.00), ("Whiteboard 120x90 cm", 1, 129.99),
      ("Dry-erase markers, 12 pack", 5, 9.80), ("Sticky notes, 24 pads", 3, 16.40)], None),
    ("brightline", "BWS/2026/081", "Sep 21, 2026", "Oct 21, 2026",
     [("Monthly website maintenance", 1, 300.00), ("Bug fixes (hours)", 6, 65.00)], None),
    ("atlas", "AF-55187", "2026-09-25", "2026-10-09",
     [("LTL shipment Columbus to Detroit", 2, 365.00), ("Residential delivery fee", 1, 38.00)],
     1205.80),  # printed total is wrong on purpose
    ("bluepeak", "BP-1071", "2026-10-01", "2026-10-31",
     [("A4 copy paper, 5-ream box", 8, 38.50), ("Shipping boxes, 25 pack", 4, 27.60),
      ("Packing tape, 6 rolls", 6, 11.90)], None),
    ("brightline", "BWS/2026/090", "Oct 6, 2026", "Nov 5, 2026",
     [("E-commerce product photos (set of 20)", 2, 240.00), ("Image retouching", 20, 6.50),
      ("Monthly website maintenance", 1, 300.00)], None),
]


def money(value: float) -> str:
    return f"${value:,.2f}"


def draw_classic(c: canvas.Canvas, vendor: dict, inv_no: str, issued: str, due: str,
                 items: list, subtotal: float, tax: float, total: float) -> None:
    width, height = A4
    c.setFont("Helvetica-Bold", 20)
    c.drawString(20 * mm, height - 25 * mm, vendor["name"])
    c.setFont("Helvetica", 9)
    for i, line in enumerate(vendor["address"]):
        c.drawString(20 * mm, height - (32 + i * 5) * mm, line)

    c.setFont("Helvetica-Bold", 14)
    c.setFillColor(colors.HexColor("#1F3A5F"))
    c.drawRightString(width - 20 * mm, height - 25 * mm, "INVOICE")
    c.setFillColor(colors.black)
    c.setFont("Helvetica", 10)
    c.drawRightString(width - 20 * mm, height - 33 * mm, f"Invoice No: {inv_no}")
    c.drawRightString(width - 20 * mm, height - 39 * mm, f"Date: {issued}")
    c.drawRightString(width - 20 * mm, height - 45 * mm, f"Due Date: {due}")

    c.setFont("Helvetica-Bold", 10)
    c.drawString(20 * mm, height - 60 * mm, "Bill To:")
    c.setFont("Helvetica", 10)
    for i, line in enumerate(BILL_TO):
        c.drawString(20 * mm, height - (66 + i * 5) * mm, line)

    # Table with grid lines
    top = height - 90 * mm
    cols = [20 * mm, 115 * mm, 135 * mm, 162 * mm, width - 20 * mm]
    row_h = 8 * mm
    c.setFillColor(colors.HexColor("#E8EEF5"))
    c.rect(cols[0], top - row_h, cols[-1] - cols[0], row_h, fill=1, stroke=0)
    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 10)
    headers = ["Description", "Qty", "Unit Price", "Amount"]
    c.drawString(cols[0] + 2 * mm, top - 5.5 * mm, headers[0])
    for j, h in enumerate(headers[1:], start=1):
        c.drawRightString(cols[j + 1] - 2 * mm, top - 5.5 * mm, h)

    c.setFont("Helvetica", 10)
    y = top - row_h
    for desc, qty, price in items:
        y -= row_h
        c.drawString(cols[0] + 2 * mm, y + 2.5 * mm, desc)
        c.drawRightString(cols[2] - 2 * mm, y + 2.5 * mm, str(qty))
        c.drawRightString(cols[3] - 2 * mm, y + 2.5 * mm, money(price))
        c.drawRightString(cols[4] - 2 * mm, y + 2.5 * mm, money(qty * price))
    c.setStrokeColor(colors.HexColor("#9AA8B8"))
    for k in range(len(items) + 2):
        c.line(cols[0], top - k * row_h, cols[-1], top - k * row_h)
    for x in cols:
        c.line(x, top, x, y)

    c.setStrokeColor(colors.black)
    ty = y - 10 * mm
    for label, value, bold in [("Subtotal", subtotal, False), (vendor["tax_label"], tax, False),
                               ("Total", total, True)]:
        c.setFont("Helvetica-Bold" if bold else "Helvetica", 10)
        c.drawString(cols[2], ty, label)
        c.drawRightString(cols[4] - 2 * mm, ty, money(value))
        ty -= 6 * mm

    c.setFont("Helvetica-Oblique", 8)
    c.drawString(20 * mm, 20 * mm, "Thank you for your business. Payment terms as stated above.")


def draw_modern(c: canvas.Canvas, vendor: dict, inv_no: str, issued: str, due: str,
                items: list, subtotal: float, tax: float, total: float) -> None:
    width, height = A4
    c.setFillColor(colors.HexColor("#0F766E"))
    c.rect(0, height - 40 * mm, width, 40 * mm, fill=1, stroke=0)
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 22)
    c.drawString(20 * mm, height - 22 * mm, vendor["name"])
    c.setFont("Helvetica", 9)
    c.drawString(20 * mm, height - 30 * mm, "  |  ".join(vendor["address"]))
    c.setFillColor(colors.black)

    c.setFont("Helvetica", 10)
    c.drawString(20 * mm, height - 55 * mm, f"Invoice # {inv_no}")
    c.drawString(20 * mm, height - 61 * mm, f"Issued: {issued}")
    c.drawString(20 * mm, height - 67 * mm, f"Due: {due}")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(120 * mm, height - 55 * mm, "Billed to")
    c.setFont("Helvetica", 10)
    for i, line in enumerate(BILL_TO):
        c.drawString(120 * mm, height - (61 + i * 6) * mm, line)

    # Table without grid lines, only a rule under the header
    top = height - 95 * mm
    cols = [20 * mm, 118 * mm, 140 * mm, 165 * mm, width - 20 * mm]
    c.setFont("Helvetica-Bold", 9)
    c.setFillColor(colors.HexColor("#0F766E"))
    c.drawString(cols[0], top, "ITEM")
    c.drawRightString(cols[2] - 2 * mm, top, "QTY")
    c.drawRightString(cols[3] - 2 * mm, top, "RATE")
    c.drawRightString(cols[4], top, "TOTAL")
    c.setFillColor(colors.black)
    c.line(cols[0], top - 3 * mm, cols[4], top - 3 * mm)

    c.setFont("Helvetica", 10)
    y = top - 10 * mm
    for desc, qty, price in items:
        c.drawString(cols[0], y, desc)
        c.drawRightString(cols[2] - 2 * mm, y, str(qty))
        c.drawRightString(cols[3] - 2 * mm, y, money(price))
        c.drawRightString(cols[4], y, money(qty * price))
        y -= 8 * mm

    c.line(cols[2], y + 2 * mm, cols[4], y + 2 * mm)
    ty = y - 5 * mm
    for label, value, bold in [("Sub-total", subtotal, False), (vendor["tax_label"], tax, False),
                               ("Amount Due", total, True)]:
        c.setFont("Helvetica-Bold" if bold else "Helvetica", 10)
        c.drawString(cols[2], ty, label)
        c.drawRightString(cols[4], ty, money(value))
        ty -= 6 * mm

    c.setFont("Helvetica", 8)
    c.drawString(20 * mm, 20 * mm, "Bank transfer details are on your account page. Late payments may incur a fee.")


def main() -> None:
    OUT_DIR.mkdir(exist_ok=True)
    for vendor_key, inv_no, issued, due, items, total_override in INVOICES:
        vendor = VENDORS[vendor_key]
        subtotal = round(sum(q * p for _, q, p in items), 2)
        tax = round(subtotal * vendor["tax_rate"], 2)
        total = total_override if total_override is not None else round(subtotal + tax, 2)

        path = OUT_DIR / f"{inv_no.replace('/', '-')}.pdf"
        c = canvas.Canvas(str(path), pagesize=A4)
        c.setTitle(f"Invoice {inv_no}")
        draw = draw_classic if vendor["layout"] == "classic" else draw_modern
        draw(c, vendor, inv_no, issued, due, items, subtotal, tax, total)
        c.showPage()
        c.save()
        print(f"Created {path.name}")


if __name__ == "__main__":
    main()
