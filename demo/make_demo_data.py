"""Generate the fictional demo inputs for the pdf-excel-cleanup pipeline.

Everything here is simulated. The buyer "Harbor & Pine Supply" and every
vendor, customer and product below are invented. A fixed random seed makes
the whole input folder reproducible.

The invoice PDFs are drawn with reportlab to look like real invoices:
per-vendor fonts, a drawn logo shape, a header block, and a bordered
line-item table that pdfplumber can extract as a table.

For every messy customer name planted in the sales exports the generator
also records the ground truth (which master customer it was derived from,
or that it is genuinely unknown) into input/matching_truth.csv. The
pipeline scores its matcher against that file.

Run directly:  python demo/make_demo_data.py [output_dir]
Or via the CLI: python -m cleanup demo  (generates into ./input)
"""
from __future__ import annotations

import csv
import datetime as dt
import json
import random
import string
import sys
from decimal import Decimal
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import Workbook
from openpyxl.styles import Font
from PIL import Image, ImageDraw
from reportlab.lib import colors
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas

SEED = 20260401
D = Decimal  # all money maths below is Decimal, never float
BUYER = "Harbor & Pine Supply"

# ---------------------------------------------------------------- vendors
VENDORS = {
    "NPP": {"name": "Northern Pine Paper Co.", "tax_id": "22-4819305", "items": [
        ("Copy paper, letter, 10-ream box", 42.50), ("Kraft envelopes 9x12, 250 ct", 31.20),
        ("Sticky notes, 24-pack", 18.75), ("Label sheets, 100 ct", 22.40),
        ("Cardstock, white, 250 sheets", 36.90), ("Binding combs, 100 ct", 14.25),
        ("File folders, letter, box of 100", 27.60), ("Index dividers, 8-tab", 9.85),
        ("Printer paper, legal, 10-ream", 45.10), ("Mailing tubes, 12 ct", 19.95),
        ("Report covers, 25 ct", 16.30), ("Notepads, 12-pack", 13.45)]},
    "BHS": {"name": "Blue Harbor Office Supply", "tax_id": "55-7102946", "items": [
        ("Toner cartridge, black", 89.99), ("Desk organizer", 24.00),
        ("Whiteboard markers, 12 ct", 15.60), ("Ergonomic mouse", 38.50),
        ("Monitor stand", 42.75), ("Keyboard, wireless", 54.20),
        ("Desk lamp, LED", 33.80), ("Cable clips, 50 ct", 8.95),
        ("USB hub, 7-port", 29.40), ("Laptop sleeve, 15in", 26.15),
        ("Surge protector, 8-outlet", 21.70), ("Webcam, HD", 47.30)]},
    "AES": {"name": "Alpen East Supplies", "tax_id": "DE881245667", "items": [
        ("Kugelschreiber, 50 Stk", 23.40), ("Notizblock A5, 10 Stk", 31.15),
        ("Ordner, breit, 10 Stk", 44.80), ("Haftnotizen, gross, 12 Stk", 17.25),
        ("Locher, Metall", 13.90), ("Tacker, Bueromodell", 22.60),
        ("Register A-Z, Karton", 11.35), ("Briefumschlaege C5, 500 Stk", 38.75),
        ("Tintenpatrone, cyan", 62.40), ("Papierkorb, Draht", 9.95),
        ("Schreibtischunterlage", 19.50), ("Textmarker, 8 Stk", 12.85)]},
    "MPS": {"name": "Maple Pine Freight & Pack", "tax_id": "33-6029814", "items": [
        ("Shipping boxes, medium, 25 ct", 34.60), ("Packing tape, 6 rolls", 21.30),
        ("Bubble wrap roll, 50 ft", 28.90), ("Poly mailers, 100 ct", 17.45),
        ("Pallet wrap, 4 rolls", 52.10), ("Shipping labels, 200 ct", 24.80),
        ("Box cutter, safety", 7.65), ("Kraft paper roll", 33.25),
        ("Fragile stickers, 500 ct", 9.40), ("Edge protectors, 50 ct", 26.70),
        ("Mail scale, 25 lb", 41.85), ("Strapping kit", 58.20)]},
    "CRS": {"name": "Cedar Ridge Supply", "tax_id": "77-2048813", "items": [
        ("Overstock credit - copy paper", 42.50), ("Returned toner credit", 89.99),
        ("Price adjustment credit", 25.00), ("Damaged goods credit", 60.00),
        ("Volume rebate credit", 150.00), ("Shipping overcharge credit", 18.75),
        ("Early payment credit", 35.40), ("Catalogue price credit", 12.30),
        ("Restock fee credit", 48.90), ("Promotional credit", 75.00),
        ("Freight refund credit", 22.15), ("Service credit", 95.60)]},
}

ACCENTS = {"NPP": "#1e6b3a", "BHS": "#1a5fa8", "AES": "#4a4a4a",
           "MPS": "#b34700", "CRS": "#5d4037"}

LINES_PER_PAGE_NPP = 12  # longer NPP invoices continue on page 2

# duplicate-invoice plant: (source invoice, alias filename)
DUP_ALIASES = [("MPS-2209", "MPS-2209 refax.pdf"),
               ("CRS-5107", "CRS-5107 duplicate.pdf")]

# ------------------------------------------------------------- customers
FIRST = ["Alder", "Beacon", "Cedar", "Drift", "Elm", "Foxglove", "Granite", "Hollow",
         "Iron", "Juniper", "Kestrel", "Larkspur", "Meadow", "Northgate", "Osprey",
         "Plover", "Quarry", "Riverbend", "Sable", "Timberline", "Upland", "Verdant",
         "Willow", "Yellowfin", "Zephyr", "Birchwood", "Copper", "Dune"]
SECOND = ["Bakery", "Clinic", "Deli", "Hardware", "Florist", "Garage", "Books",
          "Cafe", "Pharmacy", "Market", "Dental", "Veterinary", "Tattoo", "Cleaners",
          "Print Shop", "Nursery", "Bicycles", "Outfitters", "Brewing", "Cheese Shop",
          "Boutique", "Butchery", "Photo Studio"]
CITIES = ["Bayfield", "Clearwater", "Eastport", "Fairpoint", "Greyrock", "Hartland",
          "Inverness", "Kingsford", "Lakemont", "Millbrook"]

# Pairs of master customers that share a first word and the start of the
# second — typing a truncated version is genuinely ambiguous.
AMBIGUOUS_PAIRS = [("Riverbend Books", "Riverbend Boutique"),
                   ("Kestrel Bakery", "Kestrel Butchery"),
                   ("Timberline Dental", "Timberline Denim Co"),
                   ("Meadow Pharmacy", "Meadow Photo Studio")]
# Truncations planted into the online export: each maps to one pair member
# but a human should decide, not a matcher.
AMBIGUOUS_NAMES = ["Riverbend Bo", "Riverbend Bou", "Kestrel Ba",
                   "Kestrel Bu", "Timberline De", "Timberline Den",
                   "Meadow Ph", "Meadow Pho"]

PRODUCTS_POS = {"Widget A": 19.99, "Widget B": 34.50, "Gizmo Pro": 129.00,
                "Starter Kit": 59.00, "Refill Pack": 12.75, "Carry Case": 24.90}
PRODUCTS_ONLINE = {"SKU-1001": ("Travel mug, steel", 18.50), "SKU-1002": ("Canvas tote", 12.00),
                   "SKU-1003": ("Insulated bottle", 24.75), "SKU-1004": ("Notebook set", 15.20),
                   "SKU-1005": ("Desk caddy", 21.40), "SKU-1006": ("Tea sampler", 16.90)}
REGIONS = {"North": ["North", "north", "North ", "N."],
           "South": ["South", "SOUTH", "S."],
           "East": ["East", "east ", "E."],
           "West": ["West", "W.", "west"]}
REPS = ["Avery Quinn", "Jordan Blake", "Casey Morgan", "Riley Shaw", "Taylor Reed"]
# made-up names chosen to score below the review threshold against every
# master customer, so the "unknown customer" problem is deterministic
UNKNOWN_FIRST = ["Quibell", "Vexmoor", "Tarnside", "Brevik", "Solmara",
                 "Dunmarrow", "Hallowick", "Priskly", "Umberton", "Zelkara"]
UNKNOWN_SECOND = ["Works", "Supply Hut", "Trading Co", "Odd Lots", "Depot"]


def customers_master() -> list[dict]:
    rng = random.Random(SEED + 7)
    names = {n for pair in AMBIGUOUS_PAIRS for n in pair}
    while len(names) < 120:
        names.add(f"{rng.choice(FIRST)} {rng.choice(SECOND)}")
    return [{"customer_id": f"C-{i:04d}", "customer_name": n,
             "city": rng.choice(CITIES), "terms": rng.choice(["Net 15", "Net 30", "Due on receipt"])}
            for i, n in enumerate(sorted(names), 1)]


def vendor_master() -> list[dict]:
    return [{"vendor_code": k, "vendor_name": v["name"], "tax_id": v["tax_id"],
             "terms": "Net 30"} for k, v in VENDORS.items()]


# ------------------------------------------------------------- PDF drawing
def _us(x) -> str:
    return f"{x:,.2f}"


def _eu(x) -> str:
    return f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _logo(c: canvas.Canvas, code: str, x: float, y: float) -> None:
    """A small drawn mark per vendor (vector shapes, no image assets)."""
    c.saveState()
    accent = colors.HexColor(ACCENTS[code])
    c.setStrokeColor(accent)
    c.setFillColor(accent)
    c.setLineWidth(1.4)
    if code == "NPP":  # pine tree
        p = c.beginPath()
        p.moveTo(x, y + 20); p.lineTo(x + 9, y + 7); p.lineTo(x - 9, y + 7)
        p.close()
        p2 = c.beginPath()
        p2.moveTo(x, y + 12); p2.lineTo(x + 12, y - 2); p2.lineTo(x - 12, y - 2)
        p2.close()
        c.drawPath(p, fill=1, stroke=0)
        c.drawPath(p2, fill=1, stroke=0)
        c.rect(x - 2, y - 8, 4, 6, fill=1, stroke=0)
    elif code == "BHS":  # waves inside a circle (harbor)
        c.circle(x, y + 7, 11, fill=0, stroke=1)
        c.setLineWidth(1.1)
        c.line(x - 7, y + 7, x + 7, y + 7)
        c.line(x - 5, y + 4, x + 5, y + 4)
        c.line(x - 7, y + 10, x + 7, y + 10)
    elif code == "AES":  # mountain range
        p = c.beginPath()
        p.moveTo(x - 11, y - 4); p.lineTo(x - 2, y + 16); p.lineTo(x + 4, y + 4)
        p.lineTo(x + 7, y + 10); p.lineTo(x + 13, y - 4); p.close()
        c.drawPath(p, fill=0, stroke=1)
    elif code == "MPS":  # open shipping box
        c.rect(x - 9, y - 4, 18, 12, fill=0, stroke=1)
        c.line(x - 9, y + 8, x, y + 14)
        c.line(x + 9, y + 8, x, y + 14)
        c.line(x - 9, y + 8, x + 9, y + 8)
    elif code == "CRS":  # concentric diamonds
        p = c.beginPath()
        p.moveTo(x, y + 15); p.lineTo(x + 10, y + 5); p.lineTo(x, y - 5)
        p.lineTo(x - 10, y + 5); p.close()
        c.drawPath(p, fill=0, stroke=1)
        c.circle(x, y + 5, 2, fill=1, stroke=0)
    c.restoreState()


def _draw_table(c: canvas.Canvas, x: float, y_top: float, col_w: list[float],
                header: list[str], rows: list[list[str]], *,
                font: str = "Helvetica", font_bold: str = "Helvetica-Bold",
                fs: float = 9, row_h: float = 17,
                header_fill: str = "#e8edf3",
                aligns: list[str] | None = None) -> float:
    """A fully ruled table (borders on every cell) — pdfplumber's 'lines'
    table strategy reads this as a real table."""
    tw = sum(col_w)
    n = len(rows) + 1
    c.saveState()
    c.setFillColor(colors.HexColor(header_fill))
    c.rect(x, y_top - row_h, tw, row_h, fill=1, stroke=0)
    c.restoreState()
    for i in range(n + 1):
        c.line(x, y_top - i * row_h, x + tw, y_top - i * row_h)
    xx = x
    for w in col_w:
        c.line(xx, y_top, xx, y_top - n * row_h)
        xx += w
    c.line(x + tw, y_top, x + tw, y_top - n * row_h)
    xs = [x]
    for w in col_w:
        xs.append(xs[-1] + w)
    pad = 5
    c.setFont(font_bold, fs)
    for j, t in enumerate(header):
        c.drawString(xs[j] + pad, y_top - row_h + pad, str(t))
    c.setFont(font, fs)
    for i, row in enumerate(rows):
        y = y_top - (i + 2) * row_h + pad
        for j, t in enumerate(row):
            if aligns and aligns[j] == "r":
                c.drawRightString(xs[j + 1] - pad, y, str(t))
            else:
                c.drawString(xs[j] + pad, y, str(t))
    return y_top - n * row_h


def _totals_box(c: canvas.Canvas, x_right: float, y_top: float,
                entries: list[tuple[str, str, bool]], *,
                w: float = 200, lh: float = 15, pad: float = 8,
                font: str = "Helvetica",
                font_bold: str = "Helvetica-Bold", fs: float = 9) -> None:
    """Labelled totals inside a bordered box, right side of the page."""
    h = pad * 2 + lh * len(entries)
    c.rect(x_right - w, y_top - h, w, h)
    y = y_top - pad - fs + 2
    for label, val, bold in entries:
        c.setFont(font_bold if bold else font, fs)
        c.drawString(x_right - w + pad, y, label)
        c.drawRightString(x_right - pad, y, val)
        y -= lh


def _npp(c: canvas.Canvas, inv: dict) -> None:
    v = VENDORS["NPP"]
    x = 56
    _logo(c, "NPP", x + 10, LETTER[1] - 58)
    c.setFont("Helvetica-Bold", 15)
    c.drawString(x + 30, LETTER[1] - 52, v["name"])
    c.setFont("Helvetica", 8.5)
    c.drawString(x + 30, LETTER[1] - 66,
                 "1200 Alder Street, Bayfield  |  ap@npp.example")
    meta_x = LETTER[0] - 56
    c.setFont("Helvetica", 9.5)
    for i, line in enumerate([
            f"Invoice No.: {inv['number']}",
            f"Invoice Date: {inv['date']:%m/%d/%Y}",
            f"Bill To: {BUYER}", f"Tax ID: {inv['tax_id']}"]):
        c.drawRightString(meta_x, LETTER[1] - 52 - i * 13, line)
    header = ["Item Description", "Qty", "Unit Price", "Amount"]
    colw, aligns = [290, 55, 75, 75], ["l", "r", "r", "r"]
    rows = [[d, str(q), _us(p), _us(q * p)] for d, q, p in inv["lines"]]
    first, rest = rows[:LINES_PER_PAGE_NPP], rows[LINES_PER_PAGE_NPP:]
    yb = _draw_table(c, x, LETTER[1] - 110, colw, header, first,
                     font="Helvetica", font_bold="Helvetica-Bold",
                     aligns=aligns)
    if rest:
        c.setFont("Helvetica-Oblique", 9)
        c.drawString(x, yb - 18, "Continued on page 2")
        c.showPage()
        _logo(c, "NPP", x + 10, LETTER[1] - 48)
        c.setFont("Helvetica-Bold", 12)
        c.drawString(x + 30, LETTER[1] - 44, v["name"])
        c.setFont("Helvetica", 9)
        c.drawString(x + 30, LETTER[1] - 58,
                     f"Invoice {inv['number']} - continued")
        yb = _draw_table(c, x, LETTER[1] - 85, colw, header, rest,
                         font="Helvetica", font_bold="Helvetica-Bold",
                         aligns=aligns)
    _totals_box(c, x + sum(colw), yb - 16,
                [("Subtotal:", _us(inv["line_sum"]), False),
                 ("Tax (8%):", _us(inv["tax"]), False),
                 ("Total Due:", _us(inv["total"]), True)])
    c.setFont("Helvetica", 8)
    c.drawString(x, 56, "Payment terms: Net 30. Thank you for your business.")
    c.showPage()


def _bhs(c: canvas.Canvas, inv: dict) -> None:
    v = VENDORS["BHS"]
    x = 62
    _logo(c, "BHS", x + 10, LETTER[1] - 62)
    c.setFont("Helvetica-Bold", 16)
    c.drawString(x + 28, LETTER[1] - 54, v["name"])
    c.setFont("Helvetica", 8.5)
    c.drawString(x + 28, LETTER[1] - 68, "44 Quay Lane, Eastport")
    # bordered meta box, top right - the BHS signature layout feature
    bx, by, bw, bh = LETTER[0] - 210, LETTER[1] - 92, 150, 44
    c.setLineWidth(1.6)
    c.rect(bx, by, bw, bh)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(bx + 10, by + bh - 18, "INVOICE")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(bx + 10, by + 12, f"No. {inv['number']}")
    c.setFont("Helvetica", 9.5)
    for i, line in enumerate([
            f"Invoice Date: {inv['date']:%d-%b-%Y}",
            f"Customer: {BUYER}", f"Vendor Tax ID: {inv['tax_id']}"]):
        c.drawString(x, LETTER[1] - 116 - i * 13, line)
    header = ["PRODUCT", "QTY", "UNIT PRICE", "EXTENDED"]
    colw, aligns = [280, 55, 75, 75], ["l", "r", "r", "r"]
    rows = [[d, str(q), _us(p), _us(q * p)] for d, q, p in inv["lines"]]
    yb = _draw_table(c, x, LETTER[1] - 178, colw, header, rows,
                     font="Helvetica", font_bold="Helvetica-Bold",
                     header_fill="#dce7f2", aligns=aligns)
    _totals_box(c, x + sum(colw), yb - 16,
                [("AMOUNT DUE:", _us(inv["total"]), True)], w=220)
    c.setFont("Helvetica", 8)
    c.drawString(x, 56, "Terms: Net 30")
    c.showPage()


def _aes(c: canvas.Canvas, inv: dict) -> None:
    v = VENDORS["AES"]
    x = 58
    _logo(c, "AES", x + 12, LETTER[1] - 60)
    c.setFont("Times-Bold", 15)
    c.drawString(x + 32, LETTER[1] - 52, v["name"])
    c.setFont("Times-Roman", 9)
    c.drawString(x + 32, LETTER[1] - 66, "Bergstrasse 9, Inverness")
    c.setFont("Times-Bold", 13)
    c.drawRightString(LETTER[0] - 58, LETTER[1] - 52, "RECHNUNG")
    c.setFont("Times-Roman", 10)
    for i, line in enumerate([
            f"Rechnung / Invoice: {inv['number']}",
            f"Datum: {inv['date']:%d.%m.%Y}",
            f"Kunde: {BUYER}", f"USt-IdNr: {inv['tax_id']}"]):
        c.drawString(x, LETTER[1] - 112 - i * 14, line)
    header = ["Beschreibung", "Menge", "Einzelpreis", "Betrag"]
    colw, aligns = [285, 55, 75, 75], ["l", "r", "r", "r"]
    rows = [[d, str(q), _eu(p), _eu(q * p)] for d, q, p in inv["lines"]]
    yb = _draw_table(c, x, LETTER[1] - 180, colw, header, rows,
                     font="Times-Roman", font_bold="Times-Bold",
                     header_fill="#ececec", aligns=aligns)
    _totals_box(c, x + sum(colw), yb - 16,
                [("Zwischensumme:", f"{_eu(inv['line_sum'])} EUR", False),
                 ("Gesamtbetrag:", f"{_eu(inv['total'])} EUR", True)],
                font="Times-Roman", font_bold="Times-Bold", w=215)
    c.setFont("Times-Roman", 8)
    c.drawString(x, 56, "Zahlbar innerhalb von 30 Tagen.")
    c.showPage()


def _mps(c: canvas.Canvas, inv: dict) -> None:
    v = VENDORS["MPS"]
    x = 52
    _logo(c, "MPS", x + 10, LETTER[1] - 60)
    c.setFont("Helvetica-Bold", 14)
    c.drawString(x + 28, LETTER[1] - 52, v["name"])
    c.setFont("Helvetica", 8.5)
    c.drawString(x + 28, LETTER[1] - 66,
                 "7 Dock Road, Millbrook  |  freight@mps.example")
    c.setFont("Helvetica-Oblique", 9.5)
    for i, line in enumerate([
            f"Invoice #: {inv['number']}",
            f"Date: {inv['date']:%Y-%m-%d}",
            f"Bill To: {BUYER}", f"Tax ID: {inv['tax_id']}"]):
        c.drawRightString(LETTER[0] - 52, LETTER[1] - 50 - i * 13, line)
    header = ["Line Item", "Qty", "Rate", "Amount"]
    colw, aligns = [292, 55, 75, 75], ["l", "r", "r", "r"]
    rows = [[d, str(q), _us(p), _us(q * p)] for d, q, p in inv["lines"]]
    yb = _draw_table(c, x, LETTER[1] - 112, colw, header, rows,
                     font="Helvetica", font_bold="Helvetica-Bold",
                     header_fill="#f2e4da", aligns=aligns)
    _totals_box(c, x + sum(colw), yb - 16,
                [("TOTAL:", _us(inv["total"]), True)], w=210)
    c.setFont("Helvetica", 8)
    c.drawString(x, 56, "Freight terms: Net 30.")
    c.showPage()


def _crs(c: canvas.Canvas, inv: dict) -> None:
    v = VENDORS["CRS"]
    x = 58
    # dark banner across the top: credit notes announce themselves
    c.saveState()
    c.setFillColor(colors.HexColor(ACCENTS["CRS"]))
    c.rect(0, LETTER[1] - 46, LETTER[0], 34, fill=1, stroke=0)
    c.restoreState()
    c.setFillColor(colors.white)
    c.setFont("Courier-Bold", 14)
    c.drawString(x, LETTER[1] - 36, "CREDIT NOTE")
    c.setFillColor(colors.black)
    _logo(c, "CRS", x + 10, LETTER[1] - 96)
    c.setFont("Courier-Bold", 13)
    c.drawString(x + 30, LETTER[1] - 88, v["name"])
    c.setFont("Courier", 9)
    c.drawString(x + 30, LETTER[1] - 102, "19 Ridge Way, Greyrock")
    c.setFont("Courier", 9.5)
    for i, line in enumerate([
            f"Credit Note No.: {inv['number']}",
            f"Date: {inv['date']:%m/%d/%y}",
            f"Customer: {BUYER}", f"Tax ID: {inv['tax_id']}"]):
        c.drawRightString(LETTER[0] - 58, LETTER[1] - 82 - i * 13, line)
    header = ["Description", "Qty", "Unit Credit", "Amount"]
    colw, aligns = [285, 55, 80, 75], ["l", "r", "r", "r"]
    rows = [[d, str(q), _us(-abs(p)), _us(-abs(q * p))]
            for d, q, p in inv["lines"]]
    yb = _draw_table(c, x, LETTER[1] - 158, colw, header, rows,
                     font="Courier", font_bold="Courier-Bold",
                     header_fill="#e6ddd8", aligns=aligns)
    _totals_box(c, x + sum(colw), yb - 16,
                [("Total Credit:", f"-{_us(abs(inv['total']))}", True)],
                font="Courier", font_bold="Courier-Bold", w=215)
    c.setFont("Courier", 8)
    c.drawString(x, 56, "Credit applied to your account.")
    c.showPage()


LAYOUTS = {"NPP": _npp, "BHS": _bhs, "AES": _aes, "MPS": _mps, "CRS": _crs}


def _scanned_pdf(path: Path, number: str, rng: random.Random) -> None:
    """Image-only PDF: a raster picture of an invoice, no text layer."""
    img = Image.new("RGB", (612, 500), "white")
    draw = ImageDraw.Draw(img)
    draw.rectangle((30, 20, 220, 70), outline="black", width=2)
    draw.text((40, 32), "Blue Harbor Office Supply", fill="black")
    draw.text((40, 52), f"Scanned invoice {number}", fill="black")
    draw.line((30, 100, 582, 100), fill="black", width=1)
    y = 110
    for label, val in [("Item", "Amount"), ("Toner cartridge, black", "89.99"),
                       ("Desk organizer", "24.00"), ("Shipping", "15.00")]:
        draw.text((40, y), label, fill="black")
        draw.text((500, y), val, fill="black")
        y += 22
        draw.line((30, y - 4, 582, y - 4), fill="gray")
    draw.text((380, y + 10), "Amount due: $ 128.99", fill="black")
    draw.text((30, 470), "(scanned fax copy - no text layer)", fill="gray")
    # draw the PIL image straight from memory — no temp file that could
    # clobber (and then delete) a same-named file in the output folder
    c = canvas.Canvas(str(path), pagesize=LETTER, invariant=1)
    c.drawImage(ImageReader(img), 30, 250, width=552, height=450)
    c.showPage()
    c.save()


def make_invoices(inv_dir: Path, rng: random.Random) -> tuple[list[dict], dict]:
    """Write ~120 text PDFs + 3 scanned PDFs; return (invoice_specs, planted)."""
    problems = {"duplicate_invoice": [], "invoice_total_mismatch": [],
                "vendor_tax_id_mismatch": [], "future_invoice_date": [], "scanned_pdf": []}
    specs: list[dict] = []
    starts = {"NPP": 1001, "BHS": 2001, "AES": 3101, "MPS": 2201, "CRS": 5101}
    day0 = dt.date(2026, 2, 2)
    for code, start in starts.items():
        for i in range(24):
            n_lines = rng.randint(2, 6) if code != "NPP" else rng.choice([3, 5, 8, 15, 19])
            pool = VENDORS[code]["items"]
            lines = [(d, rng.randint(1, 8), D(str(p)))
                     for d, p in rng.choices(pool, k=n_lines)]
            if code == "MPS":
                ship = ("Shipping & Handling", 1,
                        D(str(rng.choice([14.50, 18.75, 22.00]))))
                disc = ("Discount - loyalty 5%", 1,
                        -(D("0.05") * sum(q * p for _, q, p in lines)
                          ).quantize(D("0.01")))   # ROUND_HALF_EVEN-free
                lines = lines + [disc, ship]
            inv = {"vendor": code, "number": f"{code}-{start + i}",
                   "date": day0 + dt.timedelta(days=int(i * 3.4) + rng.randint(0, 2)),
                   "lines": lines, "tax_id": VENDORS[code]["tax_id"]}
            inv["line_sum"] = sum(q * p for _, q, p in lines).quantize(
                D("0.01"))
            inv["tax"] = (inv["line_sum"] * D("0.08")).quantize(D("0.01")) \
                if code == "NPP" else D("0")
            inv["total"] = inv["line_sum"] + inv["tax"]
            if code == "CRS":
                inv["total"] = -abs(inv["line_sum"])
            specs.append(inv)

    # --- plant problems -------------------------------------------------
    by_no = {s["number"]: s for s in specs}
    # 1. invoice whose lines do not add up to the printed total
    s = by_no["NPP-1017"]
    s["total"] = s["total"] + D("25.00")
    problems["invoice_total_mismatch"].append(s["number"])
    # 2. invoice with the wrong vendor tax ID
    s = by_no["BHS-2011"]
    s["tax_id"] = "55-0000000"
    problems["vendor_tax_id_mismatch"].append(s["number"])
    # 3. invoice dated in the future (vs reference_date)
    s = by_no["AES-3122"]
    s["date"] = dt.date(2026, 5, 18)
    problems["future_invoice_date"].append(s["number"])

    for spec in specs:
        c = canvas.Canvas(str(inv_dir / f"{spec['number']}.pdf"),
                          pagesize=LETTER, invariant=1)
        LAYOUTS[spec["vendor"]](c, spec)
        c.save()

    # 4. duplicate invoices: same invoice sent twice (byte copies, new names)
    import shutil
    for num, alias in DUP_ALIASES:
        shutil.copyfile(inv_dir / f"{num}.pdf", inv_dir / alias)
        problems["duplicate_invoice"].append(num)

    # 5. scanned PDFs (no text layer)
    for i in range(3):
        name = f"scan_{i + 1:03d}.pdf"
        _scanned_pdf(inv_dir / name, f"BHS-{2100 + i}", rng)
        problems["scanned_pdf"].append(name)

    return specs, problems


# ------------------------------------------------------------- POS export
def _messy_date(d: dt.date, k: int):
    return [d, d.strftime("%m/%d/%y"), d.strftime("%Y-%m-%d"),
            d.strftime("%B %d %Y"), d.strftime("%d.%m.%Y")][k % 5]


def _variant(name: str, rng: random.Random) -> tuple[str, str]:
    """A messy spelling of a master customer. Returns (variant, kind).

    kinds: "typo" - an edit the matcher must fuzzy-resolve (1-char sub/del,
    or a swap on a long name); "suffix"/"punct"/"case"/"space" - variants
    normalisation alone fixes.
    """
    kinds = ["typo", "typo", "suffix", "punct", "case", "space"]
    kind = rng.choice(kinds)
    if kind == "typo":
        edit = rng.choice(["sub", "del", "swap"] if len(name) >= 14
                          else ["sub", "del"])
        if edit == "swap":
            # adjacent identical letters would produce an unchanged string
            pos = rng.choice([p for p in range(3, len(name) - 2)
                              if name[p] != name[p + 1]])
            return name[:pos] + name[pos + 1] + name[pos] + name[pos + 2:], kind
        pos = rng.randint(3, len(name) - 1)
        if edit == "del":
            return name[:pos] + name[pos + 1:], kind
        ch = rng.choice(string.ascii_lowercase)
        while ch.lower() == name[pos].lower():
            ch = rng.choice(string.ascii_lowercase)
        return name[:pos] + ch + name[pos + 1:], kind
    if kind == "suffix":
        return name + rng.choice([" Inc", " LLC", " Co.", " Ltd"]), kind
    if kind == "punct":
        return name.replace(" ", rng.choice(["-", "/", "  "]), 1), kind
    if kind == "case":
        return rng.choice([name.upper(), name.lower()]), kind
    return " " + name + "  ", kind


def make_pos_xlsx(path: Path, customers: list[dict], rng: random.Random):
    wb = Workbook()
    ws = wb.active
    ws.title = "POS Export"
    ws.merge_cells("A1:H1")
    ws["A1"] = "Harbor & Pine POS Export - FICTIONAL DEMO DATA"
    ws["A1"].font = Font(bold=True, size=13)
    ws.merge_cells("A2:H2")
    ws["A2"] = "legacy till system v2.3 (made up) - do not edit"
    ws.append([])
    # merged group header row above the column headers
    ws.merge_cells("B4:C4")
    ws.merge_cells("D4:H4")
    ws["A4"] = "Record"
    ws["B4"] = "Location / Staff"
    ws["D4"] = "Sale Details"
    ws.append(["Order Date ", "REGION", "Sales Rep", "product", "Units ",
               "Revenue", "Notes", " Customer "])

    ids = {c["customer_name"]: c["customer_id"] for c in customers}
    names = [c["customer_name"] for c in customers]
    problems = {"duplicate_sales_row": [], "negative_quantity": [],
                "unparseable_amount": [], "subtotal_row_dropped": [],
                "blank_row": [], "metadata_row_dropped": []}
    truth: dict[str, dict] = {}
    planted = {"typo": [], "ambiguous": [], "unknown": []}
    region_cycle = list(REGIONS)
    n_data = 3000
    row_i = 0
    dup_targets = set(rng.sample(range(n_data), 15))
    neg_targets = set(rng.sample(range(n_data), 8))
    bad_amt_targets = set(rng.sample(range(n_data), 6))
    subtotal_at = set(range(249, n_data, 250))
    blank_at = set(range(119, n_data, 200))
    # rows that get a non-master customer name
    unknown_targets = set(rng.sample(range(n_data), 10))
    fuzzy_targets = set(rng.sample(range(n_data), 20))
    unknown_names = [f"{rng.choice(UNKNOWN_FIRST)} {rng.choice(UNKNOWN_SECOND)}"
                     for _ in range(10)]

    def record(raw: str, truth_id: str, kind: str) -> None:
        if raw in truth:
            return  # same string planted earlier under another kind
        truth[raw] = {"truth": truth_id, "planted_kind": kind}
        if kind in planted and raw not in planted[kind]:
            planted[kind].append(raw)

    while row_i < n_data:
        d = dt.date(2026, 1, 5) + dt.timedelta(days=rng.randint(0, 110))
        region = region_cycle[row_i % 4]
        prod = rng.choice(list(PRODUCTS_POS))
        qty = rng.randint(1, 40)
        # Decimal maths: a half-cent discount (e.g. 5 x 19.99 x 0.9 =
        # 89.955) rounds HALF_UP like the pipeline, so the planted data
        # is a reliable money baseline
        rev = (D(qty) * D(str(PRODUCTS_POS[prod]))
               * D(str(rng.choice(["1.0", "1.0", "0.9"])))).quantize(
            D("0.01"), rounding="ROUND_HALF_UP")
        rep = rng.choice(REPS)
        if rng.random() < 0.7:
            cust, tid, k = "Walk-in", "WALKIN", "walkin"
        else:
            base = rng.choice(names)
            cust, tid, k = base, ids[base], "normal"

        if row_i in unknown_targets:
            cust = unknown_names[len(planted["unknown"]) % len(unknown_names)]
            record(cust, "UNKNOWN", "unknown")
        elif row_i in fuzzy_targets:
            base = rng.choice(names)
            cust, k = _variant(base, rng)
            record(cust, ids[base], k if k != "typo" else "typo")
        else:
            record(cust, tid, k)

        vals = [_messy_date(d, row_i), rng.choice(REGIONS[region]),
                rep if row_i % 7 else rep.upper(),
                prod if row_i % 5 else prod.lower() + " ",
                qty if row_i % 6 else f"{qty} pcs",
                rev if row_i % 4 else f"${rev:,.2f}",
                "", cust]
        if row_i in neg_targets:
            vals[4] = f"-{qty}"
            problems["negative_quantity"].append(f"POS-R{ws.max_row + 1}")
        if row_i in bad_amt_targets:
            vals[5] = rng.choice(["N/A", "TBD", "pending"])
            problems["unparseable_amount"].append(f"POS-R{ws.max_row + 1}")
        ws.append(vals)
        if row_i in dup_targets:
            ws.append(vals)
            problems["duplicate_sales_row"].append(f"POS-R{ws.max_row}")
        if row_i in subtotal_at:
            ws.append(["", f"Subtotal - {region}", "", "", "",
                       round(rng.uniform(5000, 9000), 2), "", ""])
            problems["subtotal_row_dropped"].append(f"POS-R{ws.max_row}")
        if row_i in blank_at:
            ws.append([])
            problems["blank_row"].append(f"POS-R{ws.max_row + 1}")
        row_i += 1

    ws.append([])
    problems["blank_row"].append(f"POS-R{ws.max_row + 1}")
    ws.append(["Note: export ends here", "", "", "", "", "", "", ""])
    problems["metadata_row_dropped"].append(f"POS-R{ws.max_row}")
    ws.append(["", "GRAND TOTAL", "", "", "", "=SUM(F6:F3000)", "", ""])
    problems["metadata_row_dropped"].append(f"POS-R{ws.max_row}")

    for col, w in zip("ABCDEFGH", (16, 10, 16, 14, 9, 12, 14, 22)):
        ws.column_dimensions[col].width = w
    wb.save(path)
    return ws.max_row, problems, truth, planted


# ---------------------------------------------------------- online export
def make_online_csv(path: Path, customers: list[dict], rng: random.Random):
    ids = {c["customer_name"]: c["customer_id"] for c in customers}
    names = [c["customer_name"] for c in customers]
    pair_ids = {name: ids[name] for pair in AMBIGUOUS_PAIRS for name in pair}
    problems = {"refund_without_sale": []}
    truth: dict[str, dict] = {}
    planted = {"typo": [], "ambiguous": [], "unknown": []}
    unknown_names = [f"{rng.choice(UNKNOWN_FIRST)} {rng.choice(UNKNOWN_SECOND)}"
                     for _ in range(8)]
    rows = []
    order_no = 90001
    t0 = dt.datetime(2026, 1, 6, 8, 0, 0)
    n_orders = 1900

    def record(raw: str, truth_id: str, kind: str) -> None:
        if raw in truth:
            return  # same string planted earlier under another kind
        truth[raw] = {"truth": truth_id, "planted_kind": kind}
        if kind in planted and raw not in planted[kind]:
            planted[kind].append(raw)

    for i in range(n_orders):
        sku = rng.choice(list(PRODUCTS_ONLINE))
        item, price = PRODUCTS_ONLINE[sku]
        qty = rng.randint(1, 12)
        oid = f"ONL-{order_no + i}"
        kind = rng.random()
        if kind < 0.50:
            base = rng.choice(names)
            cust = base
            record(cust, ids[base], "normal")
        elif kind < 0.68:
            base = rng.choice(names)
            cust = rng.choice([base.upper(), base.lower(), " " + base + " "])
            record(cust, ids[base], "case")
        elif kind < 0.82:
            base = rng.choice(names)
            cust, vk = _variant(base, rng)
            record(cust, ids[base], vk)
        elif kind < 0.88:
            cust = rng.choice(AMBIGUOUS_NAMES)
            record(cust, "AMBIGUOUS:" + pair_ids[_ambiguous_base(cust)],
                   "ambiguous")
        else:
            cust = rng.choice(unknown_names)
            record(cust, "UNKNOWN", "unknown")
        ts = t0 + dt.timedelta(minutes=rng.randint(0, 110 * 24 * 60))
        rows.append([oid, ts.strftime("%Y-%m-%dT%H:%M:%SZ"), cust, sku, item,
                     qty, f"{price:.2f}", f"{qty * price:.2f}", "sale"])

    # refunds: negative rows; some reference a real order, some do not
    refund_rows = rng.sample([r for r in rows if r[8] == "sale"], 60)
    for r in refund_rows:
        oid = r[0]
        rows.append([oid, r[1], r[2], r[3], r[4], -int(r[5]), r[6],
                     f"-{r[7]}", "refund"])
    for i in range(8):
        oid = f"ONL-{order_no + n_orders + i}"
        r = rng.choice(refund_rows)
        rows.append([oid, r[1], r[2], r[3], r[4], -int(r[5]), r[6],
                     f"-{r[7]}", "refund"])
        problems["refund_without_sale"].append(oid)
    # refunds of real orders whose amount came out positive (data bug)
    problems["refund_sign_invalid"] = []
    for r in refund_rows[:2]:
        rows.append([r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7],
                     "refund"])
        problems["refund_sign_invalid"].append(r[0])
    # a refund of an item the order never sold
    problems["refund_line_mismatch"] = []
    r = refund_rows[3]
    rows.append([r[0], r[1], r[2], "SKU-9999", "Unknown item",
                 -int(r[5]), r[6], f"-{r[7]}", "refund"])
    problems["refund_line_mismatch"].append(r[0])
    # exact re-exported rows (identical duplicates, provable by row id)
    for r in rng.sample(rows[:500], 3):
        rows.append(list(r))
        problems.setdefault("duplicate_sales_row", []).append(r[0])

    rng.shuffle(rows)
    with path.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, lineterminator="\n")
        w.writerow(["order_id", "ordered_at_utc", "customer", "sku", "item",
                    "quantity", "unit_price", "line_total", "status"])
        w.writerows(rows)
    return len(rows), problems, truth, planted


def _ambiguous_base(raw: str) -> str:
    """Which pair member a truncation was derived from (best guess by
    shared prefix length)."""
    names = [n for pair in AMBIGUOUS_PAIRS for n in pair]
    return max(names, key=lambda n: len(_prefix(raw.lower(), n.lower())))


def _prefix(a: str, b: str) -> str:
    out = ""
    for x, y in zip(a, b):
        if x != y:
            break
        out += x
    return out


# ---------------------------------------------------------------- masters
def make_masters(out_dir: Path, customers: list[dict]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Customers"
    ws.append(["customer_id", "customer_name", "city", "terms"])
    for c in customers:
        ws.append([c["customer_id"], c["customer_name"], c["city"], c["terms"]])
    wb.save(out_dir / "customer_master.xlsx")

    wb = Workbook()
    ws = wb.active
    ws.title = "Vendors"
    ws.append(["vendor_code", "vendor_name", "tax_id", "terms"])
    for v in vendor_master():
        ws.append([v["vendor_code"], v["vendor_name"], v["tax_id"], v["terms"]])
    wb.save(out_dir / "vendor_master.xlsx")


def main(out_dir: Path | None = None) -> Path:
    rng = random.Random(SEED)
    root = out_dir or Path(__file__).resolve().parent.parent / "input"
    inv_dir = root / "invoices"
    inv_dir.mkdir(parents=True, exist_ok=True)
    # deterministic PDFs: no creation-date metadata between regenerations
    import reportlab
    reportlab.rl_config.invariant = 1
    customers = customers_master()
    make_masters(root, customers)

    specs, inv_problems = make_invoices(inv_dir, rng)
    pos_rows, pos_problems, pos_truth, pos_planted = make_pos_xlsx(
        root / "pos_sales_export.xlsx", customers, rng)
    online_rows, online_problems, online_truth, online_planted = \
        make_online_csv(root / "online_sales_export.csv", customers, rng)

    # --------------------------------------------------- ground truth file
    # one row per distinct messy name; a name planted several times (or in
    # both exports) appears once
    truth_rows: dict[str, dict] = {}
    for t in (pos_truth, online_truth):
        for raw, rec in t.items():
            truth_rows.setdefault(
                raw, {"raw_name": raw, "truth": rec["truth"],
                      "planted_kind": rec["planted_kind"]})
    truth_rows = sorted(truth_rows.values(),
                        key=lambda r: r["raw_name"])
    with (root / "matching_truth.csv").open("w", newline="",
                                            encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["raw_name", "truth", "planted_kind"],
                           lineterminator="\n")
        w.writeheader()
        w.writerows(truth_rows)

    problems: dict[str, dict] = {}
    for source in (inv_problems, pos_problems, online_problems):
        for rule, ids in source.items():
            problems.setdefault(rule, {"count": 0, "ids": []})
            problems[rule]["ids"] += ids

    for rule in problems.values():
        rule["ids"] = sorted(set(rule["ids"]))
        rule["count"] = len(rule["ids"])

    # customer names: the truth planted, not a prediction of the matcher.
    # Tests score the pipeline against this (see matching_eval tests).
    matching_expectation = {
        "typo": sorted(set(pos_planted.get("typo", [])
                           + online_planted.get("typo", []))),
        "ambiguous": sorted(set(pos_planted.get("ambiguous", [])
                                + online_planted.get("ambiguous", []))),
        "unknown": sorted(set(pos_planted.get("unknown", [])
                              + online_planted.get("unknown", []))),
    }

    expected = {
        "reference_date": "2026-04-30",
        "seed": SEED,
        "counts": {
            "invoice_pdf_files": len(list(inv_dir.glob("*.pdf"))),
            "text_invoices": len(specs) + len(inv_problems["duplicate_invoice"]),
            "scanned_pdfs": len(inv_problems["scanned_pdf"]),
            "pos_sheet_rows": pos_rows,
            "online_rows": online_rows,
            "problem_rule_count": len(problems),
        },
        # per-invoice truth, generated before the PDFs were drawn: tests
        # compare parsed output against this, not against itself
        "invoices": {s["number"]: {
            "date": s["date"].isoformat(),
            "line_sum": str(s["line_sum"]),
            "tax": str(s["tax"]),
            "total": str(s["total"]),
            "n_lines": len(s["lines"]),
            "kind": "credit_note" if s["vendor"] == "CRS" else "invoice",
        } for s in specs},
        "problems": problems,
        "matching": matching_expectation,
    }
    exp_path = Path(__file__).resolve().parent / "expected.json"
    with exp_path.open("w", encoding="utf-8", newline="\n") as f:
        f.write(json.dumps(expected, indent=2) + "\n")

    # the input folder must contain exactly what this run generated:
    # remove files an older run left behind, refuse foreign files.
    # `own` comes from what was written this run, never from a listing -
    # a foreign file must not be able to adopt itself.
    own = (["matching_truth.csv", "pos_sales_export.xlsx",
            "online_sales_export.csv", "customer_master.xlsx",
            "vendor_master.xlsx"]
           + [f"invoices/{s['number']}.pdf" for s in specs]
           + [f"invoices/{alias}" for _, alias in DUP_ALIASES]
           + [f"invoices/{n}" for n in inv_problems["scanned_pdf"]])
    manifest_path = root / ".generated_files.json"
    owned_prev = set()
    if manifest_path.exists():
        try:
            owned_prev = set(json.loads(manifest_path.read_text()))
        except json.JSONDecodeError:
            owned_prev = set()
    current = {p.relative_to(root).as_posix()
               for p in root.rglob("*") if p.is_file()
               and p.name != ".generated_files.json"}
    for stale in owned_prev - set(own):
        if stale in current:
            (root / stale).unlink()
            current.discard(stale)
    foreign = current - set(own)
    if foreign:
        raise RuntimeError(
            f"input folder contains files the generator does not own: "
            f"{sorted(foreign)[:5]} - refusing to mix data sets")
    manifest_path.write_text(json.dumps(own, indent=0) + "\n")

    print(f"wrote {inv_dir} ({len(list(inv_dir.glob('*.pdf')))} pdfs), "
          f"pos {pos_rows} rows, online {online_rows} rows")
    print(f"wrote {exp_path} and {root / 'matching_truth.csv'} "
          f"({len(truth_rows)} names)")
    return root


if __name__ == "__main__":
    import sys
    main(Path(sys.argv[1]) if len(sys.argv) > 1 else None)
