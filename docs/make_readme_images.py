"""Render the README images. No manual screenshots: everything here is
drawn by code from the real input/output files.

    python docs/make_readme_images.py

Produces in docs/images/:
  invoice_detail.png        zoomed crop of a messy PDF invoice next to the
                            same invoice's rows in the clean workbook
  pos_before_after.png      8 messy POS export rows next to the same rows
                            in the clean Sales sheet (Excel-styled)
  problems_by_type_light.png / _dark.png
  match_methods_light.png / _dark.png
"""
from __future__ import annotations

import csv
import datetime as dt
import json
from collections import Counter
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pdfplumber
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
IMG = ROOT / "docs" / "images"
INPUT = ROOT / "input"
OUT = ROOT / "output"
RESULTS = ROOT / "results"

LIGHT = {"bg": "#fcfcfb", "fg": "#0b0b0b", "sec": "#52514e", "grid": "#e1e0d9",
         "cats": ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4",
                  "#008300", "#6250d6", "#e34948"]}
DARK = {"bg": "#1a1a19", "fg": "#f0efec", "sec": "#c3c2b7", "grid": "#2c2c2a",
        "cats": ["#3987e5", "#d95926", "#199e70", "#c98500", "#d55181",
                 "#008300", "#9085e9", "#e66767"]}
HDR_BG = "#2a78d6"      # same fill the workbook writer uses
HDR_FG = "#ffffff"
ROW_ALT = "#eef3fb"


def _style(ax, t):
    ax.set_facecolor(t["bg"])
    for s in ax.spines.values():
        s.set_visible(False)
    ax.tick_params(colors=t["sec"], labelsize=9)
    ax.grid(axis="x", color=t["grid"], lw=0.8)
    ax.set_axisbelow(True)


def _fit_widths(headers, data, fontsize):
    """Column widths sized to the longest rendered value per column, plus
    ~10 px padding each side at 200 dpi (PAD alone can't stop clipped
    text when a column is too narrow)."""
    fig = plt.figure(figsize=(10, 1), dpi=200)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.axis("off")
    cols = len(headers)
    texts = [ax.text(0, 0, s, fontsize=fontsize) for c in range(cols)
             for s in [headers[c]] + [str(r[c]) for r in data]]
    fig.canvas.draw()
    it = iter(texts)
    widths = []
    for c in range(cols):
        mx = max(t.get_window_extent().width for t in
                 [next(it) for _ in range(len(data) + 1)])
        widths.append(mx + 20.0)  # ~10px either side
    plt.close(fig)
    s = sum(widths)
    return [w / s for w in widths]


def _excel_table(ax, headers, data, col_widths, t, fontsize=10,
                 aligns=None):
    """Draw a table that looks like the clean workbook: blue header row,
    white bold labels, thin grid, alternating rows, wide enough that no
    cell is truncated. fig is ~200 dpi so fontsize 10 -> ~28 px glyphs."""
    ax.axis("off")
    tab = ax.table(cellText=data, colLabels=headers, colWidths=col_widths,
                   loc="center", cellLoc="center")
    tab.auto_set_font_size(False)
    tab.set_fontsize(fontsize)
    aligns = aligns or ["center"] * len(headers)
    for (r, c), cell in tab.get_celld().items():
        cell.set_edgecolor("#c9d4e4" if t is LIGHT else "#3a3a38")
        cell.set_height(0.11)
        if r == 0:
            cell.set_facecolor(HDR_BG)
            cell.set_text_props(color=HDR_FG, weight="bold")
        else:
            cell.set_facecolor(t["bg"] if r % 2 else ROW_ALT)
            cell.set_text_props(color=t["fg"])
        cell.PAD = 0.045
        cell.get_text().set_ha(aligns[c])


# ------------------------------------------------------------- image (a)
def invoice_detail() -> None:
    """NPP-1001: the drawn PDF next to the same invoice's clean rows."""
    number = "NPP-1001"
    with pdfplumber.open(INPUT / "invoices" / f"{number}.pdf") as pdf:
        page = pdf.pages[0]
        # crop tight around the letterhead, line table and totals box
        im = page.crop((55, 25, 565, 340)).to_image(resolution=200).original

    inv = next(r for r in csv.DictReader(open(OUT / "invoices.csv"))
               if r["invoice_number"] == number)
    lines = [r for r in csv.DictReader(open(OUT / "invoice_lines.csv"))
             if r["invoice_number"] == number]

    t = LIGHT
    fig = plt.figure(figsize=(13.2, 6.0), dpi=200, facecolor=t["bg"])
    ax1 = fig.add_axes([0.005, 0.02, 0.46, 0.93])
    ax1.imshow(im)
    ax1.set_title("Supplier invoice PDF (before)", color=t["fg"], fontsize=13)
    ax1.axis("off")

    ax2 = fig.add_axes([0.475, 0.02, 0.52, 0.93])
    ax2.set_title("Clean Invoices / Invoice_Lines rows (after)",
                  color=t["fg"], fontsize=13)
    meta = [[inv["invoice_number"], inv["vendor_name"], inv["invoice_date"],
             f"${float(inv['total']):,.2f}"]]
    # draw the invoice header row as its own little table, then the lines
    tab1 = ax2.table(cellText=meta, colLabels=["invoice", "vendor", "date",
                                             "total"],
                     colWidths=_fit_widths(
                         ["invoice", "vendor", "date", "total"], meta, 9.5),
                     bbox=[0.0, 0.87, 1.0, 0.12])
    tab1.auto_set_font_size(False)
    tab1.set_fontsize(9.5)
    for (r, c), cell in tab1.get_celld().items():
        cell.set_edgecolor("#c9d4e4")
        cell.PAD = 0.045
        if r == 0:
            cell.set_facecolor(HDR_BG)
            cell.set_text_props(color=HDR_FG, weight="bold")
        else:
            cell.set_facecolor(ROW_ALT)
    data = [[l["description"], l["quantity"],
             f"{float(l['unit_price']):,.2f}",
             f"{float(l['amount']):,.2f}"] for l in lines]
    rows_h = 0.115 * (len(data) + 1)
    tab2 = ax2.table(cellText=data,
                     colLabels=["item", "qty", "unit price", "amount"],
                     colWidths=_fit_widths(
                         ["item", "qty", "unit price", "amount"],
                         data, 9.5),
                     bbox=[0.0, 0.85 - rows_h, 1.0, rows_h])
    tab2.auto_set_font_size(False)
    tab2.set_fontsize(9.5)
    for (r, c), cell in tab2.get_celld().items():
        cell.set_edgecolor("#c9d4e4")
        cell.PAD = 0.045
        cell.get_text().set_ha("left" if c == 0 else "center")
        if r == 0:
            cell.set_facecolor(HDR_BG)
            cell.set_text_props(color=HDR_FG, weight="bold")
            cell.get_text().set_ha("left" if c == 0 else "center")
        else:
            cell.set_facecolor(t["bg"] if r % 2 else ROW_ALT)
    ax2.axis("off")
    fig.savefig(IMG / "invoice_detail.png", facecolor=t["bg"])
    plt.close(fig)


# ------------------------------------------------------------- image (b)
def pos_before_after() -> None:
    """8 rows x 6 cols of the messy POS export vs the same clean rows."""
    pos = load_workbook(INPUT / "pos_sales_export.xlsx").active
    sales = {r["record_id"]: r for r in csv.DictReader(open(OUT / "sales.csv"))
             if r["record_id"].startswith("POS-")}
    # walk the sheet, grab the first 8 rows that survived into sales.csv
    messy, clean = [], []
    for idx, row in enumerate(pos.iter_rows(min_row=6, values_only=True)):
        sheet_row = idx + 6
        s = sales.get(f"POS-R{sheet_row}")
        if s is None:
            continue
        vals = [row[0], row[1], row[3], row[4], row[5], row[7]]
        messy.append(["" if v is None else str(v) for v in vals])
        clean.append([s["order_date"], s["channel"], s["product"],
                      s["quantity"], f"{float(s['amount']):,.2f}",
                      s["customer_name"] or "(review)"])
        if len(messy) >= 8:
            break

    t = LIGHT
    fig = plt.figure(figsize=(13.2, 4.4), dpi=200, facecolor=t["bg"])
    headers = ["Date", "Region", "Product", "Qty", "Revenue", "Customer"]
    ax1 = fig.add_axes([0.01, 0.10, 0.47, 0.80])
    ax1.set_title("POS export, raw (before)", color=t["fg"], fontsize=13)
    _excel_table(ax1, headers, messy, _fit_widths(headers, messy, 10), t,
                 fontsize=10,
                 aligns=["left", "left", "left", "center", "right", "left"])
    ax2 = fig.add_axes([0.52, 0.10, 0.47, 0.80])
    ax2.set_title("Clean Sales sheet (after)", color=t["fg"], fontsize=13)
    _excel_table(ax2, headers, clean, _fit_widths(headers, clean, 10), t,
                 fontsize=10,
                 aligns=["center", "center", "left", "center", "right", "left"])
    fig.savefig(IMG / "pos_before_after.png", facecolor=t["bg"])
    plt.close(fig)


# ------------------------------------------------------------- charts
def _barh(title: str, labels: list[str], values: list[int],
          t: dict, out: Path) -> None:
    fig = plt.figure(figsize=(6.4, 3.2), dpi=200, facecolor=t["bg"])
    ax = fig.add_subplot(111)
    colors = t["cats"][:len(labels)]
    ax.barh(labels[::-1], values[::-1], color=colors[::-1], height=0.7)
    for i, v in enumerate(values[::-1]):
        ax.text(v + max(values) * 0.01, i, str(v), va="center",
                color=t["sec"], fontsize=9)
    ax.set_title(title, color=t["fg"], fontsize=11, loc="left", pad=10)
    ax.set_xlabel("rows", color=t["sec"], fontsize=9)
    _style(ax, t)
    fig.tight_layout()
    fig.savefig(out, facecolor=t["bg"])
    plt.close(fig)


def problems_chart() -> None:
    """Problems only — routine reformatting is a separate count, not a bar."""
    counts = Counter()
    with (RESULTS / "problems_by_type.csv").open() as f:
        for r in csv.DictReader(f):
            counts[r["rule"]] = int(r["count"])
    top = counts.most_common(6)
    other = sum(c for _, c in counts.most_common()[6:])
    labels = [r for r, _ in top] + (["other"] if other else [])
    values = [c for _, c in top] + ([other] if other else [])
    summary = json.loads((RESULTS / "summary.json").read_text())
    problems = summary["problems_found"]
    routine = summary["routine_fixes"]
    title = f"{problems:,} problems found; {routine:,} routine fixes omitted"
    for suffix, t in [("_light.png", LIGHT), ("_dark.png", DARK)]:
        _barh(title, labels, values, t, IMG / f"problems_by_type{suffix}")


def match_chart() -> None:
    methods = Counter()
    with (OUT / "customers_matched.csv").open() as f:
        for r in csv.DictReader(f):
            methods[r["method"]] += int(r["occurrences"])
    order = ["exact", "fuzzy_auto", "fuzzy_review", "unknown", "walkin"]
    labels = [m for m in order if methods[m]]
    values = [methods[m] for m in labels]
    m = json.loads((RESULTS / "matching_eval.json").read_text())
    title = (f"precision {m['precision']:.2f}, recall {m['recall']:.2f}; "
             f"{m['review_sent']} names for a human")
    for suffix, t in [("_light.png", LIGHT), ("_dark.png", DARK)]:
        _barh(title, labels, values, t, IMG / f"match_methods{suffix}")


if __name__ == "__main__":
    IMG.mkdir(parents=True, exist_ok=True)
    invoice_detail()
    pos_before_after()
    problems_chart()
    match_chart()
    print("wrote", IMG)
