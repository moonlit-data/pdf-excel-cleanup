"""Extract invoice data from PDF files, driven entirely by config.yml."""
from __future__ import annotations

import datetime as dt
import re
from decimal import Decimal
from pathlib import Path

import pdfplumber

from .config import Config, VendorConfig
from .money import cents, parse_amount as _parse_money


def parse_amount(text: str, number_format: str) -> Decimal | None:
    """Strictly parse an invoice amount per the vendor's declared format.
    Returns a Decimal parsed from the original string, or None."""
    return _parse_money(text, number_format)


def parse_date(text: str, formats: list) -> dt.date | None:
    s = text.strip()
    for fmt in formats:
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def page_texts(path: Path) -> list[str]:
    texts = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            texts.append(page.extract_text() or "")
    return texts


def _rows_from_table(table: list[list], page_no: int,
                     vendor: VendorConfig,
                     warnings: list[str]) -> list[dict]:
    """One extracted pdfplumber table -> line-item dicts.

    Rows are consumed only after a row matching the vendor's `header`
    regex (checked on the joined cell text) and stop at the `end` regex or
    an empty row. A non-empty row that fails to parse is logged as a
    warning instead of being silently dropped.
    """
    cols = vendor.lines.get("columns", ["desc", "qty", "unit", "amount"])
    items = []
    in_table = False
    for r in table:
        cells = [re.sub(r"\s+", " ", c or "").strip() for c in r]
        joined = " ".join(cells)
        if not in_table:
            if re.search(vendor.lines["header"], joined):
                in_table = True
            continue
        if re.search(vendor.lines["end"], joined) or not any(cells):
            break
        vals = dict(zip(cols, cells))
        desc = vals.get("desc", "")
        qty = unit = amount = None
        ok = True
        try:
            qty = int(vals["qty"])
        except (KeyError, ValueError, TypeError):
            ok = "qty" not in cols
        if "unit" in cols:
            unit = parse_amount(vals.get("unit", ""),
                                vendor.number_format)
            ok = ok and unit is not None
        if "amount" in cols:
            amount = parse_amount(vals.get("amount", ""),
                                  vendor.number_format)
            ok = ok and amount is not None
        if not ok or not desc:
            # a subtotal/footer row folded into the grid, or a data row
            # that failed to parse — either way, keep it auditable
            warnings.append(
                f"page {page_no}: unparseable table row {joined[:80]!r}")
            continue
        kind = "item"
        if vendor.lines.get("discount_desc") and re.search(
                vendor.lines["discount_desc"], desc):
            kind = "discount"
        elif vendor.lines.get("shipping_desc") and re.search(
                vendor.lines["shipping_desc"], desc):
            kind = "shipping"
        items.append({"page": page_no, "description": desc,
                      "quantity": qty, "unit_price": unit,
                      "amount": amount, "kind": kind})
    return items


def _rows_from_text_page(text: str, page_no: int, in_table: bool,
                         vendor: VendorConfig) -> tuple[list[dict], bool]:
    """Row-regex fallback for one page's text. Returns (items, in_table)."""
    items = []
    for raw_line in text.split("\n"):
        line = raw_line.rstrip()
        if not in_table:
            if re.search(vendor.lines["header"], line):
                in_table = True
            continue
        if re.search(vendor.lines["end"], line):
            in_table = False
            break
        m = re.match(vendor.lines.get("row", r"$^"), line)
        if m:
            unit = parse_amount(m.group("unit"), vendor.number_format)
            amount = parse_amount(m.group("amount"), vendor.number_format)
            if amount is None:
                continue
            desc = m.group("desc").strip()
            kind = "item"
            if vendor.lines.get("discount_desc") and re.search(
                    vendor.lines["discount_desc"], desc):
                kind = "discount"
            elif vendor.lines.get("shipping_desc") and re.search(
                    vendor.lines["shipping_desc"], desc):
                kind = "shipping"
            items.append({"page": page_no, "description": desc,
                          "quantity": int(m.group("qty")),
                          "unit_price": unit, "amount": amount,
                          "kind": kind})
    return items, in_table


def _rows_from_text(texts: list[str], vendor: VendorConfig) -> list[dict]:
    items = []
    in_table = False
    for page_no, text in enumerate(texts, 1):
        page_items, in_table = _rows_from_text_page(
            text, page_no, in_table, vendor)
        items += page_items
    return items


def detect_vendor(text: str, config: Config) -> VendorConfig | None:
    for v in config.vendors.values():
        if v.detect in text:
            return v
    return None


def _first_group(pattern: str, text: str) -> str | None:
    m = re.search(pattern, text)
    return m.group(1) if m else None


def _line_key(l: dict) -> tuple:
    return (l["description"], l["quantity"], l["unit_price"], l["amount"])


def parse_invoice_pdf(path: Path, config: Config) -> dict:
    """Parse one PDF. Returns a record dict; never raises for bad input.

    status: ok | needs_review | scanned | unreadable | unknown_layout.
    Missing required fields (invoice number, total) never produce a
    silent "ok" — the record is marked needs_review so the dedupe stage
    cannot drop it on an empty key.
    """
    rec = {"source_file": path.name, "pages": 0, "status": "ok",
           "vendor_code": "", "invoice_number": "", "invoice_date": None,
           "tax_id": "", "total": None, "line_sum": None, "kind": "invoice",
           "lines": [], "errors": [], "warnings": []}
    try:
        with pdfplumber.open(path) as pdf:
            pages = list(pdf.pages)
            texts = [p.extract_text() or "" for p in pages]
            # table extraction happens inside the open block: pdfplumber
            # may lazily read file content when building table objects
            strategy_pages: dict[int, list[list]] = {}
            vendor_probe = detect_vendor("\n".join(texts), config)
            if vendor_probe and vendor_probe.lines.get(
                    "strategy", "text") == "table":
                settings = vendor_probe.lines.get(
                    "table_settings",
                    {"vertical_strategy": "lines",
                     "horizontal_strategy": "lines"})
                for page_no, page in enumerate(pages, 1):
                    strategy_pages[page_no] = page.extract_tables(settings)
    except Exception as exc:  # unreadable file
        rec["status"] = "unreadable"
        rec["errors"].append(str(exc))
        return rec
    rec["pages"] = len(texts)
    full = "\n".join(texts)
    if not full.strip():
        rec["status"] = "scanned"
        return rec
    vendor = detect_vendor(full, config)
    if vendor is None:
        rec["status"] = "unknown_layout"
        return rec
    rec["vendor_code"] = vendor.code
    rec["vendor_name"] = vendor.name
    rec["kind"] = vendor.document_kind
    f = vendor.fields

    num = _first_group(f["invoice_number"], full)
    if num:
        rec["invoice_number"] = num
    else:
        rec["status"] = "needs_review"
        rec["errors"].append("invoice_number not found")
    date_s = _first_group(f["invoice_date"], full)
    d = parse_date(date_s, vendor.date_formats) if date_s else None
    rec["invoice_date"] = d.isoformat() if d else None
    rec["tax_id"] = _first_group(f.get("tax_id", r"$^"), full) or ""
    total_s = _first_group(f["total"], full)
    if total_s is not None:
        rec["total"] = parse_amount(total_s, vendor.number_format)
        if rec["total"] is None:
            rec["status"] = "needs_review"
            rec["errors"].append(f"total {total_s!r} unparseable")
    else:
        rec["status"] = "needs_review"
        rec["errors"].append("total not found")
    if f.get("tax"):
        tax_s = _first_group(f["tax"], full)
        rec["tax_amount"] = (parse_amount(tax_s, vendor.number_format)
                             if tax_s else Decimal("0"))
        if tax_s and rec["tax_amount"] is None:
            rec["warnings"].append(f"tax {tax_s!r} unparseable")
            rec["tax_amount"] = Decimal("0")
    else:
        rec["tax_amount"] = Decimal("0")

    # line items: `strategy: table` uses pdfplumber's ruled-table
    # extraction. If a page's tables yield nothing but its text contains
    # parseable rows, the text fallback runs for that page (deduped), so
    # a page whose grid isn't detected can't silently lose its lines.
    if strategy_pages:
        table_keys: set[tuple] = set()
        for page_no, tables in sorted(strategy_pages.items()):
            page_lines: list[dict] = []
            for table in tables:
                page_lines += _rows_from_table(table, page_no, vendor,
                                               rec["warnings"])
            # table rows are authoritative: identical line items on one
            # invoice are legal and must be kept. Dedupe only against
            # the text fallback so the same row is not counted twice.
            table_keys.update(_line_key(l) for l in page_lines)
            text_lines, _ = _rows_from_text_page(
                texts[page_no - 1], page_no, False, vendor)
            page_lines += [l for l in text_lines
                           if _line_key(l) not in table_keys]
            rec["lines"] += page_lines
        if not rec["lines"] and vendor.lines.get("strategy") != "table":
            rec["lines"] = _rows_from_text(texts, vendor)
    else:
        rec["lines"] = _rows_from_text(texts, vendor)
    rec["line_sum"] = cents(sum(l["amount"] for l in rec["lines"]))
    if vendor.document_kind == "credit_note" and rec["total"] is not None:
        rec["total"] = -abs(rec["total"])
    return rec
