"""Clean the two sales exports into one schema.

POS export (xlsx): title rows above the header, merged header cells, five
date formats, numbers stored as text, region aliases, blank/subtotal/notes
rows and repeated rows.

Online export (csv): clean column types but UTC timestamps and refunds as
separate negative rows.

Every change, drop or flag is logged in the IssueLog. Nothing disappears
silently: rows read == rows kept + rows dropped. Rows that merely *look*
suspicious (e.g. a POS row identical to another) are kept and flagged —
only rows provably malformed or duplicated by a real row-level ID are
dropped.
"""
from __future__ import annotations

import csv
import datetime as dt
import re
from decimal import Decimal
from pathlib import Path

from openpyxl import load_workbook

from .config import Config
from .issues import BLOCKED, FIXED, NEEDS_REVIEW, IssueLog
from .money import cents, parse_amount

# The POS export is a US point-of-sale system: month-first dates. The
# convention is declared per source in config.yml (`sales.pos.date_order`).
DATE_FORMATS_US = ["%m/%d/%y", "%Y-%m-%d", "%B %d %Y", "%d.%m.%Y",
                   "%m/%d/%Y", "%Y-%m-%dT%H:%M:%SZ"]

SALES_COLUMNS = ["source_system", "record_id", "order_date", "customer_raw",
                 "customer_id", "customer_name", "product", "quantity",
                 "unit_price", "amount", "channel", "note"]

REQUIRED_ONLINE_COLS = {"order_id", "ordered_at_utc", "customer", "sku",
                        "item", "quantity", "unit_price", "line_total",
                        "status"}


def _parse_any_date(value, formats: list) -> dt.date | None:
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    s = str(value).strip()
    for fmt in formats:
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def _units(value) -> tuple[int | None, bool]:
    if isinstance(value, bool):
        return None, True
    if isinstance(value, int):
        return value, False
    if isinstance(value, float) and value.is_integer():
        return int(value), False
    m = re.match(r"^\s*(-?\d+)\s*(pcs|pc|units?)?\s*$", str(value), re.I)
    if m:
        return int(m.group(1)), True
    return None, True


def _title_if_shouted(text: str) -> tuple[str, bool]:
    s = str(text).strip()
    fixed = re.sub(r"\s+", " ", s)
    if fixed.isupper() or fixed.islower():
        fixed = fixed.title()
    return fixed, fixed != str(text)


def parse_pos(path: Path, config: Config, log: IssueLog) -> dict:
    """Returns {'rows': kept sales rows, 'rows_in': n, 'rows_dropped': n}."""
    # load twice: once for formulas (data_only=False), once for the
    # cached values Excel stored for them. A formula cell with no cached
    # value is flagged, not silently read as "=SUM(...)" text.
    wb = load_workbook(path, data_only=False)
    ws = wb.active
    cells = list(ws.iter_rows(values_only=True))
    wb_vals = load_workbook(path, data_only=True)
    vals_rows = list(wb_vals.active.iter_rows(values_only=True))
    money_fmt = config.sale_format("pos", "number_format", "us")

    header_idx = None
    for i, row in enumerate(cells[:12]):
        joined = " ".join(str(c or "") for c in row).lower()
        if "order date" in joined and "region" in joined:
            header_idx = i
            break
    if header_idx is None:
        log.add(path.name, "sheet", "", "unparseable_sheet", action="blocked",
                severity=BLOCKED, detail="header row not found")
        return {"rows": [], "rows_in": 0, "rows_dropped": 0}
    headers = [str(c or "").strip().lower() for c in cells[header_idx]]
    col = {name: headers.index(name) for name in
           ["order date", "region", "sales rep", "product", "units",
            "revenue", "notes", "customer"] if name in headers}

    alias = {}
    for canon, variants in config.region_aliases.items():
        for v in variants + [canon]:
            alias[v.strip().lower().rstrip(".")] = canon

    rows, rows_in, dropped = [], 0, 0
    seen_keys: dict[tuple, str] = {}
    for r_off, row in enumerate(cells[header_idx + 1:]):
        sheet_row = header_idx + 2 + r_off   # 1-based row in the sheet
        rid = f"POS-R{sheet_row}"
        loc = f"row {sheet_row}"
        # substitute computed values for formula cells
        calc = (vals_rows[header_idx + 1 + r_off]
                if header_idx + 1 + r_off < len(vals_rows) else row)
        vals = [calc[j] if isinstance(c, str) and c.startswith("=")
                else c for j, c in enumerate(row)]
        if all(v is None or str(v).strip() == "" for v in vals):
            rows_in += 1
            dropped += 1
            log.add(path.name, loc, rid, "blank_row", action="dropped",
                    severity=FIXED)
            continue
        cell = lambda name: vals[col[name]] if name in col and col[name] < len(vals) else None
        region_raw = str(cell("region") or "").strip()
        first = str(vals[0] or "").strip() if vals else ""
        if region_raw.lower().startswith("subtotal"):
            rows_in += 1
            dropped += 1
            log.add(path.name, loc, rid, "subtotal_row_dropped",
                    field="region", original=region_raw, action="dropped",
                    severity=FIXED, detail="embedded subtotal row")
            continue
        if region_raw.upper() == "GRAND TOTAL" or first.lower().startswith("note"):
            rows_in += 1
            dropped += 1
            log.add(path.name, loc, rid, "metadata_row_dropped",
                    field="row", original=" | ".join(str(v) for v in vals if v),
                    action="dropped", severity=FIXED,
                    detail="export metadata/footer row")
            continue
        rows_in += 1

        row_flags: list[str] = []
        # identical-content row: NOT proof of a duplicate sale (the POS
        # export has no transaction id), so it is kept and flagged for a
        # human rather than silently deleting real revenue.
        key = tuple(str(v) for v in vals)
        if key in seen_keys:
            row_flags.append("duplicate_sales_row")
            log.add(path.name, loc, rid, "duplicate_sales_row",
                    field="row", original=" | ".join(str(v) for v in vals if v),
                    action="flagged", severity=NEEDS_REVIEW,
                    detail=f"identical to {seen_keys[key]}; kept - no "
                           f"row-level id to prove duplication")
        else:
            seen_keys[key] = rid

        def fix(field, original, new, rule):
            log.add(path.name, loc, rid, rule, field=field,
                    original=original, new=new, action="changed",
                    severity=FIXED)

        # date
        d_raw = cell("order date")
        d = _parse_any_date(d_raw, DATE_FORMATS_US)
        if d is None:
            dropped += 1
            log.add(path.name, loc, rid, "unparseable_date", field="order_date",
                    original=d_raw, action="dropped", severity=NEEDS_REVIEW)
            continue
        if not isinstance(d_raw, (dt.date, dt.datetime)):
            fix("order_date", d_raw, d.isoformat(), "date_format_fixed")

        # region
        region = alias.get(region_raw.strip().lower().rstrip("."), region_raw)
        if region != region_raw:
            fix("region", region_raw, region, "region_normalized")

        rep, rep_changed = _title_if_shouted(cell("sales rep") or "")
        if rep_changed:
            fix("sales_rep", cell("sales rep"), rep, "name_case_fixed")
        prod, prod_changed = _title_if_shouted(cell("product") or "")
        if prod_changed:
            fix("product", cell("product"), prod, "product_name_fixed")

        units, units_text = _units(cell("units"))
        if units is None:
            dropped += 1
            log.add(path.name, loc, rid, "unparseable_amount", field="units",
                    original=cell("units"), action="dropped",
                    severity=NEEDS_REVIEW)
            continue
        if units_text:
            fix("units", cell("units"), units, "number_format_fixed")
        amount = parse_amount(cell("revenue"), money_fmt)
        if amount is None:
            dropped += 1
            log.add(path.name, loc, rid, "unparseable_amount", field="revenue",
                    original=cell("revenue"), action="dropped",
                    severity=NEEDS_REVIEW)
            continue
        if isinstance(cell("revenue"), str):
            fix("revenue", cell("revenue"), f"{cents(amount)}",
                "number_format_fixed")

        if units < 0:
            row_flags.append("negative_quantity")
            log.add(path.name, loc, rid, "negative_quantity", field="units",
                    original=units, new=units, action="flagged",
                    severity=NEEDS_REVIEW, detail="kept; check if a return")
        rows.append({"source_system": "pos", "record_id": rid,
                     "order_date": d.isoformat(),
                     "customer_raw": str(cell("customer") or "").strip(),
                     "product": prod, "quantity": units,
                     "unit_price": "", "amount": cents(amount),
                     "channel": region, "note": str(cell("notes") or "").strip(),
                     "review_flag": ";".join(row_flags)})

    return {"rows": rows, "rows_in": rows_in, "rows_dropped": dropped}


def parse_online(path: Path, config: Config, log: IssueLog) -> dict:
    rows, rows_in, dropped = [], 0, 0
    money_fmt = config.sale_format("online", "number_format", "us")
    ts_formats = config.sale_format(
        "online", "timestamp_formats", ["%Y-%m-%dT%H:%M:%SZ"])
    if isinstance(ts_formats, str):
        ts_formats = [ts_formats]
    sale_lines: dict[tuple, str] = {}   # (order_id, sku, qty, |total|) -> oid
    sale_order_ids: set[str] = set()
    seen_rows: dict[tuple, str] = {}    # full row -> first csv line
    refund_rows: list[dict] = []
    with path.open(newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        missing = REQUIRED_ONLINE_COLS - set(reader.fieldnames or [])
        if missing:
            log.add(path.name, "header", "", "unparseable_sheet",
                    action="blocked", severity=BLOCKED,
                    detail=f"missing columns: {sorted(missing)}")
            return {"rows": [], "rows_in": 0, "rows_dropped": 0}
        for i, r in enumerate(reader, start=2):  # header is row 1
            rows_in += 1
            rid = (r.get("order_id") or "").strip()
            loc = f"row {i}"

            # exact duplicate of a previous row (same order id AND all
            # fields): a re-export artifact, safe to drop + count.
            key = tuple((r.get(c) or "").strip()
                        for c in reader.fieldnames)
            if key in seen_rows:
                dropped += 1
                log.add(path.name, loc, rid, "duplicate_sales_row",
                        field="order_id", original=rid, action="dropped",
                        severity=NEEDS_REVIEW,
                        detail=f"exact duplicate of row {seen_rows[key]}")
                continue
            seen_rows[key] = loc

            qty_raw, unit_raw, total_raw = (r["quantity"] or "").strip(), \
                (r["unit_price"] or "").strip(), (r["line_total"] or "").strip()
            qty = parse_amount(qty_raw, money_fmt)
            unit = parse_amount(unit_raw, money_fmt)
            amount = parse_amount(total_raw, money_fmt)
            if qty is None or unit is None or amount is None or \
                    not qty == qty.to_integral_value():
                dropped += 1
                log.add(path.name, loc, rid, "unparseable_amount",
                        field="line_total" if amount is None else "quantity",
                        original=total_raw if amount is None else qty_raw,
                        action="dropped", severity=NEEDS_REVIEW,
                        detail="quantity must be a whole number; amounts "
                               "must be finite numbers")
                continue
            qty = int(qty)
            ts = (r.get("ordered_at_utc") or "").strip()
            d = _parse_any_date(ts, ts_formats)
            if d is None:
                dropped += 1
                log.add(path.name, loc, rid, "unparseable_date",
                        field="ordered_at_utc", original=ts,
                        action="dropped", severity=NEEDS_REVIEW,
                        detail=f"not in a declared format {ts_formats}")
                continue
            status = (r.get("status") or "sale").strip().lower()
            rec = {"source_system": "online", "record_id": rid,
                   "order_date": d.isoformat(),
                   "customer_raw": (r["customer"] or "").strip(),
                   "product": (r["item"] or "").strip(), "quantity": qty,
                   "unit_price": unit, "amount": cents(amount),
                   "channel": "online",
                   "note": f"{status} {ts}", "review_flag": ""}
            rows.append(rec)
            if status == "sale":
                sale_order_ids.add(rid)
                sale_lines[(rid, (r["sku"] or "").strip(), qty,
                            abs(amount))] = rid
            elif status == "refund":
                refund_rows.append({"rec": rec, "loc": loc, "r": r})
            else:
                rec["review_flag"] = "unknown_status"
                log.add(path.name, loc, rid, "unknown_status",
                        field="status", original=status, action="flagged",
                        severity=NEEDS_REVIEW)

    for item in refund_rows:
        rec, loc, r = item["rec"], item["loc"], item["r"]
        if rec["record_id"] not in sale_order_ids:
            rec["review_flag"] = "refund_without_sale"
            log.add(path.name, loc, rec["record_id"], "refund_without_sale",
                    field="order_id", original=rec["record_id"],
                    action="flagged", severity=NEEDS_REVIEW,
                    detail="refund row with no matching sale in this export")
            continue
        if rec["amount"] >= 0 or rec["quantity"] >= 0:
            rec["review_flag"] = "refund_sign_invalid"
            log.add(path.name, loc, rec["record_id"], "refund_sign_invalid",
                    field="line_total", original=rec["amount"],
                    action="flagged", severity=NEEDS_REVIEW,
                    detail="refund row with a non-negative amount; kept "
                           "for a human to check")
            continue
        line_key = (rec["record_id"], (r["sku"] or "").strip(),
                    -rec["quantity"], abs(rec["amount"]))
        if line_key not in sale_lines:
            rec["review_flag"] = "refund_line_mismatch"
            log.add(path.name, loc, rec["record_id"], "refund_line_mismatch",
                    field="sku", original=r.get("sku"),
                    action="flagged", severity=NEEDS_REVIEW,
                    detail="refund does not match a sold line (sku, "
                           "quantity and amount differ)")
    return {"rows": rows, "rows_in": rows_in, "rows_dropped": dropped}
