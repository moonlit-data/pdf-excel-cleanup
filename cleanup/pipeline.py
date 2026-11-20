"""Run the whole clean-up: inputs -> records -> Issues ledger -> outputs."""
from __future__ import annotations

import datetime as dt
import re
import time
from pathlib import Path

from decimal import Decimal

from . import outputs
from .config import Config
from .issues import BLOCKED, FIXED, NEEDS_REVIEW, IssueLog, is_problem
from .manifest import Manifest, cache_fingerprint, sha256_file
from .matching import CustomerMatcher, evaluate
from .money import cents
from .pdf_extract import parse_invoice_pdf
from .sales import parse_online, parse_pos

_MONEY_KEYS = {"total", "line_sum", "total_diff", "tax_amount",
               "unit_price", "amount"}


def _restore_decimals(payload: dict) -> dict:
    """JSON caches store Decimals as strings; turn them back."""
    def fix(obj):
        if isinstance(obj, dict):
            return {k: (Decimal(v) if k in _MONEY_KEYS
                        and isinstance(v, str)
                        and re.fullmatch(r"-?\d+(\.\d+)?", v) else fix(v))
                    for k, v in obj.items()}
        if isinstance(obj, list):
            return [fix(x) for x in obj]
        return obj
    return fix(payload)


def _read_master(path: Path, key: str) -> list[dict]:
    from openpyxl import load_workbook
    ws = load_workbook(path).active
    rows = list(ws.iter_rows(values_only=True))
    headers = [str(h) for h in rows[0]]
    return [dict(zip(headers, r)) for r in rows[1:] if any(r)]


def _find(input_dir: Path, stem: str, suffix: str) -> Path | None:
    p = input_dir / f"{stem}{suffix}"
    return p if p.exists() else None


class _Processor:
    """Per-file cache wrapper: unchanged files replay their cached records
    and issues instead of being re-parsed."""

    def __init__(self, manifest: Manifest, log: IssueLog, rebuild: bool):
        self.manifest = manifest
        self.log = log
        self.rebuild = rebuild
        self.file_status: list[dict] = []

    def run(self, path: Path, func) -> tuple[dict, str]:
        sha = sha256_file(path)
        cached = None if self.rebuild else self.manifest.lookup(path, sha)
        if cached is None:
            payload = func()
            self.manifest.store(path, sha, payload)
            status = "processed"
        else:
            payload = _restore_decimals(cached)
            status = "skipped"
            # byte-identical files share one cache entry keyed by sha;
            # restore this file's own name on replay
            for rec in payload.get("records", []):
                if "source_file" in rec:
                    rec["source_file"] = path.name
        for iss in payload.get("issues", []):
            self.log.add(iss["source_file"], iss["location"], iss["record_id"],
                         iss["rule"], field=iss["field"],
                         original=iss["original_value"], new=iss["new_value"],
                         action=iss["action"], severity=iss["severity"],
                         detail=iss["detail"])
        self.file_status.append({"file": path.name, "sha256": sha[:16],
                                 "status": status})
        return payload, status


def run(input_dir: Path, out_dir: Path, config_path: Path,
        rebuild: bool = False, results_dir: Path | None = None) -> dict:
    t0 = time.time()
    config = Config.load(config_path)
    input_dir, out_dir = Path(input_dir), Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    log = IssueLog()
    manifest = Manifest(out_dir, cache_fingerprint(config_path))
    if rebuild:
        manifest.clear()          # ignore - and delete - any old manifest
    else:
        manifest.load()           # corrupt/stale manifests just miss
    proc = _Processor(manifest, log, rebuild)

    customers = _read_master(_find(input_dir, config.inputs["customer_master"], ".xlsx"), "customer")
    vendors = {v["vendor_code"]: v for v in
               _read_master(_find(input_dir, config.inputs["vendor_master"], ".xlsx"), "vendor")}
    matcher = CustomerMatcher(customers, config.auto_apply_threshold,
                              config.review_threshold, config.match_margin)

    # ------------------------------------------------ invoices
    invoice_payloads = []
    inv_dir = input_dir / "invoices"
    pdf_files = sorted(inv_dir.glob("*.pdf")) if inv_dir.exists() else []
    for pdf in pdf_files:
        payload, _ = proc.run(pdf, lambda p=pdf: {
            "records": [parse_invoice_pdf(p, config)], "issues": []})
        invoice_payloads.append(payload["records"][0])

    # de-duplicate on a RELIABLE key only: (vendor, invoice number) must
    # both be present. Records without one stay and are flagged, never
    # dropped on an empty key. When two records share a key, identical
    # content is a re-send (drop); conflicting content is kept both and
    # flagged — alphabetical order must not choose which invoice dies.
    seen: dict[tuple, dict] = {}
    invoices = []
    for rec in invoice_payloads:
        if rec["status"] in ("scanned", "unreadable", "unknown_layout"):
            rule = {"scanned": "scanned_pdf", "unreadable": "unreadable_file",
                    "unknown_layout": "unknown_layout"}[rec["status"]]
            detail = ("no extractable text - needs manual entry or OCR"
                      if rec["status"] == "scanned" else "; ".join(rec["errors"]))
            log.add(rec["source_file"], "file", rec["source_file"], rule,
                    action="blocked", severity=BLOCKED, detail=detail)
            continue
        if rec["status"] == "needs_review":
            log.add(rec["source_file"], "file", rec["source_file"],
                    "missing_invoice_field",
                    field=rec["errors"][0].split()[0] if rec["errors"]
                    else "invoice",
                    action="flagged", severity=NEEDS_REVIEW,
                    detail="; ".join(rec["errors"]) or
                    "required field could not be extracted")
            invoices.append(rec)
            continue
        key = (rec["vendor_code"], rec["invoice_number"])
        if key in seen:
            prev = seen[key]
            same = (prev["total"] == rec["total"]
                    and prev["line_sum"] == rec["line_sum"]
                    and prev["invoice_date"] == rec["invoice_date"]
                    and len(prev["lines"]) == len(rec["lines"]))
            if same:
                log.add(rec["source_file"], "file", rec["invoice_number"],
                        "duplicate_invoice", field="invoice_number",
                        original=rec["invoice_number"], action="dropped",
                        severity=NEEDS_REVIEW,
                        detail=f"same invoice as {prev['source_file']}")
                continue
            log.add(rec["source_file"], "file", rec["invoice_number"],
                    "duplicate_invoice_conflict",
                    field="invoice_number", original=rec["invoice_number"],
                    action="flagged", severity=NEEDS_REVIEW,
                    detail=f"same number as {prev['source_file']} but "
                           f"different content; both kept")
        else:
            seen[key] = rec
        invoices.append(rec)

    # per-invoice validation
    ref = config.ref_date()
    for rec in invoices:
        tax = rec.get("tax_amount") or Decimal("0")
        if rec["total"] is not None and rec["line_sum"] is not None:
            diff = cents(rec["total"] - (rec["line_sum"] + tax))
            rec["total_diff"] = diff
            if abs(diff) > Decimal("0.01"):
                log.add(rec["source_file"], "file", rec["invoice_number"],
                        "invoice_total_mismatch", field="total",
                        original=rec["total"],
                        new=cents(rec["line_sum"] + tax),
                        action="flagged", severity=NEEDS_REVIEW,
                        detail="lines do not add up to the printed total")
        master_tax = vendors.get(rec["vendor_code"], {}).get("tax_id")
        if master_tax and rec["tax_id"] and rec["tax_id"] != master_tax:
            log.add(rec["source_file"], "file", rec["invoice_number"],
                    "vendor_tax_id_mismatch", field="tax_id",
                    original=rec["tax_id"], new=master_tax, action="flagged",
                    severity=NEEDS_REVIEW,
                    detail="tax ID on invoice differs from vendor master")
        if rec["invoice_date"] and dt.date.fromisoformat(rec["invoice_date"]) > ref:
            log.add(rec["source_file"], "file", rec["invoice_number"],
                    "future_invoice_date", field="invoice_date",
                    original=str(rec["invoice_date"]), action="flagged",
                    severity=NEEDS_REVIEW,
                    detail=f"dated after reference date {ref.isoformat()}")
        for err in rec["errors"]:
            if rec["status"] == "needs_review":
                continue  # already logged as missing_invoice_field
            log.add(rec["source_file"], "file", rec["invoice_number"],
                    "missing_field", field=err.split()[0], action="flagged",
                    severity=NEEDS_REVIEW, detail=err)
        for warn in rec.get("warnings", []):
            log.add(rec["source_file"], "file", rec["invoice_number"],
                    "unparseable_invoice_line", field="lines",
                    action="flagged", severity=NEEDS_REVIEW, detail=warn)

    # ------------------------------------------------ sales files
    pos_path = _find(input_dir, config.inputs["pos_export"], ".xlsx")
    online_path = _find(input_dir, config.inputs["online_export"], ".csv")
    sales_rows, rows_in, rows_dropped = [], 0, 0
    for path, parser in [(pos_path, parse_pos), (online_path, parse_online)]:
        if path is None:
            continue
        payload, _ = proc.run(
            path, lambda p=path, fn=parser: _sales_payload(fn, p, config))
        sales_rows += payload["rows"]
        rows_in += payload["rows_in"]
        rows_dropped += payload["rows_dropped"]

    # ------------------------------------------------ customer matching
    classified: dict[str, dict] = {}
    for raw in sorted({r["customer_raw"] for r in sales_rows}):
        classified[raw] = matcher.classify(raw)
    for raw, m in classified.items():
        loc = "customer column"
        if m["method"] == "exact":
            if raw != m["customer_name"]:
                log.add(_who_has(raw, sales_rows), loc, raw,
                        "name_normalized", field="customer", original=raw,
                        new=m["customer_name"], action="changed",
                        severity=FIXED)
        elif m["method"] == "fuzzy_auto":
            log.add(_who_has(raw, sales_rows), loc, raw,
                    "customer_name_fuzzy", field="customer", original=raw,
                    new=m["customer_name"], action="changed", severity=FIXED,
                    detail=f"score {m['score']}")
        elif m["method"] == "fuzzy_review":
            detail = ("name normalises to several different customers - "
                      "human must pick"
                      if m.get("collision") else
                      f"suggested match {m['customer_name']} "
                      f"(score {m['score']}) - not applied")
            log.add(_who_has(raw, sales_rows), loc, raw,
                    "low_confidence_customer_match", field="customer",
                    original=raw, new=m["customer_name"], action="flagged",
                    severity=NEEDS_REVIEW, detail=detail)
        elif m["method"] == "unknown":
            log.add(_who_has(raw, sales_rows), loc, raw,
                    "unknown_customer", field="customer", original=raw,
                    action="flagged", severity=NEEDS_REVIEW,
                    detail="no match in customer master")
    for row in sales_rows:
        m = classified[row["customer_raw"]]
        if m["method"] in ("exact", "fuzzy_auto", "walkin"):
            row["customer_id"] = m["customer_id"]
            row["customer_name"] = m["customer_name"]
        else:
            row["customer_id"] = ""
            row["customer_name"] = ""

    # demo-only: when the input folder carries a ground-truth file for the
    # customer names, score the matcher against it
    match_eval = None
    truth_stem = config.inputs.get("matching_truth")
    truth_path = input_dir / f"{truth_stem}.csv" if truth_stem else None
    if truth_path and truth_path.exists():
        import csv
        with truth_path.open(newline="", encoding="utf-8") as f:
            match_eval = evaluate(list(csv.DictReader(f)), classified)

    # ------------------------------------------------ row accounting
    n_invoice_files = len(pdf_files)
    n_invoices_kept = len(invoices)
    n_invoices_dropped = n_invoice_files - n_invoices_kept
    if rows_in != len(sales_rows) + rows_dropped:
        raise RuntimeError(
            f"row accounting failed: {rows_in} in != "
            f"{len(sales_rows)} out + {rows_dropped} dropped")

    summary = {
        "files_in": n_invoice_files + int(pos_path is not None)
        + int(online_path is not None),
        "invoice_pdfs": n_invoice_files,
        "invoices_parsed": n_invoices_kept,
        "invoices_dropped_or_blocked": n_invoices_dropped,
        "sales_rows_in": rows_in,
        "sales_rows_out": len(sales_rows),
        "sales_rows_dropped": rows_dropped,
        # kept rows flagged for review still count in the totals; these
        # two views let a bookkeeper see the effect before deciding
        "sales_amount_all": f"{cents(sum(r['amount'] for r in sales_rows))}",
        "sales_rows_held": sum(1 for r in sales_rows if r["review_flag"]),
        "sales_amount_held": f"{cents(sum(r['amount'] for r in sales_rows if r['review_flag']))}",
        "sales_amount_excl_held": f"{cents(sum(r['amount'] for r in sales_rows if not r['review_flag']))}",
        "issues_logged": log.count(),
        "issues_needing_review": sum(1 for r in log.rows
                                     if r["severity"] != FIXED),
        # problems = things needing judgement; routine_fixes = pure
        # reformatting (dates, units, region spellings, name casing)
        "problems_found": sum(1 for r in log.rows if is_problem(r["rule"])),
        "routine_fixes": sum(1 for r in log.rows if not is_problem(r["rule"])),
    }
    if match_eval is not None:
        m = match_eval["metrics"]
        summary.update({
            "match_names_evaluated": m["names_evaluated"],
            "match_auto_correct": m["auto_correct"],
            "match_auto_wrong": m["auto_wrong"],
            "match_review_sent": m["review_sent"],
            "match_precision": m["precision"],
            "match_recall": m["recall"],
        })
    file_counts = {
        "files_skipped": sum(1 for f in proc.file_status
                             if f["status"] == "skipped"),
        "files_processed": sum(1 for f in proc.file_status
                               if f["status"] == "processed"),
    }
    manifest.save()

    runtime = round(time.time() - t0, 2)
    outputs.write_all(out_dir, results_dir, summary, invoices, sales_rows,
                      classified, log, proc.file_status, runtime,
                      file_counts, match_eval)
    return {**summary, **file_counts}


def _who_has(raw: str, sales_rows: list[dict]) -> str:
    for r in sales_rows:
        if r["customer_raw"] == raw:
            return f"{r['source_system']} export"
    return "sales"


def _sales_payload(func, path: Path, config: Config) -> dict:
    """Run a sales parser with its own log so issues can be cached and
    replayed on incremental runs."""
    inner = IssueLog()
    res = func(path, config, inner)
    return {"rows": res["rows"], "rows_in": res["rows_in"],
            "rows_dropped": res["rows_dropped"], "issues": inner.rows}
