"""Write every deliverable of a run: the Excel workbook, CSVs, run log,
HTML + PDF report for the business owner, and the results/ numbers."""
from __future__ import annotations

import csv
import datetime as dt
import html
import json
from collections import Counter
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .issues import FIXED, IssueLog, is_problem
from .money import cents

SALES_COLUMNS = ["source_system", "record_id", "order_date", "customer_raw",
                 "customer_id", "customer_name", "product", "quantity",
                 "unit_price", "amount", "channel", "note", "review_flag"]
INVOICE_COLUMNS = ["invoice_number", "vendor_code", "vendor_name", "kind",
                   "invoice_date", "total", "line_sum", "total_diff",
                   "num_lines", "pages", "source_file", "status"]
LINE_COLUMNS = ["invoice_number", "source_file", "page", "description",
                "kind", "quantity", "unit_price", "amount"]
MATCH_COLUMNS = ["raw_name", "normalized", "method", "score",
                 "customer_id", "customer_name", "occurrences"]
ISSUE_COLUMNS = ["issue_id", "source_file", "location", "record_id", "rule",
                 "field", "original_value", "new_value", "action",
                 "severity", "detail"]

HEADER_FILL = PatternFill("solid", fgColor="2a78d6")
HEADER_FONT = Font(bold=True, color="FFFFFF")
TOTAL_FONT = Font(bold=True)

# External text (customer names, product descriptions, notes) must never
# become a formula. In Excel we force string type; in CSV we prefix an
# apostrophe, the standard formula-injection mitigation.
_DANGEROUS_PREFIX = ("=", "+", "-", "@", "\t", "\r")


def _safe_text(v):
    """CSV cell text: escape a leading formula-like character."""
    if isinstance(v, str) and v.startswith(_DANGEROUS_PREFIX):
        return "'" + v
    return v


def _esc(v) -> str:
    return html.escape("" if v is None else str(v))


# ------------------------------------------------------------- tables
def build_tables(invoices: list[dict], sales_rows: list[dict],
                 classified: dict[str, dict], log: IssueLog) -> dict:
    flagged = {(r["source_file"], r["record_id"]) for r in log.rows
               if r["severity"] != FIXED}
    inv_rows = []
    for rec in sorted(invoices, key=lambda r: (r["vendor_code"], r["invoice_number"])):
        status = "flagged" if (rec["source_file"], rec["invoice_number"]) in flagged else "ok"
        inv_rows.append({
            "invoice_number": rec["invoice_number"] or "(missing)",
            "vendor_code": rec["vendor_code"],
            "vendor_name": rec.get("vendor_name", ""),
            "kind": rec["kind"],
            "invoice_date": rec["invoice_date"] or "",
            "total": rec["total"],
            "line_sum": rec["line_sum"],
            "total_diff": rec.get("total_diff", ""),
            "num_lines": len(rec["lines"]),
            "pages": rec["pages"],
            "source_file": rec["source_file"],
            "status": "needs_review" if rec["status"] == "needs_review"
            else status,
        })
    line_rows = []
    for rec in sorted(invoices, key=lambda r: (r["vendor_code"], r["invoice_number"])):
        for i, ln in enumerate(rec["lines"], 1):
            line_rows.append({"invoice_number": rec["invoice_number"],
                              "source_file": rec["source_file"],
                              "page": ln["page"], "description": ln["description"],
                              "kind": ln["kind"], "quantity": ln["quantity"],
                              "unit_price": ln["unit_price"], "amount": ln["amount"]})
    match_rows = []
    counts = Counter(r["customer_raw"] for r in sales_rows)
    for raw in sorted(classified):
        m = classified[raw]
        match_rows.append({"raw_name": raw, "normalized": m["normalized"],
                           "method": m["method"], "score": m["score"],
                           "customer_id": m["customer_id"],
                           "customer_name": m["customer_name"],
                           "occurrences": counts.get(raw, 0)})
    review_rows = [r for r in log.rows if r["severity"] != FIXED]
    return {"Invoices": (INVOICE_COLUMNS, inv_rows),
            "Invoice_Lines": (LINE_COLUMNS, line_rows),
            "Sales": (SALES_COLUMNS, sales_rows),
            "Customers_Matched": (MATCH_COLUMNS, match_rows),
            "Review": (ISSUE_COLUMNS, review_rows),
            "Issues": (ISSUE_COLUMNS, log.rows)}


# ------------------------------------------------------------- workbook
def _autosize(ws, widths: dict[int, float]) -> None:
    for idx, w in widths.items():
        ws.column_dimensions[get_column_letter(idx)].width = w


def _write_sheet(ws, columns: list[str], rows: list[dict],
                 money_cols: set[str] | None = None,
                 date_cols: set[str] | None = None,
                 total_col: str | None = None) -> None:
    ws.sheet_view.showGridLines = False
    money_cols = money_cols or set()
    date_cols = date_cols or set()
    for j, name in enumerate(columns, 1):
        c = ws.cell(row=1, column=j, value=name)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.alignment = Alignment(vertical="center")
    for i, r in enumerate(rows, 2):
        for j, name in enumerate(columns, 1):
            v = r.get(name, "")
            if name in date_cols and isinstance(v, str) and v:
                v = dt.date.fromisoformat(v)
            c = ws.cell(row=i, column=j, value=v)
            if isinstance(v, str) and v.startswith("="):
                # openpyxl auto-typed this string as a formula; force it
                # back to a literal string so external text can't inject
                # one
                c.data_type = "s"
            if name in money_cols:
                c.number_format = "#,##0.00"
            elif name in ("unit_price",):
                c.number_format = "0.00###"   # source precision kept
            elif name in date_cols:
                c.number_format = "yyyy-mm-dd"
            elif name in ("score",):
                c.number_format = "0.0"
    n = len(rows)
    if total_col:
        j = columns.index(total_col) + 1
        label_j = j - 1
        ws.cell(row=n + 2, column=label_j, value="TOTAL").font = TOTAL_FONT
        tc = ws.cell(row=n + 2, column=j,
                     value=(f"=SUM({get_column_letter(j)}2:"
                            f"{get_column_letter(j)}{n + 1})") if n else 0)
        tc.font = TOTAL_FONT
        tc.number_format = "#,##0.00"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{n + 1}"
    for j, name in enumerate(columns, 1):
        longest = max([len(name)] + [len(str(r.get(name, ""))) for r in rows[:200]])
        ws.column_dimensions[get_column_letter(j)].width = min(longest + 3, 42)


def write_workbook(path: Path, tables: dict, summary: dict,
                   problems: Counter) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Clean-up summary"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "Simulated demo data - Harbor & Pine Supply (fictional)"
    INT_FMT = "#,##0"
    MONEY_FMT = '"$"#,##0.00'
    PCT_FMT = "0.0%"
    # (label, value, number_format) — money stays a real numeric cell so a
    # bookkeeper can SUM it; formats only change how it displays
    rows = [
        ("Input files", summary["files_in"], INT_FMT),
        ("Invoice PDFs received", summary["invoice_pdfs"], INT_FMT),
        ("Invoices parsed", summary["invoices_parsed"], INT_FMT),
        ("Invoices dropped or blocked",
         summary["invoices_dropped_or_blocked"], INT_FMT),
        ("Sales rows read", summary["sales_rows_in"], INT_FMT),
        ("Sales rows kept", summary["sales_rows_out"], INT_FMT),
        ("Sales rows dropped", summary["sales_rows_dropped"], INT_FMT),
        ("Sales total (all rows)", float(summary["sales_amount_all"]),
         MONEY_FMT),
        ("Sales rows held for review", summary["sales_rows_held"], INT_FMT),
        ("Sales held for review", float(summary["sales_amount_held"]),
         MONEY_FMT),
        ("Sales total (excluding held)",
         float(summary["sales_amount_excl_held"]), MONEY_FMT),
        ("Audit log entries", summary["issues_logged"], INT_FMT),
        ("Problems found", summary["problems_found"], INT_FMT),
        ("Routine standardisations", summary["routine_fixes"], INT_FMT),
        ("Items needing a human look", summary["issues_needing_review"],
         INT_FMT),
    ]
    if "match_precision" in summary:
        rows += [("Match precision", float(summary["match_precision"]),
                  PCT_FMT),
                 ("Match recall", float(summary["match_recall"]), PCT_FMT)]
    for i, (k, v, fmt) in enumerate(rows, 4):
        ws.cell(row=i, column=1, value=k)
        c = ws.cell(row=i, column=2, value=v)
        c.number_format = fmt
    # live formulas: these update if the workbook is edited by hand
    n_inv = len(tables["Invoices"][1])
    n_iss = len(tables["Issues"][1])
    n_rev = len(tables["Review"][1])
    n_sales = len(tables["Sales"][1])
    r0 = 4 + len(rows) + 1
    # COUNTA over data rows only; an empty sheet reports 0, not a
    # miscounted header row
    def _counta(sheet: str, n: int):
        return f"=COUNTA({sheet}!A2:A{n + 1})" if n else 0
    ws.cell(row=r0, column=1, value="Check: invoice rows in workbook")
    ws.cell(row=r0, column=2, value=_counta("Invoices", n_inv))
    ws.cell(row=r0 + 1, column=1, value="Check: audit log rows in workbook")
    ws.cell(row=r0 + 1, column=2, value=_counta("Issues", n_iss))
    ws.cell(row=r0 + 2, column=1, value="Check: review rows in workbook")
    ws.cell(row=r0 + 2, column=2, value=_counta("Review", n_rev))
    # the workbook proves its own sales totals: J = amount, M = review_flag
    ws.cell(row=r0 + 3, column=1,
            value="Check: sales total (all rows)")
    c = ws.cell(row=r0 + 3, column=2, value="=SUM(Sales!J2:J%d)"
                % (n_sales + 1))
    c.number_format = MONEY_FMT
    ws.cell(row=r0 + 4, column=1,
            value="Check: sales total (excluding held)")
    c = ws.cell(row=r0 + 4, column=2,
                value='=SUMIFS(Sales!J2:J%d,Sales!M2:M%d,"")'
                % (n_sales + 1, n_sales + 1))
    c.number_format = MONEY_FMT
    ws.cell(row=r0 + 6, column=1, value="Problems by type").font = Font(bold=True)
    for i, (rule, cnt) in enumerate(sorted(problems.items()), r0 + 7):
        ws.cell(row=i, column=1, value=rule)
        ws.cell(row=i, column=2, value=cnt)
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 14

    for name in ("Invoices", "Invoice_Lines", "Sales", "Customers_Matched",
                 "Review", "Issues"):
        ws = wb.create_sheet(name)
        cols, rows_ = tables[name]
        _write_sheet(ws, cols, rows_,
                     money_cols={"total", "line_sum", "total_diff",
                                 "amount"},
                     date_cols={"invoice_date", "order_date"},
                     total_col={"Invoices": "total",
                                "Sales": "amount"}.get(name))
    wb.save(path)


# ------------------------------------------------------------- csv / json
MONEY_COLS = {"total", "line_sum", "total_diff", "unit_price", "amount"}


def _write_csv(path: Path, columns: list[str], rows: list[dict],
               money_cols: set[str] | None = None,
               unit_price_col: str | None = None) -> None:
    money_cols = (money_cols or set()) & set(columns)
    # lineterminator='\n' -> LF on every OS (csv's default is CRLF)
    with path.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=columns, lineterminator="\n")
        w.writeheader()
        for r in rows:
            out = {}
            for k, v in r.items():
                if k in money_cols and isinstance(
                        v, (int, float, Decimal)):
                    # formatted numbers are safe — never escape them
                    out[k] = f"{cents(Decimal(str(v)))}"
                    continue
                if k == unit_price_col and isinstance(
                        v, (int, float, Decimal)):
                    out[k] = str(Decimal(str(v)))  # keep source precision
                    continue
                out[k] = _safe_text(v)
            w.writerow(out)


def _write_json(path: Path, obj) -> None:
    # newline='\n' -> LF on every OS
    with path.open("w", encoding="utf-8", newline="\n") as f:
        f.write(json.dumps(obj, indent=2) + "\n")


# ------------------------------------------------------------- html report
def write_report_html(path: Path, summary: dict, problems: Counter,
                      review_rows: list[dict]) -> None:
    top = problems.most_common(8)
    maxc = top[0][1] if top else 1
    palette = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100",
               "#e87ba4", "#008300", "#6250d6", "#e34948"]
    bars = []
    for i, (rule, cnt) in enumerate(top):
        w = round(cnt / maxc * 100)
        bars.append(
            f'<div class="bar"><span class="lbl" title="{_esc(rule)}">'
            f'{_esc(rule)}</span><svg width="60%" height="18">'
            f'<rect x="0" y="0" width="{w}%" height="18" '
            f'fill="{palette[i % len(palette)]}"><title>{_esc(rule)}: '
            f'{cnt}</title></rect></svg><span class="val">{cnt}</span></div>')
    table_rows = "".join(
        f"<tr><td>{_esc(r['source_file'])}</td><td>{_esc(r['record_id'])}</td>"
        f"<td>{_esc(r['rule'])}</td><td>{_esc(r['severity'])}</td>"
        f"<td>{_esc(r['detail'])}</td></tr>"
        for r in review_rows[:50])
    cards = "".join(f'<div class="card"><div class="n">{v}</div>'
                    f'<div class="k">{_esc(k)}</div></div>' for k, v in [
        ("files in", summary["files_in"]),
        ("invoices parsed", summary["invoices_parsed"]),
        ("sales rows kept", summary["sales_rows_out"]),
        ("rows dropped", summary["sales_rows_dropped"]),
        ("changes logged", summary["issues_logged"]),
        ("needs a human", summary["issues_needing_review"]),
    ])
    doc = f"""<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Clean-up report</title><style>
:root{{--bg:#fcfcfb;--fg:#0b0b0b;--sec:#52514e;--card:#fff;--grid:#e1e0d9}}
@media(prefers-color-scheme:dark){{:root{{--bg:#1a1a19;--fg:#f0efec;
--sec:#c3c2b7;--card:#232322;--grid:#2c2c2a}}}}
*{{box-sizing:border-box}}body{{margin:0;font:15px/1.5 system-ui,sans-serif;
background:var(--bg);color:var(--fg)}}
main{{max-width:820px;margin:0 auto;padding:24px 16px}}
h1{{font-size:1.4rem;margin:0 0 4px}}.sub{{color:var(--sec);margin:0 0 20px}}
.cards{{display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));
gap:10px;margin-bottom:24px}}
.card{{background:var(--card);border:1px solid var(--grid);border-radius:8px;
padding:12px}}.n{{font-size:1.5rem;font-weight:700}}.k{{color:var(--sec);
font-size:.8rem}}
.bar{{display:flex;align-items:center;gap:8px;margin:4px 0}}
.lbl{{width:34%;font-size:.8rem;text-align:right;overflow:hidden;
text-overflow:ellipsis;white-space:nowrap}}
.val{{font-size:.8rem;color:var(--sec)}}
table{{border-collapse:collapse;width:100%;font-size:.85rem}}
th,td{{border:1px solid var(--grid);padding:4px 8px;text-align:left}}
th{{background:var(--card)}}
button{{padding:6px 12px;border:1px solid var(--grid);border-radius:6px;
background:var(--card);color:var(--fg);cursor:pointer;margin-bottom:8px}}
h2{{font-size:1.05rem;margin:24px 0 8px}}
.hidden{{display:none}}.foot{{color:var(--sec);font-size:.8rem;margin-top:32px}}
</style></head><body><main>
<h1>Clean-up report</h1>
<p class="sub">What came in, what came out, and what needs your attention.
Simulated demo data.</p>
<div class="cards">{cards}</div>
<h2>Problems found, by type</h2>
<button id="tgl" onclick="var c=document.getElementById('chart'),
t=document.getElementById('tbl');c.classList.toggle('hidden');
t.classList.toggle('hidden');this.textContent=
c.classList.contains('hidden')?'Show chart':'Show table view'">
Show table view</button>
<div id="chart">{''.join(bars)}</div>
<table id="tbl" class="hidden"><tr><th>rule</th><th>count</th></tr>
{''.join(f"<tr><td>{_esc(r)}</td><td>{c}</td></tr>" for r, c in top)}</table>
<h2>Items needing a human look</h2>
<table><tr><th>file</th><th>record</th><th>problem</th><th>severity</th>
<th>detail</th></tr>{table_rows}</table>
<p class="foot">Full detail is in clean_data.xlsx (Issues and Review sheets)
and issues.csv. Run metadata is in run_log.json. Everything in this report
is simulated.</p>
</main></body></html>"""
    path.write_text(doc, encoding="utf-8")


def write_report_pdf(path: Path, summary: dict, problems: Counter) -> None:
    """One-page PDF summary for the business owner (pure Python: reportlab)."""
    from reportlab.lib.pagesizes import LETTER
    from reportlab.pdfgen import canvas
    c = canvas.Canvas(str(path), pagesize=LETTER)
    w, h = LETTER
    y = h - 60
    c.setFont("Helvetica-Bold", 18)
    c.drawString(50, y, "Clean-up report")
    y -= 22
    c.setFont("Helvetica", 9)
    c.drawString(50, y, "Simulated demo data - fictional company")
    y -= 30
    c.setFont("Helvetica", 11)
    for k, v in [("Files received", summary["files_in"]),
                 ("Invoices parsed", summary["invoices_parsed"]),
                 ("Scanned PDFs (need manual entry)",
                  problems.get("scanned_pdf", 0)),
                 ("Sales rows kept", summary["sales_rows_out"]),
                 ("Sales rows dropped", summary["sales_rows_dropped"]),
                 ("Changes logged", summary["issues_logged"]),
                 ("Items needing a human", summary["issues_needing_review"])]:
        c.drawString(60, y, f"{k}: {v}")
        y -= 16
    y -= 12
    c.setFont("Helvetica-Bold", 11)
    c.drawString(50, y, "Top problems by type")
    y -= 16
    c.setFont("Helvetica", 9)
    for rule, cnt in problems.most_common(8):
        c.drawString(60, y, f"{rule}: {cnt}")
        y -= 13
    c.setFont("Helvetica", 8)
    c.drawString(50, 50, "Full detail: clean_data.xlsx (Issues and Review sheets).")
    c.save()


# ------------------------------------------------------------- entry point
def write_all(out_dir: Path, results_dir: Path | None, summary: dict,
              invoices: list[dict], sales_rows: list[dict],
              classified: dict[str, dict], log: IssueLog,
              file_status: list[dict], runtime: float,
              file_counts: dict | None = None,
              match_eval: dict | None = None) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    tables = build_tables(invoices, sales_rows, classified, log)
    problems = Counter(r["rule"] for r in log.rows if is_problem(r["rule"]))
    routine = Counter(r["rule"] for r in log.rows if not is_problem(r["rule"]))

    write_workbook(out_dir / "clean_data.xlsx", tables, summary, problems)
    for name, (cols, rows) in tables.items():
        _write_csv(out_dir / f"{name.lower()}.csv", cols, rows,
                   money_cols=MONEY_COLS, unit_price_col="unit_price")
    _write_csv(out_dir / "summary.csv", ["metric", "value"],
               [{"metric": k, "value": v} for k, v in summary.items()])

    run_log = {
        "run_at": dt.datetime.now().isoformat(timespec="seconds"),
        "runtime_seconds": runtime,
        "files": file_status,
        "summary": {**summary, **(file_counts or {})},
        "problems_by_type": dict(sorted(problems.items())),
        "routine_fixes_by_type": dict(sorted(routine.items())),
    }
    _write_json(out_dir / "run_log.json", run_log)
    review_rows = [r for r in log.rows if r["severity"] != FIXED]
    write_report_html(out_dir / "report.html", summary, problems,
                      review_rows)
    write_report_pdf(out_dir / "report.pdf", summary, problems)

    if results_dir is not None:
        results_dir.mkdir(parents=True, exist_ok=True)
        _write_csv(results_dir / "results_table.csv", ["metric", "value"],
                   [{"metric": k, "value": v} for k, v in summary.items()])
        _write_csv(results_dir / "problems_by_type.csv", ["rule", "count"],
                   [{"rule": r, "count": c} for r, c in problems.most_common()])
        _write_csv(results_dir / "routine_fixes_by_type.csv",
                   ["rule", "count"],
                   [{"rule": r, "count": c} for r, c in routine.most_common()])
        if match_eval is not None:
            _write_csv(results_dir / "matching_eval.csv",
                       ["raw_name", "planted_kind", "truth", "method",
                        "matched_id", "matched_name", "score", "verdict"],
                       match_eval["rows"])
            _write_json(results_dir / "matching_eval.json",
                        match_eval["metrics"])
        else:
            # a previous run left eval files behind; results must never
            # mix two runs
            for stale in ("matching_eval.csv", "matching_eval.json"):
                p = results_dir / stale
                if p.exists():
                    p.unlink()
        _write_json(results_dir / "summary.json",
                    {**summary,
                     "problems_by_type": dict(sorted(problems.items())),
                     "routine_fixes_by_type": dict(sorted(routine.items())),
                     "matching": match_eval["metrics"] if match_eval
                     else None})
        _write_json(results_dir / "run_meta.json",
                    {"demo_runtime_seconds": runtime,
                     "generated_at": dt.datetime.now().isoformat(
                         timespec="seconds")})
