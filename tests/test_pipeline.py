"""End-to-end tests on the generated demo inputs."""
from __future__ import annotations

import csv
import json
from pathlib import Path

from openpyxl import load_workbook

ISSUE_FIELDS = {"issue_id", "source_file", "location", "record_id", "rule",
                "field", "original_value", "new_value", "action",
                "severity", "detail"}


# ---------------------------------------------------------- planted problems
def test_every_planted_problem_is_found_by_id(issues, expected):
    for rule, spec in expected["problems"].items():
        got = sorted(r["record_id"] for r in issues if r["rule"] == rule)
        assert got == sorted(spec["ids"]), (
            f"{rule}: expected {spec['count']} ids, got {len(got)}; "
            f"missing {[i for i in spec['ids'] if i not in got][:5]}, "
            f"extra {[i for i in got if i not in spec['ids']][:5]}")


def test_problem_rule_counts(issues, expected):
    for rule, spec in expected["problems"].items():
        assert sum(1 for r in issues if r["rule"] == rule) == spec["count"], rule


# ---------------------------------------------------------- matching eval
def test_matching_eval_scored_against_ground_truth(run_dir, expected):
    """The matcher is measured, not rubber-stamped: every planted messy name
    has its true customer in input/matching_truth.csv, and the pipeline
    writes results/matching_eval.csv|json scored against it."""
    res = run_dir["results"]
    rows = {(r["raw_name"], r["planted_kind"]): r for r in csv.DictReader(
        open(res / "matching_eval.csv", encoding="utf-8"))}
    truth = list(csv.DictReader(
        open(run_dir["input"] / "matching_truth.csv", encoding="utf-8")))
    # every evaluated name has declared truth; a truth name that only
    # appears on dropped rows is legitimately not evaluated
    assert {r["raw_name"] for r in rows.values()} <= {
        r["raw_name"] for r in truth}
    for kind, raws in expected["matching"].items():
        for raw in raws:
            assert (raw, kind) in rows, raw


def test_no_wrong_auto_applied_match(run_dir):
    rows = list(csv.DictReader(
        open(run_dir["results"] / "matching_eval.csv", encoding="utf-8")))
    assert not any(r["verdict"] == "auto_wrong" for r in rows)


def test_ambiguous_names_sent_to_review(run_dir, expected):
    rows = {r["raw_name"]: r for r in csv.DictReader(
        open(run_dir["results"] / "matching_eval.csv", encoding="utf-8"))}
    for raw in expected["matching"]["ambiguous"]:
        assert rows[raw]["method"] == "fuzzy_review", raw


def test_unknown_customers_flagged(run_dir, expected):
    rows = {r["raw_name"]: r for r in csv.DictReader(
        open(run_dir["results"] / "matching_eval.csv", encoding="utf-8"))}
    for raw in expected["matching"]["unknown"]:
        assert rows[raw]["verdict"] in ("unknown_correct",
                                        "review_on_unknown"), raw
        assert rows[raw]["method"] in ("unknown", "fuzzy_review")


def test_matching_precision_and_recall(run_dir):
    m = json.loads((run_dir["results"] / "matching_eval.json").read_text())
    assert m["precision"] == 1.0
    assert m["recall"] >= 0.95
    # a useful review queue: dozens not hundreds
    assert m["review_sent"] <= 30


# ---------------------------------------------------------- row accounting
def test_row_accounting_holds(issues, sales, run_dir, expected):
    s = run_dir["summary1"]
    assert s["sales_rows_in"] == s["sales_rows_out"] + s["sales_rows_dropped"]
    dropped_sales = sum(
        1 for r in issues
        if r["action"] == "dropped" and r["source_file"].endswith((".xlsx", ".csv")))
    assert s["sales_rows_out"] == len(sales)
    assert s["sales_rows_dropped"] == dropped_sales
    # nothing disappears silently: every kept row is in sales.csv, every
    # dropped row has an Issues entry
    inv = run_dir["summary1"]
    blocked_or_dropped = sum(
        1 for r in issues
        if r["source_file"].endswith(".pdf")
        and r["action"] in ("dropped", "blocked"))
    assert inv["invoice_pdfs"] == inv["invoices_parsed"] + blocked_or_dropped


def test_row_counts_match_generator_truth(run_dir, expected):
    """rows_in comes from independently recounting the generated inputs
    (not from the pipeline's own summary), and kept + dropped covers the
    input multiset exactly."""
    import csv as csvmod
    from openpyxl import load_workbook as _lw
    wb = _lw(run_dir["input"] / "pos_sales_export.xlsx")
    ws = wb.active
    header_idx = next(i for i, r in enumerate(ws.iter_rows(values_only=True))
                      if r and str(r[0]).strip() == "Order Date")
    pos_rows = ws.max_row - header_idx - 1
    online_rows = sum(1 for _ in csvmod.DictReader(
        open(run_dir["input"] / "online_sales_export.csv")))
    s = run_dir["summary1"]
    assert s["sales_rows_in"] == pos_rows + online_rows
    assert s["invoice_pdfs"] == expected["counts"]["invoice_pdf_files"]


def test_invoice_totals_reconcile(invoices, expected):
    mismatch = set(expected["problems"]["invoice_total_mismatch"]["ids"])
    for row in invoices:
        if row["invoice_number"] in mismatch:
            continue
        assert abs(float(row["total_diff"])) <= 0.01, row["invoice_number"]
    # the planted mismatch is actually detected
    m = next(r for r in invoices if r["invoice_number"] in mismatch)
    assert abs(float(m["total_diff"])) > 0.01


def test_credit_notes_are_negative(invoices):
    for row in invoices:
        if row["kind"] == "credit_note":
            assert float(row["total"]) < 0
            assert float(row["line_sum"]) < 0


def test_scanned_pdfs_listed_not_parsed(issues, invoices, expected):
    scanned = expected["problems"]["scanned_pdf"]["ids"]
    got = sorted(r["record_id"] for r in issues if r["rule"] == "scanned_pdf")
    assert got == sorted(scanned)
    for r in issues:
        if r["rule"] == "scanned_pdf":
            assert r["severity"] == "blocked"
    assert not any(r["source_file"] in scanned for r in invoices)


# ---------------------------------------------------------- determinism
def test_incremental_run_skips_processed_files(run_dir, expected):
    s2 = run_dir["summary2"]
    assert s2["files_skipped"] == expected["counts"]["invoice_pdf_files"] + 2
    assert s2["files_processed"] == 0


def test_rerun_and_rebuild_give_identical_outputs(run_dir):
    snap1, snap2, snap3 = run_dir["snap"]
    for name in snap1:
        assert snap1[name] == snap2[name], f"incremental run differs: {name}"
        assert snap1[name] == snap3[name], f"--rebuild run differs: {name}"


def test_results_files_byte_identical(run_dir):
    res1, res2, res3 = run_dir["res"]
    for name in res1:
        assert res1[name] == res2[name] == res3[name], name


# ---------------------------------------------------------- workbook
def test_workbook_sheets_and_formulas(run_dir):
    wb = load_workbook(run_dir["out"] / "clean_data.xlsx")
    assert wb.sheetnames == ["Summary", "Invoices", "Invoice_Lines", "Sales",
                            "Customers_Matched", "Review", "Issues"]
    ws = wb["Invoices"]
    assert ws.freeze_panes == "A2"
    assert ws.auto_filter.ref.startswith("A1:")
    formulas = [c.value for row in ws.iter_rows() for c in row
                if isinstance(c.value, str) and c.value.startswith("=SUM(")]
    assert formulas, "expected a real SUM total on the Invoices sheet"
    # the SUM must span exactly the data rows, nothing else
    n_data = ws.max_row - 2            # header + total row
    assert formulas == [f"=SUM(F2:F{n_data + 1})"]
    ws = wb["Summary"]
    countas = [c.value for row in ws.iter_rows() for c in row
               if isinstance(c.value, str) and c.value.startswith("=COUNTA(")]
    assert countas and all(
        f"A2:A" in f for f in countas), countas
    # COUNTA ranges end exactly on the last data row of each sheet
    import re
    for sheet, f in zip(("Invoices", "Issues", "Review"), countas):
        m = re.match(rf"=COUNTA\({sheet}!A2:A(\d+)\)", f)
        n = sum(1 for r in wb[sheet].iter_rows(min_row=2)
                if r[0].value not in (None, ""))
        assert m and int(m.group(1)) == n + 1, (sheet, f)
    ws = wb["Sales"]
    assert any(isinstance(c.value, str) and c.value.startswith("=SUM(")
               for row in ws.iter_rows() for c in row)


def test_workbook_formats(run_dir):
    wb = load_workbook(run_dir["out"] / "clean_data.xlsx")
    ws = wb["Sales"]
    # unit_price keeps source precision (0.00###); amount is always 2dp
    assert ws.cell(row=2, column=9).number_format == "0.00###"
    assert ws.cell(row=2, column=10).number_format == "#,##0.00"
    ws = wb["Invoices"]
    assert ws.cell(row=2, column=5).number_format == "yyyy-mm-dd"
    assert not ws.sheet_view.showGridLines


# ---------------------------------------------------------- numbers files
def test_results_match_summary(run_dir):
    summary = json.loads((run_dir["results"] / "summary.json").read_text())
    with (run_dir["results"] / "results_table.csv").open() as f:
        table = {r["metric"]: r["value"] for r in csv.DictReader(f)}
    for k, v in table.items():
        if k.startswith("files_") or k.startswith("sales_") \
                or k.startswith("invoices_") or k.startswith("issues_"):
            assert str(summary[k]) == v


def test_report_files_exist(run_dir):
    for name in ["report.html", "report.pdf", "run_log.json",
                 "clean_data.xlsx"]:
        assert (run_dir["out"] / name).exists()
    html = (run_dir["out"] / "report.html").read_text()
    assert "prefers-color-scheme" in html
    assert "Show table view" in html


# ---------------------------------------------------------- log honesty
def test_actions_are_honest(issues):
    """action=changed means a value was actually rewritten; suggestions
    and human-review flags are 'flagged', real removals are
    'dropped'/'blocked'."""
    allowed = {"changed", "flagged", "dropped", "blocked"}
    assert {r["action"] for r in issues} <= allowed
    # rules that only advise must never claim a change was made
    suggestion_rules = {"invoice_total_mismatch", "vendor_tax_id_mismatch",
                        "future_invoice_date", "missing_invoice_field",
                        "low_confidence_customer_match", "unknown_customer",
                        "negative_quantity", "refund_without_sale",
                        "refund_sign_invalid", "refund_line_mismatch",
                        "duplicate_invoice_conflict", "duplicate_sales_row"}
    for r in issues:
        if r["rule"] in suggestion_rules:
            assert r["action"] in ("flagged", "dropped"), \
                (r["rule"], r["action"], r["record_id"])


def test_held_for_review_sales_totals(run_dir):
    """Rows kept but flagged for review are marked and their totals are
    shown separately: all = held + excluding_held."""
    import csv as csvmod
    from decimal import Decimal as D
    results = dict(csvmod.reader(open(run_dir["results"] /
                                      "results_table.csv")))
    with open(run_dir["out"] / "sales.csv") as f:
        sales = list(csvmod.DictReader(f))
    assert "review_flag" in sales[0]
    held = [r for r in sales if r["review_flag"]]
    # the 15 kept POS duplicates must be visible on the sheet, not hidden
    assert sum(1 for r in held if "duplicate_sales_row" in r["review_flag"]) \
        == 15
    assert len(held) == int(results["sales_rows_held"])
    total = sum(D(r["amount"]) for r in sales)
    held_amt = sum(D(r["amount"]) for r in held)
    assert f"{held_amt}" == results["sales_amount_held"]
    assert f"{total - held_amt}" == results["sales_amount_excl_held"]
    assert f"{total}" == results["sales_amount_all"]
    # the same three figures appear on the Excel Summary sheet
    from openpyxl import load_workbook as _lw
    wb = _lw(run_dir["out"] / "clean_data.xlsx")
    cells = {ws.cell(row=i, column=1).value: ws.cell(row=i, column=2).value
             for i in range(1, 40) if (ws := wb["Summary"]).cell(
                 row=i, column=1).value}
    assert cells["Sales total (all rows)"] == \
        float(results["sales_amount_all"])
    assert cells["Sales held for review"] == \
        float(results["sales_amount_held"])
    assert cells["Sales total (excluding held)"] == \
        float(results["sales_amount_excl_held"])


def test_summary_cells_are_numeric_with_formats(run_dir):
    """Every Summary value cell is a real number (never a string) with the
    right format, and the self-check formulas point at the right columns."""
    from openpyxl.utils import get_column_letter
    from cleanup.outputs import SALES_COLUMNS
    wb = load_workbook(run_dir["out"] / "clean_data.xlsx")
    ws = wb["Summary"]
    cells = {}
    for i in range(1, 60):
        k = ws.cell(row=i, column=1).value
        if k:
            cells[k] = ws.cell(row=i, column=2)
    money_labels = {"Sales total (all rows)", "Sales held for review",
                    "Sales total (excluding held)"}
    pct_labels = {"Match precision", "Match recall"}
    problems = {r["rule"]
                for r in csv.DictReader(
                    open(run_dir["results"] / "problems_by_type.csv"))}
    metric_labels = set(cells) - {
        k for k in cells if str(k).startswith("Check:")} - {
        "Clean-up summary",
        "Simulated demo data - Harbor & Pine Supply (fictional)",
        "Problems by type"} - problems
    for k in metric_labels:
        c = cells[k]
        assert isinstance(c.value, (int, float)) and \
            not isinstance(c.value, str) and \
            not isinstance(c.value, bool), f"{k} is {type(c.value)}"
        expected = '"$"#,##0.00' if k in money_labels else \
            "0.0%" if k in pct_labels else "#,##0"
        assert c.number_format == expected, (k, c.number_format)
    # check formulas target amount (J) and review_flag (M)
    amt = get_column_letter(SALES_COLUMNS.index("amount") + 1)
    flag = get_column_letter(SALES_COLUMNS.index("review_flag") + 1)
    n = sum(1 for _ in open(run_dir["out"] / "sales.csv")) - 1
    assert cells["Check: sales total (all rows)"].value == \
        f"=SUM(Sales!{amt}2:{amt}{n + 1})"
    assert cells["Check: sales total (excluding held)"].value == \
        f'=SUMIFS(Sales!{amt}2:{amt}{n + 1},Sales!{flag}2:{flag}{n + 1},"")'
    assert cells["Check: sales total (all rows)"].number_format == \
        '"$"#,##0.00'
