"""Unit tests: number/date parsing, matching, and one fixture invoice
per vendor layout."""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from decimal import Decimal  # noqa: E402

import pytest  # noqa: E402

from cleanup.config import Config  # noqa: E402
from cleanup.matching import CustomerMatcher, normalize  # noqa: E402
from cleanup.money import cents, parse_amount  # noqa: E402
from cleanup.pdf_extract import parse_date, parse_invoice_pdf  # noqa: E402


@pytest.fixture(scope="module")
def config():
    return Config.load(ROOT / "config.yml")


# ---------------------------------------------------------- primitives
@pytest.mark.parametrize("text,fmt,want", [
    ("1,234.56", "us", "1234.56"), ("$12.00", "us", "12.00"),
    ("-412.50", "us", "-412.50"), ("1.234,56", "eu", "1234.56"),
    ("699,35", "eu", "699.35"), ("-12,30", "eu", "-12.30"),
    ("12,50", "eu", "12.50"), ("2.675", "us", "2.675"),
    ("(19.95)", "us", "-19.95"), ("EUR 1.234,56", "eu", "1234.56"),
])
def test_parse_amount(text, fmt, want):
    got = parse_amount(text, fmt)
    assert isinstance(got, Decimal)
    assert got == Decimal(want)


@pytest.mark.parametrize("text,fmt", [
    # EU strings under a declared "us" format must NOT silently lose the
    # comma or be misread: they are rejected and flagged instead
    ("12,50", "us"), ("1.234,56", "us"), ("699,35", "us"),
    # and US strings are rejected under "eu" for the same reason
    ("1,234.56", "eu"), ("$1.50", "eu"),
    # malformed / non-numeric input is rejected, not coerced
    ("N/A", "us"), ("", "us"), (None, "us"), ("inf", "us"),
    (float("nan"), "us"), (float("inf"), "us"), (True, "us"),
    ("1.2.3", "us"), ("12,34,56", "us"), ("abc", "eu"),
])
def test_parse_amount_rejects(text, fmt):
    assert parse_amount(text, fmt) is None


def test_cents_rounds_half_up_once():
    # the one rounding stage: ROUND_HALF_UP to cents at storage time
    assert cents(Decimal("2.675")) == Decimal("2.68")
    assert cents(Decimal("89.955")) == Decimal("89.96")
    assert cents(Decimal("-2.675")) == Decimal("-2.68")
    assert cents(Decimal("100")) == Decimal("100.00")


@pytest.mark.parametrize("text,fmt", [
    ("02/03/26", "%m/%d/%y"), ("2026-03-14", "%Y-%m-%d"),
    ("March 4 2026", "%B %d %Y"), ("14.03.2026", "%d.%m.%Y"),
])
def test_parse_date(text, fmt):
    assert parse_date(text, [fmt]) is not None


def test_normalize():
    assert normalize("  Alder   Bakery ") == "alder bakery"
    assert normalize("Birchwood-Books!") == "birchwood books"
    assert normalize("Maple Grove LLC") == "maple grove"


# ---------------------------------------------------------- matching
def _master():
    return [{"customer_id": "C-0001", "customer_name": "Alder Grove Bakery",
             "city": "x", "terms": "Net 30"},
            {"customer_id": "C-0002", "customer_name": "Beacon Hill Clinic",
             "city": "y", "terms": "Net 30"}]


def test_match_exact_after_normalising():
    m = CustomerMatcher(_master())
    assert m.classify("  ALDER GROVE BAKERY ")["method"] == "exact"


def test_match_fuzzy_auto_applied():
    m = CustomerMatcher(_master())   # config.yml thresholds (66/55/12)
    r = m.classify("Alder Grove Bakry")
    assert r["method"] == "fuzzy_auto"
    assert r["customer_id"] == "C-0001"


def test_match_close_call_goes_to_review():
    m = CustomerMatcher([{"customer_id": "C-1", "customer_name": "Riverbend Books",
                          "city": "x", "terms": "Net 30"},
                         {"customer_id": "C-2", "customer_name": "Riverbend Boutique",
                          "city": "y", "terms": "Net 30"}])
    # one truncation, two plausible masters: high score but no clear winner
    assert m.classify("Riverbend Bo")["method"] == "fuzzy_review"


def test_match_review_and_unknown():
    m = CustomerMatcher(_master())
    # a name between two masters and below auto-apply score: human review
    r = m.classify("Alder Clinic")
    assert r["method"] in ("fuzzy_review", "unknown")
    r = m.classify("Quibell Works")
    assert r["method"] == "unknown"
    assert r["customer_id"] == ""


def test_walkin_is_a_known_pseudo_customer():
    m = CustomerMatcher(_master())
    assert m.classify("walk-in")["customer_id"] == "WALKIN"


# ---------------------------------------------------------- one fixture per layout
def _invoice(input_dir, number):
    return input_dir / "invoices" / f"{number}.pdf"


def test_each_vendor_layout_parses(run_dir, config):
    cases = [  # invoice_number -> checks
        ("NPP-1001", {"pages": 1}), ("NPP-1008", {"pages": 2}),
        ("BHS-2001", {}), ("AES-3101", {}), ("MPS-2201", {}),
        ("CRS-5101", {}),
    ]
    for number, extra in cases:
        rec = parse_invoice_pdf(_invoice(run_dir["input"], number), config)
        assert rec["status"] == "ok", number
        assert rec["invoice_number"] == number
        assert rec["invoice_date"], number
        assert rec["lines"], number
        for k, v in extra.items():
            assert rec[k] == v, (number, k)


def test_invoices_match_generator_truth(run_dir, config):
    """Every parsed invoice agrees with the generator's own truth on
    date, per-line count, line_sum and total (exact Decimal values)."""
    import json
    truth = json.loads((ROOT / "demo" / "expected.json").read_text())[
        "invoices"]
    checked = 0
    for number, want in truth.items():
        rec = parse_invoice_pdf(_invoice(run_dir["input"], number), config)
        assert rec["status"] == "ok", number
        assert str(rec["invoice_date"]) == want["date"], number
        assert len(rec["lines"]) == want["n_lines"], number
        # credit notes carry negated lines; compare magnitudes
        assert abs(rec["line_sum"]) == Decimal(want["line_sum"]), number
        assert rec["total"] == Decimal(want["total"]), number
        checked += 1
    assert checked > 100


def test_multipage_lines_all_captured(run_dir, config):
    rec = parse_invoice_pdf(_invoice(run_dir["input"], "NPP-1008"), config)
    assert rec["pages"] == 2
    assert any(l["page"] == 2 for l in rec["lines"])
    assert rec["line_sum"] > 0


def test_eu_number_format_invoice(run_dir, config):
    rec = parse_invoice_pdf(_invoice(run_dir["input"], "AES-3101"), config)
    # EU "1.234,56" must arrive as an exact Decimal; sums reconcile
    assert rec["total"] == rec["line_sum"]
    assert all(isinstance(l["amount"], Decimal) for l in rec["lines"])
    # AES rows carry no price/quantity columns; the table reads description
    # + amount only
    assert all(l["amount"] > 0 for l in rec["lines"])


def test_credit_note_invoice(run_dir, config):
    rec = parse_invoice_pdf(_invoice(run_dir["input"], "CRS-5101"), config)
    assert rec["kind"] == "credit_note"
    assert rec["total"] < 0
    assert all(l["amount"] < 0 for l in rec["lines"])


def test_mps_discount_and_shipping_lines(run_dir, config):
    rec = parse_invoice_pdf(_invoice(run_dir["input"], "MPS-2201"), config)
    kinds = {l["kind"] for l in rec["lines"]}
    assert {"item", "discount", "shipping"} <= kinds
    assert [l for l in rec["lines"] if l["kind"] == "discount"][0]["amount"] < 0


def test_scanned_pdf_detected(run_dir, config):
    rec = parse_invoice_pdf(run_dir["input"] / "invoices" / "scan_001.pdf",
                            config)
    assert rec["status"] == "scanned"
    assert rec["lines"] == []


# ------------------------------------------- review-driven regression tests
def test_best_match_reports_real_runner_up():
    from cleanup.matching import best_match
    name, score, second = best_match(
        "alder grove bakry",
        ["alder grove bakery", "alder grove clinic", "beacon hill"])
    assert name == "alder grove bakery"
    # the runner-up is a real score - a cutoff must never hide it as 0
    assert score > second > 0


def test_normalised_collision_sent_to_review():
    # two DIFFERENT customers whose names normalise identically
    m = CustomerMatcher([
        {"customer_id": "C-1", "customer_name": "Alder Grove Bakery",
         "city": "x", "terms": "Net 30"},
        {"customer_id": "C-2", "customer_name": "Alder Grove Bakery, Inc.",
         "city": "x", "terms": "Net 30"}])
    # exact raw spelling still wins deterministically
    assert m.classify("Alder Grove Bakery")["customer_id"] == "C-1"
    assert m.classify("Alder Grove Bakery, Inc.")["customer_id"] == "C-2"
    # a normalisation-only hit shared by both: review, not a silent pick
    r = m.classify("ALDER GROVE BAKERY INC")
    assert r["method"] == "fuzzy_review" and r["collision"]


def test_parse_date_explicit_values():
    import datetime as dt
    assert parse_date("02/03/26", ["%m/%d/%y"]) == dt.date(2026, 2, 3)
    assert parse_date("14.03.2026", ["%d.%m.%Y"]) == dt.date(2026, 3, 14)


def test_config_accepts_yaml_date_object(tmp_path):
    import datetime as dt
    p = tmp_path / "config.yml"
    p.write_text("reference_date: 2026-04-30\n")  # unquoted -> date object
    cfg = Config.load(p)
    assert cfg.reference_date == dt.date(2026, 4, 30)
    cfg = Config(reference_date=dt.date(2026, 1, 5))
    assert cfg.ref_date() == dt.date(2026, 1, 5)


def test_manifest_survives_corruption_and_config_change(tmp_path):
    from cleanup.manifest import Manifest, cache_fingerprint
    fp = cache_fingerprint(ROOT / "config.yml")
    man = Manifest(tmp_path, fp)
    man.files["a.pdf"] = "sha-x"
    man.save()
    man2 = Manifest(tmp_path, fp)
    man2.load()
    assert man2.files == {"a.pdf": "sha-x"}
    # a different config/parser/schema fingerprint clears the cache
    man3 = Manifest(tmp_path, "other-config")
    man3.load()
    assert man3.files == {}
    # a corrupt manifest file must not block a run or a --rebuild
    (tmp_path / ".manifest.json").write_text("{ not json !!")
    man4 = Manifest(tmp_path, fp)
    man4.clear()   # what --rebuild does first
    man4.load()
    assert man4.files == {}


def test_table_row_with_bad_cells_warns_and_skips():
    from cleanup.pdf_extract import _rows_from_table
    vendor = Config.load(ROOT / "config.yml").vendors["AES"]
    warnings = []
    table = [["Beschreibung", "Menge", "Einzelpreis", "Betrag"],
             ["Ordner", "10", "44,80", "448,00"],
             ["Ordner", "zehn", "44,80", "448,00"],   # unparseable qty
             ["Zwischensumme", "", "", "448,00"]]     # footer, not an item
    items = _rows_from_table(table, 1, vendor, warnings)
    assert len(items) == 1
    assert items[0]["description"] == "Ordner"
    # the bad data row was logged, not silently kept or silently dropped
    assert warnings and "unparseable" in warnings[0]


def test_csv_formula_injection_guard(tmp_path):
    from cleanup.outputs import _write_csv
    out = tmp_path / "t.csv"
    _write_csv(out, ["a", "b"], [
        {"a": "=1+cmd|'/c calc'!A0", "b": Decimal("-9.99")},
        {"a": "+SUM(x)", "b": Decimal("10.00")},
    ], money_cols={"b"})
    lines = out.read_text().splitlines()
    assert lines[1].startswith("'=")
    # real numbers are never escaped: a negative stays a number in Excel
    assert "-9.99" in lines[1] and "'-9.99" not in lines[1]
    assert lines[2].startswith("'+SUM")


def test_excel_external_text_stays_string():
    from cleanup.outputs import _write_sheet
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    _write_sheet(ws, ["name", "total"],
                 [{"name": "=HYPERLINK('x')", "total": Decimal("12.50")}],
                 money_cols={"total"}, total_col="total")
    cell = ws.cell(row=2, column=1)
    assert cell.data_type == "s"        # literal string, not a formula
    assert ws.cell(row=3, column=2).value.startswith("=SUM(")


def test_empty_sheet_totals_are_literal_zero(tmp_path):
    from cleanup.outputs import write_workbook
    empty_tables = {n: (["a"], []) for n in
                    ("Invoices", "Invoice_Lines", "Sales",
                     "Customers_Matched", "Review", "Issues")}
    # Summary sheet needs the column pairs it writes
    empty_tables["Invoices"] = (["invoice_number", "total"], [])
    empty_tables["Sales"] = (["record_id", "amount"], [])
    summary = {k: 0 for k in
               ("files_in", "invoice_pdfs", "invoices_parsed",
                "invoices_dropped_or_blocked", "sales_rows_in",
                "sales_rows_out", "sales_rows_dropped", "issues_logged", "sales_amount_all",
                    "sales_rows_held", "sales_amount_held", "sales_amount_excl_held",
                "problems_found", "routine_fixes", "issues_needing_review")}
    write_workbook(tmp_path / "w.xlsx", empty_tables, summary, {})
    from openpyxl import load_workbook
    wb = load_workbook(tmp_path / "w.xlsx")
    summary = wb["Summary"]
    checks = [c.value for row in summary.iter_rows() for c in row
              if isinstance(c.value, (str, int)) and
              (c.value == 0 or (isinstance(c.value, str)
                                and c.value.startswith("=COUNTA")))]
    assert 0 in checks            # empty sheets -> literal 0, not COUNTA(A2:A1)
    assert wb["Invoices"].cell(row=2, column=1).value == "TOTAL"
    assert wb["Invoices"].cell(row=2, column=2).value == 0  # not =SUM(B2:B1)


def test_stale_matching_eval_removed(tmp_path):
    from cleanup.issues import IssueLog
    from cleanup.outputs import write_all
    res = tmp_path / "results"
    res.mkdir()
    for name in ("matching_eval.csv", "matching_eval.json"):
        (res / name).write_text("stale")
    summary = {k: 0 for k in
               ("files_in", "invoice_pdfs", "invoices_parsed",
                "invoices_dropped_or_blocked", "sales_rows_in",
                "sales_rows_out", "sales_rows_dropped", "issues_logged", "sales_amount_all",
                    "sales_rows_held", "sales_amount_held", "sales_amount_excl_held",
                "problems_found", "routine_fixes", "issues_needing_review")}
    write_all(tmp_path / "out", res, summary, [], [], {},
              IssueLog(), [], 0.0, match_eval=None)
    assert not (res / "matching_eval.csv").exists()
    assert not (res / "matching_eval.json").exists()


def test_missing_invoice_number_is_needs_review(tmp_path, config):
    """An invoice PDF whose number cannot be read: status needs_review,
    never an 'ok' record that could dedupe on an empty key."""
    from reportlab.lib.pagesizes import LETTER
    from reportlab.pdfgen import canvas
    p = tmp_path / "AES-no-number.pdf"
    c = canvas.Canvas(str(p), pagesize=LETTER)
    c.drawString(60, 740, "Alpen East Supplies")   # vendor detect string
    c.drawString(60, 700, "Datum: 14.03.2026")
    c.drawString(60, 660, "Gesamtbetrag: 448,00 EUR")
    c.save()
    rec = parse_invoice_pdf(p, config)
    assert rec["status"] == "needs_review"
    assert rec["invoice_number"] == ""


def test_scanned_generator_leaves_no_pngs(run_dir):
    # scan simulation draws in memory; no temp image may survive in input/
    assert not list((run_dir["input"] / "invoices").glob("*.png"))


def test_generator_rejects_foreign_input_files(tmp_path):
    sys.path.insert(0, str(ROOT / "demo"))
    import make_demo_data
    make_demo_data.main(tmp_path)
    (tmp_path / "stray_invoice.pdf").write_bytes(b"%PDF-fake")
    with pytest.raises(RuntimeError, match="foreign|does not own"):
        make_demo_data.main(tmp_path)


def test_regenerated_invoice_pdfs_are_byte_identical(tmp_path):
    """`invariant` mode: regenerating the demo yields identical PDF bytes,
    so hash-based caching isn't fooled by embedded timestamps."""
    sys.path.insert(0, str(ROOT / "demo"))
    import make_demo_data
    import hashlib
    a = make_demo_data.main(tmp_path / "a")
    b = make_demo_data.main(tmp_path / "b")
    pa = sorted((a / "invoices").glob("*.pdf"))
    pb = sorted((b / "invoices").glob("*.pdf"))
    assert [p.name for p in pa] == [p.name for p in pb]
    for x, y in zip(pa, pb):
        assert hashlib.sha256(x.read_bytes()).hexdigest() \
            == hashlib.sha256(y.read_bytes()).hexdigest(), x.name


def test_owned_paths_are_separator_independent():
    """Windows: str(Path.relative_to(...)) yields 'invoices\\x.pdf' while the
    manifest stores 'invoices/x.pdf'; the comparison must use as_posix()."""
    from pathlib import PurePosixPath, PureWindowsPath
    for root, p in [(PureWindowsPath("C:/inputs"),
                     PureWindowsPath("C:/inputs/invoices/AES-3101.pdf")),
                    (PurePosixPath("/inputs"),
                     PurePosixPath("/inputs/invoices/AES-3101.pdf"))]:
        assert p.relative_to(root).as_posix() == "invoices/AES-3101.pdf"
    # guard the fix: no raw str(relative_to()) in stored/compared paths
    src = (ROOT / "demo" / "make_demo_data.py").read_text()
    assert "str(p.relative_to(" not in src
