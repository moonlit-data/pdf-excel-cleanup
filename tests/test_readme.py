"""The README's numbers must come from files the pipeline writes."""
from __future__ import annotations

import csv
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))


def test_readme_results_table_matches_results_files():
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    with (ROOT / "results" / "results_table.csv").open() as f:
        rows = {r["metric"]: r["value"] for r in csv.DictReader(f)}
    # each README summary-table row is bound to a specific metric by its
    # label — a stray substring elsewhere in the file does not count
    table = dict(re.findall(r"^\| ([^|]+) \| ([^|]+) \|$", readme,
                            flags=re.M))
    label_for = {
        "files_in": "Files in",
        "invoice_pdfs": "Invoice PDFs",
        "invoices_parsed": "Invoices parsed",
        "invoices_dropped_or_blocked": "Invoices dropped or blocked",
        "sales_rows_in": "Sales rows in",
        "sales_rows_out": "Sales rows out",
        "sales_rows_dropped": "Sales rows dropped (all logged)",
            "sales_rows_held": "Sales rows held for review",
            "sales_amount_all": "Sales total (all rows)",
            "sales_amount_held": "Sales held for review",
            "sales_amount_excl_held": "Sales total (excluding held)",
        "issues_logged": "Audit log entries",
        "problems_found": "– problems needing judgement",
        "routine_fixes": "– routine standardisations",
        "issues_needing_review": "Items needing a human look",
        "match_names_evaluated": "Names evaluated",
        "match_auto_correct": "Auto-applied, correct",
        "match_auto_wrong": "Auto-applied, wrong",
        "match_review_sent": "Sent to Review",
        "match_precision": "Precision",
        "match_recall": "Recall",
    }
    for metric, label in label_for.items():
        cell = table.get(label, "").strip().replace(",", "")
        assert cell == rows[metric], f"{label}: README {cell} != {rows[metric]}"
    # run_meta is regenerated each run; the README cell must match it
    meta = json.loads((ROOT / "results" / "run_meta.json").read_text())
    m = re.search(r"Demo run time\s*\|\s*([\d.]+)\s*s", readme)
    assert m and abs(float(m.group(1))
                     - meta["demo_runtime_seconds"]) < 0.5


def test_readme_planted_defect_table_matches_expected():
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    import json as _json
    problems = _json.loads(
        (ROOT / "demo" / "expected.json").read_text())["problems"]
    for rule, spec in problems.items():
        if spec["count"] == 0:
            continue
        m = re.search(rf"^\| {re.escape(rule)} \| (\d+)/(\d+) \|$",
                      readme, flags=re.M)
        assert m, f"{rule} missing from README defect table"
        assert int(m.group(1)) == int(m.group(2)) == spec["count"], rule


def test_readme_links_and_images_exist():
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    refs = re.findall(r'(?:src|srcset|href|\])\(([^)"\s]+\.(?:png|md|csv|json|yml))["\s)]',
                      readme)
    refs += re.findall(r'srcset="([^"]+)"', readme)
    refs += re.findall(r'\((docs/[^)]+)\)', readme)
    refs += re.findall(r'`(results/[^`]+|docs/[^`]+|demo/[^`]+|config\.yml)`', readme)
    missing = [r for r in set(refs) if not (ROOT / r).exists()]
    assert not missing, f"README references missing files: {missing}"


def test_no_marketing_words_or_badges():
    readme = (ROOT / "README.md").read_text(encoding="utf-8").lower()
    for w in ["leverage", "robust", "seamless", "cutting-edge", "unlock",
              "delve", "comprehensive"]:
        assert w not in readme, w
