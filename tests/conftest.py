"""Session fixture: build the demo inputs in a temp dir and run the
pipeline three times — fresh, incremental, and --rebuild — so tests can
check both correctness and determinism. Runs offline."""
from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from cleanup.pipeline import run  # noqa: E402
from demo.make_demo_data import main as make_demo  # noqa: E402

CONFIG = ROOT / "config.yml"

OUTPUT_FILES = ["invoices.csv", "invoice_lines.csv", "sales.csv",
                "customers_matched.csv", "review.csv", "issues.csv",
                "summary.csv", "report.html"]
RESULT_FILES = ["summary.json", "problems_by_type.csv", "results_table.csv",
                "routine_fixes_by_type.csv", "matching_eval.csv",
                "matching_eval.json"]


def _snapshot(dir_path: Path, names: list[str]) -> dict:
    return {n: (dir_path / n).read_bytes() for n in names}


@pytest.fixture(scope="session")
def run_dir(tmp_path_factory):
    root = tmp_path_factory.mktemp("demo_run")
    input_dir = make_demo(root / "input")
    out_dir = root / "output"
    results_dir = root / "results"

    s1 = run(input_dir, out_dir, CONFIG, rebuild=True, results_dir=results_dir)
    snap1 = _snapshot(out_dir, OUTPUT_FILES)
    res1 = _snapshot(results_dir, RESULT_FILES)

    s2 = run(input_dir, out_dir, CONFIG, results_dir=results_dir)
    snap2 = _snapshot(out_dir, OUTPUT_FILES)
    res2 = _snapshot(results_dir, RESULT_FILES)

    s3 = run(input_dir, out_dir, CONFIG, rebuild=True, results_dir=results_dir)
    snap3 = _snapshot(out_dir, OUTPUT_FILES)
    res3 = _snapshot(results_dir, RESULT_FILES)

    return {"root": root, "input": input_dir, "out": out_dir,
            "results": results_dir, "summary1": s1, "summary2": s2,
            "summary3": s3, "snap": (snap1, snap2, snap3),
            "res": (res1, res2, res3)}


@pytest.fixture(scope="session")
def expected():
    return json.loads((ROOT / "demo" / "expected.json").read_text())


@pytest.fixture(scope="session")
def issues(run_dir):
    import csv
    with (run_dir["out"] / "issues.csv").open() as f:
        return list(csv.DictReader(f))


@pytest.fixture(scope="session")
def sales(run_dir):
    import csv
    with (run_dir["out"] / "sales.csv").open() as f:
        return list(csv.DictReader(f))


@pytest.fixture(scope="session")
def invoices(run_dir):
    import csv
    with (run_dir["out"] / "invoices.csv").open() as f:
        return list(csv.DictReader(f))
