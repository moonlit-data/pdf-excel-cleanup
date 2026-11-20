"""Command line interface.

    python -m cleanup run <input_folder> --out <output_folder> \\
        [--config config.yml] [--rebuild]
    python -m cleanup demo
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path


def _default_config() -> Path:
    p = Path("config.yml")
    if p.exists():
        return p
    return Path(__file__).resolve().parent.parent / "config.yml"


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="cleanup")
    sub = ap.add_subparsers(dest="cmd", required=True)

    run_p = sub.add_parser("run", help="process a folder of input files")
    run_p.add_argument("input_folder", type=Path)
    run_p.add_argument("--out", type=Path, required=True)
    run_p.add_argument("--config", type=Path, default=None)
    run_p.add_argument("--rebuild", action="store_true",
                       help="ignore the manifest and reprocess everything")
    run_p.add_argument("--results", type=Path, default=None,
                       help="also write machine-readable result numbers here")

    sub.add_parser("demo", help="generate the demo inputs, run, print summary")

    args = ap.parse_args(argv)
    if args.cmd == "demo":
        root = Path(__file__).resolve().parent.parent
        sys.path.insert(0, str(root))
        from demo.make_demo_data import main as make_demo
        input_dir = make_demo()
        out_dir = root / "output"
        config = root / "config.yml"
        results = root / "results"
    else:
        input_dir = args.input_folder
        out_dir = args.out
        config = args.config or _default_config()
        results = args.results

    from .pipeline import run
    summary = run(input_dir, out_dir, config, rebuild=getattr(args, "rebuild", False),
                  results_dir=results)
    print(f"done: {summary['files_in']} files in, "
          f"{summary['invoices_parsed']} invoices parsed, "
          f"{summary['sales_rows_out']} sales rows kept, "
          f"{summary['sales_rows_dropped']} dropped, "
          f"{summary['issues_logged']} audit entries "
          f"({summary['issues_needing_review']} need a human look)")
    print(f"outputs in {out_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
