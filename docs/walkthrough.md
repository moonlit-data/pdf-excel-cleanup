# Walkthrough — explaining this project in an interview

A plain-English tour of the code in the order it runs, the hard design
decisions, and likely questions. Every number cited is real — written by
the pipeline into `results/` and `output/` on the demo run.

## What the project is

A command-line tool that cleans a month of messy bookkeeping inputs:
~125 supplier-invoice PDFs in five different layouts, a ~3,000-row legacy
POS export, a ~2,000-row online-store export, and two master files.
It produces one clean Excel workbook plus a complete audit log — 7,071
entries on the demo — of everything it changed, dropped or flagged.
The demo data is simulated and generated with a fixed seed, so the whole
thing is reproducible offline.

## The code, in the order it runs

- `cleanup/cli.py` — `python -m cleanup run <folder> --out <dir>
  [--config config.yml] [--rebuild]` and `python -m cleanup demo`.
- `cleanup/config.py` — loads `config.yml` into `VendorConfig` records:
  detection anchor, field regexes, line-row regex, number format, date
  formats.
- `demo/make_demo_data.py` — builds all inputs with `random.seed(20260401)`,
  plants ~12 problem types, and writes `demo/expected.json` with exact IDs.
- `cleanup/pipeline.py` — the orchestrator. For each input file it asks
  the manifest whether the file's SHA-256 is unchanged; unchanged files
  replay cached records, changed ones are parsed.
- `cleanup/pdf_extract.py` — per-PDF extraction: text per page via
  pdfplumber; empty text means a scan; vendor is detected by an anchor
  string; fields come from regexes; line items come from ruled-table
  extraction (`strategy: table`) or a row-regex text fallback.
- `cleanup/sales.py` — cleans the POS sheet (header detection under title
  rows, 5 date formats, `$`/`pcs` numbers, region aliases, blank/subtotal/
  footer/duplicate rows dropped) and the online CSV (UTC timestamps,
  refund rows checked against sales).
- `cleanup/matching.py` — customer names: normalise → exact → Levenshtein
  ratio. A fuzzy match auto-applies only if score >= 66 and the margin
  over the runner-up is >= 12; close calls go to Review, below 55 is
  unknown. `evaluate()` scores it against `input/matching_truth.csv`.
- `cleanup/issues.py` — the ledger; every change/drop/flag is one row with
  source file, location, field, old → new value, rule, severity.
- `cleanup/manifest.py` — SHA-256 manifest + per-file cache under
  `output/.cache/` for incremental runs.
- `cleanup/outputs.py` — the workbook (7 sheets, frozen headers, filters,
  `SUM`/`COUNTA` formulas), CSVs, `run_log.json`, `report.html`,
  `report.pdf`, and the `results/` files the README quotes.
- `docs/make_readme_images.py` — renders the README images (invoice-crop
  and POS before/after tables, light/dark charts) from the real files.

## The 8–10 hardest decisions

1. **Config-driven vendor layouts.** All layout knowledge lives in YAML:
  regexes, anchors, number formats. Adding a vendor is a data change, not
  a code change — that's the whole maintainability story of the project.
2. **Line items via table geometry, not text guessing.** The demo
   invoices draw fully ruled tables, and pdfplumber's `lines` strategy
   extracts them cell-by-cell; the config maps column positions to
   fields. Text regexes stay as fallback for non-ruled layouts.
3. **The audit log is the deliverable, not a by-product.** Every fix —
   including the ~6,900 routine reformats — is a row with file, row, old
   value, new value. Routine standardisations are counted apart from the
   199 real problems, so the log tells a bookkeeper what actually needs
   attention.
4. **Row accounting as an invariant.** `rows_in == rows_out + rows_dropped`
  is asserted in the pipeline and in tests. Nothing disappears silently —
  even blank rows are dropped *with a log entry*.
5. **Determinism end-to-end.** Fixed seed, sorted file iteration, a fixed
  `reference_date` for the future-date check, and no timestamps in output
  tables — two runs give byte-identical CSVs and results files.
6. **Incremental runs via content hashing.** The manifest stores SHA-256 →
  cached records. A subtle bug the determinism test caught: two
  byte-identical invoice files share one cache entry, so `source_file` had
  to be restored per file on replay.
7. **Fuzzy matching is scored, not trusted.** The generator records
   which customer each messy name came from (`matching_truth.csv`), the
   pipeline writes `matching_eval.csv|json` with per-name verdicts, and
   tests assert precision 1.0 / recall >= 0.95, zero wrong auto-applies,
   all 8 planted ambiguous names in Review.
8. **Validation is separate from parsing.** Parsing returns records;
   business checks (totals reconcile, tax IDs match the master, dates
   aren't in the future) run in the pipeline. Cached parse results still
   get re-validated every run.
9. **Auto-apply requires a clear winner.** Score alone can't separate a
   typo from a name that sits between two masters — both score ~70–84 —
   so the winner must also lead the runner-up by >= 12 points. That one
   rule is what keeps the Review queue at 17 items instead of ~200.
10. **Money stays numeric.** Amounts are parsed to floats at the edge
   (both `1,234.56` and `1.234,56`), so Excel sees real numbers, real date
   cells, and real `SUM` formulas — not text that looks like numbers.

## Likely interview questions

1. **How do you parse five different invoice layouts?** — One generic
   parser; five YAML blocks. Detection is an anchor string; fields are
   regexes; the table is a row pattern between a header and an end anchor.
2. **What happens with a scanned PDF?** — pdfplumber returns no text; the
   file is marked `scanned`, logged `blocked`, and listed on the Review
   sheet — 3 in the demo. No crash, no OCR.
3. **How do you know nothing was lost?** — 5,013 rows in = 4,962 kept +
   51 dropped, asserted in code and tests; every drop has an Issues row.
4. **How does customer matching work?** — Normalise (case, punctuation
   to spaces, whitespace, legal suffixes) → exact; else Levenshtein
   ratio with a winner's margin. On the demo truth file: 481 names
   auto-applied correctly, 0 wrong, 17 to Review, 17 unknowns flagged.
5. **Why not just fuzzy-match everything?** — A wrong auto-match is worse
   than a review queue; the thresholds are in `config.yml`, tunable per
   client.
6. **Is it deterministic?** — Fixed seed, sorted iteration, fixed
   reference date. A test runs the pipeline three times and byte-compares
   all CSVs and results files.
7. **How do re-runs stay fast?** — SHA-256 manifest + a per-file record
   cache; the second demo run parses 0 files but still rebuilds identical
   outputs.
8. **What planted problems does it catch?** — All of them: duplicates
   (2 invoices, 15 sales rows), a total mismatch, a wrong tax ID, a
   future date, 8 orphan refunds, 8 negative quantities, 6 unparseable
   amounts, 16 blanks, 12 subtotals, 2 footer rows, 3 scans, plus 124
   planted customer-name problems.
9. **What would you do with more vendors?** — One YAML block each. The
   README shows the exact block format; no Python changes.
10. **What are the limits?** — No OCR; layouts must be configured once;
    fuzzy matching is a heuristic; duplicates are caught by vendor +
    number, not by content diffing.
11. **How would you handle 100x the volume?** — The PDF parse is the slow
    part (~120 files in ~2 s); it's embarrassingly parallel over files and
    could stream. The Excel write dominates at this size.
12. **How did you test it?** — 43 pytest tests, all offline: planted-
    problem IDs, row accounting, per-layout parsing of the drawn
    invoices, EU numbers, credit notes, matching-eval verdicts,
    incremental/rebuild determinism, and workbook structure.
