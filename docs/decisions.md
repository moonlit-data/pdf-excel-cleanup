# Decisions log

Choices made where the task left room. Each entry: the choice and why.

- **One PDF per invoice file.** The client story says invoices arrive as
  files; 125 files in `input/invoices/` reads more realistically than one
  merged document, and it exercises per-file hashing for incremental runs.
- **PDFs drawn with reportlab, parsed with pdfplumber.** Both are pure
  Python, deterministic, and need no browser or external tools, unlike the
  prototype's Edge print-to-PDF step. The demo invoices look like real
  invoices — bordered line-item tables, header boxes, drawn logos, a
  different font/layout per vendor — because a pipeline demo that parses
  monospace text proves less than one that handles table geometry.
  Scanned PDFs are image-only pages (PIL render placed as a picture), so
  they have no text layer — exactly what real scans look like.
- **Line items are read by ruled-table extraction.** All five layouts draw
  a fully bordered grid, so `strategy: table` in `config.yml` maps
  pdfplumber's table rows onto configured columns ([desc, qty, unit,
  amount]). A per-vendor text-regex path (`strategy: text`) stays as the
  fallback and as the way non-ruled layouts would be configured.
- **Problems vs routine standardisation are counted separately.**
  `ROUTINE_RULES` in issues.py (date/number/region/name formatting) is
  everything a bookkeeper never needs to review; `problems_by_type.csv`
  and the README table cover the rest. The Summary sheet shows both.
- **Matching ground truth is generated, not predicted.**
  `make_demo_data.py` writes `input/matching_truth.csv` — for each messy
  name, the real customer_id it came from, UNKNOWN, WALKIN, or
  AMBIGUOUS:<id> for truncations that sit between two masters
  (Riverbend Books/Boutique, Kestrel Bakery/Butchery, Timberline
  Dental/Denim Co, Meadow Pharmacy/Photo Studio). The pipeline scores
  itself in `results/matching_eval.csv|json` (per-name verdicts,
  precision, recall) — the matcher never labels its own test.
- **Fuzzy auto-apply needs a score AND a clear winner.** Levenshtein
  `fuzz.ratio` alone cannot tell a typo from a genuinely ambiguous
  truncation (both score ~70–84), so a match is auto-applied only when
  score >= 66 and the margin over the runner-up is >= 12. On the demo
  truth: precision 1.0, recall 1.0, 12 names to Review (the 8 planted
  ambiguous ones plus a few borderline typos).
- **14 planted problem types** in `demo/expected.json` (IDs asserted in
  tests): duplicate_invoice, invoice_total_mismatch,
  vendor_tax_id_mismatch, future_invoice_date, scanned_pdf,
  duplicate_sales_row, refund_without_sale, refund_sign_invalid,
  refund_line_mismatch, negative_quantity, unparseable_amount,
  subtotal_row_dropped, blank_row, metadata_row_dropped — plus a
  `matching` section (typo/ambiguous/unknown name sets) asserted through
  the eval verdicts.
- **`reference_date` in config.yml (2026-04-30)** drives the future-date
  check so the demo is reproducible; real deployments delete the line and
  get today's date.
- **Row accounting counts every non-header sheet row**, including blank,
  subtotal and footer rows — "nothing disappears silently" is taken
  literally, and the check `rows_in == rows_out + rows_dropped` is
  enforced in code and tests.
- **Unparseable rows are dropped, not kept.** A sale with revenue `N/A`
  cannot go into a workbook a bookkeeper would trust; it is dropped and
  logged `needs review`. Negative quantities are kept but flagged,
  because a POS return is plausibly real.
- **Walk-in customers map to a pseudo-ID `WALKIN`**, not to "unknown" —
  they are cash sales, not master-file gaps.
- **Incremental cache stores records + issue rows per input SHA-256,**
  replayed on re-runs. Byte-identical duplicate files share a cache entry,
  so `source_file` is restored on replay (a real bug found by the
  determinism test).
- **`files_processed`/`files_skipped` live only in `run_log.json`,** not
  in `summary`/results files, because they are run metadata and would
  break byte-identical outputs between fresh and incremental runs.
- **`results/run_meta.json` holds run time and timestamp.** It is the one
  `results/` file excluded from the byte-identical test — every other
  results file must be reproducible byte-for-byte.
- **Report PDF via reportlab** (pure Python, already a dependency for
  generating the demo invoices) rather than an HTML-to-PDF tool that needs
  a browser.
- **Excel formulas where a bookkeeper expects them**: `SUM` totals under
  the Invoices and Sales tables, `COUNTA` cross-checks on the Summary
  sheet; the Summary numbers themselves are values so they read correctly
  even in a CSV viewer.
- **Issues are logged per unique raw customer name** for matching
  decisions (one entry per distinct spelling, not per sales row), while
  row-level problems (duplicates, unparseable amounts) log per row.
  Per-row logs for every reformatting change are kept because the audit
  trail is the deliverable.
- **MIT licence, author `moonlit-data`**, no real names anywhere. The
  distributed zip carries no `.git` so the recipient authors the history.
- **Money is Decimal end to end.** Amounts are parsed from the original
  text into `Decimal` by `cleanup/money.py` — strict per-convention
  regexes, so EU `12,50` under a declared `us` format (or vice versa) is
  rejected and flagged, never silently read as 1250. There is exactly one
  rounding stage: `cents()` quantizes ROUND_HALF_UP to 2 decimals when a
  value is stored or compared. `unit_price` keeps source precision (CSV
  `str(Decimal)`, Excel `0.00###`); every other money column shows 2
  decimals in CSVs, the Excel workbook (`#,##0.00`) and the README
  images. The generator does its arithmetic in Decimal too, so planted
  totals are a reliable baseline. All CSV/JSON outputs (and the
  generator's own input files) are written with LF line endings (`csv`
  needs `lineterminator="\n"`, text writes `newline="\n"`) so they are
  identical on Windows and Linux.
- **The incremental cache is keyed to a fingerprint** of config.yml's
  SHA-256 plus a parser-version and cache-schema constant, so a config
  change or a parser fix can't replay stale parse results. A corrupt
  manifest file never blocks a run — it simply matches nothing — and
  `--rebuild` clears it before loading.
- **Dedupe needs a full key and a reason.** An invoice whose number
  can't be read becomes `needs_review`, never joins dedupe on an empty
  key; two records sharing (vendor, number) are dropped only when the
  content is identical, otherwise both are kept and flagged
  `duplicate_invoice_conflict`. In the POS export an identical row is
  kept and flagged `duplicate_sales_row` — with no row-level id, an
  identical row can be two real sales — while the online export has a
  full-row identity, so an exact re-exported row is dropped and counted.
- **Matching: both sides are normalised before scoring**, the runner-up
  score is always reported (no cutoff hiding it), and two master
  customers that normalise to the same string go to Review instead of a
  silent pick unless the raw spelling is an exact hit. The eval counts
  WALKIN names separately so they can never inflate precision/recall.
- **Suggestion rules are `action=flagged`, not `changed`.** The Issues
  log only claims a change where a value was actually rewritten
  (`changed`), dropped (`dropped`) or couldn't be read at all
  (`blocked`); everything sent to a human is `flagged`.
- **External text can't inject formulas.** CSV cells starting with
  `=`, `+`, `-`, `@`, tab or CR are prefixed with `'`, and openpyxl
  cells holding external text are forced back to string type. Empty
  sheets write literal 0 totals instead of invalid `SUM`/`COUNTA`
  ranges, and a run without a matching-truth file deletes stale
  `matching_eval.*` outputs.
- **Date order is declared per source**, not guessed: `sales.pos
  .date_order: mdy` means `01/02/26` is Jan 2 2026 in the US POS export;
  the online export's ISO timestamps are unambiguous.

- **Review-2026-11-20 follow-ups:** all relative paths stored or compared
  in the generator use `PurePath.as_posix()`, so the ownership check is
  separator-independent (Windows `\\` vs `/`). Kept-but-flagged sales rows
  carry a `review_flag` column, and the Summary sheet reports sales totals
  both for all rows and excluding rows held for review, so a bookkeeper
  sees the effect of the 34 held rows before deciding on them.
