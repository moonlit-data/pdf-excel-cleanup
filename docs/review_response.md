# Response to the 40-item review

Each item is either fixed (with a regression test that fails on the old
code and passes now) or answered with a reason. IDs follow the review's
order: H1–H8 (high), M1–M26 (medium), L1–L6 (low).

| id | severity | file | action | test |
|----|----------|------|--------|------|
| H1 | high | README.md, pipeline.py | fixed — incremental cache is keyed on a fingerprint of config.yml's SHA-256 + `PARSER_VERSION` + `CACHE_SCHEMA`; config change makes normal run == --rebuild | `test_manifest_survives_corruption_and_config_change` |
| H2 | high | manifest.py | fixed — manifest stores the fingerprint and drops the whole cache dir on mismatch (also covers config edits) | `test_manifest_survives_corruption_and_config_change` |
| H3 | high | matching.py | fixed — `by_norm` keeps every candidate; exact raw spelling wins, otherwise a normalised-name collision between different IDs is sent to Review, never silently mapped | `test_normalised_collision_sent_to_review` |
| H4 | high | outputs.py | fixed — external text forced to Excel string type; CSV cells starting `= + - @ \t \r` get a `'` prefix; numbers are never escaped | `test_excel_external_text_stays_string`, `test_csv_formula_injection_guard` |
| H5 | high | pdf_extract.py | fixed — a missing invoice number yields `status="needs_review"` and the record is kept out of deduplication entirely | `test_missing_invoice_number_is_needs_review` |
| H6 | high | pipeline.py | fixed — dedupe only on a full (vendor, number) key; identical content drops with `duplicate_invoice`, conflicting content keeps both + `duplicate_invoice_conflict` flag | `test_every_planted_problem_is_found_by_id` (planted duplicate pair) |
| H7 | high | pipeline.py | fixed — same fingerprint fix as H1/H2; `--rebuild` clears the manifest before load | `test_manifest_survives_corruption_and_config_change` |
| H8 | high | sales.py | fixed — `parse_amount` uses strict per-convention regexes; `12,50` under `us` or `1,234.56` under `eu` is rejected and flagged, never coerced | `test_parse_amount`, `test_parse_amount_rejects` |
| M1 | medium | config.py | fixed — `reference_date` accepts a YAML date object or ISO string; anything else is a config error | `test_config_accepts_yaml_date_object` |
| M2 | medium | manifest.py, pipeline.py | fixed — corrupt `.manifest.json` degrades to an empty cache instead of crashing, and `--rebuild` calls `clear()` before `load()` | `test_manifest_survives_corruption_and_config_change` |
| M3 | medium | matching.py | fixed — both sides are normalised before `fuzz.ratio` scoring and mapped back to the master row | `test_match_fuzzy_auto_applied`, `test_matching_eval_scored_against_ground_truth` |
| M4 | medium | matching.py | fixed — `best_match` uses `score_cutoff=0.0`, so the runner-up score is always real, never hidden as 0 | `test_best_match_reports_real_runner_up` |
| M5 | medium | matching.py | fixed — `auto_correct` counts only real, unambiguous customers; WALKIN verdicts are tallied separately (`walkin_correct`/`walkin_wrong`) | `test_matching_precision_and_recall` |
| M6 | medium | outputs.py | fixed — `unit_price` keeps source precision (CSV `str(Decimal)`, Excel `0.00###`); only settlement money is quantized to cents | `test_workbook_formats`, `test_csv_formula_injection_guard` |
| M7 | medium | outputs.py | fixed — zero-row sheets write a literal `0`, never `COUNTA(...!A2:A1)` or `SUM` over an inverted range | `test_empty_sheet_totals_are_literal_zero` |
| M8 | medium | outputs.py | fixed — a run without a matching-truth file deletes stale `matching_eval.csv`/`.json` | `test_stale_matching_eval_removed` |
| M9 | medium | pdf_extract.py | fixed — table extraction runs per page inside the `with` block; the text fallback runs per page and merges without dropping real lines; unparseable rows are warned and logged `unparseable_invoice_line` | `test_table_row_with_bad_cells_warns_and_skips`, `test_multipage_lines_all_captured` |
| M10 | medium | pdf_extract.py, money.py | fixed — money is parsed straight from the source string into `Decimal`; the single rounding stage is `cents()` (ROUND_HALF_UP), so `2.675 → 2.68` | `test_cents_rounds_half_up_once`, `test_parse_amount` |
| M11 | medium | pipeline.py | fixed — suggestions and review flags log `action="flagged"`; only real rewrites claim `changed` | `test_actions_are_honest` |
| M12 | medium | sales.py | fixed — identical POS rows are kept and flagged `duplicate_sales_row` (no row-level id to prove a dup); identical legit sales are never deleted | `test_every_planted_problem_is_found_by_id` |
| M13 | medium | sales.py | fixed — POS and online share `parse_amount` → `Decimal` → `cents()` at the same stage | `test_row_accounting_holds`, `test_parse_amount` |
| M14 | medium | sales.py | fixed — exact full-row online duplicates are detected, dropped and counted in Issues; 3 planted in the demo | `test_every_planted_problem_is_found_by_id` |
| M15 | medium | sales.py | fixed — header is validated (`unparseable_sheet`/blocked); per-row bad qty/amount is dropped+logged; NaN/Inf/bool rejected by `parse_amount` | `test_parse_amount_rejects`, planted `unparseable_amount` ids |
| M16 | medium | sales.py | fixed — an unparseable online timestamp drops the row with `unparseable_date` logged, never a silent blank date | planted `unparseable_date` ids in `test_every_planted_problem_is_found_by_id` |
| M17 | medium | sales.py | fixed — refunds must be negative AND match a sold line (order, sku, qty, amount); `refund_sign_invalid` and `refund_line_mismatch` are planted and asserted | `test_every_planted_problem_is_found_by_id` |
| M18 | medium | demo/make_demo_data.py | fixed — the generator does all money maths in `Decimal` with ROUND_HALF_UP to cents (5 × 19.99 × 0.9 = 89.955 → 89.96) | `test_invoices_match_generator_truth`, `test_cents_rounds_half_up_once` |
| M19 | medium | demo/make_demo_data.py | fixed — the scanned-PDF image is drawn via `ImageReader` in memory; no temp `.png` is ever created or deleted | `test_scanned_generator_leaves_no_pngs` |
| M20 | medium | demo/make_demo_data.py | fixed — the generator keeps a `.generated_files.json` manifest, removes its own stale files, and refuses foreign files in the input folder | `test_generator_rejects_foreign_input_files` |
| M21 | medium | tests/test_pipeline.py | fixed — `sales_rows_in` is verified by independently recounting POS sheet rows below the header + online CSV rows; `invoice_pdfs` is checked against generator counts | `test_row_counts_match_generator_truth` |
| M22 | medium | tests/test_pipeline.py | fixed — parsed invoices are compared field-by-field to generator truth; matching verdicts compare `customer_id`/`method` per name and the metrics are asserted from the eval artifact | `test_invoices_match_generator_truth`, `test_matching_eval_scored_against_ground_truth` |
| M23 | medium | tests/test_pipeline.py | fixed — the workbook test now asserts the exact `=SUM(F2:F…)` range and every `=COUNTA(...!A2:A…)` range against the real data rows | `test_workbook_sheets_and_formulas` |
| M24 | medium | tests/test_readme.py | fixed — every README summary-table cell is bound to its metric by label (no free substring match), the defect table is checked against expected.json, and run time is checked against `run_meta.json` | `test_readme_results_table_matches_results_files`, `test_readme_planted_defect_table_matches_expected` |
| M25 | medium | tests/test_units.py | fixed — `expected.json` now carries per-invoice truth (date, n_lines, line_sum, tax, total); all 120 invoices are compared as exact `Decimal` values | `test_invoices_match_generator_truth` |
| M26 | medium | tests/test_units.py | fixed — date tests assert explicit values (`02/03/26 → 2026-02-03` under mdy; `14.03.2026 → 2026-03-14` under dmy); per-invoice truth asserts real dates | `test_parse_date_explicit_values`, `test_invoices_match_generator_truth` |
| L1 | low | README.md, conftest.py | fixed — the claim is narrowed: the byte-identical comparison covers the CSV/JSON deliverables and results files; `clean_data.xlsx` and `run_meta.json` carry timestamps and are compared by content | `test_results_files_byte_identical`, README wording |
| L2 | low | pdf_extract.py | fixed — `extract_tables` runs inside the `pdfplumber.open` `with` block, inside the per-file try/except | `test_invoices_match_generator_truth` (120 PDFs parse clean) |
| L3 | low | sales.py | fixed — the POS workbook is opened twice (`data_only` False+True); formula cells are replaced by their cached computed values, formula-less results flag for review | `test_row_counts_match_generator_truth` |
| L4 | low | sales.py, config.yml | fixed — `sales.pos.date_order: mdy` is declared in config.yml and used to parse POS dates; documented in decisions.md | `test_parse_date_explicit_values` |
| L5 | low | demo/make_demo_data.py | fixed — reportlab runs with `invariant=1`, so regenerated invoice PDFs are byte-identical; the reproducibility note is narrowed to that | `test_regenerated_invoice_pdfs_are_byte_identical` |
| L6 | low | docs/decisions.md | fixed — the count now says 14 problem rules plus a `matching` section, and `expected.json` carries `problem_rule_count` so the number is checkable | `test_readme_planted_defect_table_matches_expected` |
