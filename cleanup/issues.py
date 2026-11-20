"""The Issues ledger: every change, drop or flag is one row."""
from __future__ import annotations

FIXED = "fixed"
NEEDS_REVIEW = "needs review"
BLOCKED = "blocked"

COLUMNS = ["issue_id", "source_file", "location", "record_id", "rule",
           "field", "original_value", "new_value", "action", "severity",
           "detail"]

# Rules that are routine standardisation (reformatting a value a human
# would have fixed by reflex) as opposed to real problems (duplicates,
# mismatches, things that need judgement). Summaries and the README count
# the two groups separately.
ROUTINE_RULES = {
    "date_format_fixed",      # 01/05/26 -> 2026-01-05 and friends
    "number_format_fixed",    # "$1,234.56" -> 1234.56, "12 pcs" -> 12
    "region_normalized",      # "N." -> North
    "name_case_fixed",        # "AVERY QUINN" -> Avery Quinn
    "product_name_fixed",     # "widget a" -> Widget A
    "name_normalized",        # customer name differing only by case/spacing
}                             #  or punctuation, matched exactly afterwards


def is_problem(rule: str) -> bool:
    return rule not in ROUTINE_RULES


class IssueLog:
    def __init__(self) -> None:
        self.rows: list[dict] = []

    def add(self, source_file: str, location: str, record_id: str, rule: str,
            field: str = "", original: object = "", new: object = "",
            action: str = "changed", severity: str = FIXED,
            detail: str = "") -> dict:
        row = {
            "issue_id": f"I{len(self.rows) + 1:05d}",
            "source_file": str(source_file),
            "location": str(location),
            "record_id": str(record_id),
            "rule": rule,
            "field": str(field),
            "original_value": "" if original is None else str(original),
            "new_value": "" if new is None else str(new),
            "action": action,
            "severity": severity,
            "detail": detail,
        }
        self.rows.append(row)
        return row

    def count(self, rule: str | None = None) -> int:
        if rule is None:
            return len(self.rows)
        return sum(1 for r in self.rows if r["rule"] == rule)

    def dropped(self) -> list[dict]:
        return [r for r in self.rows if r["action"] == "dropped"]
