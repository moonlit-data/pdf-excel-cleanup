"""Customer name normalisation and fuzzy matching.

Exact match after normalising is applied silently-corrected; a fuzzy match
above `auto_apply_threshold` is applied and logged; between the review and
auto thresholds it is suggested only (Review sheet); below review it is
recorded as an unknown customer.

Both sides are normalised before scoring (a master written "ACME TOOLS"
must compare fairly against "acme toolz"). If two *different* master
customers normalise to the same key ("Acme LLC" vs "Acme Inc"), the raw
spelling decides; if it can't, the name goes to review — the collision is
never silently assigned to whichever customer loaded last.

`evaluate` scores the classifier against a ground-truth file (the demo
ships `matching_truth.csv` in the input folder) and produces the per-name
rows and aggregate metrics written to results/matching_eval.*.
"""
from __future__ import annotations

import re
import string
from collections import Counter

from rapidfuzz import fuzz, process

_PUNCT_TO_SPACE = {ord(c): " " for c in string.punctuation}
WALKIN_TOKENS = {"walkin", "walk in", "walk-in", "cash", "cash sale"}

# Legal-form suffixes a bookkeeper may add or drop freely; removing them is
# a normalisation fix, not a fuzzy match.
LEGAL_SUFFIXES = {"inc", "llc", "ltd", "co", "corp", "corporation",
                  "company", "gmbh", "lp", "llp", "plc", "sa", "bv"}


def normalize(name: str) -> str:
    s = name.strip().lower().translate(_PUNCT_TO_SPACE)
    tokens = re.sub(r"\s+", " ", s).strip().split(" ")
    while len(tokens) > 1 and tokens[-1] in LEGAL_SUFFIXES:
        tokens.pop()
    return " ".join(tokens)


def best_match(raw_name: str, canonical: list[str],
               ) -> tuple[str | None, float, float]:
    """Return (name, score, runner_up_score) of the best fuzzy match.

    Plain Levenshtein ratio: a truncated or subset name must NOT score
    high (token_set would call "Alder" a perfect match for "Alder
    Bakery"). No score cutoff: the real runner-up score is always
    reported, never hidden as 0.
    """
    hits = process.extract(raw_name, canonical, scorer=fuzz.ratio,
                           limit=2, score_cutoff=0.0)
    if not hits:
        return None, 0.0, 0.0
    second = float(hits[1][1]) if len(hits) > 1 else 0.0
    return hits[0][0], float(hits[0][1]), second


class CustomerMatcher:
    def __init__(self, master: list[dict], auto: float = 66.0,
                 review: float = 55.0, margin: float = 16.0) -> None:
        self.auto = auto
        self.review = review
        self.margin = margin
        self.by_norm: dict[str, list[dict]] = {}
        self.names: list[str] = []          # display names for the report
        self.norm_names: list[str] = []     # normalised candidates
        self._norm_to_rows: dict[str, list[dict]] = {}
        for row in master:
            n = normalize(row["customer_name"])
            self.by_norm.setdefault(n, []).append(row)
            self.names.append(row["customer_name"])
            if n not in self._norm_to_rows:
                self._norm_to_rows[n] = self.by_norm[n]
                self.norm_names.append(n)

    def classify(self, raw: str) -> dict:
        """One of: walkin | exact | fuzzy_auto | fuzzy_review | unknown."""
        n = normalize(raw)
        if n in WALKIN_TOKENS:
            return {"method": "walkin", "customer_id": "WALKIN",
                    "customer_name": "Walk-in", "score": 100.0,
                    "normalized": n, "collision": False}
        if n in self.by_norm:
            rows = self.by_norm[n]
            if len(rows) > 1:
                # normalised-name collision between different customers
                # (e.g. "Acme LLC" and "Acme Inc"): a case/whitespace-
                # exact raw spelling picks the unique winner; anything
                # else goes to a human.
                exact = [r for r in rows if r["customer_name"] == raw]
                if len(exact) != 1:
                    exact = [r for r in rows
                             if r["customer_name"].casefold()
                             == raw.strip().casefold()]
                if len(exact) == 1:
                    rows = exact
                else:
                    return {"method": "fuzzy_review", "customer_id": "",
                            "customer_name": " | ".join(
                                r["customer_name"] for r in rows),
                            "score": 100.0, "normalized": n,
                            "collision": True}
            row = rows[0]
            return {"method": "exact", "customer_id": row["customer_id"],
                    "customer_name": row["customer_name"], "score": 100.0,
                    "normalized": n, "collision": False}
        # score the normalised input against normalised candidates so
        # case/suffix noise doesn't push an obvious typo below threshold
        name, score, second = best_match(n, self.norm_names)
        if name is None:
            return {"method": "unknown", "customer_id": "",
                    "customer_name": "", "score": score, "normalized": n,
                    "collision": False}
        rows = self.by_norm[name]
        # a fuzzy match is applied only when the winner is clear: plausible
        # score AND well ahead of the runner-up. A close call — two masters
        # it could plausibly be — is exactly what a human should see.
        if score >= self.auto and score - second >= self.margin \
                and len(rows) == 1:
            method = "fuzzy_auto"
        elif score >= self.review:
            method = "fuzzy_review"
        else:
            method = "unknown"
        row = rows[0]
        if method == "unknown":
            # the best candidate is below the review floor — not a
            # plausible suggestion, so don't present it as one
            return {"method": "unknown", "customer_id": "",
                    "customer_name": "", "score": round(score, 1),
                    "normalized": n, "collision": False}
        return {"method": method, "customer_id": row["customer_id"],
                "customer_name": row["customer_name"],
                "score": round(score, 1), "normalized": n,
                "collision": len(rows) > 1}


# ------------------------------------------------------- ground-truth eval
def evaluate(truth_rows: list[dict], classified: dict[str, dict]) -> dict:
    """Score the matcher against planted ground truth.

    truth_rows: {raw_name, truth, planted_kind}; truth is a customer_id,
    "UNKNOWN", "WALKIN", or "AMBIGUOUS:<id>" (a name meant for a human).
    Returns {"rows": [...], "metrics": {...}}.

    precision = auto_correct / applied; recall = auto_correct /
    matchable, where matchable counts only real, unambiguous customers —
    walk-ins are tallied separately so a correct WALKIN can never inflate
    either side of the ratio.
    """
    rows_out = []
    counts = Counter()
    matchable = 0   # truths that are real, unambiguous customers
    for t in sorted(truth_rows, key=lambda r: r["raw_name"]):
        raw, truth, kind = t["raw_name"], t["truth"], t.get("planted_kind", "")
        m = classified.get(raw)
        if m is None:
            continue   # raw name never appeared in the sales rows
        method, mid = m["method"], m["customer_id"]
        ambiguous = truth.startswith("AMBIGUOUS:")
        real_customer = not ambiguous and truth not in ("UNKNOWN", "WALKIN")
        if real_customer:
            matchable += 1
        if method in ("exact", "fuzzy_auto", "walkin"):
            if ambiguous or truth == "UNKNOWN":
                verdict = "auto_wrong"
            elif truth == "WALKIN":
                verdict = "walkin_correct" if mid == "WALKIN" \
                    else "walkin_wrong"
            elif mid == truth:
                verdict = "auto_correct"
            else:
                verdict = "auto_wrong"
        elif method == "fuzzy_review":
            verdict = ("review_correct" if ambiguous
                       else "review_on_unknown" if truth == "UNKNOWN"
                       else "review_on_walkin" if truth == "WALKIN"
                       else "review_missed")
        else:
            verdict = ("unknown_correct" if truth == "UNKNOWN"
                       else "unknown_miss")
        counts[verdict] += 1
        rows_out.append({"raw_name": raw, "planted_kind": kind,
                         "truth": truth, "method": method,
                         "matched_id": mid,
                         "matched_name": m["customer_name"],
                         "score": m["score"], "verdict": verdict})
    applied = counts["auto_correct"] + counts["auto_wrong"]
    metrics = {
        "names_evaluated": len(rows_out),
        "applied": applied,
        "auto_correct": counts["auto_correct"],
        "auto_wrong": counts["auto_wrong"],
        "review_sent": (counts["review_correct"] + counts["review_missed"]
                        + counts["review_on_unknown"]
                        + counts["review_on_walkin"]),
        "ambiguous_sent_to_review": counts["review_correct"],
        "review_extra": (counts["review_missed"]
                         + counts["review_on_unknown"]
                         + counts["review_on_walkin"]),
        "unknown_flagged": counts["unknown_correct"]
                           + counts["review_on_unknown"],
        "unknown_misses": counts["unknown_miss"],
        "walkin_correct": counts["walkin_correct"],
        "walkin_wrong": counts["walkin_wrong"],
        "precision": round(counts["auto_correct"] / applied, 4)
                   if applied else 1.0,
        "recall": round(counts["auto_correct"] / matchable, 4)
                  if matchable else 1.0,
    }
    return {"rows": rows_out, "metrics": metrics}
