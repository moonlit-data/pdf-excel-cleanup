"""Money handling: Decimal everywhere, one rounding rule, one stage.

Rules (documented in docs/decisions.md):
- Amounts are parsed from their original decimal text into Decimal —
  never through float.
- `number_format` declares per-source conventions: "us" = '.' decimal /
  ',' thousands; "eu" = ',' decimal / '.' thousands. A string that does
  not strictly match the declared format is rejected (returns None) so
  the caller can send it to review instead of guessing.
- Line amounts are summed unquantized; only when a monetary result is
  stored/compared is it quantized to cents with ROUND_HALF_UP (see
  `cents`). Unit prices keep the precision they were parsed at.
"""
from __future__ import annotations

import math
import re
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

CENTS = Decimal("0.01")

_US_RE = re.compile(r"^\$?-?\d{1,3}(,\d{3})*(\.\d+)?$|^\$?-?\d+(\.\d+)?$")
_EU_RE = re.compile(r"^-?\d{1,3}(\.\d{3})*(,\d+)?$|^-?\d+(,\d+)?$")


def cents(v) -> Decimal:
    """Quantize a monetary amount to cents, ROUND_HALF_UP."""
    if not isinstance(v, Decimal):
        v = Decimal(str(v))
    return v.quantize(CENTS, rounding=ROUND_HALF_UP)


def parse_amount(text, number_format: str) -> Decimal | None:
    """Parse an amount string strictly per the declared format.

    Accepts int/float numerics (finite only) and strings. Returns None
    when the text does not match the declared convention, so "12,50"
    under 'us' is rejected rather than silently read as 1250.
    """
    if isinstance(text, Decimal):
        return text
    if isinstance(text, bool):
        return None
    if isinstance(text, int):
        return Decimal(text)
    if isinstance(text, float):
        return Decimal(str(text)) if math.isfinite(text) else None
    s = str(text).strip()
    if not s:
        return None
    negative = s.startswith("-") or (s.startswith("(") and s.endswith(")"))
    t = s.lstrip("-").strip("()").strip()
    t = t.replace("$", "").replace("EUR", "").strip()
    if number_format == "eu":
        if not re.fullmatch(r"\d{1,3}(\.\d{3})*(,\d+)?|\d+(,\d+)?", t):
            return None
        t = t.replace(".", "").replace(",", ".")
    else:
        if not re.fullmatch(r"\d{1,3}(,\d{3})*(\.\d+)?|\d+(\.\d+)?", t):
            return None
        t = t.replace(",", "")
    try:
        val = Decimal(t)
    except InvalidOperation:
        return None
    return -val if negative else val


def money_to_json(v):
    """Decimal -> canonical string for JSON caches/results."""
    return str(v) if isinstance(v, Decimal) else v
