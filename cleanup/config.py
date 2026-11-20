"""Load and validate config.yml."""
from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field
from pathlib import Path

import yaml


@dataclass
class VendorConfig:
    code: str
    name: str
    detect: str
    number_format: str            # "us" | "eu"
    document_kind: str            # "invoice" | "credit_note"
    fields: dict
    date_formats: list
    lines: dict


@dataclass
class Config:
    vendors: dict = field(default_factory=dict)
    auto_apply_threshold: float = 66.0
    review_threshold: float = 55.0
    match_margin: float = 16.0
    reference_date: dt.date | None = None
    inputs: dict = field(default_factory=dict)
    region_aliases: dict = field(default_factory=dict)
    # per-source conventions for the two sales exports, e.g.
    # {"pos": {"number_format": "us", "date_order": "mdy"},
    #  "online": {"number_format": "us"}}
    sales: dict = field(default_factory=dict)

    @staticmethod
    def load(path: str | Path) -> "Config":
        raw = yaml.safe_load(Path(path).read_text(encoding="utf-8")) or {}
        match = raw.get("customer_match", {})
        ref = raw.get("reference_date")
        vendors = {}
        for code, v in (raw.get("vendors") or {}).items():
            vendors[code] = VendorConfig(
                code=code,
                name=v["name"],
                detect=v["detect"],
                number_format=v.get("number_format", "us"),
                document_kind=v.get("document_kind", "invoice"),
                fields=v["fields"],
                date_formats=v.get("date_formats", ["%Y-%m-%d"]),
                lines=v["lines"],
            )
        # YAML turns an unquoted `reference_date: 2026-01-05` into a
        # datetime.date; an ISO string is fine too. Anything else is a
        # config error, not a crash.
        if isinstance(ref, dt.datetime):
            ref = ref.date()
        elif ref is not None and not isinstance(ref, dt.date):
            ref = str(ref).strip()
            try:
                ref = dt.date.fromisoformat(ref)
            except ValueError as exc:
                raise ValueError(
                    f"config reference_date {ref!r} is not an ISO date"
                ) from exc
        return Config(
            vendors=vendors,
            auto_apply_threshold=float(match.get("auto_apply_threshold", 66)),
            review_threshold=float(match.get("review_threshold", 55)),
            match_margin=float(match.get("min_margin", 16)),
            reference_date=ref if isinstance(ref, dt.date) else None,
            inputs=raw.get("inputs") or {},
            region_aliases=raw.get("region_aliases") or {},
            sales=raw.get("sales") or {},
        )

    def sale_format(self, source: str, key: str, default):
        return (self.sales.get(source) or {}).get(key, default)

    def ref_date(self) -> dt.date:
        return self.reference_date or dt.date.today()
