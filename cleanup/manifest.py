"""Incremental runs: a manifest of input SHA-256 hashes plus a per-file
cache of the records extracted last time. Unchanged files are skipped on
re-run; --rebuild ignores the manifest entirely.

Cache validity is keyed on three things, not just file bytes:
- the SHA-256 of the config file that produced the cached records,
- PARSER_VERSION (bumped when the parsing code's behaviour changes),
- CACHE_SCHEMA (bumped when the payload layout changes).
Any of them changing makes every cached entry miss, so a config fix can
never be silently masked by stale cache.
"""
from __future__ import annotations

import hashlib
import json
from pathlib import Path

PARSER_VERSION = 2
CACHE_SCHEMA = 1


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def cache_fingerprint(config_path: Path) -> str:
    """Fingerprint of everything that affects how files are parsed."""
    h = hashlib.sha256()
    h.update(sha256_file(config_path).encode())
    h.update(f"|parser:{PARSER_VERSION}|schema:{CACHE_SCHEMA}".encode())
    return h.hexdigest()


class Manifest:
    def __init__(self, out_dir: Path, fingerprint: str = "") -> None:
        self.out_dir = out_dir
        self.path = out_dir / ".manifest.json"
        self.cache_dir = out_dir / ".cache"
        self.fingerprint = fingerprint
        self.files: dict[str, str] = {}

    def load(self) -> None:
        """Load the manifest. A corrupt or stale manifest does not crash
        the run — it simply matches nothing and is rebuilt."""
        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError, AttributeError):
            data = {}
        if data.get("fingerprint") == self.fingerprint:
            self.files = data.get("files", {})
        else:
            # different config/parser/schema: every cached record is stale
            self.files = {}
            if self.cache_dir.exists():
                import shutil
                shutil.rmtree(self.cache_dir)

    def clear(self) -> None:
        self.files = {}
        if self.path.exists():
            self.path.unlink()
        if self.cache_dir.exists():
            import shutil
            shutil.rmtree(self.cache_dir)

    def save(self) -> None:
        self.path.write_text(
            json.dumps({"fingerprint": self.fingerprint,
                        "parser_version": PARSER_VERSION,
                        "cache_schema": CACHE_SCHEMA,
                        "files": self.files}, indent=1),
            encoding="utf-8")

    def lookup(self, path: Path, sha: str) -> dict | None:
        """Cached parse result for an unchanged file, or None."""
        if self.files.get(str(path.resolve())) != sha:
            return None
        cache_file = self.cache_dir / f"{sha}.json"
        try:
            return json.loads(cache_file.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return None   # corrupt entry: reprocess the file

    def store(self, path: Path, sha: str, payload: dict) -> None:
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        (self.cache_dir / f"{sha}.json").write_text(
            json.dumps(payload, default=str), encoding="utf-8")
        self.files[str(path.resolve())] = sha
