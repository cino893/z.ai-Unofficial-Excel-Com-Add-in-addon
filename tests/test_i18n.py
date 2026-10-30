#!/usr/bin/env python3
"""i18n key-set parity and version 5.0 checks."""
from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src" / "i18n"
LEGACY = ROOT / "legacy" / "i18n"
LANGS = ("en", "pl", "de", "fr", "es", "uk", "zh", "ja")


def load(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def main() -> int:
    failed = False
    src_files = {lang: SRC / f"{lang}.json" for lang in LANGS}
    missing = [p.name for p in src_files.values() if not p.exists()]
    if missing:
        print("FAIL missing src i18n files:", missing)
        return 1

    keys = {lang: set(load(p).keys()) for lang, p in src_files.items()}
    en_keys = keys["en"]
    for lang, kset in keys.items():
        if kset != en_keys:
            print(f"FAIL src/{lang}.json keys differ from en")
            print("  only en:", sorted(en_keys - kset))
            print("  only", lang + ":", sorted(kset - en_keys))
            failed = True

    for lang, p in src_files.items():
        about = load(p).get("about.text", "")
        if "5.0" not in about:
            print(f"FAIL src/{lang}.json about.text does not contain 5.0: {about[:80]!r}")
            failed = True

    if not LEGACY.exists():
        print("FAIL legacy/i18n/ is missing")
        return 1

    for lang in LANGS:
        lp = LEGACY / f"{lang}.json"
        sp = src_files[lang]
        if not lp.exists():
            print(f"FAIL missing legacy/i18n/{lang}.json")
            failed = True
            continue
        src_obj = load(sp)
        leg_obj = load(lp)
        if set(src_obj.keys()) != set(leg_obj.keys()):
            print(f"FAIL legacy/i18n/{lang}.json key set differs from src")
            failed = True
        if src_obj != leg_obj:
            print(f"FAIL legacy/i18n/{lang}.json content differs from src")
            failed = True
        if "5.0" not in str(leg_obj.get("about.text", "")):
            print(f"FAIL legacy/i18n/{lang}.json about.text missing 5.0")
            failed = True

    if failed:
        return 1
    print(f"OK: {len(LANGS)} languages, {len(en_keys)} keys, about.text contains 5.0")
    return 0


if __name__ == "__main__":
    sys.exit(main())
