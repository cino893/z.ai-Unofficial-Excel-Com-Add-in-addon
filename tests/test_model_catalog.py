#!/usr/bin/env python3
"""COM and VBA text-model catalogs must stay identical (ids, order, prices)."""
from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_MODEL = "glm-4.7-flash"
REQUIRED = ("glm-5.3", "glm-5.3-flash", "glm-5.3-flashx")
BANNED = ("glm-5-code",)
BUDGET_NEW = ("glm-5.3-flash", "glm-5.3-flashx")
STANDARD_NEW = ("glm-5.1", "glm-5.2", "glm-5.3")
VISION_BANNED = (
    "glm-4.6v",
    "glm-image",
    "cogview",
    "cogvideox",
    "glm-asr",
    "glm-5.3-flashx-legal",
)


def _read(rel: str) -> str:
    return (ROOT / rel).read_text(encoding="utf-8", errors="replace")


def parse_com_catalog(text: str) -> list[tuple[str, str, str, int]]:
    block = re.search(r"ModelCatalog\s*=\s*\[(.*?)\]\s*;", text, re.S)
    if not block:
        raise SystemExit("Could not find ModelCatalog in ZaiApiService.cs")
    rows = re.findall(
        r'new\(\s*"([^"]+)"\s*,\s*"([^"]+)"\s*,\s*"([^"]+)"\s*,\s*(\d+)\s*\)',
        block.group(1),
    )
    if not rows:
        raise SystemExit("ModelCatalog has no new(...) entries")
    return [(i, inp, out, int(sort)) for i, inp, out, sort in rows]


def parse_vba_ids(text: str) -> list[str]:
    fn = re.search(
        r"Public Function GetModelCatalogId\(.*?\nEnd Function",
        text,
        re.S,
    )
    if not fn:
        raise SystemExit("Could not find GetModelCatalogId")
    pairs = re.findall(r'Case\s+(\d+):\s*GetModelCatalogId\s*=\s*"([^"]+)"', fn.group(0))
    if not pairs:
        raise SystemExit("GetModelCatalogId has no Case N: id assignments")
    pairs.sort(key=lambda p: int(p[0]))
    expected = list(range(1, len(pairs) + 1))
    got = [int(p[0]) for p in pairs]
    if got != expected:
        raise SystemExit(f"GetModelCatalogId cases are not 1..n contiguous: {got}")
    return [p[1] for p in pairs]


def parse_vba_count(text: str) -> int:
    m = re.search(r"GetModelCatalogCount\s*=\s*(\d+)", text)
    if not m:
        raise SystemExit("Could not find GetModelCatalogCount")
    return int(m.group(1))


def parse_vba_price_tags(text: str) -> dict[int, str]:
    fn = re.search(r"Private Function PriceTag\(.*?\nEnd Function", text, re.S)
    if not fn:
        raise SystemExit("Could not find VBA PriceTag")
    tags: dict[int, str] = {}
    for cases, tag in re.findall(
        r'^\s*Case\s+([^:]+):\s*PriceTag\s*=\s*"([^"]+)"',
        fn.group(0),
        re.M,
    ):
        if cases.strip() == "Else":
            continue
        for part in cases.split(","):
            part = part.strip()
            if part.isdigit():
                tags[int(part)] = tag
    return tags


def parse_vba_display_cutoffs(text: str) -> tuple[int, int, int]:
    fn = re.search(
        r"Public Function GetModelCatalogDisplay\(.*?\nEnd Function",
        text,
        re.S,
    )
    if not fn:
        raise SystemExit("Could not find GetModelCatalogDisplay")
    body = fn.group(0)
    free = re.search(r"If index <= (\d+) Then", body)
    bounds = re.findall(r"ElseIf index < (\d+) Then", body)
    if not free or len(bounds) < 2:
        raise SystemExit("Could not parse GetModelCatalogDisplay index bands")
    return int(free.group(1)), int(bounds[0]), int(bounds[1])


def vba_band(index: int, free_le: int, budget_lt: int, standard_lt: int) -> str:
    if index <= free_le:
        return "free"
    if index < budget_lt:
        return "budget"
    if index < standard_lt:
        return "standard"
    return "premium"


def com_band(sort: int) -> str:
    if sort < 10:
        return "free"
    if sort < 30:
        return "budget"
    if sort < 50:
        return "standard"
    return "premium"


def com_price_tag(inp: str, out: str, sort: int) -> str:
    if sort < 10:
        return "FREE"
    return f"${inp}→${out}/MTok"


def parse_default_models(text: str) -> list[str]:
    m = re.search(r"DefaultModels\s*=\s*\[(.*?)\]\s*;", text, re.S)
    if not m:
        raise SystemExit("Could not find DefaultModels")
    ids = re.findall(r'"([^"]+)"', m.group(1))
    if not ids:
        raise SystemExit("DefaultModels is empty")
    return ids


def fail(msg: str) -> None:
    print("FAIL", msg)


def main() -> int:
    failed = False
    com_src = _read("src/Services/ZaiApiService.cs")
    vba_src = _read("legacy/modZaiAPI.bas")
    auth_src = _read("src/Services/AuthService.cs")
    conv_src = _read("src/Services/ConversationService.cs")
    test_api = _read("test_api.py")

    com = parse_com_catalog(com_src)
    com_ids = [row[0] for row in com]
    vba_ids = parse_vba_ids(vba_src)
    vba_count = parse_vba_count(vba_src)
    vba_tags = parse_vba_price_tags(vba_src)
    free_le, budget_lt, standard_lt = parse_vba_display_cutoffs(vba_src)

    print(f"COM catalog ({len(com_ids)}): {com_ids}")
    print(f"VBA catalog ({len(vba_ids)}): {vba_ids}")

    if com_ids != vba_ids:
        fail(f"COM and VBA catalog ids differ\n  COM {com_ids}\n  VBA {vba_ids}")
        failed = True
    if vba_count != len(vba_ids):
        fail(f"GetModelCatalogCount={vba_count} but {len(vba_ids)} GetModelCatalogId cases")
        failed = True
    if len(com_ids) != len(set(i.lower() for i in com_ids)):
        fail("COM catalog has duplicate ids")
        failed = True

    for req in REQUIRED:
        if req not in com_ids:
            fail(f"catalog missing required id {req}")
            failed = True
    for ban in BANNED:
        if ban in com_ids or ban in vba_ids:
            fail(f"catalog still lists dropped id {ban}")
            failed = True
    for vis in VISION_BANNED:
        if any(vis in i.lower() for i in com_ids):
            fail(f"catalog includes non-text model {vis}")
            failed = True

    for idx, (mid, inp, out, sort) in enumerate(com, start=1):
        expected = com_price_tag(inp, out, sort)
        got = vba_tags.get(idx)
        if got != expected:
            fail(f"PriceTag mismatch for {mid} (index {idx}): COM {expected!r} VBA {got!r}")
            failed = True
        cb = com_band(sort)
        vb = vba_band(idx, free_le, budget_lt, standard_lt)
        if cb != vb:
            fail(f"display band mismatch for {mid}: COM {cb} (sort {sort}) VBA {vb} (index {idx})")
            failed = True

    lookup = {mid: sort for mid, _, _, sort in com}
    for mid in BUDGET_NEW:
        if mid in lookup and not (10 <= lookup[mid] < 30):
            fail(f"{mid} SortOrder={lookup[mid]} is not in the budget band (10-29)")
            failed = True
    glm5_sort = lookup.get("glm-5")
    for mid in STANDARD_NEW:
        if mid not in lookup:
            fail(f"missing {mid}")
            failed = True
            continue
        if com_band(lookup[mid]) != com_band(glm5_sort or 41):
            fail(f"{mid} band {com_band(lookup[mid])} is not consistent with glm-5")
            failed = True

    defaults = parse_default_models(com_src)
    if defaults[0] != DEFAULT_MODEL:
        fail(f"DefaultModels[0] is {defaults[0]!r}, expected {DEFAULT_MODEL!r}")
        failed = True
    for req in REQUIRED:
        if req not in defaults:
            fail(f"DefaultModels missing {req}")
            failed = True

    load_model = re.search(r"public string LoadModel\(\)\s*\{(.*?)\n    \}", auth_src, re.S)
    if not load_model:
        fail("Could not find AuthService.LoadModel")
        failed = True
    elif f'?? "{DEFAULT_MODEL}"' not in load_model.group(1) or f'return "{DEFAULT_MODEL}"' not in load_model.group(1):
        fail("AuthService.LoadModel fallback is not glm-4.7-flash")
        failed = True
    if not re.search(rf'ZAI_DEFAULT_MODEL As String = "{DEFAULT_MODEL}"', vba_src):
        fail("ZAI_DEFAULT_MODEL is not glm-4.7-flash")
        failed = True
    if not re.search(rf'MODEL = os\.environ\.get\("ZAI_MODEL", "{DEFAULT_MODEL}"\)', test_api):
        fail("test_api.py MODEL default is not glm-4.7-flash")
        failed = True

    en_prompt = __import__("json").loads(_read("src/i18n/en.json")).get("model.prompt", "")
    if DEFAULT_MODEL not in en_prompt or "free" not in en_prompt.lower():
        fail("en.json model.prompt must say glm-4.7-flash is free")
        failed = True

    max_rounds = re.search(r"^MAX_ROUNDS = (\d+)", test_api, re.M)
    tool_rounds = re.search(r"private const int MaxToolRounds = (\d+)", conv_src)
    vba_rounds = re.search(r"Public Const MAX_TOOL_ROUNDS As Long = (\d+)", _read("legacy/modConversation.bas"))
    if not (max_rounds and tool_rounds and vba_rounds):
        fail("could not parse MAX_ROUNDS / MaxToolRounds / MAX_TOOL_ROUNDS")
        failed = True
    elif not (max_rounds.group(1) == tool_rounds.group(1) == vba_rounds.group(1) == "45"):
        fail(
            "tool-round mismatch: "
            f"test_api={max_rounds.group(1) if max_rounds else None} "
            f"COM={tool_rounds.group(1) if tool_rounds else None} "
            f"VBA={vba_rounds.group(1) if vba_rounds else None} (want 45)"
        )
        failed = True

    if failed:
        return 1
    print("OK: COM/VBA catalogs match; glm-5.3 line present; glm-5-code absent; default glm-4.7-flash")
    return 0


if __name__ == "__main__":
    sys.exit(main())
