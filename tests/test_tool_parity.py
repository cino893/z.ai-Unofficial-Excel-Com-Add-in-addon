#!/usr/bin/env python3
"""Assert COM and VBA Excel tool name sets are identical."""
from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def _read(rel: str) -> str:
    return (ROOT / rel).read_text(encoding="utf-8", errors="replace")


def com_make_tool_names() -> set[str]:
    text = _read("src/Services/ExcelSkillService.cs")
    return set(re.findall(r'MakeTool\(\s*"([a-z0-9_]+)"', text))


def com_execute_names() -> set[str]:
    text = _read("src/Services/ExcelSkillService.cs")
    m = re.search(r"toolName switch\s*\{(.*?)\n\s*_ =>", text, re.S)
    if not m:
        raise SystemExit("Could not find ExecuteInner switch in ExcelSkillService.cs")
    return set(re.findall(r'"([a-z0-9_]+)"\s*=>', m.group(1)))


def vba_case_names() -> set[str]:
    text = _read("legacy/modExcelSkills.bas")
    m = re.search(
        r"Private Function ExecuteInner.*?Select Case toolName(.*?)Case Else",
        text,
        re.S | re.I,
    )
    if not m:
        raise SystemExit("Could not find ExecuteInner Select Case in modExcelSkills.bas")
    return set(re.findall(r'Case\s+"([a-z0-9_]+)"', m.group(1)))


def vba_tooldef_names() -> set[str]:
    names: set[str] = set()
    for rel in (
        "legacy/modExcelSkills.bas",
        "legacy/modExcelSkillsOps.bas",
        "legacy/modExcelSkillsPivot.bas",
    ):
        path = ROOT / rel
        if not path.exists():
            continue
        text = path.read_text(encoding="utf-8", errors="replace")
        names.update(re.findall(r'TD\(\s*"([a-z0-9_]+)"', text))
        names.update(re.findall(r'"""name"":""([a-z0-9_]+)"""', text))
    return names


def _diff(a: set[str], b: set[str], left: str, right: str) -> str:
    only_a = sorted(a - b)
    only_b = sorted(b - a)
    lines = []
    if only_a:
        lines.append(f"  only in {left}: {only_a}")
    if only_b:
        lines.append(f"  only in {right}: {only_b}")
    return "\n".join(lines)


def main() -> int:
    com_defs = com_make_tool_names()
    com_exec = com_execute_names()
    vba_case = vba_case_names()
    vba_defs = vba_tooldef_names()

    print(f"COM MakeTool: {len(com_defs)} {sorted(com_defs)}")
    print(f"COM ExecuteInner: {len(com_exec)} {sorted(com_exec)}")
    print(f"VBA Case: {len(vba_case)} {sorted(vba_case)}")
    print(f"VBA ToolDef: {len(vba_defs)} {sorted(vba_defs)}")

    failed = False
    if com_defs != com_exec:
        print("FAIL COM defs vs ExecuteInner:\n" + _diff(com_defs, com_exec, "MakeTool", "ExecuteInner"))
        failed = True
    if com_defs != vba_case:
        print("FAIL COM vs VBA Case:\n" + _diff(com_defs, vba_case, "COM", "VBA Case"))
        failed = True
    if com_defs != vba_defs:
        print("FAIL COM vs VBA ToolDef:\n" + _diff(com_defs, vba_defs, "COM", "VBA ToolDef"))
        failed = True
    if vba_case != vba_defs:
        print("FAIL VBA Case vs ToolDef:\n" + _diff(vba_case, vba_defs, "Case", "ToolDef"))
        failed = True

    if failed:
        return 1
    print("OK: COM and VBA tool names are identical")
    return 0


if __name__ == "__main__":
    sys.exit(main())
