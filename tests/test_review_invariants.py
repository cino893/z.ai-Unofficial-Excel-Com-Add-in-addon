#!/usr/bin/env python3
"""Source-level invariants from the v5 review (busy retry, coerce, pivot args)."""
from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def _read(rel: str) -> str:
    return (ROOT / rel).read_text(encoding="utf-8", errors="replace")


def main() -> int:
    failed = False

    skills = _read("src/Services/ExcelSkillService.cs")
    if "TryCoerceNumericCell" not in skills:
        print("FAIL COM missing TryCoerceNumericCell")
        failed = True
    if "if (IsExcelBusy(ex))" not in skills or "throw;" not in skills:
        print("FAIL ExecuteInner must rethrow VBA_E_IGNORE so Execute can retry")
        failed = True
    if "NumberStyles.Any" in skills:
        print("FAIL Int/Dbl still use NumberStyles.Any")
        failed = True

    data = _read("src/Services/ExcelSkillService.Data.cs")
    if "GetValue<int>()" in data:
        print("FAIL remove_duplicates still uses GetValue<int>()")
        failed = True
    if "TryCoerceNumericCell" not in data:
        print("FAIL write_cell does not call TryCoerceNumericCell")
        failed = True

    charts = _read("src/Services/ExcelSkillService.Charts.cs")
    if "GetValue<string>()" in charts:
        print("FAIL Charts.cs still uses GetValue<string>() on pivot fields")
        failed = True
    if "PivotFields(Str(f))" not in charts:
        print("FAIL Charts.cs should pass Str(f) to PivotFields")
        failed = True

    vba_json = _read("legacy/modJSON.bas")
    if "TryCoerceNumericCell" not in vba_json:
        print("FAIL VBA missing TryCoerceNumericCell")
        failed = True

    ops = _read("legacy/modExcelSkillsOps.bas")
    if "values(c)" in ops and "RangeValueAt" not in ops:
        print("FAIL VBA read_range still indexes a 1-row range as 1D values(c)")
        failed = True
    write_range = ops.split("Public Function SkillWriteRange", 1)[-1].split(
        "Public Function SkillGetSheetInfo", 1
    )[0]
    if "TryParseInvNumber" in write_range or "TryCoerceNumericCell" in write_range:
        print("FAIL VBA write_range must not coerce JSON strings to numbers")
        failed = True
    if "RangeValueAt" not in ops:
        print("FAIL VBA read_range missing RangeValueAt 2D-with-1D-fallback")
        failed = True

    conv = _read("legacy/modConversation.bas")
    tool_loop = conv.split("For tc = 1 To toolCalls.Count", 1)[-1].split("Next tc", 1)[0]
    if "DoEvents" not in tool_loop:
        print("FAIL VBA tool loop has no DoEvents after each tool")
        failed = True

    frm = _read("legacy/frmZaiChat.frm")
    build = _read("legacy/build.vbs")
    if "If w < 240 Or h < 240 Then Exit Sub" not in frm:
        print("FAIL frmZaiChat UserForm_Resize min size is not 240")
        failed = True
    if "If w < 240 Or h < 240 Then Exit Sub" not in build:
        print("FAIL build.vbs generated UserForm_Resize min size is not 240")
        failed = True

    addin = _read("src/AddIn.cs")
    if "v5.0.0" not in addin:
        print("FAIL AddIn log is not v5.0.0")
        failed = True

    if failed:
        return 1
    print("OK: review invariants hold (busy retry, coerce, pivot Str, 2D read_range)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
