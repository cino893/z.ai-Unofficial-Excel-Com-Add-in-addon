"""
Z.AI API communication test.

Live calls run only when ZAI_API_KEY is set. Offline VWB unit tests live in
tests/test_vwb.py and always run in CI.
"""
from __future__ import annotations

import json
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from tests.vwb import VWB  # noqa: E402

API_BASE = "https://api.z.ai/api/paas/v4"
API_KEY = os.environ.get("ZAI_API_KEY", "").strip()
MODEL = os.environ.get("ZAI_MODEL", "glm-4.7-flash")
MAX_ROUNDS = 30
MAX_SAME_REPEATS = 2


def tools():
    def mt(n, d, p, r):
        return {"type": "function", "function": {"name": n, "description": d,
                                                 "parameters": {"type": "object", "properties": p, "required": r}}}

    def ps(d):
        return {"type": "string", "description": d}

    def pn(d):
        return {"type": "number", "description": d}

    def pb(d):
        return {"type": "boolean", "description": d}

    def pa(d, t="string"):
        return {"type": "array", "description": d, "items": {"type": t}}

    return [
        mt("read_cell", "Read cell", {"cell": ps("Cell"), "sheet": ps("Sheet")}, ["cell"]),
        mt("write_cell", "Write cell", {"cell": ps("Cell"), "value": ps("Value"), "sheet": ps("Sheet")}, ["cell", "value"]),
        mt("read_range", "Read range", {"range": ps("Range"), "sheet": ps("Sheet")}, ["range"]),
        mt("write_range", "Write 2D array", {"start_cell": ps("Start cell"),
            "data": {"type": "array", "description": "2D array", "items": {"type": "array", "items": {"type": "string"}}},
            "sheet": ps("Sheet")}, ["start_cell", "data"]),
        mt("get_sheet_info", "Sheet info", {"sheet": ps("Sheet")}, []),
        mt("get_workbook_info", "Workbook info", {}, []),
        mt("format_range", "Format cells", {"range": ps("Range"), "bold": pb("Bold"), "italic": pb("Italic"),
            "underline": pb("Underline"), "font_name": ps("Font"), "font_size": pn("Size"),
            "font_color": pn("Font color"), "bg_color": pn("BG color"),
            "number_format": ps("Num fmt"), "h_align": ps("Align"), "wrap_text": pb("Wrap"),
            "borders": pb("Borders"), "column_width": pn("Col width"), "row_height": pn("Row height"),
            "autofit": pb("Autofit"), "merge": pb("Merge"), "sheet": ps("Sheet")}, ["range"]),
        mt("insert_formula", "Insert formula", {"cell": ps("Cell"), "formula": ps("Formula"),
            "fill_to": ps("Fill to"), "sheet": ps("Sheet")}, ["cell", "formula"]),
        mt("sort_range", "Sort range", {"range": ps("Range"), "sort_column": ps("Col"),
            "order": ps("asc/desc"), "has_headers": pb("Headers"), "sheet": ps("Sheet")}, ["range", "sort_column"]),
        mt("add_sheet", "Add sheet", {"name": ps("Name")}, []),
        mt("delete_rows", "Delete rows", {"start_row": pn("Start row"), "count": pn("Count"), "sheet": ps("Sheet")}, ["start_row"]),
        mt("insert_rows", "Insert rows", {"at_row": pn("At row"), "count": pn("Count"), "sheet": ps("Sheet")}, ["at_row"]),
        mt("delete_columns", "Delete columns", {"start_column": ps("Col"), "count": pn("Count"), "sheet": ps("Sheet")}, ["start_column"]),
        mt("insert_columns", "Insert columns", {"start_column": ps("Col"), "count": pn("Count"), "sheet": ps("Sheet")}, ["start_column"]),
        mt("create_chart", "Create chart", {"data_range": ps("Range"), "chart_type": ps("Type"),
            "title": ps("Title"), "sheet": ps("Sheet")}, ["data_range"]),
        mt("delete_chart", "Delete chart", {"chart_name": ps("Name"), "sheet": ps("Sheet")}, ["chart_name"]),
        mt("list_charts", "List charts", {"sheet": ps("Sheet")}, []),
        mt("create_pivot_table", "Create PivotTable", {"source_range": ps("Source"), "dest_cell": ps("Dest cell"),
            "dest_sheet": ps("Dest sheet"), "name": ps("Name"),
            "row_fields": pa("Row fields"), "column_fields": pa("Col fields"),
            "value_fields": pa("Value fields"), "value_function": ps("Function"), "sheet": ps("Sheet")},
            ["source_range", "row_fields", "value_fields"]),
        mt("move_pivot_table", "Move a PivotTable to another sheet",
            {"name": ps("PivotTable name"), "source_range": ps("Source range"),
             "dest_sheet": ps("Dest sheet"), "dest_cell": ps("Dest cell"), "sheet": ps("Source sheet")}, []),
        mt("modify_pivot_table", "Configure fields on an existing PivotTable",
            {"name": ps("Name"), "row_fields": pa("Rows"), "column_fields": pa("Cols"),
             "value_fields": pa("Values"), "page_fields": pa("Page"), "value_function": ps("Fn"),
             "clear_fields": pb("Clear"), "refresh": pb("Refresh"), "sheet": ps("Sheet")}, []),
        mt("auto_filter", "Auto filter", {"range": ps("Range"), "field": pn("Field"), "criteria": ps("Criteria"),
            "operator": ps("Op"), "clear": pb("Clear"), "sheet": ps("Sheet")}, ["range"]),
        mt("find_replace", "Find/replace", {"find": ps("Find"), "replace": ps("Replace"), "range": ps("Range"),
            "match_case": pb("Case"), "match_entire": pb("Entire"), "sheet": ps("Sheet")}, ["find", "replace"]),
        mt("conditional_format", "Conditional format", {"range": ps("Range"), "rule_type": ps("Rule type"),
            "operator": ps("Op"), "value1": ps("V1"), "value2": ps("V2"),
            "format_color": pn("Color"), "sheet": ps("Sheet")}, ["range", "rule_type"]),
        mt("copy_range", "Copy range", {"source": ps("Source"), "destination": ps("Dest"), "dest_sheet": ps("Dest sheet"),
            "values_only": pb("Values only"), "sheet": ps("Sheet")}, ["source", "destination"]),
        mt("rename_sheet", "Rename sheet", {"sheet": ps("Sheet"), "new_name": ps("New name")}, ["new_name"]),
        mt("delete_sheet", "Delete sheet", {"sheet": ps("Sheet name")}, ["sheet"]),
        mt("freeze_panes", "Freeze panes", {"cell": ps("Cell"), "unfreeze": pb("Unfreeze"), "sheet": ps("Sheet")}, []),
        mt("remove_duplicates", "Remove duplicates", {"range": ps("Range"), "columns": pa("Cols", "number"),
            "has_headers": pb("Headers"), "sheet": ps("Sheet")}, ["range"]),
        mt("set_validation", "Data validation", {"range": ps("Range"), "type": ps("Type"), "formula1": ps("F1"),
            "sheet": ps("Sheet")}, ["range", "type", "formula1"]),
        mt("list_pivot_tables", "List PivotTables", {"sheet": ps("Sheet")}, []),
        mt("clear_range", "Clear range", {"range": ps("Range"), "what": ps("contents/formats/all"), "sheet": ps("Sheet")}, ["range"]),
    ]


SYSTEM_PROMPT = """You are an AI assistant integrated into Microsoft Excel through the Z.AI add-in. You have access to tools that can read and modify Excel workbooks.

Rules you must follow:
1. Always call get_sheet_info or get_workbook_info first to understand the current state of the workbook before taking any action.
2. Always read data before modifying it. Never assume the contents of cells.
3. After making changes, confirm what was done by reading back the affected cells.
4. Write all formulas using English function names (SUM, AVERAGE, IF, VLOOKUP, COUNT, MAX, MIN, etc.).
5. When setting colors, use RGB Long values: Red=255, Green=65280, Blue=16711680, Yellow=65535, White=16777215, Black=0.
6. Default to the active sheet unless the user specifies otherwise.
7. Before creating charts, call list_charts first. If a similar chart already exists, delete it before creating a new one.
8. When creating PivotTables, read headers first with get_sheet_info to know exact field names.
9. If a pivot table blocks an operation (e.g. delete_rows, insert_rows), use move_pivot_table to move it to a separate sheet. If move_pivot_table returns a COM error, use create_pivot_table then modify_pivot_table.
10. Do not repeat operations that have already been completed successfully.
11. Communicate with the user in Polish.
12. Plan your actions efficiently — try to stay within {max_rounds} tool rounds per request and batch related operations when possible.
13. Do not explain what you are doing step by step. Instead, after completing work, write a short summary listing which cells/ranges were changed and what was done."""


def call_api(msgs, tls):
    import requests
    r = requests.post(
        f"{API_BASE}/chat/completions",
        json={"model": MODEL, "messages": msgs, "max_tokens": 4096, "temperature": 0.7,
              "tools": tls, "tool_choice": "auto"},
        headers={"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"},
        timeout=120,
    )
    if r.status_code != 200:
        print(f"  HTTP {r.status_code}: {r.text[:300]}")
        return None
    return r.json()


def run(prompt, wb, label, max_rounds=MAX_ROUNDS):
    print(f"\n{'=' * 60}\n  TEST: {label}\n  User: {prompt}\n{'=' * 60}")
    msgs = [{"role": "system", "content": SYSTEM_PROMPT.format(max_rounds=max_rounds)},
            {"role": "user", "content": prompt}]
    prev_sig, reps, total_tc = None, 0, 0
    ri_idx = None

    for rnd in range(1, max_rounds + 1):
        ri_msg = {"role": "system", "content": f"[Round {rnd}/{max_rounds}]"}
        if ri_idx is not None and ri_idx < len(msgs):
            msgs[ri_idx] = ri_msg
        else:
            msgs.append(ri_msg)
            ri_idx = len(msgs) - 1

        if rnd == max_rounds:
            msgs.append({"role": "user", "content": (
                f"This is your final response before reaching the {max_rounds} tool-round limit. "
                "Write yourself a summary of what you did and what's left."
            )})

        print(f"  R{rnd}/{max_rounds}...", end="", flush=True)
        data = call_api(msgs, tools())
        if not data:
            return {"stop": "api_error", "rounds": rnd, "tools": total_tc}

        msg = data["choices"][0]["message"]
        if msg.get("tool_calls"):
            tcl = msg["tool_calls"]
            print(f" {len(tcl)} tool(s)")
            msgs.append(msg)
            sig = "|".join(f"{t['function']['name']}:{t['function']['arguments']}" for t in tcl)
            if sig == prev_sig:
                reps += 1
                if reps >= MAX_SAME_REPEATS:
                    print("  LOOP DETECTED")
                    return {"stop": "loop", "rounds": rnd, "tools": total_tc}
            else:
                reps = 0
            prev_sig = sig

            for tc in tcl:
                fn = tc["function"]["name"]
                fa = tc["function"]["arguments"]
                total_tc += 1
                try:
                    args = json.loads(fa)
                except json.JSONDecodeError:
                    args = {}
                result = wb.exec(fn, args)
                rstr = json.dumps(result, ensure_ascii=False)
                print(f"    {fn}({fa[:70]}) -> {rstr[:90]}")
                msgs.append({"role": "tool", "content": rstr, "tool_call_id": tc["id"]})
            continue

        content = msg.get("content", "")
        if not content:
            print(" EMPTY")
            return {"stop": "empty", "rounds": rnd, "tools": total_tc}
        print(" done")
        for line in content.split("\n"):
            print(f"    {line}")
        print(f"  {rnd} rounds, {total_tc} tool calls")
        return {"stop": "ok", "rounds": rnd, "tools": total_tc, "response": content}

    print("  MAX ROUNDS")
    return {"stop": "max_rounds", "rounds": max_rounds, "tools": total_tc}


def main() -> int:
    print("Z.AI API Test")
    if not API_KEY:
        print("ZAI_API_KEY is not set; skipping live API tests.")
        print("Offline coverage: python3 tests/test_vwb.py tests/test_tool_parity.py tests/test_i18n.py")
        return 0

    print(f"  Model: {MODEL}, Max rounds: {MAX_ROUNDS}\n")
    import requests
    print("API check...", end="", flush=True)
    try:
        r = requests.post(
            f"{API_BASE}/chat/completions",
            json={"model": MODEL, "messages": [{"role": "user", "content": "test"}], "max_tokens": 5},
            headers={"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"},
            timeout=30,
        )
        if r.status_code != 200:
            print(f" HTTP {r.status_code}")
            return 1
        print(" OK")
    except Exception as e:
        print(f" {e}")
        return 1

    results = {}
    wb1 = VWB()
    results["T1: tabela+pivot"] = run(
        "Stworz mala tabele (5 wierszy danych + naglowek) z kolumnami: Produkt, Kategoria, Ilosc, Cena. "
        "Oblicz kolumne Wartosc (Ilosc*Cena). Nastepnie stworz tabele przestawna podsumowujaca Wartosc wg Kategorii.",
        wb1, "Mala tabela + pivot (5 wierszy)", max_rounds=20)

    if results.get("T1: tabela+pivot", {}).get("stop") == "ok":
        results["T2: ladniejsza"] = run(
            "Sformatuj tabele — pogrubiony naglowek, obramowanie, format walutowy dla Cena i Wartosc.",
            wb1, "Upiekszenie", max_rounds=15)

    wb3 = VWB()
    for r in range(1, 7):
        wb3.sheets["Arkusz1"][(r, 1)] = ["Produkt", "A", "B", "C", "A", "B"][r - 1]
        wb3.sheets["Arkusz1"][(r, 2)] = ["Wartosc", 100, 200, 300, 150, 250][r - 1]
    wb3.pivots["PT1"] = {"source": "'Arkusz1'!A1:B6", "dest_sheet": "Arkusz1", "dest_cell": "D1",
                         "row_fields": ["Produkt"], "value_fields": ["Wartosc"]}
    results["T3: przenies pivot"] = run(
        "Na arkuszu Arkusz1 mam dane w A1:B6 i tabele przestawna PT1 w D1. "
        "Przenies tabele przestawna na osobny arkusz (move_pivot_table) i usun puste wiersze 8-20 z danych.",
        wb3, "Przeniesienie pivot + delete_rows", max_rounds=20)

    wb4 = VWB()
    results["T4: pusty"] = run("Podsumuj dane w tym arkuszu", wb4, "Pusty arkusz", max_rounds=10)

    print(f"\n{'=' * 60}\n  PODSUMOWANIE\n{'=' * 60}")
    for name, res in results.items():
        if res:
            print(f"  {name}: stop={res['stop']}, rounds={res['rounds']}, tools={res['tools']}")
        else:
            print(f"  {name}: brak wyniku")
    print(f"\n  WB po T3: sheets={list(wb3.sheets.keys())} pivots={wb3.pivots}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
