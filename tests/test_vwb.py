#!/usr/bin/env python3
"""Offline virtual-workbook tests for every Excel tool name."""
from __future__ import annotations

import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from tests.vwb import VWB  # noqa: E402


ALL_TOOLS = [
    "read_cell", "write_cell", "read_range", "write_range",
    "get_sheet_info", "get_workbook_info", "format_range", "insert_formula",
    "sort_range", "add_sheet", "delete_rows", "insert_rows",
    "delete_columns", "insert_columns", "create_chart", "delete_chart",
    "list_charts", "create_pivot_table", "move_pivot_table", "modify_pivot_table",
    "auto_filter", "find_replace", "conditional_format", "copy_range",
    "rename_sheet", "delete_sheet", "freeze_panes", "remove_duplicates",
    "set_validation", "list_pivot_tables", "clear_range",
]


def _seed(wb: VWB) -> None:
    wb.exec("write_range", {
        "start_cell": "A1",
        "data": [
            ["Product", "Category", "Qty", "Price"],
            ["A", "X", 2, 10],
            ["B", "Y", 3, 20],
            ["A", "X", 1, 10],
        ],
    })


class VwbTests(unittest.TestCase):
    def test_every_tool_success_path(self):
        wb = VWB()
        _seed(wb)
        calls = {
            "write_cell": {"cell": "Z1", "value": "1.0"},
            "read_cell": {"cell": "A1"},
            "read_range": {"range": "A1:D4"},
            "write_range": {"start_cell": "F1", "data": [["h"], ["1"]]},
            "get_sheet_info": {},
            "get_workbook_info": {},
            "format_range": {"range": "A1:D1", "bold": True, "underline": True, "font_name": "Arial"},
            "insert_formula": {"cell": "E2", "formula": "C2*D2", "fill_to": "E4"},
            "sort_range": {"range": "A1:D4", "sort_column": "C"},
            "add_sheet": {"name": "Data"},
            "delete_rows": {"start_row": 4, "count": 1},
            "insert_rows": {"at_row": 2, "count": 1},
            "delete_columns": {"start_column": "Z", "count": 1},
            "insert_columns": {"start_column": "C", "count": 1},
            "create_chart": {"data_range": "A1:B4", "chart_type": "column", "title": "T"},
            "list_charts": {},
            "create_pivot_table": {
                "source_range": "A1:D4",
                "row_fields": ["Category"],
                "value_fields": ["Qty"],
                "value_function": "sum",
                "dest_sheet": "Pivot",
                "dest_cell": "A3",
            },
            "list_pivot_tables": {},
            "auto_filter": {"range": "A1:D4", "field": 2, "criteria": "X", "operator": "equals"},
            "find_replace": {"find": "A", "replace": "AA", "match_case": False},
            "conditional_format": {"range": "C2:C4", "rule_type": "color_scale"},
            "copy_range": {"source": "A1:B2", "destination": "H1", "values_only": True},
            "freeze_panes": {"cell": "A2"},
            "remove_duplicates": {"range": "A1:D4", "columns": [1], "has_headers": True},
            "set_validation": {"range": "B2:B10", "type": "list", "formula1": "X,Y"},
            "clear_range": {"range": "Z1", "what": "contents"},
        }
        # Tools that need prior results
        chart = wb.exec("create_chart", calls["create_chart"])
        self.assertTrue(chart.get("success"), chart)
        calls["delete_chart"] = {"chart_name": chart["chart_name"]}

        pivot = wb.exec("create_pivot_table", calls["create_pivot_table"])
        self.assertTrue(pivot.get("success"), pivot)
        calls["move_pivot_table"] = {"name": pivot["name"], "dest_sheet": "PivotMoved", "dest_cell": "A1"}
        calls["modify_pivot_table"] = {
            "name": pivot["name"],
            "row_fields": ["Category"],
            "value_fields": ["Qty"],
            "clear_fields": True,
        }

        extra = wb.exec("add_sheet", {"name": "TempDelete"})
        self.assertTrue(extra.get("success"), extra)
        calls["delete_sheet"] = {"sheet": extra["name"]}
        calls["rename_sheet"] = {"sheet": "Data", "new_name": "Data2"}
        wb.exec("add_sheet", {"name": "Data"})

        seen = set()
        for tool in ALL_TOOLS:
            args = calls.get(tool, {})
            result = wb.exec(tool, args)
            seen.add(tool)
            self.assertNotIn("error", result, f"{tool} failed: {result}")
            if tool not in ("get_sheet_info", "get_workbook_info", "read_cell", "read_range", "list_charts", "list_pivot_tables"):
                self.assertTrue(result.get("success", True), f"{tool}: {result}")
        self.assertEqual(seen, set(ALL_TOOLS))

    def test_write_cell_numeric_coerce(self):
        wb = VWB()
        r = wb.exec("write_cell", {"cell": "A1", "value": "1.0"})
        self.assertTrue(r["success"])
        self.assertEqual(r["value"], 1)
        r2 = wb.exec("write_cell", {"cell": "A2", "value": "hello"})
        self.assertEqual(r2["value"], "hello")

    def test_read_range_truncation(self):
        wb = VWB()
        data = [[i] for i in range(1, 602)]
        wb.exec("write_range", {"start_cell": "A1", "data": data})
        r = wb.exec("read_range", {"range": "A1:A601"})
        self.assertTrue(r.get("truncated"))
        self.assertEqual(r["rows"], 500)
        self.assertEqual(r["total_rows"], 601)

    def test_delete_last_sheet_error(self):
        wb = VWB()
        r = wb.exec("delete_sheet", {"sheet": "Arkusz1"})
        self.assertIn("error", r)
        self.assertIn("last sheet", r["error"].lower())

    def test_unknown_tool(self):
        wb = VWB()
        r = wb.exec("move_table", {})
        self.assertEqual(r.get("error"), "Unknown tool: move_table")

    def test_start_row_lt_1(self):
        wb = VWB()
        r = wb.exec("delete_rows", {"start_row": 0})
        self.assertIn("error", r)
        self.assertIn("start_row", r["error"])

    def test_insert_rows_at_row_lt_1(self):
        wb = VWB()
        r = wb.exec("insert_rows", {"at_row": 0})
        self.assertIn("error", r)

    def test_get_sheet_info_sample_and_flags(self):
        wb = VWB()
        _seed(wb)
        info = wb.exec("get_sheet_info", {})
        self.assertEqual(info["headers"][0], "Product")
        self.assertEqual(len(info["sample_data"]), 4)
        self.assertIn("has_charts", info)
        self.assertIn("has_pivot_tables", info)
        self.assertIn("has_filters", info)

    def test_conditional_format_rule_types(self):
        wb = VWB()
        _seed(wb)
        for rule in ("color_scale", "data_bars", "icon_set", "duplicate", "unique",
                     "above_average", "top_bottom", "formula", "cell_value"):
            r = wb.exec("conditional_format", {"range": "C2:C4", "rule_type": rule, "value1": "10"})
            self.assertTrue(r.get("success"), rule)


def main() -> int:
    suite = unittest.defaultTestLoader.loadTestsFromTestCase(VwbTests)
    result = unittest.TextTestRunner(verbosity=2).run(suite)
    return 0 if result.wasSuccessful() else 1


if __name__ == "__main__":
    sys.exit(main())
