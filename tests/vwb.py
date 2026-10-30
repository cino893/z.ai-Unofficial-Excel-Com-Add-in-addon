"""Virtual workbook used by offline skill tests and optional live API tests."""
from __future__ import annotations

from typing import Any


class VWB:
    def __init__(self):
        self.sheets: dict[str, dict[tuple[int, int], Any]] = {"Arkusz1": {}}
        self.active = "Arkusz1"
        self.pivots: dict[str, dict] = {}
        self.charts: dict[str, dict] = {}
        self.filters: dict[str, dict] = {}
        self.frozen: dict[str, Any] = {}
        self.validations: dict[str, dict] = {}
        self.formats: list[dict] = []
        self.pivot_caches: list[str] = []

    def _gs(self, s=None):
        return s or self.active

    def _pc(self, a: str):
        col, row = "", ""
        for c in a.upper().replace("$", ""):
            if c.isalpha():
                col += c
            else:
                row += c
        ci = 0
        for c in col:
            ci = ci * 26 + ord(c) - 64
        return int(row or 0), ci

    def _cl(self, i: int) -> str:
        r = ""
        while i > 0:
            i -= 1
            r = chr(i % 26 + 65) + r
            i //= 26
        return r

    def _pr(self, rng: str):
        p = rng.replace("$", "").split(":")
        r1, c1 = self._pc(p[0])
        if len(p) == 1:
            return r1, c1, r1, c1
        r2, c2 = self._pc(p[1])
        return r1, c1, r2, c2

    def _coerce(self, v):
        if isinstance(v, bool) or v is None:
            return v
        if isinstance(v, (int, float)):
            return v
        s = str(v)
        try:
            n = float(s)
            if n == int(n):
                return int(n)
            return n
        except (TypeError, ValueError):
            return v

    def exec(self, name: str, args: dict) -> dict:
        s = args.get("sheet")
        sn = self._gs(s)
        try:
            if name == "get_workbook_info":
                result = {
                    "name": "Test.xlsx",
                    "path": r"C:\Test.xlsx",
                    "sheets": list(self.sheets.keys()),
                    "active_sheet": self.active,
                }
                if self.pivot_caches or self.pivots:
                    result["pivot_cache_count"] = max(len(self.pivot_caches), len(self.pivots))
                return result
            if name == "get_sheet_info":
                cells = self.sheets.get(sn, {})
                if not cells:
                    return {
                        "name": sn, "index": list(self.sheets).index(sn) + 1 if sn in self.sheets else 1,
                        "used_range": "", "used_rows": 0, "used_cols": 0,
                        "headers": [], "sample_data": [],
                        "has_charts": any(c.get("sheet") == sn for c in self.charts.values()),
                        "has_pivot_tables": any(p.get("dest_sheet") == sn for p in self.pivots.values()),
                        "has_filters": sn in self.filters,
                    }
                mr = max(r for r, c in cells)
                mc = max(c for r, c in cells)
                mnr = min(r for r, c in cells)
                mnc = min(c for r, c in cells)
                max_cols = min(mc - mnc + 1, 26)
                hdrs = [str(cells.get((mnr, c), "") or "") for c in range(mnc, mnc + max_cols)]
                sample = []
                for r in range(mnr, min(mnr + 5, mr + 1)):
                    sample.append([str(cells.get((r, c), "") or "") for c in range(mnc, mnc + max_cols)])
                return {
                    "name": sn, "index": list(self.sheets).index(sn) + 1,
                    "used_range": f"{self._cl(mnc)}{mnr}:{self._cl(mc)}{mr}",
                    "used_rows": mr - mnr + 1, "used_cols": mc - mnc + 1,
                    "first_cell": f"{self._cl(mnc)}{mnr}",
                    "last_cell": f"{self._cl(mc)}{mr}",
                    "headers": hdrs, "sample_data": sample,
                    "has_charts": any(c.get("sheet") == sn for c in self.charts.values()),
                    "has_pivot_tables": any(p.get("dest_sheet") == sn for p in self.pivots.values()),
                    "has_filters": sn in self.filters,
                }
            if name == "read_cell":
                r, c = self._pc(args["cell"])
                v = self.sheets.get(sn, {}).get((r, c), "")
                return {"cell": args["cell"], "value": v, "formula": "", "type": type(v).__name__, "sheet": sn}
            if name == "write_cell":
                r, c = self._pc(args["cell"])
                v = self._coerce(args["value"])
                self.sheets.setdefault(sn, {})[(r, c)] = v
                return {"success": True, "cell": args["cell"], "value": v}
            if name == "read_range":
                r1, c1, r2, c2 = self._pr(args["range"])
                total_rows = r2 - r1 + 1
                truncated = total_rows > 500
                rows = 500 if truncated else total_rows
                data = []
                for r in range(r1, r1 + rows):
                    data.append([self.sheets.get(sn, {}).get((r, c)) for c in range(c1, c2 + 1)])
                result = {"range": args["range"], "sheet": sn, "rows": rows, "cols": c2 - c1 + 1, "data": data}
                if truncated:
                    result["truncated"] = True
                    result["total_rows"] = total_rows
                return result
            if name == "write_range":
                r, c = self._pc(args["start_cell"])
                data = args.get("data") or []
                w = 0
                for ri, row in enumerate(data):
                    for ci, val in enumerate(row):
                        self.sheets.setdefault(sn, {})[(r + ri, c + ci)] = self._coerce(val)
                        w += 1
                return {"success": True, "start_cell": args["start_cell"], "rows_written": len(data), "cells_written": w}
            if name == "format_range":
                self.formats.append({"range": args["range"], "sheet": sn, **args})
                return {"success": True, "range": args["range"]}
            if name == "insert_formula":
                r, c = self._pc(args["cell"])
                formula = args["formula"]
                if not str(formula).startswith("="):
                    formula = "=" + str(formula)
                self.sheets.setdefault(sn, {})[(r, c)] = f"[F:{formula}]"
                filled = 1
                if args.get("fill_to"):
                    r2, c2 = self._pc(args["fill_to"])
                    filled = max(1, abs(r2 - r) + 1) * max(1, abs(c2 - c) + 1)
                return {"success": True, "cell": args["cell"], "formula": formula, "result": "", "cells_filled": filled}
            if name == "sort_range":
                return {"success": True, "range": args["range"], "sort_column": args.get("sort_column", "A")}
            if name == "add_sheet":
                n = args.get("name") or f"Arkusz{len(self.sheets) + 1}"
                if n in self.sheets:
                    actual = f"Sheet{len(self.sheets) + 1}"
                    self.sheets[actual] = {}
                    return {"success": True, "name": actual, "warning": f"Could not rename sheet to '{n}' (duplicate or invalid name). Created as '{actual}'."}
                self.sheets[n] = {}
                return {"success": True, "name": n}
            if name == "delete_rows":
                sr = int(args["start_row"])
                if sr < 1:
                    return {"error": "start_row must be >= 1"}
                return {"success": True, "deleted_from": sr, "count": args.get("count", 1)}
            if name == "insert_rows":
                ar = int(args["at_row"])
                if ar < 1:
                    return {"error": "at_row must be >= 1"}
                return {"success": True, "at_row": ar, "count": args.get("count", 1)}
            if name == "delete_columns":
                sc = str(args.get("start_column") or "")
                if not sc:
                    return {"error": "start_column is required (e.g. 'C')"}
                count = int(args.get("count", 1) or 1)
                start_n = self._pc(sc + "1")[1]
                end = self._cl(start_n + count - 1)
                return {"success": True, "deleted_from": sc.upper(), "deleted_to": end, "count": count}
            if name == "insert_columns":
                sc = str(args.get("start_column") or "")
                if not sc:
                    return {"error": "start_column is required (e.g. 'C')"}
                count = int(args.get("count", 1) or 1)
                start_n = self._pc(sc + "1")[1]
                end = self._cl(start_n + count - 1)
                return {"success": True, "inserted_at": sc.upper(), "inserted_to": end, "count": count}
            if name == "create_chart":
                n = f"Chart {len(self.charts) + 1}"
                self.charts[n] = {"type": args.get("chart_type", "column"), "data": args["data_range"], "sheet": sn}
                return {"success": True, "chart_name": n, "type": args.get("chart_type", "column")}
            if name == "delete_chart":
                cn = args["chart_name"]
                if cn in self.charts:
                    del self.charts[cn]
                    return {"success": True, "deleted": cn}
                return {"error": f"Chart not found: {cn}"}
            if name == "list_charts":
                items = [{"name": k, "type": v["type"], "source": v.get("data", "")}
                         for k, v in self.charts.items() if v.get("sheet") == sn]
                return {"sheet": sn, "charts": items, "count": len(items)}
            if name == "create_pivot_table":
                pn = args.get("name") or f"PivotTable{len(self.pivots) + 1}"
                ds = args.get("dest_sheet") or sn
                dc = args.get("dest_cell")
                if not dc and not args.get("dest_sheet"):
                    ds = f"Pivot_{pn}"
                    self.sheets[ds] = {}
                    dc = "A3"
                elif ds not in self.sheets:
                    self.sheets[ds] = {}
                    dc = dc or "A3"
                self.pivots[pn] = {
                    "source": f"'{sn}'!{args['source_range']}",
                    "dest_sheet": ds, "dest_cell": dc or "A1",
                    "row_fields": args.get("row_fields") or [],
                    "column_fields": args.get("column_fields") or [],
                    "value_fields": args.get("value_fields") or [],
                    "value_function": args.get("value_function", "sum"),
                    "page_fields": [],
                }
                self.pivot_caches.append(f"{sn}!{args['source_range']}")
                return {"success": True, "name": pn, "dest_sheet": ds,
                        "row_fields": len(self.pivots[pn]["row_fields"]),
                        "value_fields": len(self.pivots[pn]["value_fields"])}
            if name == "list_pivot_tables":
                pts = [{"name": k, "sheet": v["dest_sheet"], "location": v["dest_cell"], "source": v["source"]}
                       for k, v in self.pivots.items()
                       if not s or v["dest_sheet"] == s]
                return {"pivot_tables": pts, "count": len(pts)}
            if name == "move_pivot_table":
                ds = args.get("dest_sheet") or f"Moved_{len(self.sheets) + 1}"
                if ds not in self.sheets:
                    self.sheets[ds] = {}
                pn = args.get("name")
                src_range = args.get("source_range")
                found = None
                if pn and pn in self.pivots:
                    found = pn
                elif src_range:
                    for k, v in self.pivots.items():
                        if v.get("dest_sheet") == sn:
                            found = k
                            break
                if not found and not self.pivots:
                    return {
                        "error": "No pivot table found. Check the name or source_range.",
                        "hint": "Use list_pivot_tables to find pivot tables, or use copy_range to move plain data.",
                    }
                if found:
                    self.pivots[found]["dest_sheet"] = ds
                    self.pivots[found]["dest_cell"] = args.get("dest_cell", "A1")
                    return {"success": True, "method": "location_move", "name": found,
                            "from_sheet": sn, "to_sheet": ds, "dest_cell": args.get("dest_cell", "A1")}
                new_name = (pn or "PivotTable") + "_moved"
                self.pivots[new_name] = {"source": "", "dest_sheet": ds, "dest_cell": args.get("dest_cell", "A1"),
                                         "row_fields": [], "value_fields": []}
                return {"success": True, "method": "pivot_recreated", "new_name": new_name,
                        "to_sheet": ds, "dest_cell": args.get("dest_cell", "A1"),
                        "next_step": "Pivot recreated empty — use modify_pivot_table to configure row_fields, value_fields, column_fields."}
            if name == "modify_pivot_table":
                pn = args.get("name")
                pt = None
                if pn and pn in self.pivots:
                    pt = self.pivots[pn]
                elif self.pivots:
                    pt = next(iter(self.pivots.values()))
                    pn = next(iter(self.pivots))
                if pt is None:
                    return {"error": f"Pivot table '{pn}' not found. Use list_pivot_tables to find available pivots."}
                if args.get("clear_fields", True):
                    if args.get("row_fields") is not None or args.get("value_fields") is not None \
                            or args.get("column_fields") is not None or args.get("page_fields") is not None:
                        pt["row_fields"] = []
                        pt["column_fields"] = []
                        pt["value_fields"] = []
                        pt["page_fields"] = []
                changes = 0
                for key in ("row_fields", "column_fields", "value_fields", "page_fields"):
                    if args.get(key) is not None:
                        pt[key] = list(args[key])
                        changes += len(args[key])
                if args.get("value_function"):
                    pt["value_function"] = args["value_function"]
                layout = {
                    "row_fields": pt.get("row_fields", []),
                    "column_fields": pt.get("column_fields", []),
                    "value_fields": pt.get("value_fields", []),
                }
                if pt.get("page_fields"):
                    layout["page_fields"] = pt["page_fields"]
                return {"success": True, "name": pn, "changes": changes, "current_layout": layout}
            if name == "auto_filter":
                if args.get("clear"):
                    self.filters.pop(sn, None)
                    return {"success": True, "action": "cleared"}
                field = args.get("field") or 0
                self.filters[sn] = {"range": args["range"], "field": field, "operator": args.get("operator", "equals")}
                return {"success": True, "range": args["range"], "action": "filtered" if field else "toggled"}
            if name == "find_replace":
                return {"success": True, "find": args["find"], "replace": args["replace"]}
            if name == "conditional_format":
                return {"success": True, "range": args["range"], "rule_type": args.get("rule_type", "cell_value")}
            if name == "copy_range":
                src = args["source"]
                dest = args["destination"]
                r1, c1, r2, c2 = self._pr(src)
                dr, dc = self._pc(dest)
                dsn = args.get("dest_sheet") or sn
                self.sheets.setdefault(dsn, {})
                for r in range(r1, r2 + 1):
                    for c in range(c1, c2 + 1):
                        self.sheets[dsn][(dr + (r - r1), dc + (c - c1))] = self.sheets.get(sn, {}).get((r, c))
                return {"success": True, "source": src, "destination": dest}
            if name == "rename_sheet":
                old = self._gs(s)
                nn = args["new_name"]
                self.sheets[nn] = self.sheets.pop(old)
                if self.active == old:
                    self.active = nn
                return {"success": True, "old_name": old, "new_name": nn}
            if name == "delete_sheet":
                dn = args["sheet"]
                if len(self.sheets) <= 1:
                    return {"error": "Cannot delete the last sheet in the workbook"}
                if dn in self.sheets:
                    del self.sheets[dn]
                    if self.active == dn:
                        self.active = next(iter(self.sheets))
                    return {"success": True, "deleted": dn}
                return {"error": f"Sheet '{dn}' not found"}
            if name == "freeze_panes":
                if args.get("unfreeze"):
                    self.frozen.pop(sn, None)
                    return {"success": True, "action": "unfrozen"}
                cell = args.get("cell", "A2")
                self.frozen[sn] = cell
                return {"success": True, "action": "frozen", "cell": cell}
            if name == "remove_duplicates":
                return {"success": True, "range": args["range"], "rows_before": 3, "rows_after": 2, "rows_removed": 1}
            if name == "set_validation":
                self.validations[f"{sn}:{args['range']}"] = args
                return {"success": True, "range": args["range"], "type": args.get("type")}
            if name == "clear_range":
                what = args.get("what", "contents")
                r1, c1, r2, c2 = self._pr(args["range"])
                if what != "formats":
                    for r in range(r1, r2 + 1):
                        for c in range(c1, c2 + 1):
                            self.sheets.get(sn, {}).pop((r, c), None)
                return {"success": True, "range": args["range"], "cleared": what}
            return {"error": f"Unknown tool: {name}"}
        except Exception as ex:
            return {"error": str(ex)}
