"""
Z.AI API Communication Test (simplified)
Tests: random table + pivot, prettify, move_table with pivot
"""
import requests, json, sys, time

API_BASE = "https://api.z.ai/api/paas/v4"
API_KEY = "8cf9f0dda0b147f88eba639767510300.jZoc956GGNMKrdtO"
MODEL = "glm-4.7-flash"
MAX_ROUNDS = 30  # cap for test speed
MAX_SAME_REPEATS = 2

# ── Virtual Workbook (minimal) ─────────────────────────────────────
class VWB:
    def __init__(self):
        self.sheets = {"Arkusz1": {}}
        self.active = "Arkusz1"
        self.pivots = {}
        self.charts = {}

    def _gs(self, s=None): return s or self.active
    def _pc(self, a):
        col, row = "", ""
        for c in a.upper().replace("$",""):
            if c.isalpha(): col += c
            else: row += c
        ci = 0
        for c in col: ci = ci*26 + ord(c)-64
        return int(row), ci

    def _cl(self, i):
        r = ""
        while i > 0:
            i -= 1; r = chr(i%26+65) + r; i //= 26
        return r

    def _pr(self, rng):
        p = rng.replace("$","").split(":")
        r1,c1 = self._pc(p[0])
        if len(p)==1: return r1,c1,r1,c1
        r2,c2 = self._pc(p[1])
        return r1,c1,r2,c2

    def exec(self, name, args):
        s = args.get("sheet")
        sn = self._gs(s)
        try:
            if name == "get_workbook_info":
                return {"file_name":"Test.xlsx","path":"C:\\Test.xlsx",
                        "sheets":list(self.sheets.keys()),"active_sheet":self.active}
            elif name == "get_sheet_info":
                cells = self.sheets.get(sn, {})
                if not cells:
                    return {"name":sn,"used_range":"","rows":0,"cols":0,"headers":[]}
                mr = max(r for r,c in cells); mc = max(c for r,c in cells)
                mnr = min(r for r,c in cells); mnc = min(c for r,c in cells)
                hdrs = [str(cells.get((mnr,c),"")) for c in range(mnc, mc+1)]
                return {"name":sn,"used_range":f"{self._cl(mnc)}{mnr}:{self._cl(mc)}{mr}",
                        "rows":mr-mnr+1,"cols":mc-mnc+1,"headers":hdrs}
            elif name == "read_cell":
                r,c = self._pc(args["cell"])
                v = self.sheets.get(sn,{}).get((r,c),"")
                return {"cell":args["cell"],"value":v,"formula":"","type":"string" if v else "empty","sheet":sn}
            elif name == "write_cell":
                r,c = self._pc(args["cell"])
                v = args["value"]
                try: v = float(v);
                except: pass
                if v != args["value"] and v == int(v): v = int(v)
                self.sheets.setdefault(sn,{})[(r,c)] = v
                return {"success":True,"cell":args["cell"],"value":v}
            elif name == "read_range":
                r1,c1,r2,c2 = self._pr(args["range"])
                data = []
                for r in range(r1,r2+1):
                    row = [self.sheets.get(sn,{}).get((r,c)) for c in range(c1,c2+1)]
                    data.append(row)
                return {"range":args["range"],"sheet":sn,"rows":r2-r1+1,"cols":c2-c1+1,"data":data}
            elif name == "write_range":
                r,c = self._pc(args["start_cell"])
                w = 0
                for ri, row in enumerate(args["data"]):
                    for ci, val in enumerate(row):
                        try:
                            val = float(val)
                            if val == int(val): val = int(val)
                        except: pass
                        self.sheets.setdefault(sn,{})[(r+ri, c+ci)] = val
                        w += 1
                return {"success":True,"start_cell":args["start_cell"],"rows_written":len(args["data"]),"cells_written":w}
            elif name == "format_range":
                return {"success":True,"range":args["range"],"sheet":sn}
            elif name == "insert_formula":
                r,c = self._pc(args["cell"])
                self.sheets.setdefault(sn,{})[(r,c)] = f"[F:{args['formula']}]"
                return {"success":True,"cell":args["cell"],"formula":args["formula"]}
            elif name == "add_sheet":
                n = args.get("name") or f"Arkusz{len(self.sheets)+1}"
                self.sheets[n] = {}
                return {"success":True,"name":n}
            elif name == "delete_rows":
                sr = args["start_row"]
                if sr < 1: return {"error":"start_row must be >= 1"}
                return {"success":True,"deleted_from":sr,"count":args.get("count",1)}
            elif name == "insert_rows":
                ar = args["at_row"]
                if ar < 1: return {"error":"at_row must be >= 1"}
                return {"success":True,"at_row":ar,"count":args.get("count",1)}
            elif name == "create_chart":
                n = f"Chart {len(self.charts)+1}"
                self.charts[n] = {"type":args.get("chart_type","column"),"data":args["data_range"]}
                return {"success":True,"chart_name":n,"type":args.get("chart_type","column")}
            elif name == "delete_chart":
                cn = args["chart_name"]
                if cn in self.charts: del self.charts[cn]; return {"success":True,"deleted":cn}
                return {"error":f"Chart '{cn}' not found"}
            elif name == "list_charts":
                return {"charts":[{"name":k,"type":v["type"],"data_range":v["data"]} for k,v in self.charts.items()],"count":len(self.charts)}
            elif name == "create_pivot_table":
                pn = args.get("name") or f"PT{len(self.pivots)+1}"
                ds = sn
                dc = args.get("dest_cell")
                if not dc:
                    ds = f"Pivot_{pn}"; self.sheets[ds] = {}; dc = "A1"
                self.pivots[pn] = {"source":f"'{sn}'!{args['source_range']}","dest_sheet":ds,"dest_cell":dc,
                                   "row_fields":args.get("row_fields",[]),"value_fields":args.get("value_fields",[])}
                return {"success":True,"name":pn,"dest_sheet":ds,"dest_cell":dc}
            elif name == "list_pivot_tables":
                pts = [{"name":k,"sheet":v["dest_sheet"],"location":v["dest_cell"],"source":v["source"]}
                       for k,v in self.pivots.items()
                       if not s or v["dest_sheet"]==s]
                return {"pivot_tables":pts,"count":len(pts)}
            elif name == "move_table":
                ds = args.get("dest_sheet")
                if not ds:
                    ds = f"Moved_{len(self.sheets)+1}"
                if ds not in self.sheets:
                    self.sheets[ds] = {}
                pn = args.get("name")
                if pn and pn in self.pivots:
                    self.pivots[pn]["dest_sheet"] = ds
                    self.pivots[pn]["dest_cell"] = args.get("dest_cell","A1")
                    return {"success":True,"moved":"pivot_table","name":pn,"to_sheet":ds,"dest_cell":args.get("dest_cell","A1")}
                return {"success":True,"moved":"data","to_sheet":ds,"dest_cell":args.get("dest_cell","A1")}
            elif name == "clear_range":
                return {"success":True,"range":args["range"],"cleared":args.get("what","contents")}
            elif name == "sort_range":
                return {"success":True,"range":args["range"],"sorted_by":args.get("sort_column","A")}
            elif name == "auto_filter":
                return {"success":True,"range":args["range"]}
            elif name == "find_replace":
                return {"success":True,"find":args["find"],"replace":args["replace"],"replacements_made":0}
            elif name == "conditional_format":
                return {"success":True,"range":args["range"],"rule_type":args.get("rule_type","")}
            elif name == "copy_range":
                return {"success":True,"source":args["source"],"destination":args["destination"]}
            elif name == "rename_sheet":
                old = self._gs(s)
                nn = args["new_name"]
                self.sheets[nn] = self.sheets.pop(old)
                if self.active == old: self.active = nn
                return {"success":True,"old_name":old,"new_name":nn}
            elif name == "delete_sheet":
                dn = args["sheet"]
                if len(self.sheets) <= 1: return {"error":"Cannot delete the last sheet"}
                if dn in self.sheets: del self.sheets[dn]; return {"success":True,"deleted":dn}
                return {"error":f"Sheet '{dn}' not found"}
            elif name == "freeze_panes":
                return {"success":True}
            elif name == "remove_duplicates":
                return {"success":True,"range":args["range"],"rows_removed":0}
            elif name == "set_validation":
                return {"success":True,"range":args["range"]}
            else:
                return {"error":f"Unknown tool: {name}"}
        except Exception as ex:
            return {"error":str(ex)}


# ── Tool definitions (compact) ──────────────────────────────────────
def tools():
    def mt(n,d,p,r):
        return {"type":"function","function":{"name":n,"description":d,"parameters":{"type":"object","properties":p,"required":r}}}
    def ps(d): return {"type":"string","description":d}
    def pn(d): return {"type":"number","description":d}
    def pb(d): return {"type":"boolean","description":d}
    return [
        mt("read_cell","Read cell",{"cell":ps("Cell"),"sheet":ps("Sheet")},["cell"]),
        mt("write_cell","Write cell",{"cell":ps("Cell"),"value":ps("Value"),"sheet":ps("Sheet")},["cell","value"]),
        mt("read_range","Read range",{"range":ps("Range"),"sheet":ps("Sheet")},["range"]),
        mt("write_range","Write 2D array",{"start_cell":ps("Start cell"),
            "data":{"type":"array","description":"2D array","items":{"type":"array","items":{"type":"string"}}},
            "sheet":ps("Sheet")},["start_cell","data"]),
        mt("get_sheet_info","Sheet info",{"sheet":ps("Sheet")},[]),
        mt("get_workbook_info","Workbook info",{},{}),
        mt("format_range","Format cells",{"range":ps("Range"),"bold":pb("Bold"),"italic":pb("Italic"),
            "font_size":pn("Size"),"font_color":pn("Font color"),"bg_color":pn("BG color"),
            "number_format":ps("Num fmt"),"h_align":ps("Align"),"wrap_text":pb("Wrap"),
            "borders":pb("Borders"),"column_width":pn("Col width"),"row_height":pn("Row height"),
            "autofit":pb("Autofit"),"merge":pb("Merge"),"sheet":ps("Sheet")},["range"]),
        mt("insert_formula","Insert formula",{"cell":ps("Cell"),"formula":ps("Formula"),"sheet":ps("Sheet")},["cell","formula"]),
        mt("sort_range","Sort range",{"range":ps("Range"),"sort_column":ps("Col"),"order":ps("asc/desc"),"has_headers":pb("Headers"),"sheet":ps("Sheet")},["range","sort_column"]),
        mt("add_sheet","Add sheet",{"name":ps("Name")},[]),
        mt("delete_rows","Delete rows",{"start_row":pn("Start row"),"count":pn("Count"),"sheet":ps("Sheet")},["start_row"]),
        mt("insert_rows","Insert rows",{"at_row":pn("At row"),"count":pn("Count"),"sheet":ps("Sheet")},["at_row"]),
        mt("create_chart","Create chart",{"data_range":ps("Range"),"chart_type":ps("Type"),"title":ps("Title"),"sheet":ps("Sheet")},["data_range"]),
        mt("delete_chart","Delete chart",{"chart_name":ps("Name"),"sheet":ps("Sheet")},["chart_name"]),
        mt("list_charts","List charts",{"sheet":ps("Sheet")},[]),
        mt("create_pivot_table","Create PivotTable",{"source_range":ps("Source"),"dest_cell":ps("Dest cell"),
            "name":ps("Name"),"row_fields":{"type":"array","items":{"type":"string"},"description":"Row fields"},
            "column_fields":{"type":"array","items":{"type":"string"},"description":"Col fields"},
            "value_fields":{"type":"array","items":{"type":"string"},"description":"Value fields"},
            "value_function":ps("Function"),"sheet":ps("Sheet")},["source_range","row_fields","value_fields"]),
        mt("move_table","Move data/PivotTable to another sheet. Use when pivot blocks delete_rows/insert_rows.",
            {"name":ps("PivotTable name"),"source_range":ps("Source range"),
             "dest_sheet":ps("Dest sheet"),"dest_cell":ps("Dest cell"),"sheet":ps("Source sheet")},[]),
        mt("auto_filter","Auto filter",{"range":ps("Range"),"field":pn("Field"),"criteria":ps("Criteria"),
            "clear":pb("Clear"),"sheet":ps("Sheet")},["range"]),
        mt("find_replace","Find/replace",{"find":ps("Find"),"replace":ps("Replace"),"range":ps("Range"),
            "sheet":ps("Sheet")},["find","replace"]),
        mt("conditional_format","Conditional format",{"range":ps("Range"),"rule_type":ps("Rule type"),
            "value1":ps("V1"),"format_color":pn("Color"),"sheet":ps("Sheet")},["range","rule_type"]),
        mt("copy_range","Copy range",{"source":ps("Source"),"destination":ps("Dest"),"dest_sheet":ps("Dest sheet"),
            "values_only":pb("Values only"),"sheet":ps("Sheet")},["source","destination"]),
        mt("rename_sheet","Rename sheet",{"sheet":ps("Sheet"),"new_name":ps("New name")},["new_name"]),
        mt("delete_sheet","Delete sheet",{"sheet":ps("Sheet name")},["sheet"]),
        mt("freeze_panes","Freeze panes",{"cell":ps("Cell"),"unfreeze":pb("Unfreeze"),"sheet":ps("Sheet")},[]),
        mt("remove_duplicates","Remove duplicates",{"range":ps("Range"),"columns":{"type":"array","items":{"type":"number"}},
            "has_headers":pb("Headers"),"sheet":ps("Sheet")},["range"]),
        mt("set_validation","Data validation",{"range":ps("Range"),"type":ps("Type"),"formula1":ps("F1"),
            "sheet":ps("Sheet")},["range","type","formula1"]),
        mt("list_pivot_tables","List PivotTables",{"sheet":ps("Sheet")},[]),
        mt("clear_range","Clear range",{"range":ps("Range"),"what":ps("contents/formats/all"),"sheet":ps("Sheet")},["range"]),
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
9. If a pivot table blocks an operation (e.g. delete_rows, insert_rows), use the move_table tool to move it to a separate sheet first, then perform the operation on the data.
10. Do not repeat operations that have already been completed successfully.
11. Communicate with the user in Polish.
12. Plan your actions efficiently — try to stay within {max_rounds} tool rounds per request and batch related operations when possible.
13. Do not explain what you are doing step by step. Instead, after completing work, write a short summary listing which cells/ranges were changed and what was done."""


# ── API + loop ──────────────────────────────────────────────────────
def call_api(msgs, tls):
    r = requests.post(f"{API_BASE}/chat/completions",
        json={"model":MODEL,"messages":msgs,"max_tokens":4096,"temperature":0.7,"tools":tls,"tool_choice":"auto"},
        headers={"Authorization":f"Bearer {API_KEY}","Content-Type":"application/json"},
        timeout=120)
    if r.status_code != 200:
        print(f"  ❌ HTTP {r.status_code}: {r.text[:300]}")
        return None
    return r.json()


def run(prompt, wb, label, max_rounds=MAX_ROUNDS):
    print(f"\n{'='*60}\n  TEST: {label}\n  User: {prompt}\n{'='*60}")
    msgs = [{"role":"system","content":SYSTEM_PROMPT.format(max_rounds=max_rounds)},
            {"role":"user","content":prompt}]
    prev_sig, reps, total_tc = None, 0, 0
    # Track round_info index for replacement (like the fixed C# code)
    ri_idx = None

    for rnd in range(1, max_rounds+1):
        # Replace round info (not accumulate)
        ri_msg = {"role":"system","content":f"[Round {rnd}/{max_rounds}]"}
        if ri_idx is not None and ri_idx < len(msgs):
            msgs[ri_idx] = ri_msg
        else:
            msgs.append(ri_msg)
            ri_idx = len(msgs) - 1

        if rnd == max_rounds:
            msgs.append({"role":"user","content":f"This is your final response before reaching the {max_rounds} tool-round limit. Write yourself a summary of what you did and what's left."})

        print(f"  📡 R{rnd}/{max_rounds}...", end="", flush=True)
        data = call_api(msgs, tools())
        if not data:
            return {"stop":"api_error","rounds":rnd,"tools":total_tc}

        ch = data["choices"][0]
        msg = ch["message"]

        if msg.get("tool_calls"):
            tcl = msg["tool_calls"]
            print(f" {len(tcl)} tool(s)")
            msgs.append(msg)
            # shift ri_idx since we inserted assistant msg after it
            # actually ri_idx was before this append, but since we replaced at ri_idx,
            # the round info stays at same position. New messages go after.

            sig = "|".join(f"{t['function']['name']}:{t['function']['arguments']}" for t in tcl)
            if sig == prev_sig:
                reps += 1
                if reps >= MAX_SAME_REPEATS:
                    print(f"  🔄 LOOP DETECTED"); return {"stop":"loop","rounds":rnd,"tools":total_tc}
            else: reps = 0
            prev_sig = sig

            for tc in tcl:
                fn = tc["function"]["name"]
                fa = tc["function"]["arguments"]
                tid = tc["id"]
                total_tc += 1
                try: args = json.loads(fa)
                except: args = {}
                result = wb.exec(fn, args)
                rstr = json.dumps(result, ensure_ascii=False)
                short_a = fa if len(fa)<70 else fa[:67]+"..."
                short_r = rstr if len(rstr)<90 else rstr[:87]+"..."
                print(f"    🔧 {fn}({short_a}) → {short_r}")
                msgs.append({"role":"tool","content":rstr,"tool_call_id":tid})
            continue

        content = msg.get("content","")
        if not content:
            print(" ⚠️ EMPTY")
            return {"stop":"empty","rounds":rnd,"tools":total_tc}
        print(f" 💬 done")
        print(f"  {'─'*50}")
        for l in content.split("\n"): print(f"    {l}")
        print(f"  {'─'*50}")
        print(f"  ✅ {rnd} rounds, {total_tc} tool calls")
        return {"stop":"ok","rounds":rnd,"tools":total_tc,"response":content}

    print(f"  ⚠️ MAX ROUNDS")
    return {"stop":"max_rounds","rounds":max_rounds,"tools":total_tc}


# ── MAIN ────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("🧪 Z.AI API Test (simplified)")
    print(f"   Model: {MODEL}, Max rounds: {MAX_ROUNDS}\n")

    # Connectivity check
    print("📡 API check...", end="", flush=True)
    try:
        r = requests.post(f"{API_BASE}/chat/completions",
            json={"model":MODEL,"messages":[{"role":"user","content":"test"}],"max_tokens":5},
            headers={"Authorization":f"Bearer {API_KEY}","Content-Type":"application/json"}, timeout=30)
        if r.status_code == 200: print(" ✅")
        else: print(f" ❌ {r.status_code}"); sys.exit(1)
    except Exception as e: print(f" ❌ {e}"); sys.exit(1)

    results = {}

    # TEST 1: Build small random table + pivot (limit to 10 rows for speed)
    wb1 = VWB()
    results["T1: tabela+pivot"] = run(
        "Stwórz małą tabelę (5 wierszy danych + nagłówek) z kolumnami: Produkt, Kategoria, Ilość, Cena. "
        "Oblicz kolumnę Wartość (Ilość*Cena). Następnie stwórz tabelę przestawną podsumowującą Wartość wg Kategorii.",
        wb1, "Mała tabela + pivot (5 wierszy)", max_rounds=20)

    # TEST 2: Prettify
    if results.get("T1: tabela+pivot",{}).get("stop") == "ok":
        results["T2: ładniejsza"] = run("Sformatuj tabelę — pogrubiony nagłówek, obramowanie, format walutowy dla Cena i Wartość.",
            wb1, "Upiększenie", max_rounds=15)

    # TEST 3: Move pivot table (the key test!)
    wb3 = VWB()
    # Pre-populate data + pivot on same sheet
    for r in range(1, 7):
        wb3.sheets["Arkusz1"][(r, 1)] = ["Produkt","A","B","C","A","B"][r-1]
        wb3.sheets["Arkusz1"][(r, 2)] = ["Wartość",100,200,300,150,250][r-1]
    wb3.pivots["PT1"] = {"source":"'Arkusz1'!A1:B6","dest_sheet":"Arkusz1","dest_cell":"D1",
                         "row_fields":["Produkt"],"value_fields":["Wartość"]}
    results["T3: przenieś pivot"] = run(
        "Na arkuszu Arkusz1 mam dane w A1:B6 i tabelę przestawną PT1 w D1. "
        "Przenieś tabelę przestawną na osobny arkusz i usuń puste wiersze 8-20 z danych.",
        wb3, "Przeniesienie pivot + delete_rows", max_rounds=20)

    # TEST 4: Empty workbook edge case
    wb4 = VWB()
    results["T4: pusty"] = run("Podsumuj dane w tym arkuszu", wb4, "Pusty arkusz", max_rounds=10)

    # Summary
    print(f"\n{'='*60}\n  📊 PODSUMOWANIE\n{'='*60}")
    for name, res in results.items():
        if res:
            st = "✅" if res["stop"]=="ok" else ("⚠️" if res["stop"] in ("max_rounds","loop") else "❌")
            print(f"  {st} {name}: stop={res['stop']}, rounds={res['rounds']}, tools={res['tools']}")
        else:
            print(f"  ❌ {name}: brak wyniku")

    # Workbook states
    print(f"\n  📋 WB po T3 (move pivot):")
    print(f"     Sheets: {list(wb3.sheets.keys())}")
    print(f"     Pivots: {wb3.pivots}")
