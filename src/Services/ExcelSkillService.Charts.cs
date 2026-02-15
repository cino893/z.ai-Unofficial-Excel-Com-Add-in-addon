using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.Json.Nodes;
using ExcelDna.Integration;

namespace ZaiExcelAddin.Services;

public partial class ExcelSkillService
{
    // --- create_chart ---
    private string SkillCreateChart(JsonNode args)
    {
        dynamic ws = GetTargetSheet(args);
        string dataRange = Str(args["data_range"]);
        string chartTypeStr = Str(args["chart_type"], "column").ToLowerInvariant();
        string title = Str(args["title"]);

        dynamic srcRange = ws.Range[dataRange];
        dynamic chartObj = ws.ChartObjects.Add(
            srcRange.Left + srcRange.Width + 20,
            srcRange.Top,
            400, 300);

        dynamic chart = chartObj.Chart;
        chart.SetSourceData(srcRange);

        chart.ChartType = chartTypeStr switch
        {
            "bar" => 57,        // xlBarClustered
            "line" => 4,        // xlLine
            "pie" => 5,         // xlPie
            "scatter" => -4169, // xlXYScatter
            "area" => 1,        // xlArea
            _ => 51             // xlColumnClustered
        };

        if (!string.IsNullOrEmpty(title))
        {
            chart.HasTitle = true;
            chart.ChartTitle.Text = title;
        }

        var result = new JsonObject
        {
            ["success"] = true,
            ["chart_name"] = (string)chartObj.Name,
            ["type"] = chartTypeStr
        };
        return result.ToJsonString();
    }

    // --- delete_chart ---
    private string SkillDeleteChart(JsonNode args)
    {
        dynamic ws = GetTargetSheet(args);
        string chartName = Str(args["chart_name"]);

        int count = ws.ChartObjects.Count;
        for (int i = count; i >= 1; i--)
        {
            if ((string)ws.ChartObjects[i].Name == chartName)
            {
                ws.ChartObjects[i].Delete();
                return new JsonObject
                {
                    ["success"] = true,
                    ["deleted"] = chartName
                }.ToJsonString();
            }
        }

        return JsonSerializer.Serialize(new { error = $"Chart not found: {chartName}" });
    }

    // --- list_charts ---
    private string SkillListCharts(JsonNode args)
    {
        dynamic ws = GetTargetSheet(args);
        int count = ws.ChartObjects.Count;

        var charts = new JsonArray();
        for (int i = 1; i <= count; i++)
        {
            dynamic co = ws.ChartObjects[i];
            string typeName = "unknown";
            string source = "";

            try
            {
                int ct = co.Chart.ChartType;
                typeName = ct switch
                {
                    51 => "column",
                    57 => "bar",
                    4 => "line",
                    5 => "pie",
                    -4169 => "scatter",
                    1 => "area",
                    _ => $"other({ct})"
                };
            }
            catch { /* ignore */ }

            try
            {
                if (co.Chart.SeriesCollection.Count > 0)
                    source = (string)co.Chart.SeriesCollection[1].Formula;
            }
            catch { /* ignore */ }

            charts.Add(new JsonObject
            {
                ["name"] = (string)co.Name,
                ["type"] = typeName,
                ["source"] = source
            });
        }

        var result = new JsonObject
        {
            ["sheet"] = (string)ws.Name,
            ["count"] = count,
            ["charts"] = charts
        };
        return result.ToJsonString();
    }

    // --- create_pivot_table ---
    private string SkillCreatePivotTable(JsonNode args)
    {
        dynamic app = GetApp();
        dynamic ws = GetTargetSheet(args);
        dynamic wb = app.ActiveWorkbook;
        string sourceRange = Str(args["source_range"]);
        string destCell = Str(args["dest_cell"]);
        string name = Str(args["name"]);

        if (string.IsNullOrEmpty(name))
            name = "PivotTable" + (wb.PivotCaches().Count + 1);

        dynamic srcRange = ws.Range[sourceRange];
        string qualifiedSource = $"'{ws.Name}'!{srcRange.Address}";

        // xlDatabase = 1
        dynamic cache = wb.PivotCaches().Create(SourceType: 1, SourceData: qualifiedSource);

        dynamic destSheet;
        dynamic destRange;
        string destSheetName = Str(args["dest_sheet"]);

        if (!string.IsNullOrEmpty(destSheetName) || string.IsNullOrEmpty(destCell))
        {
            destSheet = ResolveDestSheet(wb, string.IsNullOrEmpty(destSheetName) ? null : destSheetName);
            destRange = destSheet.Range[string.IsNullOrEmpty(destCell) ? "A3" : destCell];
        }
        else
        {
            destSheet = ws;
            destRange = destSheet.Range[destCell];
        }

        dynamic pt = cache.CreatePivotTable(
            TableDestination: destRange,
            TableName: name);

        // Add row fields
        var rowFields = args["row_fields"]?.AsArray();
        if (rowFields != null)
        {
            foreach (var f in rowFields)
            {
                dynamic pf = pt.PivotFields(f!.GetValue<string>());
                pf.Orientation = 1; // xlRowField
            }
        }

        // Add column fields
        var colFields = args["column_fields"]?.AsArray();
        if (colFields != null)
        {
            foreach (var f in colFields)
            {
                dynamic pf = pt.PivotFields(f!.GetValue<string>());
                pf.Orientation = 2; // xlColumnField
            }
        }

        // Add value fields
        var valFields = args["value_fields"]?.AsArray();
        string funcStr = Str(args["value_function"], "sum").ToLowerInvariant();
        int xlFunc = funcStr switch
        {
            "count" => -4112,    // xlCount
            "average" => -4106,  // xlAverage
            "max" => -4136,      // xlMax
            "min" => -4139,      // xlMin
            _ => -4157           // xlSum
        };

        if (valFields != null)
        {
            foreach (var f in valFields)
            {
                dynamic pf = pt.PivotFields(f!.GetValue<string>());
                pf.Orientation = 4; // xlDataField
                pf.Function = xlFunc;
            }
        }

        var result = new JsonObject
        {
            ["success"] = true,
            ["name"] = name,
            ["dest_sheet"] = (string)destSheet.Name,
            ["row_fields"] = rowFields?.Count ?? 0,
            ["value_fields"] = valFields?.Count ?? 0
        };
        return result.ToJsonString();
    }

    // --- list_pivot_tables ---
    private string SkillListPivotTables(JsonNode args)
    {
        dynamic app = GetApp();
        dynamic wb = app.ActiveWorkbook;
        var sheetFilter = args["sheet"]?.GetValue<string>();

        var pivots = new JsonArray();
        var errors = new JsonArray();
        foreach (dynamic ws in wb.Worksheets)
        {
            string wsName = ws.Name;
            if (sheetFilter != null && !wsName.Equals(sheetFilter, StringComparison.OrdinalIgnoreCase))
                continue;

            try
            {
                foreach (dynamic pt in ws.PivotTables())
                {
                    string location = "";
                    try { location = pt.TableRange2.Address; } catch { }

                    string source = "";
                    try { source = pt.SourceData?.ToString() ?? ""; } catch { }

                    pivots.Add(new JsonObject
                    {
                        ["name"] = (string)pt.Name,
                        ["sheet"] = wsName,
                        ["location"] = location,
                        ["source"] = source
                    });
                }
            }
            catch
            {
                // PivotTables() can throw TYPE_E_INVDATAREAD on some workbooks
                errors.Add(wsName);
            }
        }

        var result = new JsonObject
        {
            ["pivot_tables"] = pivots,
            ["count"] = pivots.Count
        };
        if (errors.Count > 0)
        {
            result["errors"] = errors;

            // Try PivotCaches to provide source data info even when enumeration fails
            try
            {
                var cacheInfo = new JsonArray();
                foreach (var src in GetPivotCacheSourcesA1(wb))
                    cacheInfo.Add(src);
                if (cacheInfo.Count > 0)
                    result["pivot_cache_sources"] = cacheInfo;
            }
            catch { }

            // Build a clear hint with the actual data source range
            var srcList = result["pivot_cache_sources"]?.AsArray();
            string srcInfo = srcList != null && srcList.Count > 0
                ? $" The pivot data source is: {srcList[0]}."
                : "";
            result["hint"] = "PivotTable enumeration failed (COM type library issue, common with .xlsb files). " +
                "Do NOT retry list_pivot_tables or move_pivot_table." + srcInfo +
                " To recreate a pivot on another sheet: call create_pivot_table with the source data range, then modify_pivot_table to configure fields.";
        }
        return result.ToJsonString();
    }

    // --- move_pivot_table ---
    private string SkillMovePivotTable(JsonNode args)
    {
        dynamic app = GetApp();
        dynamic wb = app.ActiveWorkbook;
        string sourceName = Str(args["name"]);
        string sourceRange = Str(args["source_range"]);
        string destSheetName = Str(args["dest_sheet"]);
        string destCell = Str(args["dest_cell"], "A1");
        dynamic sourceSheet = GetTargetSheet(args);

        dynamic? foundPivot = null;
        dynamic? pivotSheet = null;
        bool comTypeLibError = false;

        // ---- TIER 1: PivotTables() enumeration ----
        if (!string.IsNullOrEmpty(sourceName))
        {
            foreach (dynamic sheet in wb.Worksheets)
            {
                try
                {
                    foreach (dynamic pt in sheet.PivotTables())
                    {
                        if ((string)pt.Name == sourceName)
                        { foundPivot = pt; pivotSheet = sheet; break; }
                    }
                    if (foundPivot != null) break;
                }
                catch (COMException ex) when (ex.HResult == unchecked((int)0x80028018))
                { comTypeLibError = true; }
                catch { }
            }
        }

        // ---- TIER 2: Direct name lookup ----
        if (foundPivot == null && !string.IsNullOrEmpty(sourceName))
        {
            foreach (dynamic sheet in wb.Worksheets)
            {
                try
                {
                    dynamic pt = sheet.PivotTables(sourceName);
                    if (pt != null) { foundPivot = pt; pivotSheet = sheet; break; }
                }
                catch (COMException ex) when (ex.HResult == unchecked((int)0x80028018))
                { comTypeLibError = true; }
                catch { }
            }
        }

        // ---- TIER 3: Range.PivotTable ----
        if (foundPivot == null && !string.IsNullOrEmpty(sourceRange))
        {
            try
            {
                dynamic firstCell = sourceSheet.Range[sourceRange].Cells[1, 1];
                foundPivot = firstCell.PivotTable;
                pivotSheet = sourceSheet;
            }
            catch (COMException) { comTypeLibError = true; }
            catch { }
        }

        // ---- TIER 3b: Intersection check (normal workbooks only) ----
        if (foundPivot == null && !string.IsNullOrEmpty(sourceRange) && !comTypeLibError)
        {
            try
            {
                dynamic srcRange = sourceSheet.Range[sourceRange];
                foreach (dynamic pt in sourceSheet.PivotTables())
                {
                    dynamic ptRange = pt.TableRange2;
                    dynamic overlap = app.Intersect(srcRange, ptRange);
                    if (overlap != null)
                    { foundPivot = pt; pivotSheet = sourceSheet; break; }
                }
            }
            catch { }
        }

        // ---- Safety net: PivotCaches exist but pivot wasn't found → COM issue ----
        if (foundPivot == null && !comTypeLibError)
        {
            try { if (wb.PivotCaches().Count > 0) comTypeLibError = true; }
            catch { }
        }

        AddIn.Logger.Debug($"move_pivot_table: name={sourceName}, range={sourceRange}, found={foundPivot != null}, comErr={comTypeLibError}");

        // ---- No pivot detected at all ----
        if (foundPivot == null && !comTypeLibError)
        {
            return JsonSerializer.Serialize(new
            {
                error = "No pivot table found. Check the name or source_range.",
                hint = "Use list_pivot_tables to find pivot tables, or use copy_range to move plain data."
            });
        }

        // ---- Resolve destination sheet ----
        dynamic destSheet = ResolveDestSheet(wb, destSheetName);
        dynamic destRng = destSheet.Range[destCell];

        // ---- A. Move using Location property (normal workbooks) ----
        if (foundPivot != null)
        {
            string fromSheet = (string)(pivotSheet?.Name ?? "?");
            string ptName = (string)foundPivot.Name;
            string destSheetActual = (string)destSheet.Name;
            try
            {
                foundPivot.Location = $"'{destSheetActual}'!{destCell}";
                return new JsonObject
                {
                    ["success"] = true,
                    ["method"] = "location_move",
                    ["name"] = ptName,
                    ["from_sheet"] = fromSheet,
                    ["to_sheet"] = destSheetActual,
                    ["dest_cell"] = destCell
                }.ToJsonString();
            }
            catch (Exception ex)
            {
                AddIn.Logger.Debug($"move_pivot_table Location failed: 0x{ex.HResult:X} {ex.Message}");
            }
        }

        // ---- B. Recreate from PivotCaches (.xlsb fallback) ----
        try
        {
            foreach (dynamic cache in wb.PivotCaches())
            {
                try
                {
                    string sd = cache.SourceData?.ToString() ?? "";
                    if (string.IsNullOrEmpty(sd)) continue;

                    string newName = (!string.IsNullOrEmpty(sourceName) ? sourceName : "PivotTable") + "_moved";
                    cache.CreatePivotTable(TableDestination: destRng, TableName: newName);

                    string clearNote = "";
                    if (!string.IsNullOrEmpty(sourceRange))
                    {
                        try { sourceSheet.Range[sourceRange].Clear(); clearNote = " Old range cleared."; }
                        catch { clearNote = " Clear old range manually with clear_range."; }
                    }

                    return new JsonObject
                    {
                        ["success"] = true,
                        ["method"] = "pivot_recreated",
                        ["new_name"] = newName,
                        ["source_data"] = ConvertR1C1toA1(sd),
                        ["to_sheet"] = (string)destSheet.Name,
                        ["dest_cell"] = destCell,
                        ["next_step"] = $"Pivot recreated empty — use modify_pivot_table to configure row_fields, value_fields, column_fields.{clearNote}"
                    }.ToJsonString();
                }
                catch { }
            }
        }
        catch (Exception ex)
        {
            AddIn.Logger.Debug($"move_pivot_table PivotCaches failed: 0x{ex.HResult:X}");
        }

        // ---- C. All methods failed — guide model to create_pivot_table ----
        var cacheSources = GetPivotCacheSourcesA1(wb);
        string srcHint = cacheSources.Count > 0 ? $" (source_range: {cacheSources[0]})" : "";
        return JsonSerializer.Serialize(new
        {
            error = "Cannot move pivot table — COM access denied on this workbook.",
            alternative = $"Use create_pivot_table with source data{srcHint} on dest_sheet='{destSheetName ?? "new sheet"}', then clear_range on old pivot area."
        });
    }

    // --- modify_pivot_table ---
    private string SkillModifyPivotTable(JsonNode args)
    {
        dynamic app = GetApp();
        dynamic wb = app.ActiveWorkbook;
        string name = Str(args["name"]);
        dynamic ws = GetTargetSheet(args);

        dynamic? pt = null;

        // Find the pivot table
        if (!string.IsNullOrEmpty(name))
        {
            // Try target sheet first, then all sheets
            try { pt = ws.PivotTables(name); } catch { }
            if (pt == null)
            {
                foreach (dynamic sheet in wb.Worksheets)
                {
                    try { pt = sheet.PivotTables(name); break; } catch { }
                }
            }
        }
        else
        {
            // No name → use first pivot on target sheet
            try
            {
                if (ws.PivotTables().Count > 0) pt = ws.PivotTables(1);
            }
            catch { }
        }

        if (pt == null)
            return JsonSerializer.Serialize(new { error = $"Pivot table '{name}' not found. Use list_pivot_tables to find available pivots." });

        int changes = 0;

        // Clear existing field orientations if we're setting new ones
        bool hasNewFields = args["row_fields"] != null || args["column_fields"] != null ||
                           args["value_fields"] != null || args["page_fields"] != null;

        if (hasNewFields && Bool(args["clear_fields"], true))
        {
            try
            {
                // Reset all fields to hidden (orientation = 0)
                foreach (dynamic pf in pt.PivotFields())
                {
                    try
                    {
                        int orient = (int)pf.Orientation;
                        if (orient != 0) // not xlHidden
                            pf.Orientation = 0; // xlHidden
                    }
                    catch { }
                }
            }
            catch { }
        }

        // Set row fields
        var rowFields = args["row_fields"]?.AsArray();
        if (rowFields != null)
        {
            foreach (var f in rowFields)
            {
                try
                {
                    dynamic pf = pt.PivotFields(f!.GetValue<string>());
                    pf.Orientation = 1; // xlRowField
                    changes++;
                }
                catch (Exception ex)
                {
                    AddIn.Logger.Debug($"modify_pivot: row field '{f}' failed: {ex.Message}");
                }
            }
        }

        // Set column fields
        var colFields = args["column_fields"]?.AsArray();
        if (colFields != null)
        {
            foreach (var f in colFields)
            {
                try
                {
                    dynamic pf = pt.PivotFields(f!.GetValue<string>());
                    pf.Orientation = 2; // xlColumnField
                    changes++;
                }
                catch (Exception ex)
                {
                    AddIn.Logger.Debug($"modify_pivot: col field '{f}' failed: {ex.Message}");
                }
            }
        }

        // Set value fields
        var valFields = args["value_fields"]?.AsArray();
        if (valFields != null)
        {
            string funcStr = Str(args["value_function"], "sum").ToLowerInvariant();
            int xlFunc = funcStr switch
            {
                "count" => -4112,
                "average" => -4106,
                "max" => -4136,
                "min" => -4139,
                _ => -4157 // xlSum
            };

            foreach (var f in valFields)
            {
                try
                {
                    dynamic pf = pt.PivotFields(f!.GetValue<string>());
                    pf.Orientation = 4; // xlDataField
                    pf.Function = xlFunc;
                    changes++;
                }
                catch (Exception ex)
                {
                    AddIn.Logger.Debug($"modify_pivot: value field '{f}' failed: {ex.Message}");
                }
            }
        }

        // Set page/filter fields
        var pageFields = args["page_fields"]?.AsArray();
        if (pageFields != null)
        {
            foreach (var f in pageFields)
            {
                try
                {
                    dynamic pf = pt.PivotFields(f!.GetValue<string>());
                    pf.Orientation = 3; // xlPageField
                    changes++;
                }
                catch (Exception ex)
                {
                    AddIn.Logger.Debug($"modify_pivot: page field '{f}' failed: {ex.Message}");
                }
            }
        }

        // Refresh
        if (Bool(args["refresh"]))
        {
            try { pt.RefreshTable(); } catch { }
        }

        // Get current field layout for confirmation
        var fieldInfo = new JsonObject();
        try
        {
            var rows = new JsonArray();
            var cols = new JsonArray();
            var vals = new JsonArray();
            var pages = new JsonArray();
            foreach (dynamic pf in pt.PivotFields())
            {
                try
                {
                    string fn = (string)pf.Name;
                    int o = (int)pf.Orientation;
                    if (o == 1) rows.Add(fn);
                    else if (o == 2) cols.Add(fn);
                    else if (o == 4) vals.Add(fn);
                    else if (o == 3) pages.Add(fn);
                }
                catch { }
            }
            fieldInfo["row_fields"] = rows;
            fieldInfo["column_fields"] = cols;
            fieldInfo["value_fields"] = vals;
            if (pages.Count > 0) fieldInfo["page_fields"] = pages;
        }
        catch { }

        var result = new JsonObject
        {
            ["success"] = true,
            ["name"] = (string)pt.Name,
            ["changes"] = changes,
            ["current_layout"] = fieldInfo
        };
        return result.ToJsonString();
    }

    /// <summary>Convert R1C1 reference (Polish W/K or English R/C) to A1 format. E.g. "Arkusz1!W1K1:W26K5" → "Arkusz1!A1:E26"</summary>
    private static string ConvertR1C1toA1(string r1c1)
    {
        // Split on '!' to handle sheet names with spaces (may be quoted)
        string sheet = "";
        string refPart = r1c1;
        int bangIdx = r1c1.LastIndexOf('!');
        if (bangIdx >= 0)
        {
            sheet = r1c1[..(bangIdx + 1)]; // includes '!'
            refPart = r1c1[(bangIdx + 1)..];
        }

        var match = System.Text.RegularExpressions.Regex.Match(refPart,
            @"^[WRwr](\d+)[KCkc](\d+)(?::?[WRwr](\d+)[KCkc](\d+))?$");
        if (!match.Success) return r1c1;

        int r1 = int.Parse(match.Groups[1].Value);
        int c1 = int.Parse(match.Groups[2].Value);
        string result = $"{sheet}{ColToLetter(c1)}{r1}";
        if (match.Groups[3].Success)
        {
            int r2 = int.Parse(match.Groups[3].Value);
            int c2 = int.Parse(match.Groups[4].Value);
            result += $":{ColToLetter(c2)}{r2}";
        }
        return result;
    }

    private static string ColToLetter(int col)
    {
        string result = "";
        while (col > 0)
        {
            col--;
            result = (char)('A' + col % 26) + result;
            col /= 26;
        }
        return result;
    }

    /// <summary>Get all PivotCache source data ranges in A1 format.</summary>
    private static List<string> GetPivotCacheSourcesA1(dynamic wb)
    {
        var sources = new List<string>();
        try
        {
            foreach (dynamic cache in wb.PivotCaches())
            {
                try
                {
                    string sd = cache.SourceData?.ToString() ?? "";
                    if (!string.IsNullOrEmpty(sd))
                        sources.Add(ConvertR1C1toA1(sd));
                }
                catch { }
            }
        }
        catch { }
        return sources;
    }

    /// <summary>Get or create a destination worksheet by name.</summary>
    private static dynamic ResolveDestSheet(dynamic wb, string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            return wb.Worksheets.Add();
        try { return wb.Worksheets[sheetName]; }
        catch
        {
            dynamic ws = wb.Worksheets.Add();
            ws.Name = sheetName;
            return ws;
        }
    }
}
