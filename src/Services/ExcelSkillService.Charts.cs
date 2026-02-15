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
        if (string.IsNullOrEmpty(destCell))
        {
            var destSheetName = Str(args["dest_sheet"]);
            if (!string.IsNullOrEmpty(destSheetName))
            {
                try { destSheet = wb.Worksheets[destSheetName]; }
                catch
                {
                    destSheet = wb.Worksheets.Add();
                    destSheet.Name = destSheetName;
                }
            }
            else
            {
                destSheet = wb.Worksheets.Add();
            }
            destRange = destSheet.Range["A3"];
        }
        else
        {
            var destSheetName = Str(args["dest_sheet"]);
            if (!string.IsNullOrEmpty(destSheetName))
            {
                try { destSheet = wb.Worksheets[destSheetName]; }
                catch
                {
                    destSheet = wb.Worksheets.Add();
                    destSheet.Name = destSheetName;
                }
            }
            else
            {
                destSheet = ws;
            }
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
                foreach (dynamic cache in wb.PivotCaches())
                {
                    try
                    {
                        string sd = cache.SourceData?.ToString() ?? "";
                        if (!string.IsNullOrEmpty(sd))
                            cacheInfo.Add(sd);
                    }
                    catch { }
                }
                if (cacheInfo.Count > 0)
                    result["pivot_cache_sources"] = cacheInfo;
            }
            catch { }

            result["hint"] = "PivotTable enumeration failed (COM type library issue, common with .xlsb files). " +
                "Do NOT retry list_pivot_tables. To move a pivot, use read_range to find its location, " +
                "then move_table with source_range parameter, or use create_pivot_table + clear_range.";
        }
        return result.ToJsonString();
    }

    // --- move_table ---
    private string SkillMoveTable(JsonNode args)
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

        // ---- TIER 3: Range.PivotTable — bypasses PivotTables() collection ----
        if (foundPivot == null && !string.IsNullOrEmpty(sourceRange))
        {
            try
            {
                dynamic firstCell = sourceSheet.Range[sourceRange].Cells[1, 1];
                foundPivot = firstCell.PivotTable;
                pivotSheet = sourceSheet;
            }
            catch (COMException ex) when (ex.HResult == unchecked((int)0x80028018))
            { comTypeLibError = true; }
            catch { /* cell is not in a pivot table — expected */ }
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

        AddIn.Logger.Debug($"move_table: name={sourceName}, range={sourceRange}, found={foundPivot != null}, comErr={comTypeLibError}");

        // ---- Determine destination sheet ----
        dynamic destSheet;
        if (string.IsNullOrEmpty(destSheetName))
        {
            destSheet = wb.Worksheets.Add();
        }
        else
        {
            try { destSheet = wb.Worksheets[destSheetName]; }
            catch
            {
                destSheet = wb.Worksheets.Add();
                destSheet.Name = destSheetName;
            }
        }

        dynamic destRng = destSheet.Range[destCell];

        // ---- 4a. Move pivot using Location property ----
        if (foundPivot != null)
        {
            string fromSheet = (string)(pivotSheet?.Name ?? "?");
            string ptName = (string)foundPivot.Name;
            string destSheetActual = (string)destSheet.Name;
            string qualifiedDest = $"'{destSheetActual}'!{destCell}";

            try
            {
                foundPivot.Location = qualifiedDest;
                return new JsonObject
                {
                    ["success"] = true,
                    ["moved"] = "pivot_table",
                    ["name"] = ptName,
                    ["from_sheet"] = fromSheet,
                    ["to_sheet"] = destSheetActual,
                    ["dest_cell"] = destCell
                }.ToJsonString();
            }
            catch (Exception ex)
            {
                AddIn.Logger.Debug($"move_table Location failed: 0x{ex.HResult:X} {ex.Message}");
                // Fall through to PivotCaches recreation
            }
        }

        // ---- 4b. PivotCaches recreation (when pivot exists but can't be moved) ----
        if (foundPivot != null || comTypeLibError)
        {
            try
            {
                foreach (dynamic cache in wb.PivotCaches())
                {
                    try
                    {
                        string sd = cache.SourceData?.ToString() ?? "";
                        if (string.IsNullOrEmpty(sd)) continue;

                        string newName = (!string.IsNullOrEmpty(sourceName) ? sourceName : "PivotTable") + "_moved";
                        dynamic newPt = cache.CreatePivotTable(
                            TableDestination: destRng,
                            TableName: newName);

                        string clearNote = "";
                        if (!string.IsNullOrEmpty(sourceRange))
                        {
                            try { sourceSheet.Range[sourceRange].Clear(); clearNote = " Old range cleared."; }
                            catch { clearNote = " Could not clear old range — clear it manually with clear_range."; }
                        }

                        AddIn.Logger.Debug($"move_table: recreated pivot from cache, source={sd}");

                        return new JsonObject
                        {
                            ["success"] = true,
                            ["moved"] = "pivot_recreated",
                            ["new_name"] = newName,
                            ["source_data"] = sd,
                            ["to_sheet"] = (string)destSheet.Name,
                            ["dest_cell"] = destCell,
                            ["warning"] = $"Pivot recreated from cache — field layout (row_fields, value_fields, column_fields) needs reconfiguration via create_pivot_table or manual setup.{clearNote}"
                        }.ToJsonString();
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                AddIn.Logger.Debug($"move_table PivotCaches failed: 0x{ex.HResult:X}");
            }

            // PivotCaches also failed — return actionable error
            string cacheSourceHint = "";
            try
            {
                foreach (dynamic c in wb.PivotCaches())
                {
                    try { cacheSourceHint = c.SourceData?.ToString() ?? ""; break; } catch { }
                }
            }
            catch { }

            return JsonSerializer.Serialize(new
            {
                error = "Cannot move pivot table — COM type library error on this workbook (.xlsb issue).",
                alternative_steps = new[]
                {
                    $"1. create_pivot_table on dest_sheet='{destSheetName ?? "new sheet"}' with same source data{(string.IsNullOrEmpty(cacheSourceHint) ? "" : $" (source: {cacheSourceHint})")} and field configuration",
                    $"2. clear_range on sheet='{sourceSheet.Name}' to remove old pivot"
                },
                hint = "Do NOT retry move_table — use create_pivot_table + clear_range instead."
            });
        }

        // ---- 4c. Move regular data range ----
        if (!string.IsNullOrEmpty(sourceRange) && !comTypeLibError)
        {
            dynamic srcRange = sourceSheet.Range[sourceRange];
            srcRange.Copy(destRng);
            srcRange.Clear();

            return new JsonObject
            {
                ["success"] = true,
                ["moved"] = "data_range",
                ["source_range"] = sourceRange,
                ["from_sheet"] = (string)sourceSheet.Name,
                ["to_sheet"] = (string)destSheet.Name,
                ["dest_cell"] = destCell
            }.ToJsonString();
        }

        // ---- 4d. source_range on problematic workbook — don't blindly copy, might be pivot ----
        if (!string.IsNullOrEmpty(sourceRange) && comTypeLibError)
        {
            return JsonSerializer.Serialize(new
            {
                error = "Cannot determine if source_range contains a pivot table (COM type library error).",
                hint = "To move data: use copy_range + clear_range. To move a pivot: use create_pivot_table on the destination sheet + clear_range on the source."
            });
        }

        // ---- Final error ----
        if (!string.IsNullOrEmpty(sourceName))
            return JsonSerializer.Serialize(new
            {
                error = $"Pivot table '{sourceName}' not found.",
                hint = "Try using source_range parameter pointing to the pivot's cell range instead."
            });

        return JsonSerializer.Serialize(new { error = "Provide 'name' (pivot table name) or 'source_range' to move." });
    }
}
