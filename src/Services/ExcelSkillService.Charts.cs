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
            destSheet = wb.Worksheets.Add();
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
        foreach (dynamic ws in wb.Worksheets)
        {
            string wsName = ws.Name;
            if (sheetFilter != null && !wsName.Equals(sheetFilter, StringComparison.OrdinalIgnoreCase))
                continue;

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

        return new JsonObject
        {
            ["pivot_tables"] = pivots,
            ["count"] = pivots.Count
        }.ToJsonString();
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

        // 1. Try to find a PivotTable by name
        dynamic? foundPivot = null;
        dynamic? pivotSheet = null;

        if (!string.IsNullOrEmpty(sourceName))
        {
            foreach (dynamic sheet in wb.Worksheets)
            {
                try
                {
                    foreach (dynamic pt in sheet.PivotTables())
                    {
                        if ((string)pt.Name == sourceName)
                        {
                            foundPivot = pt;
                            pivotSheet = sheet;
                            break;
                        }
                    }
                    if (foundPivot != null) break;
                }
                catch { /* sheet may not have pivot tables */ }
            }
        }

        // 2. Also check if source_range overlaps a pivot table
        if (foundPivot == null && !string.IsNullOrEmpty(sourceRange))
        {
            try
            {
                dynamic srcRange = sourceSheet.Range[sourceRange];
                foreach (dynamic pt in sourceSheet.PivotTables())
                {
                    dynamic ptRange = pt.TableRange2;
                    dynamic overlap = app.Intersect(srcRange, ptRange);
                    if (overlap != null)
                    {
                        foundPivot = pt;
                        pivotSheet = sourceSheet;
                        break;
                    }
                }
            }
            catch { /* no pivot tables on sheet */ }
        }

        // Determine destination sheet
        dynamic destSheet;
        if (string.IsNullOrEmpty(destSheetName))
        {
            destSheet = wb.Worksheets.Add();
        }
        else
        {
            try
            {
                destSheet = wb.Worksheets[destSheetName];
            }
            catch
            {
                destSheet = wb.Worksheets.Add();
                destSheet.Name = destSheetName;
            }
        }

        dynamic destRange = destSheet.Range[destCell];

        // 3a. Move pivot table using Location property (string form for cross-sheet)
        if (foundPivot != null)
        {
            string fromSheet = (string)(pivotSheet?.Name ?? "?");
            string ptName = (string)foundPivot.Name;

            // Location requires qualified range string: 'Sheet Name'!A1
            string destSheetActual = (string)destSheet.Name;
            string qualifiedDest = $"'{destSheetActual}'!{destCell}";
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

        // 3b. Move regular data range using Copy + clear source
        if (!string.IsNullOrEmpty(sourceRange))
        {
            dynamic srcRange = sourceSheet.Range[sourceRange];
            srcRange.Copy(destRange);
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

        return JsonSerializer.Serialize(new { error = "Provide either 'name' (pivot table name) or 'source_range' to move." });
    }
}
