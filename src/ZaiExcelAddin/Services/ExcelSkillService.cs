using System.Text.Json;
using System.Text.Json.Nodes;
using ExcelDna.Integration;

namespace ZaiExcelAddin.Services;

public class ExcelSkillService
{
    // ======================== TOOL DEFINITIONS ========================

    public JsonArray GetToolDefinitions()
    {
        return new JsonArray
        {
            MakeTool("read_cell",
                "Read value, formula and type from a single cell",
                new JsonObject
                {
                    ["cell"] = PropString("Cell address e.g. A1, B2, C10"),
                    ["sheet"] = PropString("Sheet name (optional, defaults to active sheet)")
                },
                new JsonArray { "cell" }),

            MakeTool("write_cell",
                "Write a value to a single cell",
                new JsonObject
                {
                    ["cell"] = PropString("Cell address e.g. A1"),
                    ["value"] = PropString("Value to write"),
                    ["sheet"] = PropString("Sheet name (optional)")
                },
                new JsonArray { "cell", "value" }),

            MakeTool("read_range",
                "Read all values from a range of cells. Returns 2D array of values.",
                new JsonObject
                {
                    ["range"] = PropString("Range address e.g. A1:D10"),
                    ["sheet"] = PropString("Sheet name (optional)")
                },
                new JsonArray { "range" }),

            MakeTool("write_range",
                "Write a 2D array of values starting from a cell",
                new JsonObject
                {
                    ["start_cell"] = PropString("Top-left cell to start writing e.g. A1"),
                    ["data"] = new JsonObject
                    {
                        ["type"] = "array",
                        ["description"] = "2D array of values, e.g. [[1,2],[3,4]]",
                        ["items"] = new JsonObject
                        {
                            ["type"] = "array",
                            ["items"] = new JsonObject { ["type"] = "string" }
                        }
                    },
                    ["sheet"] = PropString("Sheet name (optional)")
                },
                new JsonArray { "start_cell", "data" }),

            MakeTool("get_sheet_info",
                "Get information about a worksheet: name, used range, dimensions, headers (first row)",
                new JsonObject
                {
                    ["sheet"] = PropString("Sheet name (optional, defaults to active sheet)")
                },
                new JsonArray()),

            MakeTool("get_workbook_info",
                "Get workbook information: file name, path, list of all sheet names, active sheet",
                new JsonObject(),
                new JsonArray()),

            MakeTool("format_range",
                "Format cells: bold, italic, font size/color, background color, number format, alignment, borders, column width, row height, autofit, merge",
                new JsonObject
                {
                    ["range"] = PropString("Range to format e.g. A1:D1"),
                    ["bold"] = PropBool("Set bold"),
                    ["italic"] = PropBool("Set italic"),
                    ["font_size"] = PropNumber("Font size in points"),
                    ["font_color"] = PropNumber("Font color as RGB long (e.g. 255 for red, 65280 for green, 16711680 for blue)"),
                    ["bg_color"] = PropNumber("Background color as RGB long"),
                    ["number_format"] = PropString("Number format string e.g. #,##0.00 or 0% or yyyy-mm-dd"),
                    ["h_align"] = PropString("Horizontal alignment: left, center, right"),
                    ["wrap_text"] = PropBool("Enable text wrapping"),
                    ["borders"] = PropBool("Add/remove thin borders"),
                    ["column_width"] = PropNumber("Set column width"),
                    ["row_height"] = PropNumber("Set row height"),
                    ["autofit"] = PropBool("Auto-fit column width"),
                    ["merge"] = PropBool("Merge/unmerge cells"),
                    ["sheet"] = PropString("Sheet name (optional)")
                },
                new JsonArray { "range" }),

            MakeTool("insert_formula",
                "Insert an Excel formula into a cell. Use English function names (SUM, AVERAGE, VLOOKUP, IF, COUNT, etc.)",
                new JsonObject
                {
                    ["cell"] = PropString("Target cell e.g. B10"),
                    ["formula"] = PropString("Formula e.g. =SUM(B1:B9) or =IF(A1>0,A1*2,0)"),
                    ["sheet"] = PropString("Sheet name (optional)")
                },
                new JsonArray { "cell", "formula" }),

            MakeTool("sort_range",
                "Sort a range of cells by a specified column",
                new JsonObject
                {
                    ["range"] = PropString("Range to sort e.g. A1:D20"),
                    ["sort_column"] = PropString("Column letter to sort by e.g. B"),
                    ["order"] = PropString("Sort order: asc or desc (default: asc)"),
                    ["has_headers"] = PropBool("Whether first row contains headers (default: true)"),
                    ["sheet"] = PropString("Sheet name (optional)")
                },
                new JsonArray { "range", "sort_column" }),

            MakeTool("add_sheet",
                "Add a new worksheet to the workbook",
                new JsonObject
                {
                    ["name"] = PropString("Name for the new sheet (optional)")
                },
                new JsonArray()),

            MakeTool("delete_rows",
                "Delete one or more rows from the worksheet",
                new JsonObject
                {
                    ["start_row"] = PropNumber("First row number to delete"),
                    ["count"] = PropNumber("Number of rows to delete (default: 1)"),
                    ["sheet"] = PropString("Sheet name (optional)")
                },
                new JsonArray { "start_row" }),

            MakeTool("insert_rows",
                "Insert blank rows at a specified position",
                new JsonObject
                {
                    ["at_row"] = PropNumber("Row number where to insert"),
                    ["count"] = PropNumber("Number of rows to insert (default: 1)"),
                    ["sheet"] = PropString("Sheet name (optional)")
                },
                new JsonArray { "at_row" }),

            MakeTool("create_chart",
                "Create a chart from data range. Use list_charts first to check existing charts, and delete_chart to remove old ones before creating new.",
                new JsonObject
                {
                    ["data_range"] = PropString("Data range for the chart e.g. A1:B10"),
                    ["chart_type"] = PropString("Chart type: column, bar, line, pie, scatter, area (default: column)"),
                    ["title"] = PropString("Chart title (optional)"),
                    ["sheet"] = PropString("Sheet name (optional)")
                },
                new JsonArray { "data_range" }),

            MakeTool("delete_chart",
                "Delete a chart from the worksheet by name. Use list_charts to find chart names.",
                new JsonObject
                {
                    ["chart_name"] = PropString("Name of the chart to delete e.g. Chart 1"),
                    ["sheet"] = PropString("Sheet name (optional)")
                },
                new JsonArray { "chart_name" }),

            MakeTool("list_charts",
                "List all charts on a worksheet with their names, types and data ranges",
                new JsonObject
                {
                    ["sheet"] = PropString("Sheet name (optional, defaults to active sheet)")
                },
                new JsonArray())
        };
    }

    // ======================== TOOL EXECUTION ========================

    public string Execute(string toolName, string argsJson)
    {
        try
        {
            var args = JsonNode.Parse(argsJson) ?? new JsonObject();

            var result = toolName switch
            {
                "read_cell" => SkillReadCell(args),
                "write_cell" => SkillWriteCell(args),
                "read_range" => SkillReadRange(args),
                "write_range" => SkillWriteRange(args),
                "get_sheet_info" => SkillGetSheetInfo(args),
                "get_workbook_info" => SkillGetWorkbookInfo(args),
                "format_range" => SkillFormatRange(args),
                "insert_formula" => SkillInsertFormula(args),
                "sort_range" => SkillSortRange(args),
                "add_sheet" => SkillAddSheet(args),
                "delete_rows" => SkillDeleteRows(args),
                "insert_rows" => SkillInsertRows(args),
                "create_chart" => SkillCreateChart(args),
                "delete_chart" => SkillDeleteChart(args),
                "list_charts" => SkillListCharts(args),
                _ => JsonSerializer.Serialize(new { error = $"Unknown tool: {toolName}" })
            };

            AddIn.Logger.ToolCall(toolName, argsJson, result);
            return result;
        }
        catch (Exception ex)
        {
            var error = JsonSerializer.Serialize(new { error = ex.Message });
            AddIn.Logger.ToolCall(toolName, argsJson, error);
            return error;
        }
    }

    // ======================== SKILL IMPLEMENTATIONS ========================

    private static dynamic GetApp() => ExcelDnaUtil.Application;

    private static dynamic GetTargetSheet(JsonNode args)
    {
        dynamic app = GetApp();
        var sheetName = args["sheet"]?.GetValue<string>();
        if (!string.IsNullOrEmpty(sheetName))
            return app.ActiveWorkbook.Worksheets[sheetName];
        return app.ActiveSheet;
    }

    private static string Str(JsonNode? node, string fallback = "")
        => node?.GetValue<string>() ?? fallback;

    private static int Int(JsonNode? node, int fallback = 0)
        => node?.GetValue<int>() ?? fallback;

    private static double Dbl(JsonNode? node, double fallback = 0)
        => node?.GetValue<double>() ?? fallback;

    private static bool Bool(JsonNode? node, bool fallback = false)
        => node?.GetValue<bool>() ?? fallback;

    private static string EscapeJson(string s)
        => JsonSerializer.Serialize(s)[1..^1]; // strips surrounding quotes

    // --- read_cell ---
    private string SkillReadCell(JsonNode args)
    {
        dynamic ws = GetTargetSheet(args);
        string cell = Str(args["cell"]);
        dynamic rng = ws.Range[cell];

        object? val = rng.Value;
        var result = new JsonObject
        {
            ["cell"] = cell,
            ["value"] = val?.ToString() ?? "",
            ["formula"] = (string)rng.Formula,
            ["type"] = val?.GetType().Name ?? "Empty",
            ["sheet"] = (string)ws.Name
        };
        return result.ToJsonString();
    }

    // --- write_cell ---
    private string SkillWriteCell(JsonNode args)
    {
        dynamic ws = GetTargetSheet(args);
        string cell = Str(args["cell"]);
        string value = Str(args["value"]);

        ws.Range[cell].Value = value;

        var result = new JsonObject
        {
            ["success"] = true,
            ["cell"] = cell,
            ["value"] = value
        };
        return result.ToJsonString();
    }

    // --- read_range ---
    private string SkillReadRange(JsonNode args)
    {
        dynamic ws = GetTargetSheet(args);
        string rangeAddr = Str(args["range"]);
        dynamic rng = ws.Range[rangeAddr];

        int rows = rng.Rows.Count;
        int cols = rng.Columns.Count;

        var data = new JsonArray();
        for (int r = 1; r <= rows; r++)
        {
            var row = new JsonArray();
            for (int c = 1; c <= cols; c++)
            {
                object? cellVal = rng.Cells[r, c].Value;
                if (cellVal == null)
                    row.Add(null);
                else if (cellVal is double d)
                    row.Add(d);
                else
                    row.Add(cellVal.ToString());
            }
            data.Add(row);
        }

        var result = new JsonObject
        {
            ["range"] = rangeAddr,
            ["sheet"] = (string)ws.Name,
            ["rows"] = rows,
            ["cols"] = cols,
            ["data"] = data
        };
        return result.ToJsonString();
    }

    // --- write_range ---
    private string SkillWriteRange(JsonNode args)
    {
        dynamic ws = GetTargetSheet(args);
        string startCell = Str(args["start_cell"]);
        var data = args["data"]!.AsArray();

        dynamic startRange = ws.Range[startCell];
        int rowsWritten = 0;

        for (int r = 0; r < data.Count; r++)
        {
            var rowData = data[r]!.AsArray();
            for (int c = 0; c < rowData.Count; c++)
            {
                var val = rowData[c];
                startRange.Offset[r, c].Value = val?.GetValue<string>() ?? "";
            }
            rowsWritten++;
        }

        var result = new JsonObject
        {
            ["success"] = true,
            ["start_cell"] = startCell,
            ["rows_written"] = rowsWritten
        };
        return result.ToJsonString();
    }

    // --- get_sheet_info ---
    private string SkillGetSheetInfo(JsonNode args)
    {
        dynamic ws = GetTargetSheet(args);
        dynamic usedRng = ws.UsedRange;

        int usedRows = usedRng.Rows.Count;
        int usedCols = usedRng.Columns.Count;
        int maxCols = Math.Min(usedCols, 26);

        var headers = new JsonArray();
        for (int c = 1; c <= maxCols; c++)
        {
            object? v = usedRng.Cells[1, c].Value;
            headers.Add(v?.ToString() ?? "");
        }

        var result = new JsonObject
        {
            ["name"] = (string)ws.Name,
            ["index"] = (int)ws.Index,
            ["used_range"] = (string)usedRng.Address,
            ["used_rows"] = usedRows,
            ["used_cols"] = usedCols,
            ["first_cell"] = (string)usedRng.Cells[1, 1].Address,
            ["last_cell"] = (string)usedRng.Cells[usedRows, usedCols].Address,
            ["headers"] = headers
        };
        return result.ToJsonString();
    }

    // --- get_workbook_info ---
    private string SkillGetWorkbookInfo(JsonNode args)
    {
        dynamic app = GetApp();
        dynamic wb = app.ActiveWorkbook;

        var sheets = new JsonArray();
        int count = wb.Worksheets.Count;
        for (int i = 1; i <= count; i++)
            sheets.Add((string)wb.Worksheets[i].Name);

        var result = new JsonObject
        {
            ["name"] = (string)wb.Name,
            ["path"] = (string)wb.FullName,
            ["sheets"] = sheets,
            ["active_sheet"] = (string)app.ActiveSheet.Name
        };
        return result.ToJsonString();
    }

    // --- format_range ---
    private string SkillFormatRange(JsonNode args)
    {
        dynamic ws = GetTargetSheet(args);
        string rangeAddr = Str(args["range"]);
        dynamic rng = ws.Range[rangeAddr];

        if (args["bold"] != null) rng.Font.Bold = Bool(args["bold"]);
        if (args["italic"] != null) rng.Font.Italic = Bool(args["italic"]);
        if (args["font_size"] != null) rng.Font.Size = Dbl(args["font_size"]);
        if (args["font_color"] != null) rng.Font.Color = Int(args["font_color"]);
        if (args["bg_color"] != null) rng.Interior.Color = Int(args["bg_color"]);
        if (args["number_format"] != null) rng.NumberFormat = Str(args["number_format"]);

        if (args["h_align"] != null)
        {
            rng.HorizontalAlignment = Str(args["h_align"]).ToLowerInvariant() switch
            {
                "left" => -4131,
                "center" => -4108,
                "right" => -4152,
                _ => -4131
            };
        }

        if (args["wrap_text"] != null) rng.WrapText = Bool(args["wrap_text"]);

        if (args["borders"] != null)
        {
            if (Bool(args["borders"]))
            {
                rng.Borders.LineStyle = 1;  // xlContinuous
                rng.Borders.Weight = 2;     // xlThin
            }
            else
            {
                rng.Borders.LineStyle = -4142; // xlNone
            }
        }

        if (args["column_width"] != null) rng.ColumnWidth = Dbl(args["column_width"]);
        if (args["row_height"] != null) rng.RowHeight = Dbl(args["row_height"]);
        if (args["autofit"] != null && Bool(args["autofit"])) rng.Columns.AutoFit();
        if (args["merge"] != null)
        {
            if (Bool(args["merge"]))
                rng.Merge();
            else
                rng.UnMerge();
        }

        var result = new JsonObject
        {
            ["success"] = true,
            ["range"] = rangeAddr
        };
        return result.ToJsonString();
    }

    // --- insert_formula ---
    private string SkillInsertFormula(JsonNode args)
    {
        dynamic ws = GetTargetSheet(args);
        string cell = Str(args["cell"]);
        string formula = Str(args["formula"]);

        if (!formula.StartsWith('='))
            formula = "=" + formula;

        ws.Range[cell].Formula = formula;
        object? resultVal = ws.Range[cell].Value;

        var result = new JsonObject
        {
            ["success"] = true,
            ["cell"] = cell,
            ["formula"] = formula,
            ["result"] = resultVal?.ToString() ?? ""
        };
        return result.ToJsonString();
    }

    // --- sort_range ---
    private string SkillSortRange(JsonNode args)
    {
        dynamic ws = GetTargetSheet(args);
        string rangeAddr = Str(args["range"]);
        string sortCol = Str(args["sort_column"]);

        string orderStr = Str(args["order"], "asc").ToLowerInvariant();
        int sortOrder = (orderStr == "desc" || orderStr == "descending") ? 2 : 1; // 1=xlAscending, 2=xlDescending

        bool hasHeaders = Bool(args["has_headers"], true);
        int header = hasHeaders ? 1 : 2; // 1=xlYes, 2=xlNo

        dynamic rng = ws.Range[rangeAddr];
        int startRow = rng.Row;
        rng.Sort(
            Key1: ws.Range[$"{sortCol}{startRow}"],
            Order1: sortOrder,
            Header: header);

        var result = new JsonObject
        {
            ["success"] = true,
            ["range"] = rangeAddr,
            ["sort_column"] = sortCol
        };
        return result.ToJsonString();
    }

    // --- add_sheet ---
    private string SkillAddSheet(JsonNode args)
    {
        dynamic app = GetApp();
        string name = Str(args["name"]);

        dynamic ws = app.ActiveWorkbook.Worksheets.Add();
        if (!string.IsNullOrEmpty(name))
        {
            try { ws.Name = name; } catch { /* ignore rename failures */ }
        }

        var result = new JsonObject
        {
            ["success"] = true,
            ["name"] = (string)ws.Name
        };
        return result.ToJsonString();
    }

    // --- delete_rows ---
    private string SkillDeleteRows(JsonNode args)
    {
        dynamic ws = GetTargetSheet(args);
        int startRow = Int(args["start_row"]);
        int count = Int(args["count"], 1);

        ws.Rows[$"{startRow}:{startRow + count - 1}"].Delete();

        var result = new JsonObject
        {
            ["success"] = true,
            ["deleted_from"] = startRow,
            ["count"] = count
        };
        return result.ToJsonString();
    }

    // --- insert_rows ---
    private string SkillInsertRows(JsonNode args)
    {
        dynamic ws = GetTargetSheet(args);
        int atRow = Int(args["at_row"]);
        int count = Int(args["count"], 1);

        for (int i = 0; i < count; i++)
            ws.Rows[atRow].Insert(-4121); // xlDown

        var result = new JsonObject
        {
            ["success"] = true,
            ["at_row"] = atRow,
            ["count"] = count
        };
        return result.ToJsonString();
    }

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

    // ======================== HELPERS ========================

    private static JsonObject MakeTool(string name, string description,
        JsonObject properties, JsonArray required)
    {
        return new JsonObject
        {
            ["type"] = "function",
            ["function"] = new JsonObject
            {
                ["name"] = name,
                ["description"] = description,
                ["parameters"] = new JsonObject
                {
                    ["type"] = "object",
                    ["properties"] = properties,
                    ["required"] = required
                }
            }
        };
    }

    private static JsonObject PropString(string description) =>
        new() { ["type"] = "string", ["description"] = description };

    private static JsonObject PropNumber(string description) =>
        new() { ["type"] = "number", ["description"] = description };

    private static JsonObject PropBool(string description) =>
        new() { ["type"] = "boolean", ["description"] = description };
}
