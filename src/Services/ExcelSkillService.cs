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
                new JsonArray()),

            MakeTool("create_pivot_table",
                "Create a PivotTable from a data range. Specify row, column and value fields.",
                new JsonObject
                {
                    ["source_range"] = PropString("Source data range e.g. A1:D100"),
                    ["dest_cell"] = PropString("Top-left cell for PivotTable placement e.g. F1 (default: new sheet)"),
                    ["name"] = PropString("PivotTable name (optional)"),
                    ["row_fields"] = new JsonObject
                    {
                        ["type"] = "array",
                        ["description"] = "Fields to use as row labels e.g. [\"Category\",\"Region\"]",
                        ["items"] = new JsonObject { ["type"] = "string" }
                    },
                    ["column_fields"] = new JsonObject
                    {
                        ["type"] = "array",
                        ["description"] = "Fields to use as column labels (optional)",
                        ["items"] = new JsonObject { ["type"] = "string" }
                    },
                    ["value_fields"] = new JsonObject
                    {
                        ["type"] = "array",
                        ["description"] = "Fields to aggregate as values e.g. [\"Sales\",\"Quantity\"]",
                        ["items"] = new JsonObject { ["type"] = "string" }
                    },
                    ["value_function"] = PropString("Aggregation function: sum, count, average, max, min (default: sum)"),
                    ["sheet"] = PropString("Source sheet name (optional)")
                },
                new JsonArray { "source_range", "row_fields", "value_fields" }),

            MakeTool("auto_filter",
                "Apply or remove auto-filter on a range. Call without criteria to toggle filter on/off.",
                new JsonObject
                {
                    ["range"] = PropString("Range to filter e.g. A1:D20"),
                    ["field"] = PropNumber("Column number within range to filter (1-based, optional)"),
                    ["criteria"] = PropString("Filter criteria value (optional, omit to show all or toggle)"),
                    ["operator"] = PropString("Filter operator: equals, not_equals, greater, less, contains, top10, bottom10 (default: equals)"),
                    ["clear"] = PropBool("Set true to clear all filters"),
                    ["sheet"] = PropString("Sheet name (optional)")
                },
                new JsonArray { "range" }),

            MakeTool("find_replace",
                "Find and replace values in a range or entire sheet",
                new JsonObject
                {
                    ["find"] = PropString("Text to find"),
                    ["replace"] = PropString("Replacement text"),
                    ["range"] = PropString("Range to search (optional, defaults to entire sheet)"),
                    ["match_case"] = PropBool("Case-sensitive matching (default: false)"),
                    ["match_entire"] = PropBool("Match entire cell contents (default: false)"),
                    ["sheet"] = PropString("Sheet name (optional)")
                },
                new JsonArray { "find", "replace" }),

            MakeTool("conditional_format",
                "Add conditional formatting to a range. Supports color scale, data bars, icon sets, and value-based rules.",
                new JsonObject
                {
                    ["range"] = PropString("Range to format e.g. B2:B100"),
                    ["rule_type"] = PropString("Rule type: color_scale, data_bars, icon_set, cell_value, top_bottom, above_average, duplicate, unique"),
                    ["operator"] = PropString("For cell_value: greater, less, equal, between, not_between, greater_equal, less_equal (optional)"),
                    ["value1"] = PropString("Threshold value or formula (for cell_value/top_bottom rules)"),
                    ["value2"] = PropString("Second value (for between/not_between operator)"),
                    ["format_color"] = PropNumber("Fill color as RGB long for matching cells (e.g. 65280 for green)"),
                    ["font_color"] = PropNumber("Font color as RGB long for matching cells"),
                    ["clear_existing"] = PropBool("Clear existing conditional formats before adding (default: false)"),
                    ["sheet"] = PropString("Sheet name (optional)")
                },
                new JsonArray { "range", "rule_type" }),

            MakeTool("copy_range",
                "Copy a range of cells to another location (values, formulas and formatting)",
                new JsonObject
                {
                    ["source"] = PropString("Source range e.g. A1:D10"),
                    ["destination"] = PropString("Destination top-left cell e.g. F1"),
                    ["dest_sheet"] = PropString("Destination sheet name (optional, defaults to same sheet)"),
                    ["values_only"] = PropBool("Paste values only, no formulas/formatting (default: false)"),
                    ["sheet"] = PropString("Source sheet name (optional)")
                },
                new JsonArray { "source", "destination" }),

            MakeTool("rename_sheet",
                "Rename a worksheet",
                new JsonObject
                {
                    ["sheet"] = PropString("Current sheet name (optional, defaults to active sheet)"),
                    ["new_name"] = PropString("New name for the sheet")
                },
                new JsonArray { "new_name" }),

            MakeTool("delete_sheet",
                "Delete a worksheet from the workbook",
                new JsonObject
                {
                    ["sheet"] = PropString("Sheet name to delete")
                },
                new JsonArray { "sheet" }),

            MakeTool("freeze_panes",
                "Freeze or unfreeze rows/columns for scrolling. Freezes above and to the left of the specified cell.",
                new JsonObject
                {
                    ["cell"] = PropString("Cell below and to the right of the freeze point e.g. A2 (freeze row 1), B3 (freeze row 1-2 and column A)"),
                    ["unfreeze"] = PropBool("Set true to remove all freeze panes"),
                    ["sheet"] = PropString("Sheet name (optional)")
                },
                new JsonArray()),

            MakeTool("remove_duplicates",
                "Remove duplicate rows from a range based on specified columns",
                new JsonObject
                {
                    ["range"] = PropString("Range to deduplicate e.g. A1:D100"),
                    ["columns"] = new JsonObject
                    {
                        ["type"] = "array",
                        ["description"] = "Column numbers (1-based within range) to check for duplicates e.g. [1,3]. Defaults to all columns.",
                        ["items"] = new JsonObject { ["type"] = "number" }
                    },
                    ["has_headers"] = PropBool("Whether first row contains headers (default: true)"),
                    ["sheet"] = PropString("Sheet name (optional)")
                },
                new JsonArray { "range" }),

            MakeTool("set_validation",
                "Add data validation to cells (dropdown lists, number ranges, etc.)",
                new JsonObject
                {
                    ["range"] = PropString("Range to validate e.g. C2:C100"),
                    ["type"] = PropString("Validation type: list, whole_number, decimal, date, text_length"),
                    ["formula1"] = PropString("For list: comma-separated values e.g. Yes,No,Maybe. For numbers: min value. For date: min date."),
                    ["formula2"] = PropString("For numbers/dates: max value (optional, for between operator)"),
                    ["operator"] = PropString("Operator: between, not_between, equal, not_equal, greater, less, greater_equal, less_equal (default: between)"),
                    ["show_dropdown"] = PropBool("Show in-cell dropdown for list type (default: true)"),
                    ["error_message"] = PropString("Custom error message (optional)"),
                    ["sheet"] = PropString("Sheet name (optional)")
                },
                new JsonArray { "range", "type", "formula1" })
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
                "create_pivot_table" => SkillCreatePivotTable(args),
                "auto_filter" => SkillAutoFilter(args),
                "find_replace" => SkillFindReplace(args),
                "conditional_format" => SkillConditionalFormat(args),
                "copy_range" => SkillCopyRange(args),
                "rename_sheet" => SkillRenameSheet(args),
                "delete_sheet" => SkillDeleteSheet(args),
                "freeze_panes" => SkillFreezePanes(args),
                "remove_duplicates" => SkillRemoveDuplicates(args),
                "set_validation" => SkillSetValidation(args),
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
        string qualifiedSource = $"{ws.Name}!{srcRange.Address}";

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
            destSheet = !string.IsNullOrEmpty(destSheetName)
                ? wb.Worksheets[destSheetName] : ws;
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

    // --- auto_filter ---
    private string SkillAutoFilter(JsonNode args)
    {
        dynamic ws = GetTargetSheet(args);
        string rangeAddr = Str(args["range"]);
        dynamic rng = ws.Range[rangeAddr];

        if (Bool(args["clear"]))
        {
            if (ws.AutoFilterMode) ws.AutoFilterMode = false;
            return new JsonObject { ["success"] = true, ["action"] = "cleared" }.ToJsonString();
        }

        int field = Int(args["field"]);
        string criteria = Str(args["criteria"]);

        if (field > 0 && !string.IsNullOrEmpty(criteria))
        {
            string op = Str(args["operator"], "equals").ToLowerInvariant();

            if (op == "contains")
            {
                rng.AutoFilter(Field: field, Criteria1: $"*{criteria}*");
            }
            else if (op == "not_equals")
            {
                rng.AutoFilter(Field: field, Criteria1: $"<>{criteria}");
            }
            else if (op == "greater")
            {
                rng.AutoFilter(Field: field, Criteria1: $">{criteria}");
            }
            else if (op == "less")
            {
                rng.AutoFilter(Field: field, Criteria1: $"<{criteria}");
            }
            else if (op == "top10")
            {
                rng.AutoFilter(Field: field, Criteria1: "10",
                    Operator: 1); // xlTop10Items
            }
            else if (op == "bottom10")
            {
                rng.AutoFilter(Field: field, Criteria1: "10",
                    Operator: 4); // xlBottom10Items
            }
            else
            {
                rng.AutoFilter(Field: field, Criteria1: criteria);
            }
        }
        else
        {
            // Toggle auto-filter
            rng.AutoFilter();
        }

        var result = new JsonObject
        {
            ["success"] = true,
            ["range"] = rangeAddr,
            ["action"] = field > 0 ? "filtered" : "toggled"
        };
        return result.ToJsonString();
    }

    // --- find_replace ---
    private string SkillFindReplace(JsonNode args)
    {
        dynamic ws = GetTargetSheet(args);
        string findStr = Str(args["find"]);
        string replaceStr = Str(args["replace"]);
        string rangeAddr = Str(args["range"]);
        bool matchCase = Bool(args["match_case"]);
        bool matchEntire = Bool(args["match_entire"]);

        dynamic searchRange = string.IsNullOrEmpty(rangeAddr)
            ? ws.Cells : ws.Range[rangeAddr];

        int lookAt = matchEntire ? 1 : 2; // 1=xlWhole, 2=xlPart
        bool replaced = searchRange.Replace(
            What: findStr,
            Replacement: replaceStr,
            LookAt: lookAt,
            MatchCase: matchCase);

        var result = new JsonObject
        {
            ["success"] = replaced,
            ["find"] = findStr,
            ["replace"] = replaceStr
        };
        return result.ToJsonString();
    }

    // --- conditional_format ---
    private string SkillConditionalFormat(JsonNode args)
    {
        dynamic ws = GetTargetSheet(args);
        string rangeAddr = Str(args["range"]);
        dynamic rng = ws.Range[rangeAddr];
        string ruleType = Str(args["rule_type"]).ToLowerInvariant();

        if (Bool(args["clear_existing"]))
            rng.FormatConditions.Delete();

        string action = ruleType;

        switch (ruleType)
        {
            case "color_scale":
                rng.FormatConditions.AddColorScale(3);
                break;

            case "data_bars":
                rng.FormatConditions.AddDatabar();
                break;

            case "icon_set":
                rng.FormatConditions.AddIconSetCondition();
                break;

            case "duplicate":
                // xlDuplicateValues = 1 (FormatConditions.AddUniqueValues)
                dynamic dupRule = rng.FormatConditions.AddUniqueValues();
                dupRule.DupeUnique = 1; // xlDuplicate
                if (args["format_color"] != null)
                    dupRule.Interior.Color = Int(args["format_color"]);
                else
                    dupRule.Interior.Color = 13551615; // light red
                break;

            case "unique":
                dynamic uniqRule = rng.FormatConditions.AddUniqueValues();
                uniqRule.DupeUnique = 0; // xlUnique
                if (args["format_color"] != null)
                    uniqRule.Interior.Color = Int(args["format_color"]);
                else
                    uniqRule.Interior.Color = 5296274; // light green
                break;

            case "above_average":
                // xlAboveAverage = 0
                dynamic aaRule = rng.FormatConditions.AddAboveAverage();
                if (args["format_color"] != null)
                    aaRule.Interior.Color = Int(args["format_color"]);
                break;

            case "top_bottom":
                string tbVal = Str(args["value1"], "10");
                // xlTop10Top = 1
                dynamic tbRule = rng.FormatConditions.AddTop10();
                tbRule.TopBottom = 1; // xlTop10Top
                tbRule.Rank = int.Parse(tbVal);
                if (args["format_color"] != null)
                    tbRule.Interior.Color = Int(args["format_color"]);
                break;

            case "cell_value":
            default:
                string opStr = Str(args["operator"], "greater").ToLowerInvariant();
                string val1 = Str(args["value1"], "0");
                string val2 = Str(args["value2"]);

                int xlOp = opStr switch
                {
                    "less" => 6,           // xlLess
                    "equal" => 3,          // xlEqual
                    "between" => 1,        // xlBetween
                    "not_between" => 2,    // xlNotBetween
                    "greater_equal" => 7,  // xlGreaterEqual
                    "less_equal" => 8,     // xlLessEqual
                    _ => 5                 // xlGreater
                };

                dynamic cf;
                if (xlOp == 1 || xlOp == 2) // between/not_between
                    cf = rng.FormatConditions.Add(Type: 1, Operator: xlOp, Formula1: val1, Formula2: val2);
                else
                    cf = rng.FormatConditions.Add(Type: 1, Operator: xlOp, Formula1: val1);

                if (args["format_color"] != null)
                    cf.Interior.Color = Int(args["format_color"]);
                if (args["font_color"] != null)
                    cf.Font.Color = Int(args["font_color"]);
                action = "cell_value";
                break;
        }

        var result = new JsonObject
        {
            ["success"] = true,
            ["range"] = rangeAddr,
            ["rule_type"] = action
        };
        return result.ToJsonString();
    }

    // --- copy_range ---
    private string SkillCopyRange(JsonNode args)
    {
        dynamic ws = GetTargetSheet(args);
        string sourceAddr = Str(args["source"]);
        string destAddr = Str(args["destination"]);
        string destSheetName = Str(args["dest_sheet"]);
        bool valuesOnly = Bool(args["values_only"]);

        dynamic srcRange = ws.Range[sourceAddr];
        dynamic destSheet = !string.IsNullOrEmpty(destSheetName)
            ? ((dynamic)GetApp()).ActiveWorkbook.Worksheets[destSheetName] : ws;
        dynamic destRange = destSheet.Range[destAddr];

        if (valuesOnly)
        {
            srcRange.Copy();
            // xlPasteValues = -4163
            destRange.PasteSpecial(Paste: -4163);
            ((dynamic)GetApp()).CutCopyMode = false;
        }
        else
        {
            srcRange.Copy(destRange);
        }

        var result = new JsonObject
        {
            ["success"] = true,
            ["source"] = sourceAddr,
            ["destination"] = destAddr
        };
        return result.ToJsonString();
    }

    // --- rename_sheet ---
    private string SkillRenameSheet(JsonNode args)
    {
        dynamic ws = GetTargetSheet(args);
        string oldName = (string)ws.Name;
        string newName = Str(args["new_name"]);

        ws.Name = newName;

        var result = new JsonObject
        {
            ["success"] = true,
            ["old_name"] = oldName,
            ["new_name"] = newName
        };
        return result.ToJsonString();
    }

    // --- delete_sheet ---
    private string SkillDeleteSheet(JsonNode args)
    {
        dynamic app = GetApp();
        string sheetName = Str(args["sheet"]);
        dynamic ws = app.ActiveWorkbook.Worksheets[sheetName];

        app.DisplayAlerts = false;
        try
        {
            ws.Delete();
        }
        finally
        {
            app.DisplayAlerts = true;
        }

        var result = new JsonObject
        {
            ["success"] = true,
            ["deleted"] = sheetName
        };
        return result.ToJsonString();
    }

    // --- freeze_panes ---
    private string SkillFreezePanes(JsonNode args)
    {
        dynamic app = GetApp();

        if (Bool(args["unfreeze"]))
        {
            app.ActiveWindow.FreezePanes = false;
            return new JsonObject { ["success"] = true, ["action"] = "unfrozen" }.ToJsonString();
        }

        string cell = Str(args["cell"], "A2");
        string sheetName = Str(args["sheet"]);

        if (!string.IsNullOrEmpty(sheetName))
            app.ActiveWorkbook.Worksheets[sheetName].Activate();

        app.ActiveWindow.FreezePanes = false;
        app.ActiveSheet.Range[cell].Select();
        app.ActiveWindow.FreezePanes = true;

        var result = new JsonObject
        {
            ["success"] = true,
            ["action"] = "frozen",
            ["cell"] = cell
        };
        return result.ToJsonString();
    }

    // --- remove_duplicates ---
    private string SkillRemoveDuplicates(JsonNode args)
    {
        dynamic ws = GetTargetSheet(args);
        string rangeAddr = Str(args["range"]);
        dynamic rng = ws.Range[rangeAddr];
        bool hasHeaders = Bool(args["has_headers"], true);

        int rowsBefore = rng.Rows.Count;

        var colsNode = args["columns"]?.AsArray();
        if (colsNode != null && colsNode.Count > 0)
        {
            var colArray = colsNode.Select(c => c!.GetValue<int>()).ToArray();
            rng.RemoveDuplicates(Columns: colArray, Header: hasHeaders ? 1 : 2);
        }
        else
        {
            int colCount = rng.Columns.Count;
            var allCols = Enumerable.Range(1, colCount).ToArray();
            rng.RemoveDuplicates(Columns: allCols, Header: hasHeaders ? 1 : 2);
        }

        int rowsAfter = ws.Range[rangeAddr].Rows.Count;

        var result = new JsonObject
        {
            ["success"] = true,
            ["range"] = rangeAddr,
            ["rows_before"] = rowsBefore,
            ["rows_removed"] = rowsBefore - rowsAfter
        };
        return result.ToJsonString();
    }

    // --- set_validation ---
    private string SkillSetValidation(JsonNode args)
    {
        dynamic ws = GetTargetSheet(args);
        string rangeAddr = Str(args["range"]);
        dynamic rng = ws.Range[rangeAddr];
        string typeStr = Str(args["type"]).ToLowerInvariant();
        string formula1 = Str(args["formula1"]);
        string formula2 = Str(args["formula2"]);
        string opStr = Str(args["operator"], "between").ToLowerInvariant();

        // Clear existing validation
        try { rng.Validation.Delete(); } catch { /* ignore */ }

        int xlType = typeStr switch
        {
            "list" => 3,            // xlValidateList
            "whole_number" => 1,    // xlValidateWholeNumber
            "decimal" => 2,         // xlValidateDecimal
            "date" => 4,            // xlValidateDate
            "text_length" => 6,     // xlValidateTextLength
            _ => 3
        };

        int xlOp = opStr switch
        {
            "not_between" => 2,
            "equal" => 3,
            "not_equal" => 4,
            "greater" => 5,
            "less" => 6,
            "greater_equal" => 7,
            "less_equal" => 8,
            _ => 1 // xlBetween
        };

        int xlAlertStop = 1; // xlValidAlertStop

        if (xlType == 3) // list
        {
            rng.Validation.Add(Type: xlType, AlertStyle: xlAlertStop, Formula1: formula1);
            bool showDropdown = Bool(args["show_dropdown"], true);
            rng.Validation.InCellDropdown = showDropdown;
        }
        else if (!string.IsNullOrEmpty(formula2))
        {
            rng.Validation.Add(Type: xlType, AlertStyle: xlAlertStop,
                Operator: xlOp, Formula1: formula1, Formula2: formula2);
        }
        else
        {
            rng.Validation.Add(Type: xlType, AlertStyle: xlAlertStop,
                Operator: xlOp, Formula1: formula1);
        }

        string errorMsg = Str(args["error_message"]);
        if (!string.IsNullOrEmpty(errorMsg))
        {
            rng.Validation.ErrorMessage = errorMsg;
            rng.Validation.ShowError = true;
        }

        var result = new JsonObject
        {
            ["success"] = true,
            ["range"] = rangeAddr,
            ["type"] = typeStr
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
