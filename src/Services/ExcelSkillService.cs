using System.Text.Json;
using System.Text.Json.Nodes;
using ExcelDna.Integration;

namespace ZaiExcelAddin.Services;

public partial class ExcelSkillService
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
                "Write a 2D array of values starting from a cell. PREFERRED over multiple write_cell calls for bulk data.",
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
                "Get worksheet info: name, used range, dimensions, headers, data sample (first 5 rows), and flags for charts/pivot tables/filters",
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
                "Format cells: bold, italic, underline, font name/size/color, background color, number format, alignment, borders, column width, row height, autofit, merge",
                new JsonObject
                {
                    ["range"] = PropString("Range to format e.g. A1:D1"),
                    ["bold"] = PropBool("Set bold"),
                    ["italic"] = PropBool("Set italic"),
                    ["underline"] = PropBool("Set underline"),
                    ["font_name"] = PropString("Font name e.g. Arial, Calibri, Times New Roman"),
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
                "Insert an Excel formula into a cell. Use English function names (SUM, AVERAGE, VLOOKUP, IF, COUNT, etc.). Use fill_to to auto-fill the formula down/across a range.",
                new JsonObject
                {
                    ["cell"] = PropString("Target cell e.g. B10"),
                    ["formula"] = PropString("Formula e.g. =SUM(B1:B9) or =IF(A1>0,A1*2,0)"),
                    ["fill_to"] = PropString("End cell to auto-fill formula to e.g. B100 (optional — fills from cell to fill_to adjusting references)"),
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

            MakeTool("delete_columns",
                "Delete one or more columns from the worksheet",
                new JsonObject
                {
                    ["start_column"] = PropString("First column letter to delete e.g. C"),
                    ["count"] = PropNumber("Number of columns to delete (default: 1)"),
                    ["sheet"] = PropString("Sheet name (optional)")
                },
                new JsonArray { "start_column" }),

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
                    ["dest_sheet"] = PropString("Destination sheet name (optional, defaults to source sheet or new sheet)"),
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

            MakeTool("move_pivot_table",
                "Move a PivotTable to another sheet. Tries Location property, then recreates from PivotCache. On .xlsb files, recreated pivot is empty — use modify_pivot_table to configure fields afterward.",
                new JsonObject
                {
                    ["name"] = PropString("PivotTable name to move (use list_pivot_tables to find it)"),
                    ["source_range"] = PropString("Range containing the pivot e.g. H1:M30 (alternative to name, recommended for .xlsb)"),
                    ["dest_sheet"] = PropString("Destination sheet name (creates if not exists)"),
                    ["dest_cell"] = PropString("Destination cell e.g. A1 (default: A1)"),
                    ["sheet"] = PropString("Source sheet name (optional)")
                },
                new JsonArray()),

            MakeTool("modify_pivot_table",
                "Configure fields on an existing PivotTable: set row/column/value/page fields, change aggregation function, refresh. Use after create_pivot_table or move_pivot_table to set up field layout.",
                new JsonObject
                {
                    ["name"] = PropString("PivotTable name (optional, defaults to first pivot on sheet)"),
                    ["row_fields"] = new JsonObject
                    {
                        ["type"] = "array",
                        ["description"] = "Fields for row labels e.g. [\"Category\",\"Region\"]",
                        ["items"] = new JsonObject { ["type"] = "string" }
                    },
                    ["column_fields"] = new JsonObject
                    {
                        ["type"] = "array",
                        ["description"] = "Fields for column labels (optional)",
                        ["items"] = new JsonObject { ["type"] = "string" }
                    },
                    ["value_fields"] = new JsonObject
                    {
                        ["type"] = "array",
                        ["description"] = "Fields to aggregate e.g. [\"Sales\"]",
                        ["items"] = new JsonObject { ["type"] = "string" }
                    },
                    ["value_function"] = PropString("Aggregation: sum, count, average, max, min (default: sum)"),
                    ["page_fields"] = new JsonObject
                    {
                        ["type"] = "array",
                        ["description"] = "Fields for page/filter area (optional)",
                        ["items"] = new JsonObject { ["type"] = "string" }
                    },
                    ["clear_fields"] = PropBool("Reset all field orientations before applying new ones (default: true)"),
                    ["refresh"] = PropBool("Refresh the PivotTable after changes (default: false)"),
                    ["sheet"] = PropString("Sheet name (optional)")
                },
                new JsonArray()),

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
                "Add conditional formatting to a range. Supports color scale, data bars, icon sets, formula-based rules, and value-based rules.",
                new JsonObject
                {
                    ["range"] = PropString("Range to format e.g. B2:B100"),
                    ["rule_type"] = PropString("Rule type: color_scale, data_bars, icon_set, cell_value, formula, top_bottom, above_average, duplicate, unique. Use 'formula' for alternating row colors (=MOD(ROW(),2)=0), row highlighting, date comparisons."),
                    ["operator"] = PropString("For cell_value: greater, less, equal, between, not_between, greater_equal, less_equal. For top_bottom: top or bottom (default: top)"),
                    ["value1"] = PropString("Threshold value, or Excel formula for rule_type=formula (e.g. =MOD(ROW(),2)=0)"),
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
                new JsonArray { "range", "type", "formula1" }),

            MakeTool("list_pivot_tables",
                "List all PivotTables in the workbook or a specific sheet with their names, source ranges and locations",
                new JsonObject
                {
                    ["sheet"] = PropString("Sheet name (optional, lists all sheets if omitted)")
                },
                new JsonArray()),

            MakeTool("clear_range",
                "Clear contents, formatting, or everything from a range of cells",
                new JsonObject
                {
                    ["range"] = PropString("Range to clear e.g. A1:D10"),
                    ["what"] = PropString("What to clear: contents, formats, all (default: contents)"),
                    ["sheet"] = PropString("Sheet name (optional)")
                },
                new JsonArray { "range" })
        };
    }

    // ======================== TOOL EXECUTION ========================

    private const int VBA_E_IGNORE = unchecked((int)0x800AC472);

    public string Execute(string toolName, string argsJson)
    {
        // Retry once on VBA_E_IGNORE (Excel is busy / recalculating)
        for (int attempt = 0; attempt < 2; attempt++)
        {
            try { return ExecuteInner(toolName, argsJson); }
            catch (System.Runtime.InteropServices.COMException cx) when (cx.HResult == VBA_E_IGNORE && attempt == 0)
            {
                AddIn.Logger.Warn($"{toolName}: Excel busy (0x800AC472), retrying in 300ms...");
                System.Threading.Thread.Sleep(300);
            }
        }
        return ExecuteInner(toolName, argsJson); // final attempt — let exceptions propagate
    }

    private string ExecuteInner(string toolName, string argsJson)
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
                "delete_columns" => SkillDeleteColumns(args),
                "create_chart" => SkillCreateChart(args),
                "delete_chart" => SkillDeleteChart(args),
                "list_charts" => SkillListCharts(args),
                "create_pivot_table" => SkillCreatePivotTable(args),
                "move_pivot_table" => SkillMovePivotTable(args),
                "modify_pivot_table" => SkillModifyPivotTable(args),
                "auto_filter" => SkillAutoFilter(args),
                "find_replace" => SkillFindReplace(args),
                "conditional_format" => SkillConditionalFormat(args),
                "copy_range" => SkillCopyRange(args),
                "rename_sheet" => SkillRenameSheet(args),
                "delete_sheet" => SkillDeleteSheet(args),
                "freeze_panes" => SkillFreezePanes(args),
                "remove_duplicates" => SkillRemoveDuplicates(args),
                "set_validation" => SkillSetValidation(args),
                "list_pivot_tables" => SkillListPivotTables(args),
                "clear_range" => SkillClearRange(args),
                _ => JsonSerializer.Serialize(new { error = $"Unknown tool: {toolName}" })
            };

            AddIn.Logger.ToolCall(toolName, argsJson, result);
            return result;
        }
        catch (Exception ex)
        {
            var errorObj = new JsonObject { ["error"] = ex.Message };

            // Contextual hints for common COM errors
            if (ex.Message.Contains("pivot", StringComparison.OrdinalIgnoreCase) ||
                ex.Message.Contains("tabela przestawna", StringComparison.OrdinalIgnoreCase) ||
                ex.Message.Contains("PivotTable", StringComparison.OrdinalIgnoreCase))
            {
                errorObj["hint"] = "A PivotTable is blocking this operation. Use move_pivot_table to relocate it to another sheet first, then retry.";
            }
            else if (ex.HResult == unchecked((int)0x800A03EC))
            {
                errorObj["hint"] = "COM error — the range may be invalid, protected, or overlap an existing object. Check the range address and try again.";
            }
            else if (ex.Message.Contains("merge", StringComparison.OrdinalIgnoreCase) ||
                     ex.Message.Contains("scal", StringComparison.OrdinalIgnoreCase))
            {
                errorObj["hint"] = "Merged cells may be interfering. Try clear_range with what='formats' first, or work with a different range.";
            }
            else if (ex.HResult == unchecked((int)0x80028018))
            {
                errorObj["hint"] = "TYPE_E_INVDATAREAD — Excel type library error. This sheet may have COM interop issues. Try a different approach or skip this sheet.";
            }

            var error = errorObj.ToJsonString();
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
