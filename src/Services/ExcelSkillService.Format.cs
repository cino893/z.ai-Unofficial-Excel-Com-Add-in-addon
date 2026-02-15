using System.Text.Json;
using System.Text.Json.Nodes;
using ExcelDna.Integration;

namespace ZaiExcelAddin.Services;

public partial class ExcelSkillService
{
    // --- format_range ---
    private string SkillFormatRange(JsonNode args)
    {
        dynamic ws = GetTargetSheet(args);
        string rangeAddr = Str(args["range"]);
        dynamic rng = ws.Range[rangeAddr];

        if (args["bold"] != null) rng.Font.Bold = Bool(args["bold"]);
        if (args["italic"] != null) rng.Font.Italic = Bool(args["italic"]);
        if (args["underline"] != null)
            rng.Font.Underline = Bool(args["underline"]) ? 2 : -4142; // xlUnderlineStyleSingle / xlUnderlineStyleNone
        if (args["font_name"] != null) rng.Font.Name = Str(args["font_name"]);
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
        string fillTo = Str(args["fill_to"]);

        if (!formula.StartsWith('='))
            formula = "=" + formula;

        ws.Range[cell].Formula = formula;

        int cellsFilled = 1;
        if (!string.IsNullOrEmpty(fillTo))
        {
            dynamic sourceRange = ws.Range[cell];
            dynamic fillRange = ws.Range[$"{cell}:{fillTo}"];
            sourceRange.AutoFill(Destination: fillRange, Type: 0); // xlFillDefault
            cellsFilled = fillRange.Cells.Count;
        }

        object? resultVal = ws.Range[cell].Value;

        var result = new JsonObject
        {
            ["success"] = true,
            ["cell"] = cell,
            ["formula"] = formula,
            ["result"] = resultVal?.ToString() ?? "",
            ["cells_filled"] = cellsFilled
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
                string tbDir = Str(args["operator"], "top").ToLowerInvariant();
                dynamic tbRule = rng.FormatConditions.AddTop10();
                tbRule.TopBottom = (tbDir == "bottom") ? 0 : 1; // 0=xlTop10Bottom, 1=xlTop10Top
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
}
