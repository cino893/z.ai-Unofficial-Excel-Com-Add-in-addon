using System.Text.Json;
using System.Text.Json.Nodes;
using ExcelDna.Integration;

namespace ZaiExcelAddin.Services;

public partial class ExcelSkillService
{
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
    private const int MaxReadRows = 500;

    private string SkillReadRange(JsonNode args)
    {
        dynamic ws = GetTargetSheet(args);
        string rangeAddr = Str(args["range"]);
        dynamic rng = ws.Range[rangeAddr];

        int totalRows = rng.Rows.Count;
        int cols = rng.Columns.Count;
        bool truncated = totalRows > MaxReadRows;
        int rows = truncated ? MaxReadRows : totalRows;

        var data = new JsonArray();

        if (rows == 1 && cols == 1)
        {
            // Single cell — Value is scalar, not array
            object? cellVal = rng.Value;
            var row = new JsonArray();
            row.Add(cellVal is double d ? JsonValue.Create(d) : JsonValue.Create(cellVal?.ToString() ?? ""));
            data.Add(row);
        }
        else
        {
            // Bulk read — single COM call (truncated range if needed)
            dynamic readRange = truncated
                ? ws.Range[rng.Cells[1, 1], rng.Cells[rows, cols]]
                : rng;
            object?[,] values = readRange.Value;
            for (int r = 1; r <= rows; r++)
            {
                var row = new JsonArray();
                for (int c = 1; c <= cols; c++)
                {
                    object? cellVal = values[r, c];
                    if (cellVal == null)
                        row.Add(null);
                    else if (cellVal is double d)
                        row.Add(d);
                    else
                        row.Add(cellVal.ToString());
                }
                data.Add(row);
            }
        }

        var result = new JsonObject
        {
            ["range"] = rangeAddr,
            ["sheet"] = (string)ws.Name,
            ["rows"] = rows,
            ["cols"] = cols,
            ["data"] = data
        };

        if (truncated)
        {
            result["truncated"] = true;
            result["total_rows"] = totalRows;
        }

        return result.ToJsonString();
    }

    // --- write_range ---
    private string SkillWriteRange(JsonNode args)
    {
        dynamic ws = GetTargetSheet(args);
        string startCell = Str(args["start_cell"]);
        var data = args["data"]!.AsArray();

        int rows = data.Count;
        if (rows == 0)
            return new JsonObject { ["success"] = true, ["start_cell"] = startCell, ["rows_written"] = 0 }.ToJsonString();

        int cols = data[0]!.AsArray().Count;

        // Build 2D object array for bulk write (single COM call)
        var values = new object?[rows, cols];
        for (int r = 0; r < rows; r++)
        {
            var rowData = data[r]!.AsArray();
            for (int c = 0; c < Math.Min(rowData.Count, cols); c++)
            {
                values[r, c] = rowData[c]?.GetValue<string>() ?? "";
            }
        }

        dynamic startRange = ws.Range[startCell];
        dynamic destRange = ws.Range[startRange, startRange.Offset[rows - 1, cols - 1]];
        destRange.Value = values;

        var result = new JsonObject
        {
            ["success"] = true,
            ["start_cell"] = startCell,
            ["rows_written"] = rows,
            ["cells_written"] = rows * cols
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

        // Include first 5 rows as data sample so model understands structure without read_range
        int sampleRows = Math.Min(usedRows, 5);
        var sample = new JsonArray();
        if (sampleRows > 0 && maxCols > 0)
        {
            if (sampleRows == 1 && maxCols == 1)
            {
                object? v = usedRng.Cells[1, 1].Value;
                var row = new JsonArray { v?.ToString() ?? "" };
                sample.Add(row);
            }
            else
            {
                dynamic sampleRange = ws.Range[usedRng.Cells[1, 1], usedRng.Cells[sampleRows, maxCols]];
                object?[,] values = sampleRange.Value;
                for (int r = 1; r <= sampleRows; r++)
                {
                    var row = new JsonArray();
                    for (int c = 1; c <= maxCols; c++)
                    {
                        object? cv = values[r, c];
                        row.Add(cv?.ToString() ?? "");
                    }
                    sample.Add(row);
                }
            }
        }

        // Quick presence flags — saves the model from calling list_charts/list_pivot_tables just to check
        bool hasCharts = false;
        bool hasPivots = false;
        bool hasFilters = false;
        try { hasCharts = ws.ChartObjects.Count > 0; } catch { }
        try { hasPivots = ws.PivotTables().Count > 0; } catch { }
        try { hasFilters = (bool)ws.AutoFilterMode; } catch { }

        var result = new JsonObject
        {
            ["name"] = (string)ws.Name,
            ["index"] = (int)ws.Index,
            ["used_range"] = (string)usedRng.Address,
            ["used_rows"] = usedRows,
            ["used_cols"] = usedCols,
            ["first_cell"] = (string)usedRng.Cells[1, 1].Address,
            ["last_cell"] = (string)usedRng.Cells[usedRows, usedCols].Address,
            ["headers"] = headers,
            ["sample_data"] = sample,
            ["has_charts"] = hasCharts,
            ["has_pivot_tables"] = hasPivots,
            ["has_filters"] = hasFilters
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

    // --- remove_duplicates ---
    private string SkillRemoveDuplicates(JsonNode args)
    {
        dynamic ws = GetTargetSheet(args);
        string rangeAddr = Str(args["range"]);
        dynamic rng = ws.Range[rangeAddr];
        bool hasHeaders = Bool(args["has_headers"], true);

        int rowsBefore = rng.Rows.Count;
        int startRow = rng.Row;
        int col = rng.Column;
        dynamic app = GetApp();

        // Count non-empty rows in first column before dedup (single COM call)
        dynamic countRange = ws.Range[ws.Cells[startRow, col], ws.Cells[startRow + rowsBefore - 1, col]];
        int dataBefore = (int)app.WorksheetFunction.CountA(countRange);

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

        // Count non-empty rows after dedup
        int dataAfter = (int)app.WorksheetFunction.CountA(countRange);

        var result = new JsonObject
        {
            ["success"] = true,
            ["range"] = rangeAddr,
            ["rows_before"] = dataBefore,
            ["rows_after"] = dataAfter,
            ["rows_removed"] = dataBefore - dataAfter
        };
        return result.ToJsonString();
    }

    // --- clear_range ---
    private string SkillClearRange(JsonNode args)
    {
        dynamic ws = GetTargetSheet(args);
        string rangeAddr = Str(args["range"]);
        string what = Str(args["what"], "contents").ToLowerInvariant();
        dynamic rng = ws.Range[rangeAddr];

        switch (what)
        {
            case "formats":
                rng.ClearFormats();
                break;
            case "all":
                rng.Clear();
                break;
            default: // "contents"
                rng.ClearContents();
                break;
        }

        return new JsonObject
        {
            ["success"] = true,
            ["range"] = rangeAddr,
            ["cleared"] = what
        }.ToJsonString();
    }
}
