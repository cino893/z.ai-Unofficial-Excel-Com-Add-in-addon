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
}
