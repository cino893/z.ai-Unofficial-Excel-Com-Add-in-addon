using System.Text.Json;
using System.Text.Json.Nodes;
using ExcelDna.Integration;

namespace ZaiExcelAddin.Services;

public partial class ExcelSkillService
{
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
}
