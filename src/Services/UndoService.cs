using ExcelDna.Integration;

namespace ZaiExcelAddin.Services;

/// <summary>
/// Captures a full workbook snapshot before the AI agent starts editing
/// and registers Application.OnUndo so the user can Ctrl+Z the entire batch.
/// </summary>
public static class UndoService
{
    private record SheetSnapshot(object?[,] Formulas, int Rows, int Cols);

    private static Dictionary<string, SheetSnapshot>? _snapshot;
    private static List<string>? _snapshotSheetNames;
    private static bool _hasSnapshot;

    /// <summary>Take a snapshot of every sheet's used range (formulas) before the agent edits.</summary>
    public static void CaptureSnapshot()
    {
        try
        {
            dynamic app = ExcelDnaUtil.Application;
            dynamic wb = app.ActiveWorkbook;
            if (wb == null) return;

            _snapshot = new Dictionary<string, SheetSnapshot>();
            _snapshotSheetNames = new List<string>();

            foreach (dynamic ws in wb.Worksheets)
            {
                string name = ws.Name;
                _snapshotSheetNames.Add(name);

                dynamic? used = ws.UsedRange;
                if (used == null) continue;
                int rows = used.Rows.Count;
                int cols = used.Columns.Count;
                if (rows == 0 || cols == 0) continue;

                // Capture formulas (preserves =SUM(...) etc., falls back to value for plain cells)
                object? raw;
                if (rows == 1 && cols == 1)
                {
                    // Single cell — Formula returns scalar string
                    var formulas = new object?[1, 1];
                    formulas[0, 0] = used.Formula;
                    _snapshot![name] = new SheetSnapshot(formulas, 1, 1);
                }
                else
                {
                    raw = used.Formula;
                    if (raw is object?[,] formulas)
                        _snapshot![name] = new SheetSnapshot(formulas, rows, cols);
                }
            }

            _hasSnapshot = true;
            AddIn.Logger.Info($"Undo snapshot captured ({_snapshot.Count} sheets)");
        }
        catch (Exception ex)
        {
            _hasSnapshot = false;
            AddIn.Logger.Error($"Failed to capture undo snapshot: {ex.Message}");
        }
    }

    /// <summary>Register Application.OnUndo so Ctrl+Z restores the snapshot.</summary>
    public static void RegisterUndo()
    {
        if (!_hasSnapshot) return;

        try
        {
            dynamic app = ExcelDnaUtil.Application;
            app.OnUndo("Z.AI: Undo AI changes", "ZaiUndoRestore");
            AddIn.Logger.Info("OnUndo registered for AI changes");
        }
        catch (Exception ex)
        {
            AddIn.Logger.Error($"Failed to register OnUndo: {ex.Message}");
        }
    }

    /// <summary>Called by Excel when user presses Ctrl+Z after agent edits.</summary>
    public static void RestoreSnapshot()
    {
        if (!_hasSnapshot || _snapshot == null)
        {
            AddIn.Logger.Warn("No snapshot to restore");
            return;
        }

        try
        {
            dynamic app = ExcelDnaUtil.Application;
            dynamic wb = app.ActiveWorkbook;

            bool oldScreenUpdating = app.ScreenUpdating;
            bool oldEnableEvents = app.EnableEvents;
            app.ScreenUpdating = false;
            app.EnableEvents = false;

            try
            {
                foreach (var kvp in _snapshot)
                {
                    try
                    {
                        dynamic ws = wb.Worksheets[kvp.Key];
                        // Clear everything on the sheet first
                        ws.Cells.Clear();

                        var snap = kvp.Value;
                        dynamic destRange = ws.Range[
                            ws.Cells[1, 1],
                            ws.Cells[snap.Rows, snap.Cols]
                        ];
                        // Restore formulas — this restores both formulas and plain values
                        destRange.Formula = snap.Formulas;
                    }
                    catch (Exception ex)
                    {
                        AddIn.Logger.Error($"Failed to restore sheet '{kvp.Key}': {ex.Message}");
                    }
                }

                // Delete sheets that were added by the agent
                if (_snapshotSheetNames != null)
                {
                    app.DisplayAlerts = false;
                    try
                    {
                        var currentSheets = new List<string>();
                        foreach (dynamic ws in wb.Worksheets)
                            currentSheets.Add((string)ws.Name);

                        foreach (var name in currentSheets)
                        {
                            if (!_snapshotSheetNames.Contains(name) && currentSheets.Count > 1)
                            {
                                try { wb.Worksheets[name].Delete(); }
                                catch { /* ignore */ }
                            }
                        }
                    }
                    finally
                    {
                        app.DisplayAlerts = true;
                    }
                }

                AddIn.Logger.Info("Undo snapshot restored successfully");
            }
            finally
            {
                app.ScreenUpdating = oldScreenUpdating;
                app.EnableEvents = oldEnableEvents;
            }

            _hasSnapshot = false;
            _snapshot = null;
            _snapshotSheetNames = null;
        }
        catch (Exception ex)
        {
            AddIn.Logger.Error($"RestoreSnapshot error: {ex.Message}");
        }
    }
}
