using ExcelDna.Integration;

namespace ZaiExcelAddin.Services;

/// <summary>
/// Captures a full workbook snapshot before the AI agent starts editing
/// and registers Application.OnUndo so the user can Ctrl+Z the entire batch.
/// </summary>
public static class UndoService
{
    private static Dictionary<string, object?[,]>? _snapshot;
    private static List<string>? _snapshotSheetNames;
    private static bool _hasSnapshot;

    /// <summary>Take a snapshot of every sheet's used range before the agent edits.</summary>
    public static void CaptureSnapshot()
    {
        try
        {
            dynamic app = ExcelDnaUtil.Application;
            dynamic wb = app.ActiveWorkbook;
            if (wb == null) return;

            _snapshot = new Dictionary<string, object?[,]>();
            _snapshotSheetNames = new List<string>();

            foreach (dynamic ws in wb.Worksheets)
            {
                string name = ws.Name;
                _snapshotSheetNames.Add(name);

                dynamic? used = ws.UsedRange;
                if (used == null) continue;
                int cnt = used.Count;
                if (cnt == 0) continue;

                object? raw = used.Value;
                if (raw is object?[,] values)
                    _snapshot![name] = values;
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
                        dynamic used = ws.UsedRange;
                        // Clear current content first
                        used.ClearContents();

                        // Restore saved values
                        var values = kvp.Value;
                        int rows = values.GetLength(0);
                        int cols = values.GetLength(1);
                        dynamic destRange = ws.Range[
                            ws.Cells[1, 1],
                            ws.Cells[rows, cols]
                        ];
                        destRange.Value = values;
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
