using ExcelDna.Integration;
using ZaiExcelAddin.Services;

namespace ZaiExcelAddin;

/// <summary>
/// ExcelDNA commands exposed as macros callable by Excel (e.g. Application.OnUndo).
/// </summary>
public static class ExcelCommands
{
    [ExcelCommand(Name = "ZaiUndoRestore")]
    public static void ZaiUndoRestore()
    {
        UndoService.RestoreSnapshot();
    }
}
