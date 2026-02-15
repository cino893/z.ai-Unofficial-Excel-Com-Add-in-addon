using System.IO;

namespace ZaiExcelAddin.Services;

public class DebugLogger
{
    private readonly string _logPath;
    private readonly object _lock = new();

    public DebugLogger()
    {
        var dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "ZaiExcelAddin");
        Directory.CreateDirectory(dir);
        _logPath = Path.Combine(dir, "debug.log");
    }

    public void Info(string msg) => Log("INFO", msg);
    public void Warn(string msg) => Log("WARN", msg);
    public void Error(string msg) => Log("ERROR", msg);
    public void Debug(string msg) => Log("DEBUG", msg);

    public void ApiRequest(string method, string url, string body)
        => Log("API-REQ", $"{method} {url}\n{Truncate(body, 2000)}");

    public void ApiResponse(int status, string body)
        => Log("API-RES", $"HTTP {status}\n{Truncate(body, 2000)}");

    public void ToolCall(string name, string args, string result)
        => Log("TOOL", $"{name}({Truncate(args, 500)}) => {Truncate(result, 500)}");

    private void Log(string level, string msg)
    {
        try
        {
            lock (_lock)
            {
                File.AppendAllText(_logPath,
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [{level}] {msg}\n");
            }
        }
        catch { /* ignore logging failures */ }
    }

    private static string Truncate(string s, int max)
        => s.Length <= max ? s : s[..max] + "...";

    public void ViewLog()
    {
        if (File.Exists(_logPath))
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = _logPath,
                UseShellExecute = true
            });
        }
        else
        {
            System.Windows.Forms.MessageBox.Show(
                AddIn.I18n.T("debug.no_log"), "Z.AI",
                System.Windows.Forms.MessageBoxButtons.OK);
        }
    }

    public void ClearLog()
    {
        try { if (File.Exists(_logPath)) File.Delete(_logPath); } catch { }
    }
}
