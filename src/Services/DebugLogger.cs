using System.IO;

namespace ZaiExcelAddin.Services;

public class DebugLogger
{
    private readonly string _logPath;
    private readonly object _lock = new();
    private int _writeCount;
    private const long MaxLogSize = 2 * 1024 * 1024; // 2 MB
    private const int TrimCheckInterval = 50;

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
    {
        // Compact summary: model + message count — NOT the full body
        try
        {
            var json = System.Text.Json.JsonDocument.Parse(body);
            var model = "";
            var msgCount = 0;
            var toolCount = 0;
            if (json.RootElement.TryGetProperty("model", out var m)) model = m.GetString() ?? "";
            if (json.RootElement.TryGetProperty("messages", out var msgs)) msgCount = msgs.GetArrayLength();
            if (json.RootElement.TryGetProperty("tools", out var tools)) toolCount = tools.GetArrayLength();
            Log("API-REQ", $"{method} {url} model={model} msgs={msgCount} tools={toolCount}");
        }
        catch
        {
            Log("API-REQ", $"{method} {url} body_len={body.Length}");
        }
    }

    public void ApiResponse(int status, string body)
    {
        // Compact summary: status + finish_reason + token usage
        try
        {
            var json = System.Text.Json.JsonDocument.Parse(body);
            var root = json.RootElement;

            if (root.TryGetProperty("error", out var err))
            {
                var code = err.TryGetProperty("code", out var c) ? c.GetString() : "";
                var msg = err.TryGetProperty("message", out var em) ? em.GetString() : "";
                Log("API-RES", $"HTTP {status} error={code} {Truncate(msg ?? "", 200)}");
                return;
            }

            var finish = "";
            var promptTokens = 0;
            var compTokens = 0;
            var toolCallNames = "";
            if (root.TryGetProperty("choices", out var ch) && ch.GetArrayLength() > 0)
            {
                finish = ch[0].TryGetProperty("finish_reason", out var fr) ? fr.GetString() ?? "" : "";
                // Extract tool call names for debugging
                if (ch[0].TryGetProperty("message", out var msg) &&
                    msg.TryGetProperty("tool_calls", out var tcs))
                {
                    var names = new System.Collections.Generic.List<string>();
                    foreach (var tc in tcs.EnumerateArray())
                    {
                        if (tc.TryGetProperty("function", out var fn) &&
                            fn.TryGetProperty("name", out var n))
                            names.Add(n.GetString() ?? "?");
                    }
                    if (names.Count > 0)
                        toolCallNames = $" tools=[{string.Join(",", names)}]";
                }
            }
            if (root.TryGetProperty("usage", out var u))
            {
                if (u.TryGetProperty("prompt_tokens", out var pt)) promptTokens = pt.GetInt32();
                if (u.TryGetProperty("completion_tokens", out var ct2)) compTokens = ct2.GetInt32();
            }
            Log("API-RES", $"HTTP {status} finish={finish} tokens={promptTokens}+{compTokens}={promptTokens + compTokens}{toolCallNames}");
        }
        catch
        {
            Log("API-RES", $"HTTP {status} body_len={body.Length}");
        }
    }

    public void ToolCall(string name, string args, string result)
        => Log("TOOL", $"{name}({Truncate(args, 500)}) => {Truncate(result, 800)}");

    private void Log(string level, string msg)
    {
        try
        {
            lock (_lock)
            {
                if (++_writeCount % TrimCheckInterval == 0)
                    TrimIfNeeded();

                File.AppendAllText(_logPath,
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [{level}] {msg}\n");
            }
        }
        catch { /* ignore logging failures */ }
    }

    private void TrimIfNeeded()
    {
        try
        {
            var fi = new FileInfo(_logPath);
            if (!fi.Exists || fi.Length <= MaxLogSize) return;

            // Keep the last half of the file (line-aligned)
            var text = File.ReadAllText(_logPath);
            int mid = text.Length / 2;
            int cutAt = text.IndexOf('\n', mid);
            if (cutAt > 0 && cutAt < text.Length - 1)
                File.WriteAllText(_logPath, text[(cutAt + 1)..]);
            else
                File.WriteAllText(_logPath, ""); // fallback: clear
        }
        catch { /* ignore trim failures */ }
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
