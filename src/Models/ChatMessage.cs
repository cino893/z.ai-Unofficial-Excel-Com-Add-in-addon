namespace ZaiExcelAddin.Models;

public class ChatMessage
{
    public string Role { get; set; } = "";
    public string Content { get; set; } = "";
    public DateTime Timestamp { get; set; } = DateTime.Now;
    public bool IsUser => Role == "user";
    public bool IsAssistant => Role == "assistant";
    public bool IsSystem => Role == "system" || Role == "info";
}
