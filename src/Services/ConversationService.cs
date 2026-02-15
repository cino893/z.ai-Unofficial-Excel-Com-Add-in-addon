using System.Text.Json;
using System.Text.Json.Nodes;

namespace ZaiExcelAddin.Services;

public enum StopReason { None, Completed, MaxRounds, LoopDetected, Cancelled, Error }

public class ConversationService
{
    private JsonArray _messages = null!;
    private bool _isProcessing;
    private CancellationTokenSource? _cts;
    private const int MaxToolRounds = 15;
    private const int MaxSameToolRepeats = 2;

    public bool IsProcessing => _isProcessing;
    public StopReason LastStopReason { get; private set; } = StopReason.None;
    public bool CanContinue => LastStopReason is StopReason.MaxRounds or StopReason.Cancelled;

    public int MessageCount => _messages?.Count ?? 0;

    public void Init()
    {
        _messages = new JsonArray
        {
            new JsonObject
            {
                ["role"] = "system",
                ["content"] = AddIn.I18n.T("system.prompt")
            }
        };
        LastStopReason = StopReason.None;
        AddIn.Logger.Info("Conversation initialized");
    }

    public void Reset() => Init();

    public void Cancel()
    {
        _cts?.Cancel();
        AddIn.Logger.Info("Cancellation requested");
    }

    public string SendMessage(string userMessage, Func<string, string, string> toolExecutor)
    {
        _isProcessing = true;
        _cts = new CancellationTokenSource();
        LastStopReason = StopReason.None;
        try
        {
            if (_messages == null)
                Init();

            _messages!.Add(new JsonObject
            {
                ["role"] = "user",
                ["content"] = userMessage
            });

            AddIn.Logger.Info($"User message added, history size: {_messages.Count}");
            return RunToolLoop(toolExecutor);
        }
        catch (Exception ex)
        {
            LastStopReason = StopReason.Error;
            AddIn.Logger.Error($"SendMessage error: {ex}");
            return $"Error: {ex.Message}";
        }
        finally
        {
            _isProcessing = false;
            _cts?.Dispose();
            _cts = null;
        }
    }

    /// <summary>Continue execution after max rounds or cancellation.</summary>
    public string Continue(Func<string, string, string> toolExecutor)
    {
        _isProcessing = true;
        _cts = new CancellationTokenSource();
        LastStopReason = StopReason.None;
        try
        {
            if (_messages == null)
                Init();

            // Build context summary from recent tool calls
            var recentContext = BuildRecentContext();
            var continueMsg = AddIn.I18n.T("conv.continue_prompt");
            if (!string.IsNullOrEmpty(recentContext))
                continueMsg += "\n\nLast completed actions:\n" + recentContext;

            _messages!.Add(new JsonObject
            {
                ["role"] = "user",
                ["content"] = continueMsg
            });

            AddIn.Logger.Info("Continue requested, resuming tool loop");
            return RunToolLoop(toolExecutor);
        }
        catch (Exception ex)
        {
            LastStopReason = StopReason.Error;
            AddIn.Logger.Error($"Continue error: {ex}");
            return $"Error: {ex.Message}";
        }
        finally
        {
            _isProcessing = false;
            _cts?.Dispose();
            _cts = null;
        }
    }

    private string RunToolLoop(Func<string, string, string> toolExecutor)
    {
        string? previousSignature = null;
        int repeatCount = 0;

        for (int round = 1; round <= MaxToolRounds; round++)
        {
            if (_cts?.IsCancellationRequested == true)
            {
                AddIn.Logger.Info("Cancelled by user");
                LastStopReason = StopReason.Cancelled;
                return AddIn.I18n.T("conv.cancelled");
            }

            AddIn.Logger.Info($"Tool-calling loop round {round}/{MaxToolRounds}");

            var (success, data, error) = AddIn.Api.SendCompletion(
                _messages, AddIn.Skills.GetToolDefinitions());

            if (!success || data == null)
            {
                LastStopReason = StopReason.Error;
                AddIn.Logger.Error($"API call failed: {error}");
                return error ?? AddIn.I18n.T("conv.api_error");
            }

            var finishReason = ZaiApiService.GetFinishReason(data);
            AddIn.Logger.Info($"Finish reason: {finishReason}");

            if (ZaiApiService.HasToolCalls(data))
            {
                var assistantMsg = ZaiApiService.GetAssistantMessage(data);
                _messages.Add(assistantMsg!.DeepClone());

                var toolCalls = ZaiApiService.GetToolCalls(data)!;

                // Build signature for loop detection
                var signature = string.Join("|", toolCalls.Select(tc =>
                    $"{tc!["function"]!["name"]!.GetValue<string>()}:" +
                    $"{tc["function"]!["arguments"]!.GetValue<string>()}"));

                if (signature == previousSignature)
                {
                    repeatCount++;
                    AddIn.Logger.Warn($"Same tool signature repeated ({repeatCount}/{MaxSameToolRepeats})");
                    if (repeatCount >= MaxSameToolRepeats)
                    {
                        AddIn.Logger.Warn("Tool loop detected, breaking out");
                        LastStopReason = StopReason.LoopDetected;
                        return AddIn.I18n.T("conv.loop_detected");
                    }
                }
                else
                {
                    repeatCount = 0;
                }
                previousSignature = signature;

                foreach (var toolCall in toolCalls)
                {
                    if (_cts?.IsCancellationRequested == true)
                    {
                        AddIn.Logger.Info("Cancelled by user during tool execution");
                        LastStopReason = StopReason.Cancelled;
                        return AddIn.I18n.T("conv.cancelled");
                    }

                    var id = toolCall!["id"]!.GetValue<string>();
                    var name = toolCall["function"]!["name"]!.GetValue<string>();
                    var arguments = toolCall["function"]!["arguments"]!.GetValue<string>();

                    AddIn.Logger.Info($"Executing tool: {name}");
                    var result = toolExecutor(name, arguments);
                    AddIn.Logger.ToolCall(name, arguments, result);

                    _messages.Add(new JsonObject
                    {
                        ["role"] = "tool",
                        ["content"] = result,
                        ["tool_call_id"] = id
                    });
                }

                continue;
            }

            // No tool calls — final assistant response
            var content = ZaiApiService.GetResponseContent(data) ?? "";
            _messages.Add(new JsonObject
            {
                ["role"] = "assistant",
                ["content"] = content
            });

            AddIn.Logger.Info($"Assistant response received ({content.Length} chars)");
            LastStopReason = StopReason.Completed;
            return content;
        }

        AddIn.Logger.Warn($"Max tool rounds ({MaxToolRounds}) reached");
        LastStopReason = StopReason.MaxRounds;
        return AddIn.I18n.TFormat("conv.max_rounds", MaxToolRounds);
    }

    /// <summary>Build a short summary of recent tool calls for context on continue.</summary>
    private string BuildRecentContext()
    {
        if (_messages == null) return "";

        var lines = new List<string>();
        // Scan last messages for tool calls (up to 10 most recent)
        for (int i = _messages.Count - 1; i >= 0 && lines.Count < 10; i--)
        {
            var msg = _messages[i]?.AsObject();
            if (msg == null) continue;
            var role = msg["role"]?.GetValue<string>();

            if (role == "tool")
            {
                var content = msg["content"]?.GetValue<string>() ?? "";
                // Truncate long tool results
                if (content.Length > 100)
                    content = content[..100] + "...";
                lines.Add($"  → result: {content}");
            }
            else if (role == "assistant" && msg.ContainsKey("tool_calls"))
            {
                var toolCalls = msg["tool_calls"]?.AsArray();
                if (toolCalls != null)
                {
                    foreach (var tc in toolCalls)
                    {
                        var name = tc?["function"]?["name"]?.GetValue<string>() ?? "?";
                        lines.Add($"  • called: {name}");
                    }
                }
            }
            else if (role == "user")
            {
                break; // Stop at the previous user message
            }
        }

        lines.Reverse();
        return string.Join("\n", lines);
    }
}
