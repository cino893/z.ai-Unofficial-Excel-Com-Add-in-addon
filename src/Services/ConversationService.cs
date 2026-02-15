using System.Text.Json;
using System.Text.Json.Nodes;

namespace ZaiExcelAddin.Services;

public enum StopReason { None, Completed, MaxRounds, LoopDetected, Cancelled, Error }

public class ConversationService
{
    private JsonArray _messages = null!;
    private bool _isProcessing;
    private CancellationTokenSource? _cts;
    private string _lastAssistantResponse = "";
    private const int MaxToolRounds = 45;
    private const int MaxSameToolRepeats = 2;
    private static readonly IReadOnlyDictionary<string, string> MaxToolRoundPlaceholders =
        new Dictionary<string, string> { ["MaxToolRounds"] = MaxToolRounds.ToString() };

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
                ["content"] = AddIn.I18n.T("system.prompt", MaxToolRoundPlaceholders)
            }
        };
        _lastAssistantResponse = "";
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

            TrimOldTurns();

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
        try
        {
            if (_messages == null)
                Init();

            TrimOldTurns();

            // Check before resetting — did we hit max rounds with a summary?
            var useLastSummary = LastStopReason == StopReason.MaxRounds
                && !string.IsNullOrWhiteSpace(_lastAssistantResponse);
            LastStopReason = StopReason.None;

            // Build context summary from recent tool calls
            var recentContext = BuildRecentContext();
            var continueMsg = useLastSummary ? _lastAssistantResponse
                : AddIn.I18n.T("conv.continue_prompt");
            if (!useLastSummary && !string.IsNullOrEmpty(recentContext))
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
        int roundInfoIndex = -1;
        int finalPromptIndex = -1;

        for (int round = 1; round <= MaxToolRounds; round++)
        {
            if (_cts?.IsCancellationRequested == true)
            {
                AddIn.Logger.Info("Cancelled by user");
                LastStopReason = StopReason.Cancelled;
                RemoveTransientMessages(ref roundInfoIndex, ref finalPromptIndex);
                return AddIn.I18n.T("conv.cancelled");
            }

            AddIn.Logger.Info($"Tool-calling loop round {round}/{MaxToolRounds}");

            // Replace (not accumulate) round info — always at end of messages
            var roundReplacements = new Dictionary<string, string>
            {
                ["CurrentRound"] = round.ToString(),
                ["MaxToolRounds"] = MaxToolRounds.ToString()
            };
            var roundMsg = new JsonObject
            {
                ["role"] = "system",
                ["content"] = AddIn.I18n.T("conv.round_info", roundReplacements)
            };
            if (roundInfoIndex >= 0 && roundInfoIndex < _messages.Count)
            {
                _messages.RemoveAt(roundInfoIndex);
                if (finalPromptIndex > roundInfoIndex) finalPromptIndex--;
            }
            _messages.Add(roundMsg);
            roundInfoIndex = _messages.Count - 1;

            if (round == MaxToolRounds)
            {
                _messages.Add(new JsonObject
                {
                    ["role"] = "user",
                    ["content"] = AddIn.I18n.T("conv.final_round_prompt", MaxToolRoundPlaceholders)
                });
                finalPromptIndex = _messages.Count - 1;
            }

            var (success, data, error) = AddIn.Api.SendCompletion(
                _messages, AddIn.Skills.GetToolDefinitions());

            if (!success || data == null)
            {
                LastStopReason = StopReason.Error;
                AddIn.Logger.Error($"API call failed: {error}");
                RemoveTransientMessages(ref roundInfoIndex, ref finalPromptIndex);
                return error ?? AddIn.I18n.T("conv.api_error");
            }

            var finishReason = ZaiApiService.GetFinishReason(data);
            // Log token usage for cost monitoring
            var usage = data["usage"];
            if (usage != null)
            {
                var prompt = usage["prompt_tokens"]?.GetValue<int>() ?? 0;
                var completion = usage["completion_tokens"]?.GetValue<int>() ?? 0;
                AddIn.Logger.Info($"Round {round}: finish={finishReason}, tokens={prompt}+{completion}={prompt + completion}");
            }
            else
            {
                AddIn.Logger.Info($"Finish reason: {finishReason}");
            }

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
                        // Add dummy results for orphaned tool_calls
                        foreach (var tc in toolCalls)
                        {
                            _messages.Add(new JsonObject
                            {
                                ["role"] = "tool",
                                ["content"] = "{\"error\":\"Loop detected, execution stopped\"}",
                                ["tool_call_id"] = tc!["id"]!.GetValue<string>()
                            });
                        }
                        LastStopReason = StopReason.LoopDetected;
                        RemoveTransientMessages(ref roundInfoIndex, ref finalPromptIndex);
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
                    var id = toolCall!["id"]!.GetValue<string>();
                    var name = toolCall["function"]!["name"]!.GetValue<string>();
                    var arguments = toolCall["function"]!["arguments"]!.GetValue<string>();

                    string result;
                    if (_cts?.IsCancellationRequested == true)
                    {
                        // Must provide result for every tool_call in the batch
                        result = "{\"error\":\"Cancelled by user\"}";
                    }
                    else
                    {
                        AddIn.Logger.Info($"Executing tool: {name}");
                        result = toolExecutor(name, arguments);
                    }

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
            if (string.IsNullOrWhiteSpace(content))
            {
                LastStopReason = StopReason.Error;
                AddIn.Logger.Error("Assistant returned empty response");
                RemoveTransientMessages(ref roundInfoIndex, ref finalPromptIndex);
                return AddIn.I18n.T("conv.generic_failure");
            }
            RemoveTransientMessages(ref roundInfoIndex, ref finalPromptIndex);
            _messages.Add(new JsonObject
            {
                ["role"] = "assistant",
                ["content"] = content
            });
            _lastAssistantResponse = content;

            AddIn.Logger.Info($"Assistant response received ({content.Length} chars)");
            LastStopReason = StopReason.Completed;
            return content;
        }

        RemoveTransientMessages(ref roundInfoIndex, ref finalPromptIndex);
        AddIn.Logger.Warn($"Max tool rounds ({MaxToolRounds}) reached");
        LastStopReason = StopReason.MaxRounds;
        return AddIn.I18n.TFormat("conv.max_rounds", MaxToolRounds);
    }

    /// <summary>Remove transient round-info and final-prompt messages from history.</summary>
    private void RemoveTransientMessages(ref int roundInfoIndex, ref int finalPromptIndex)
    {
        // Remove in descending order to avoid index shifting issues
        if (finalPromptIndex >= 0 && finalPromptIndex < _messages.Count)
        {
            _messages.RemoveAt(finalPromptIndex);
            if (roundInfoIndex > finalPromptIndex) roundInfoIndex--;
            finalPromptIndex = -1;
        }
        if (roundInfoIndex >= 0 && roundInfoIndex < _messages.Count)
        {
            _messages.RemoveAt(roundInfoIndex);
            roundInfoIndex = -1;
        }
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

    /// <summary>
    /// Remove oldest conversation turns when message history gets too long.
    /// Preserves system prompt (index 0) and the most recent turns.
    /// Removes complete turns (user → assistant → tools → ... → next user) to keep API happy.
    /// </summary>
    private void TrimOldTurns()
    {
        const int softLimit = 80;
        if (_messages == null || _messages.Count <= softLimit) return;

        // Find all user message indices (skip system prompt at 0)
        var userIndices = new List<int>();
        for (int i = 1; i < _messages.Count; i++)
        {
            if (_messages[i]?.AsObject()?["role"]?.GetValue<string>() == "user")
                userIndices.Add(i);
        }

        // Need at least 2 turns to trim (always keep the latest turn)
        while (_messages.Count > softLimit && userIndices.Count >= 2)
        {
            int from = userIndices[0];
            int to = userIndices[1];
            int count = to - from;

            for (int i = 0; i < count; i++)
                _messages.RemoveAt(from);

            int trimmed = count;
            userIndices.RemoveAt(0);
            for (int i = 0; i < userIndices.Count; i++)
                userIndices[i] -= trimmed;

            AddIn.Logger.Info($"Trimmed {trimmed} old messages, history now: {_messages.Count}");
        }
    }
}
