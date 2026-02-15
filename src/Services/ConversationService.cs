using System.Text.Json;
using System.Text.Json.Nodes;

namespace ZaiExcelAddin.Services;

public class ConversationService
{
    private JsonArray _messages = null!;
    private bool _isProcessing;
    private const int MaxToolRounds = 15;
    private const int MaxSameToolRepeats = 2;

    public bool IsProcessing => _isProcessing;

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
        AddIn.Logger.Info("Conversation initialized");
    }

    public void Reset() => Init();

    public string SendMessage(string userMessage, Func<string, string, string> toolExecutor)
    {
        _isProcessing = true;
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

            string? previousSignature = null;
            int repeatCount = 0;

            for (int round = 1; round <= MaxToolRounds; round++)
            {
                AddIn.Logger.Info($"Tool-calling loop round {round}/{MaxToolRounds}");

                var (success, data, error) = AddIn.Api.SendCompletion(
                    _messages, AddIn.Skills.GetToolDefinitions());

                if (!success || data == null)
                {
                    AddIn.Logger.Error($"API call failed: {error}");
                    return AddIn.I18n.T("conv.api_error");
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
                return content;
            }

            AddIn.Logger.Warn($"Max tool rounds ({MaxToolRounds}) reached");
            return AddIn.I18n.TFormat("conv.max_rounds", MaxToolRounds);
        }
        catch (Exception ex)
        {
            AddIn.Logger.Error($"SendMessage error: {ex}");
            return $"Error: {ex.Message}";
        }
        finally
        {
            _isProcessing = false;
        }
    }
}
