using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace ZaiExcelAddin.Services;

public class ZaiApiService
{
    private const string ApiBase = "https://api.z.ai/api/paas/v4";
    private const int MaxTokens = 4096;
    private const double Temperature = 0.7;
    private static readonly HttpClient _http = new() { Timeout = TimeSpan.FromSeconds(120) };

    public (bool Success, JsonNode? Data, string? Error) SendCompletion(
        JsonArray messages, JsonArray? tools = null, string? model = null)
    {
        var apiKey = AddIn.Auth.LoadApiKey();
        if (string.IsNullOrEmpty(apiKey))
        {
            AddIn.Logger.Error("SendCompletion: No API key");
            return (false, null, "No API key configured");
        }

        model ??= AddIn.Auth.LoadModel();

        var body = new JsonObject
        {
            ["model"] = model,
            ["messages"] = messages.DeepClone(),
            ["max_tokens"] = MaxTokens,
            ["temperature"] = Temperature
        };

        if (tools != null && tools.Count > 0)
        {
            body["tools"] = tools.DeepClone();
            body["tool_choice"] = "auto";
        }

        var json = body.ToJsonString();
        AddIn.Logger.ApiRequest("POST", $"{ApiBase}/chat/completions", json);

        try
        {
            using var request = new HttpRequestMessage(HttpMethod.Post, $"{ApiBase}/chat/completions");
            request.Content = new StringContent(json, Encoding.UTF8, "application/json");
            request.Headers.Add("Authorization", $"Bearer {apiKey}");
            request.Headers.Add("Accept", "application/json");

            var response = _http.Send(request);
            var responseBody = response.Content.ReadAsStringAsync().Result;
            AddIn.Logger.ApiResponse((int)response.StatusCode, responseBody);

            if (response.IsSuccessStatusCode)
            {
                var data = JsonNode.Parse(responseBody);
                return (true, data, null);
            }
            else
            {
                var friendlyError = TranslateApiError((int)response.StatusCode, responseBody);
                AddIn.Logger.Error($"API error: HTTP {(int)response.StatusCode}");
                return (false, null, friendlyError);
            }
        }
        catch (Exception ex)
        {
            AddIn.Logger.Error($"Network error: {ex.Message}");
            return (false, null, $"Network error: {ex.Message}");
        }
    }

    public static string? GetResponseContent(JsonNode data)
    {
        return data?["choices"]?[0]?["message"]?["content"]?.GetValue<string>();
    }

    public static JsonArray? GetToolCalls(JsonNode data)
    {
        return data?["choices"]?[0]?["message"]?["tool_calls"]?.AsArray();
    }

    public static bool HasToolCalls(JsonNode data)
    {
        var tc = GetToolCalls(data);
        return tc != null && tc.Count > 0;
    }

    public static string GetFinishReason(JsonNode data)
    {
        return data?["choices"]?[0]?["finish_reason"]?.GetValue<string>() ?? "unknown";
    }

    public static JsonObject? GetAssistantMessage(JsonNode data)
    {
        return data?["choices"]?[0]?["message"]?.AsObject();
    }

    /// <summary>Translate z.ai Chinese error messages to user-friendly language.</summary>
    private string TranslateApiError(int httpCode, string body)
    {
        var t = AddIn.I18n;
        try
        {
            var json = JsonNode.Parse(body);
            var code = json?["error"]?["code"]?.GetValue<string>() ?? "";
            return code switch
            {
                "1261" => t.T("error.balance_empty"),     // 额度已用完
                "1301" => t.T("error.content_filter"),     // 内容审核
                "1302" => t.T("error.content_filter"),
                _ => httpCode switch
                {
                    401 => t.T("error.invalid_key"),
                    429 => t.T("error.rate_limit"),
                    _ => t.T("error.api_generic") + $" (HTTP {httpCode})"
                }
            };
        }
        catch
        {
            return httpCode switch
            {
                401 => t.T("error.invalid_key"),
                429 => t.T("error.rate_limit"),
                _ => t.T("error.api_generic") + $" (HTTP {httpCode})"
            };
        }
    }
}
