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
                var error = $"HTTP {(int)response.StatusCode}: {responseBody[..Math.Min(500, responseBody.Length)]}";
                AddIn.Logger.Error($"API error: {error}");
                return (false, null, error);
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
}
