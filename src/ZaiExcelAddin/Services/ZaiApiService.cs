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
        var apiKey = AddIn.Auth.LoadApiKey().Trim();
        // Ensure key only contains ASCII (copy-paste can introduce invisible Unicode)
        apiKey = new string(apiKey.Where(c => c >= 0x20 && c <= 0x7E).ToArray());
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
                "1261" => t.T("error.balance_empty"),
                "1301" => t.T("error.content_filter"),
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

    // ═══ Model catalog with pricing (from docs.z.ai/guides/overview/pricing) ═══
    public record ModelInfo(string Id, string InputPrice, string OutputPrice, int SortOrder)
    {
        public string Emoji => SortOrder switch
        {
            < 10  => "⚡",   // FREE
            < 30  => "💚",   // Budget
            < 50  => "🔷",   // Standard
            _     => "💎",   // Premium
        };
        public string PriceTag => SortOrder < 10
            ? "FREE"
            : $"${InputPrice}→${OutputPrice}/MTok";
        public string DisplayLine => $"{Emoji}  {Id}   —   {PriceTag}";
    }

    public static readonly ModelInfo[] ModelCatalog =
    [
        new("glm-4.7-flash",       "0",    "0",    0),
        new("glm-4.5-flash",       "0",    "0",    1),
        new("glm-4.7-flashx",      "0.07", "0.4",  10),
        new("glm-4-32b-0414-128k", "0.1",  "0.1",  11),
        new("glm-4.5-air",         "0.2",  "1.1",  20),
        new("glm-4.5",             "0.6",  "2.2",  30),
        new("glm-4.6",             "0.6",  "2.2",  31),
        new("glm-4.7",             "0.6",  "2.2",  32),
        new("glm-4.5-airx",        "1.1",  "4.5",  40),
        new("glm-5",               "1",    "3.2",  41),
        new("glm-5-code",          "1.2",  "5",    42),
        new("glm-4.5-x",           "2.2",  "8.9",  50),
    ];

    private static readonly Dictionary<string, ModelInfo> _catalogLookup =
        ModelCatalog.ToDictionary(m => m.Id, StringComparer.OrdinalIgnoreCase);

    /// <summary>Get models for display in select dialog, sorted by price.</summary>
    public (string Id, string Display)[] GetModelsForDisplay()
    {
        // Fetch from API + merge known models
        var apiModels = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            var fetched = GetAvailableModels();
            foreach (var m in fetched) apiModels.Add(m);
        }
        catch { }
        // Always include free flash models (API doesn't list them)
        apiModels.Add("glm-4.7-flash");
        apiModels.Add("glm-4.5-flash");

        var result = new List<(string Id, string Display)>();
        // First add catalog models in order
        foreach (var cat in ModelCatalog)
        {
            if (apiModels.Remove(cat.Id) || true) // show all catalog models
                result.Add((cat.Id, cat.DisplayLine));
        }
        // Add any API models not in catalog
        foreach (var id in apiModels.OrderBy(x => x))
        {
            result.Add((id, $"❓  {id}   —   unknown pricing"));
        }
        return result.ToArray();
    }

    // ═══ Models API ═══
    private string[]? _cachedModels;
    private DateTime _modelsFetchedAt = DateTime.MinValue;

    /// <summary>Fetch available models from API. Caches for 5 minutes.</summary>
    public string[] GetAvailableModels(bool forceRefresh = false)
    {
        if (!forceRefresh && _cachedModels != null && (DateTime.Now - _modelsFetchedAt).TotalMinutes < 5)
            return _cachedModels;

        var apiKey = AddIn.Auth.LoadApiKey().Trim();
        apiKey = new string(apiKey.Where(c => c >= 0x20 && c <= 0x7E).ToArray());
        if (string.IsNullOrEmpty(apiKey)) return DefaultModels;

        try
        {
            using var request = new HttpRequestMessage(HttpMethod.Get, $"{ApiBase}/models");
            request.Headers.Add("Authorization", $"Bearer {apiKey}");
            var response = _http.Send(request);
            var body = response.Content.ReadAsStringAsync().Result;

            if (response.IsSuccessStatusCode)
            {
                var json = JsonNode.Parse(body);
                var data = json?["data"]?.AsArray();
                if (data != null && data.Count > 0)
                {
                    _cachedModels = data
                        .Select(m => m?["id"]?.GetValue<string>())
                        .Where(id => !string.IsNullOrEmpty(id))
                        .Cast<string>()
                        .OrderBy(id => id)
                        .ToArray();
                    _modelsFetchedAt = DateTime.Now;
                    AddIn.Logger.Info($"Fetched {_cachedModels.Length} models from API");
                    return _cachedModels;
                }
            }
        }
        catch (Exception ex)
        {
            AddIn.Logger.Debug($"Models fetch error: {ex.Message}");
        }
        return DefaultModels;
    }

    public static readonly string[] DefaultModels =
        ["glm-4.7-flash", "glm-4.5-flash", "glm-4.5-air", "glm-4.5", "glm-4.6", "glm-4.7", "glm-5"];
}
