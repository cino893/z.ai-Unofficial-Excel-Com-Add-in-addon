using System.Net.Http;
using System.Security.Cryptography;
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

    // ═══ Balance API (requires JWT auth) ═══
    private const string BalanceUrl = "https://api.z.ai/api/platform-charge-zai/business/accountBalance";
    private string? _cachedBalance;
    private DateTime _balanceFetchedAt = DateTime.MinValue;

    /// <summary>Generate JWT token from API key (format: id.secret) using HS256.</summary>
    private static string? GenerateJwt(string apiKey, int expireSeconds = 300)
    {
        try
        {
            var parts = apiKey.Split('.');
            if (parts.Length != 2) return null;
            var id = parts[0];
            var secret = parts[1];

            // Header
            var header = Convert.ToBase64String(Encoding.UTF8.GetBytes("""{"alg":"HS256","sign_type":"SIGN","typ":"JWT"}"""))
                .TrimEnd('=').Replace('+', '-').Replace('/', '_');

            // Payload
            var now = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
            var exp = DateTimeOffset.UtcNow.AddSeconds(expireSeconds).ToUnixTimeMilliseconds();
            var payload = Convert.ToBase64String(Encoding.UTF8.GetBytes(
                    $$$"""{"api_key":"{{{id}}}","exp":{{{exp}}},"timestamp":{{{now}}}}"""))
                .TrimEnd('=').Replace('+', '-').Replace('/', '_');

            // Signature
            var data = $"{header}.{payload}";
            using var hmac = new HMACSHA256(Encoding.UTF8.GetBytes(secret));
            var sig = Convert.ToBase64String(hmac.ComputeHash(Encoding.UTF8.GetBytes(data)))
                .TrimEnd('=').Replace('+', '-').Replace('/', '_');

            return $"{header}.{payload}.{sig}";
        }
        catch { return null; }
    }

    /// <summary>Fetch account balance. Caches for 60 seconds.</summary>
    public string GetBalance(bool forceRefresh = false)
    {
        if (!forceRefresh && _cachedBalance != null && (DateTime.Now - _balanceFetchedAt).TotalSeconds < 60)
            return _cachedBalance;

        var apiKey = AddIn.Auth.LoadApiKey().Trim();
        apiKey = new string(apiKey.Where(c => c >= 0x20 && c <= 0x7E).ToArray());
        if (string.IsNullOrEmpty(apiKey)) return "—";

        try
        {
            var jwt = GenerateJwt(apiKey);
            var token = jwt ?? apiKey; // fallback to raw key if JWT fails

            using var request = new HttpRequestMessage(HttpMethod.Post, BalanceUrl);
            request.Content = new StringContent("{}", Encoding.UTF8, "application/json");
            request.Headers.Add("Authorization", $"Bearer {token}");
            var response = _http.Send(request);
            var body = response.Content.ReadAsStringAsync().Result;
            AddIn.Logger.Debug($"Balance response: {body[..Math.Min(200, body.Length)]}");

            if (response.IsSuccessStatusCode)
            {
                var json = JsonNode.Parse(body);
                var balance = json?["data"]?["balance"]?.GetValue<decimal>()
                    ?? json?["data"]?["available"]?.GetValue<decimal>();
                if (balance.HasValue)
                {
                    _cachedBalance = $"{balance.Value:F2} CNY";
                    _balanceFetchedAt = DateTime.Now;
                    return _cachedBalance;
                }
                // Try to extract any numeric from data
                var dataStr = json?["data"]?.ToJsonString() ?? body;
                _cachedBalance = dataStr.Length > 40 ? dataStr[..40] + "..." : dataStr;
                _balanceFetchedAt = DateTime.Now;
                return _cachedBalance;
            }
            AddIn.Logger.Debug($"Balance error: HTTP {(int)response.StatusCode}");
            return "—";
        }
        catch (Exception ex)
        {
            AddIn.Logger.Debug($"Balance fetch error: {ex.Message}");
            return "—";
        }
    }

    public void InvalidateBalanceCache() { _cachedBalance = null; }

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

    public static readonly string[] DefaultModels = ["glm-4.5-air", "glm-4.5", "glm-4.6", "glm-4.7", "glm-5"];
}
