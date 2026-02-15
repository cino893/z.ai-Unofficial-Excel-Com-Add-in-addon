using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text.Json;
using Microsoft.Win32;

namespace ZaiExcelAddin.Services;

public class I18nService
{
    private const string RegKeyPath = @"SOFTWARE\ZaiExcelAddin";
    private const string RegValueName = "Language";
    private const string DefaultLanguage = "en";

    public static Dictionary<string, string> SupportedLanguages { get; } = new()
    {
        { "en", "English" },
        { "pl", "Polski" },
        { "de", "Deutsch" },
        { "fr", "Français" },
        { "es", "Español" },
        { "uk", "Українська" },
        { "zh", "简体中文" },
        { "ja", "日本語" }
    };

    public string CurrentLanguage { get; private set; } = DefaultLanguage;

    private readonly Dictionary<string, Dictionary<string, string>> _translations = new();

    public I18nService()
    {
        LoadAllTranslations();
        var saved = LoadLanguageFromRegistry();
        if (saved != null && SupportedLanguages.ContainsKey(saved))
        {
            CurrentLanguage = saved;
        }
        else
        {
            CurrentLanguage = DetectLanguage();
        }
        try { AddIn.Logger?.Info($"I18nService initialized, language: {CurrentLanguage}"); } catch { }
    }

    public string T(string key, IReadOnlyDictionary<string, string>? replacements = null)
    {
        string value;
        if (_translations.TryGetValue(CurrentLanguage, out var lang) && lang.TryGetValue(key, out var val))
            value = val;
        else if (_translations.TryGetValue(DefaultLanguage, out var en) && en.TryGetValue(key, out var fallback))
            value = fallback;
        else
            value = key;

        return ApplyReplacements(value, replacements);
    }

    public string TFormat(string key, object arg0)
    {
        return T(key).Replace("{0}", arg0?.ToString() ?? "");
    }

    private static string ApplyReplacements(string value, IReadOnlyDictionary<string, string>? replacements)
    {
        if (string.IsNullOrEmpty(value) || replacements == null || replacements.Count == 0)
            return value;

        foreach (var kv in replacements)
        {
            if (string.IsNullOrEmpty(kv.Key))
                continue;

            var placeholder = "{{" + kv.Key + "}}";
            value = value.Replace(placeholder, kv.Value ?? string.Empty);
        }

        return value;
    }

    public void SetLanguage(string code)
    {
        if (!SupportedLanguages.ContainsKey(code)) return;
        CurrentLanguage = code;
        SaveLanguageToRegistry(code);
        try { AddIn.Logger?.Info($"Language changed to: {code}"); } catch { }
    }

    private string DetectLanguage()
    {
        var culture = CultureInfo.CurrentUICulture;
        var twoLetter = culture.TwoLetterISOLanguageName.ToLowerInvariant();
        if (SupportedLanguages.ContainsKey(twoLetter))
            return twoLetter;
        return DefaultLanguage;
    }

    private string? LoadLanguageFromRegistry()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(RegKeyPath);
            return key?.GetValue(RegValueName) as string;
        }
        catch
        {
            return null;
        }
    }

    private void SaveLanguageToRegistry(string code)
    {
        try
        {
            using var key = Registry.CurrentUser.CreateSubKey(RegKeyPath);
            key.SetValue(RegValueName, code);
        }
        catch (Exception ex)
        {
            try { AddIn.Logger?.Error($"Failed to save language to registry: {ex.Message}"); } catch { }
        }
    }

    private void LoadAllTranslations()
    {
        var assembly = Assembly.GetExecutingAssembly();
        foreach (var lang in SupportedLanguages.Keys)
        {
            var resourceName = $"ZaiExcelAddin.i18n.{lang}.json";
            using var stream = assembly.GetManifestResourceStream(resourceName);
            if (stream == null) continue;
            using var reader = new StreamReader(stream);
            var json = reader.ReadToEnd();
            var dict = JsonSerializer.Deserialize<Dictionary<string, string>>(json);
            if (dict != null)
            {
                // Normalize values: decode any JSON-style escapes that may be present
                var normalized = new Dictionary<string, string>();
                foreach (var kv in dict)
                {
                    normalized[kv.Key] = UnescapeJsonEscapes(kv.Value);
                }
                _translations[lang] = normalized;
            }
        }
    }

    private string UnescapeJsonEscapes(string value)
    {
        if (string.IsNullOrEmpty(value)) return value;
        if (!value.Contains("\\u") && !value.Contains("\\n") && !value.Contains("\\r") && !value.Contains("\\t"))
            return value;
        try
        {
            // Wrap the value as a JSON string literal so the parser decodes any \u escapes
            var jsonLiteral = "\"" + value.Replace("\"", "\\\"") + "\"";
            var decoded = JsonSerializer.Deserialize<string>(jsonLiteral);
            return decoded ?? value;
        }
        catch
        {
            try
            {
                return System.Text.RegularExpressions.Regex.Unescape(value);
            }
            catch
            {
                return value;
            }
        }
    }
}
