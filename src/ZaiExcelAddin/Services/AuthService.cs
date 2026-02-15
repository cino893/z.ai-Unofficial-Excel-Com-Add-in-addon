using System.Drawing;
using System.Net.Http;
using System.Windows.Forms;
using Microsoft.Win32;

namespace ZaiExcelAddin.Services;

public class AuthService
{
    private const string RegKeyPath = @"SOFTWARE\ZaiExcelAddin";
    private const string ApiUrl = "https://api.z.ai/api/paas/v4/chat/completions";
    private string? _cachedApiKey;

    public void SaveApiKey(string key)
    {
        key = SanitizeKey(key);
        using var reg = Registry.CurrentUser.CreateSubKey(RegKeyPath);
        reg.SetValue("ApiKey", key);
        _cachedApiKey = key;
        AddIn.Logger.Info("API key saved");
    }

    public string LoadApiKey()
    {
        if (_cachedApiKey != null) return _cachedApiKey;
        try
        {
            using var reg = Registry.CurrentUser.OpenSubKey(RegKeyPath);
            _cachedApiKey = reg?.GetValue("ApiKey")?.ToString() ?? "";
        }
        catch { _cachedApiKey = ""; }
        return _cachedApiKey;
    }

    public void ClearApiKey()
    {
        try
        {
            using var reg = Registry.CurrentUser.OpenSubKey(RegKeyPath, true);
            reg?.DeleteValue("ApiKey", false);
        }
        catch { }
        _cachedApiKey = null;
        AddIn.Logger.Info("API key cleared");
    }

    public bool IsLoggedIn() => !string.IsNullOrEmpty(LoadApiKey());

    public void SaveModel(string model)
    {
        using var reg = Registry.CurrentUser.CreateSubKey(RegKeyPath);
        reg.SetValue("Model", model);
    }

    public string LoadModel()
    {
        try
        {
            using var reg = Registry.CurrentUser.OpenSubKey(RegKeyPath);
            return reg?.GetValue("Model")?.ToString() ?? "glm-4.5-air";
        }
        catch { return "glm-4.5-air"; }
    }

    /// <summary>Strip non-ASCII characters from API key (copy-paste can include invisible Unicode).</summary>
    private static string SanitizeKey(string key)
    {
        return new string(key.Where(c => c >= 0x20 && c <= 0x7E).ToArray());
    }

    public bool ValidateApiKey(string key)
    {
        key = SanitizeKey(key);
        AddIn.Logger.Info("Validating API key...");
        try
        {
            using var client = new HttpClient();
            client.DefaultRequestHeaders.Add("Authorization", $"Bearer {key}");
            var body = new StringContent(
                """{"model":"glm-4.5-air","messages":[{"role":"user","content":"Hi"}],"max_tokens":5}""",
                System.Text.Encoding.UTF8, "application/json");
            var response = client.PostAsync(ApiUrl, body).Result;
            AddIn.Logger.Debug($"Validation: HTTP {(int)response.StatusCode}");
            return response.StatusCode != System.Net.HttpStatusCode.Unauthorized
                && response.StatusCode != System.Net.HttpStatusCode.Forbidden;
        }
        catch (Exception ex)
        {
            AddIn.Logger.Error($"Validation failed: {ex.Message}");
            return false;
        }
    }

    public void ShowLogin()
    {
        var i18n = AddIn.I18n;
        var current = LoadApiKey();

        var dlg = new UI.WpfLoginDialog();
        dlg.SetCurrentKey(current);

        if (dlg.ShowDialog() != true) return;

        var key = dlg.ApiKey?.Trim() ?? "";
        if (string.IsNullOrEmpty(key)) return;

        if (ValidateApiKey(key))
        {
            SaveApiKey(key);
            MessageBox.Show(i18n.T("auth.success"), "Z.AI",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        else
        {
            var result = MessageBox.Show(i18n.T("auth.failed"), "Z.AI",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes) SaveApiKey(key);
        }
    }

    public void ShowLogout()
    {
        var i18n = AddIn.I18n;
        if (!IsLoggedIn())
        {
            MessageBox.Show(i18n.T("auth.not_logged"), "Z.AI");
            return;
        }
        var result = MessageBox.Show(i18n.T("auth.confirm_logout"), "Z.AI",
            MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        if (result == DialogResult.Yes)
        {
            ClearApiKey();
            MessageBox.Show(i18n.T("auth.logged_out"), "Z.AI");
        }
    }
}
