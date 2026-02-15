using System.Net.Http;
using Microsoft.Win32;

namespace ZaiExcelAddin.Services;

public class AuthService
{
    private const string RegKeyPath = @"SOFTWARE\ZaiExcelAddin";
    private const string ApiUrl = "https://api.z.ai/api/paas/v4/chat/completions";
    private string? _cachedApiKey;

    public void SaveApiKey(string key)
    {
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
            return reg?.GetValue("Model")?.ToString() ?? "glm-4-plus";
        }
        catch { return "glm-4-plus"; }
    }

    public bool ValidateApiKey(string key)
    {
        AddIn.Logger.Info("Validating API key...");
        try
        {
            using var client = new HttpClient();
            client.DefaultRequestHeaders.Add("Authorization", $"Bearer {key}");
            var body = new StringContent(
                """{"model":"glm-4-plus","messages":[{"role":"user","content":"Hi"}],"max_tokens":5}""",
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
        var prompt = i18n.T("auth.prompt");
        if (!string.IsNullOrEmpty(current))
            prompt += $"\n\n{i18n.T("auth.current_key")}{current[..Math.Min(8, current.Length)]}...";

        var key = Microsoft.VisualBasic.Interaction.InputBox(prompt, i18n.T("auth.login_title"));
        if (string.IsNullOrWhiteSpace(key))
        {
            if (string.IsNullOrEmpty(current))
                System.Windows.Forms.MessageBox.Show(
                    i18n.T("auth.cancelled"), "Z.AI",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
            return;
        }

        key = key.Trim();
        if (ValidateApiKey(key))
        {
            SaveApiKey(key);
            System.Windows.Forms.MessageBox.Show(
                i18n.T("auth.success"), "Z.AI",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Information);
        }
        else
        {
            var result = System.Windows.Forms.MessageBox.Show(
                i18n.T("auth.failed"), "Z.AI",
                System.Windows.Forms.MessageBoxButtons.YesNo,
                System.Windows.Forms.MessageBoxIcon.Question);
            if (result == System.Windows.Forms.DialogResult.Yes)
                SaveApiKey(key);
        }
    }

    public void ShowLogout()
    {
        var i18n = AddIn.I18n;
        if (!IsLoggedIn())
        {
            System.Windows.Forms.MessageBox.Show(i18n.T("auth.not_logged"), "Z.AI");
            return;
        }
        var result = System.Windows.Forms.MessageBox.Show(
            i18n.T("auth.confirm_logout"), "Z.AI",
            System.Windows.Forms.MessageBoxButtons.YesNo,
            System.Windows.Forms.MessageBoxIcon.Question);
        if (result == System.Windows.Forms.DialogResult.Yes)
        {
            ClearApiKey();
            System.Windows.Forms.MessageBox.Show(i18n.T("auth.logged_out"), "Z.AI");
        }
    }
}
