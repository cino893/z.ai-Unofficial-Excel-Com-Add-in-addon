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

        using var dlg = new Form
        {
            Text = i18n.T("auth.login_title"),
            Size = new Size(420, 250),
            StartPosition = FormStartPosition.CenterScreen,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            MaximizeBox = false, MinimizeBox = false,
            BackColor = Color.White,
            Font = new Font("Segoe UI", 9.5f)
        };

        var lbl = new Label
        {
            Text = i18n.T("auth.prompt"),
            Location = new Point(16, 16), Size = new Size(370, 40)
        };
        var txtKey = new TextBox
        {
            Location = new Point(16, 62), Size = new Size(370, 28),
            UseSystemPasswordChar = true,
            Text = current
        };
        var chkShow = new CheckBox
        {
            Text = i18n.T("auth.show_key"), AutoSize = true,
            Location = new Point(16, 96)
        };
        chkShow.CheckedChanged += (_, _) => txtKey.UseSystemPasswordChar = !chkShow.Checked;

        var btnOpenSite = new Button
        {
            Text = "\U0001f310 " + i18n.T("auth.open_site"),
            Location = new Point(16, 135), Size = new Size(180, 32),
            FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand,
            BackColor = Color.FromArgb(240, 240, 240)
        };
        btnOpenSite.Click += (_, _) =>
        {
            try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                { FileName = "https://open.z.ai/", UseShellExecute = true }); }
            catch { }
        };

        var btnOk = new Button
        {
            Text = "OK",
            Location = new Point(220, 135), Size = new Size(80, 32),
            DialogResult = DialogResult.OK,
            BackColor = Color.FromArgb(102, 126, 234),
            ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand
        };
        btnOk.FlatAppearance.BorderSize = 0;

        var btnCancel = new Button
        {
            Text = i18n.T("select.cancel") ?? "Cancel",
            Location = new Point(306, 135), Size = new Size(80, 32),
            DialogResult = DialogResult.Cancel,
            FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand
        };

        dlg.Controls.AddRange(new Control[] { lbl, txtKey, chkShow, btnOpenSite, btnOk, btnCancel });
        dlg.AcceptButton = btnOk;
        dlg.CancelButton = btnCancel;

        if (dlg.ShowDialog() != DialogResult.OK) return;

        var key = txtKey.Text.Trim();
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
