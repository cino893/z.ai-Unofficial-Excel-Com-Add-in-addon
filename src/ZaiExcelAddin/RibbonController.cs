using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;

namespace ZaiExcelAddin;

[ComVisible(true)]
public class RibbonController : ExcelRibbon
{
    private CustomTaskPane? _chatPane;
    private IRibbonUI? _ribbon;

    public override string GetCustomUI(string ribbonID)
    {
        return @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>
  <ribbon>
    <tabs>
      <tab id='zaiTab' label='Z.AI'>
        <group id='grpMain' getLabel='GetGroupLabel'>
          <button id='btnChat' getLabel='GetLabel' size='large'
                  onAction='OnToggleChat' imageMso='BlogOpenExisting'
                  screentip='Toggle AI Chat Panel'/>
          <separator id='sep1'/>
          <button id='btnLogin' getLabel='GetLabel' size='normal'
                  onAction='OnLogin' imageMso='ProtectForm'
                  getEnabled='GetLoginEnabled'/>
          <button id='btnLogout' getLabel='GetLabel' size='normal'
                  onAction='OnLogout' imageMso='ReviewDeleteComment'
                  getEnabled='GetLogoutEnabled'/>
          <separator id='sep2'/>
          <button id='btnModel' getLabel='GetLabel' size='normal'
                  onAction='OnSelectModel' imageMso='ServerSettings'
                  getEnabled='GetLogoutEnabled'/>
        </group>
        <group id='grpInfo' getLabel='GetGroupLabel'>
          <labelControl id='lblStatus' getLabel='GetStatusLabel'/>
          <button id='btnAddTokens' getLabel='GetLabel' size='normal'
                  onAction='OnAddTokens' imageMso='CurrencyFormatGallery'/>
        </group>
        <group id='grpTools' getLabel='GetGroupLabel'>
          <button id='btnLang' getLabel='GetLabel' size='normal'
                  onAction='OnLanguage' imageMso='ReviewTranslate'/>
          <button id='btnLog' getLabel='GetLabel' size='normal'
                  onAction='OnViewLog' imageMso='VisualBasicModule'/>
          <button id='btnAbout' getLabel='GetLabel' size='normal'
                  onAction='OnAbout' imageMso='Info'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
    }

    public void OnRibbonLoad(IRibbonUI ribbonUI) => _ribbon = ribbonUI;
    public void RefreshRibbon() => _ribbon?.Invalidate();

    // ═══ Dynamic labels ═══
    public string GetLabel(IRibbonControl control)
    {
        var t = AddIn.I18n;
        return control.Id switch
        {
            "btnChat"      => t.T("menu.chat"),
            "btnLogin"     => t.T("menu.login"),
            "btnLogout"    => t.T("menu.logout"),
            "btnModel"     => t.T("menu.model"),
            "btnAddTokens" => t.T("menu.add_tokens"),
            "btnLang"      => t.T("menu.language"),
            "btnLog"       => t.T("menu.viewlog"),
            "btnAbout"     => t.T("menu.about"),
            _ => control.Id
        };
    }

    public string GetGroupLabel(IRibbonControl control)
    {
        var t = AddIn.I18n;
        return control.Id switch
        {
            "grpMain"  => t.T("ribbon.group_main"),
            "grpInfo"  => t.T("ribbon.group_info"),
            "grpTools" => t.T("ribbon.group_tools"),
            _ => control.Id
        };
    }

    public string GetStatusLabel(IRibbonControl control)
    {
        if (!AddIn.Auth.IsLoggedIn())
            return AddIn.I18n.T("ribbon.not_logged");

        var balance = AddIn.Api.GetBalance();
        return AddIn.I18n.T("ribbon.logged_in") + " | " + balance;
    }

    // ═══ Enabled states ═══
    public bool GetLoginEnabled(IRibbonControl control) => !AddIn.Auth.IsLoggedIn();
    public bool GetLogoutEnabled(IRibbonControl control) => AddIn.Auth.IsLoggedIn();

    // ═══ Actions ═══
    public void OnToggleChat(IRibbonControl control)
    {
        try
        {
            if (_chatPane == null)
            {
                _chatPane = CustomTaskPaneFactory.CreateCustomTaskPane(
                    typeof(UI.ChatPaneHost),
                    "Z.AI - " + AddIn.I18n.T("chat.title"));
                _chatPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                _chatPane.Width = 420;
                _chatPane.VisibleStateChange += _ => _ribbon?.InvalidateControl("btnChat");
            }
            _chatPane.Visible = !_chatPane.Visible;
        }
        catch (Exception ex)
        {
            AddIn.Logger.Error($"Toggle chat error: {ex.Message}");
            System.Windows.Forms.MessageBox.Show(
                $"Error opening chat panel:\n{ex.Message}\n\n" +
                AddIn.I18n.T("error.ctp_hint"),
                "Z.AI", System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Error);
        }
    }

    public void OnLogin(IRibbonControl control)
    {
        try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            { FileName = "https://z.ai/manage-apikey/apikey-list", UseShellExecute = true }); }
        catch { }

        AddIn.Auth.ShowLogin();
        AddIn.Api.InvalidateBalanceCache();
        _ribbon?.Invalidate();
    }

    public void OnLogout(IRibbonControl control)
    {
        AddIn.Auth.ShowLogout();
        AddIn.Api.InvalidateBalanceCache();
        _ribbon?.Invalidate();
    }

    public void OnSelectModel(IRibbonControl control)
    {
        var current = AddIn.Auth.LoadModel();
        var models = AddIn.Api.GetAvailableModels();
        var items = models.Select(m => m).ToArray();
        var keyMap = models.ToDictionary(m => m, m => m);

        var dlg = new UI.WpfSelectDialog(
            AddIn.I18n.T("model.title"),
            AddIn.I18n.T("model.prompt"),
            items, current, keyMap);

        if (dlg.ShowDialog() == true && !string.IsNullOrEmpty(dlg.SelectedKey))
        {
            AddIn.Auth.SaveModel(dlg.SelectedKey);
            _ribbon?.Invalidate();
        }
    }

    public void OnLanguage(IRibbonControl control)
    {
        var langs = Services.I18nService.SupportedLanguages;
        var current = AddIn.I18n.CurrentLanguage;
        var items = langs.Select(l => $"{l.Key}  —  {l.Value}").ToArray();
        var keyMap = langs.ToDictionary(l => l.Key, l => l.Value);

        var dlg = new UI.WpfSelectDialog(
            AddIn.I18n.T("lang.title"),
            AddIn.I18n.T("lang.select_prompt"),
            items, current, keyMap);

        if (dlg.ShowDialog() == true && !string.IsNullOrEmpty(dlg.SelectedKey))
        {
            AddIn.I18n.SetLanguage(dlg.SelectedKey);
            _ribbon?.Invalidate();
            // Refresh chat panel labels
            RefreshChatPanel();
        }
    }

    public void OnAddTokens(IRibbonControl control)
    {
        try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            { FileName = "https://z.ai/manage-apikey/billing", UseShellExecute = true }); }
        catch { }
    }

    public void OnViewLog(IRibbonControl control) => AddIn.Logger.ViewLog();

    public void OnAbout(IRibbonControl control)
    {
        System.Windows.Forms.MessageBox.Show(
            AddIn.I18n.T("about.text"),
            AddIn.I18n.T("about.title"),
            System.Windows.Forms.MessageBoxButtons.OK,
            System.Windows.Forms.MessageBoxIcon.Information);
    }

    private void RefreshChatPanel()
    {
        try
        {
            var host = _chatPane?.ContentControl as UI.ChatPaneHost;
            host?.ChatPanel?.Dispatcher.Invoke(() => host.ChatPanel.RefreshLabels());
        }
        catch { }
    }
}
