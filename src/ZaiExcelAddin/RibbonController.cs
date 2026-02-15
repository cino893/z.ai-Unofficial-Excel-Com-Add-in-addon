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
        <group id='grpMain' label='Assistant'>
          <button id='btnChat' label='💬 Chat' size='large'
                  onAction='OnToggleChat' imageMso='BlogOpenExisting'
                  screentip='Toggle AI Chat Panel' supertip='Open or close the Z.AI assistant chat panel'/>
          <separator id='sep1'/>
          <button id='btnLogin' label='Login' size='normal'
                  onAction='OnLogin' imageMso='ProtectForm'
                  screentip='Login with API Key'/>
          <button id='btnLogout' label='Logout' size='normal'
                  onAction='OnLogout' imageMso='ReviewDeleteComment'
                  screentip='Remove API Key'/>
          <button id='btnModel' label='Model' size='normal'
                  onAction='OnSelectModel' imageMso='ServerSettings'
                  screentip='Select AI Model'/>
        </group>
        <group id='grpTools' label='Tools'>
          <button id='btnLang' label='Language' size='normal'
                  onAction='OnLanguage' imageMso='ReviewTranslate'
                  screentip='Change UI Language'/>
          <button id='btnLog' label='View Log' size='normal'
                  onAction='OnViewLog' imageMso='VisualBasicModule'
                  screentip='Open Debug Log'/>
          <button id='btnAbout' label='About' size='normal'
                  onAction='OnAbout' imageMso='Info'
                  screentip='About Z.AI Add-in'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
    }

    public void OnRibbonLoad(IRibbonUI ribbonUI)
    {
        _ribbon = ribbonUI;
    }

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
                $"Error opening chat panel:\n{ex.Message}",
                "Z.AI", System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Error);
        }
    }

    public void OnLogin(IRibbonControl control) => AddIn.Auth.ShowLogin();
    public void OnLogout(IRibbonControl control) => AddIn.Auth.ShowLogout();

    public void OnSelectModel(IRibbonControl control)
    {
        var i18n = AddIn.I18n;
        var current = AddIn.Auth.LoadModel();
        var input = Microsoft.VisualBasic.Interaction.InputBox(
            i18n.T("model.prompt") + "\n\n" + i18n.T("model.current") + current,
            i18n.T("model.title"), current);
        if (!string.IsNullOrWhiteSpace(input))
        {
            AddIn.Auth.SaveModel(input.Trim());
            System.Windows.Forms.MessageBox.Show(
                i18n.T("model.changed") + input.Trim(),
                "Z.AI", System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Information);
        }
    }

    public void OnLanguage(IRibbonControl control)
    {
        var langs = Services.I18nService.SupportedLanguages;
        var current = AddIn.I18n.CurrentLanguage;
        var list = string.Join("\n", langs.Select(l => $"  {l.Key} - {l.Value}"));
        var input = Microsoft.VisualBasic.Interaction.InputBox(
            $"Select language / Wybierz język:\n\n{list}\n\nCurrent: {current}",
            "Z.AI - Language", current);
        if (!string.IsNullOrWhiteSpace(input))
        {
            var code = input.Trim().ToLower();
            if (langs.ContainsKey(code))
            {
                AddIn.I18n.SetLanguage(code);
                _ribbon?.Invalidate();
                System.Windows.Forms.MessageBox.Show(
                    AddIn.I18n.T("lang.changed"),
                    "Z.AI", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Information);
            }
        }
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
}
