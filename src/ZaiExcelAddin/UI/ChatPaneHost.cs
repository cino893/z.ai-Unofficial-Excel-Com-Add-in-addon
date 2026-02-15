using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace ZaiExcelAddin.UI;

[ComVisible(true)]
[ProgId("ZaiExcelAddin.ChatPaneHost")]
[Guid("F7A3B2C1-4D5E-6F78-9A0B-C1D2E3F4A5B6")]
[ClassInterface(ClassInterfaceType.AutoDispatch)]
public class ChatPaneHost : UserControl
{
    private ElementHost? _host;
    private ChatPanel? _chatPanel;

    public ChatPaneHost()
    {
        // Minimal constructor for COM/ActiveX compatibility
        BackColor = System.Drawing.Color.FromArgb(240, 242, 245);
        Dock = DockStyle.Fill;
    }

    protected override void OnLoad(EventArgs e)
    {
        base.OnLoad(e);
        try
        {
            // Ensure WPF Application exists
            if (System.Windows.Application.Current == null)
                new System.Windows.Application { ShutdownMode = System.Windows.ShutdownMode.OnExplicitShutdown };

            _host = new ElementHost { Dock = DockStyle.Fill };
            _chatPanel = new ChatPanel(this);
            _host.Child = _chatPanel;
            Controls.Add(_host);
        }
        catch (Exception ex)
        {
            AddIn.Logger?.Error($"ChatPaneHost load error: {ex.Message}");
            var lbl = new Label
            {
                Text = $"Failed to initialize chat:\n{ex.Message}",
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                ForeColor = System.Drawing.Color.DarkRed,
                Padding = new Padding(20)
            };
            Controls.Add(lbl);
        }
    }

    public ChatPanel? ChatPanel => _chatPanel;
}
