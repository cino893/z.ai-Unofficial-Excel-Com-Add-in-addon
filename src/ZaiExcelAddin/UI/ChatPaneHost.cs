using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace ZaiExcelAddin.UI;

[ComVisible(true)]
public interface IChatPaneHost { }

[ComVisible(true)]
[ComDefaultInterface(typeof(IChatPaneHost))]
[ProgId("ZaiExcelAddin.ChatPaneHost")]
[Guid("F7A3B2C1-4D5E-6F78-9A0B-C1D2E3F4A5B6")]
public class ChatPaneHost : UserControl, IChatPaneHost
{
    private ElementHost? _host;
    private ChatPanel? _chatPanel;

    public ChatPaneHost()
    {
        // Keep constructor absolutely minimal for COM/ActiveX instantiation
    }

    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        InitWpfContent();
    }

    private void InitWpfContent()
    {
        if (_host != null) return; // already initialized
        try
        {
            if (System.Windows.Application.Current == null)
                new System.Windows.Application { ShutdownMode = System.Windows.ShutdownMode.OnExplicitShutdown };

            _host = new ElementHost { Dock = DockStyle.Fill };
            _chatPanel = new ChatPanel(this);
            _host.Child = _chatPanel;
            Controls.Add(_host);
        }
        catch (Exception ex)
        {
            AddIn.Logger?.Error($"ChatPaneHost init error: {ex.Message}");
            Controls.Add(new Label
            {
                Text = $"Failed to initialize chat:\n{ex.Message}",
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                ForeColor = System.Drawing.Color.DarkRed,
                Padding = new Padding(20)
            });
        }
    }

    public ChatPanel? ChatPanel => _chatPanel;
}
