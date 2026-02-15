using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace ZaiExcelAddin.UI;

[ComVisible(true)]
public class ChatPaneHost : UserControl
{
    private readonly ElementHost _host;
    private readonly ChatPanel _chatPanel;

    public ChatPaneHost()
    {
        _host = new ElementHost { Dock = DockStyle.Fill };
        _chatPanel = new ChatPanel(this);
        _host.Child = _chatPanel;
        Controls.Add(_host);
    }

    public ChatPanel ChatPanel => _chatPanel;
}
