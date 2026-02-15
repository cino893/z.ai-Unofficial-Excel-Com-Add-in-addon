using System.Drawing;
using System.Windows.Forms;

namespace ZaiExcelAddin.UI;

public class SelectDialog : Form
{
    private readonly ComboBox _combo;
    public string SelectedValue => _combo.SelectedItem?.ToString() ?? "";
    public string SelectedKey { get; private set; } = "";

    private readonly Dictionary<string, string>? _keyMap;

    public SelectDialog(string title, string prompt, string[] displayItems, string defaultItem,
        Dictionary<string, string>? keyToDisplay = null)
    {
        _keyMap = keyToDisplay;

        Text = title;
        Size = new Size(360, 200);
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        BackColor = Color.White;
        Font = new Font("Segoe UI", 9.5f);

        var lbl = new Label
        {
            Text = prompt,
            Location = new Point(16, 16),
            Size = new Size(310, 40),
            AutoSize = false
        };

        _combo = new ComboBox
        {
            Location = new Point(16, 60),
            Size = new Size(310, 28),
            DropDownStyle = ComboBoxStyle.DropDownList,
            FlatStyle = FlatStyle.Flat
        };
        _combo.Items.AddRange(displayItems);

        // Select default
        for (int i = 0; i < displayItems.Length; i++)
        {
            if (displayItems[i].Contains(defaultItem))
            {
                _combo.SelectedIndex = i;
                break;
            }
        }
        if (_combo.SelectedIndex < 0 && displayItems.Length > 0)
            _combo.SelectedIndex = 0;

        var btnOk = new Button
        {
            Text = "OK",
            Location = new Point(165, 110),
            Size = new Size(80, 32),
            DialogResult = DialogResult.OK,
            BackColor = Color.FromArgb(102, 126, 234),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        btnOk.FlatAppearance.BorderSize = 0;

        var btnCancel = new Button
        {
            Text = AddIn.I18n?.T("select.cancel") ?? "Cancel",
            Location = new Point(250, 110),
            Size = new Size(80, 32),
            DialogResult = DialogResult.Cancel,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };

        Controls.AddRange(new Control[] { lbl, _combo, btnOk, btnCancel });
        AcceptButton = btnOk;
        CancelButton = btnCancel;
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        base.OnFormClosing(e);
        if (DialogResult == DialogResult.OK && _keyMap != null)
        {
            var selected = SelectedValue;
            foreach (var kv in _keyMap)
            {
                if (selected.Contains(kv.Key))
                {
                    SelectedKey = kv.Key;
                    break;
                }
            }
        }
    }
}
