using System.Windows;

namespace ZaiExcelAddin.UI;

public partial class WpfSelectDialog : Window
{
    public string SelectedValue { get; private set; } = "";
    public string SelectedKey { get; private set; } = "";

    private readonly Dictionary<string, string>? _keyToDisplay;

    public WpfSelectDialog(string title, string prompt, string[] displayItems,
        string defaultItem, Dictionary<string, string>? keyToDisplay = null)
    {
        InitializeComponent();
        _keyToDisplay = keyToDisplay;

        lblTitle.Text = title;
        Title = title;
        lblPrompt.Text = prompt;
        btnCancel.Content = AddIn.I18n?.T("select.cancel") ?? "Cancel";

        foreach (var item in displayItems)
            cboItems.Items.Add(item);

        // Select default
        for (int i = 0; i < displayItems.Length; i++)
        {
            if (displayItems[i].Contains(defaultItem))
            {
                cboItems.SelectedIndex = i;
                break;
            }
        }
        if (cboItems.SelectedIndex < 0 && displayItems.Length > 0)
            cboItems.SelectedIndex = 0;
    }

    private void OnOkClick(object sender, RoutedEventArgs e)
    {
        SelectedValue = cboItems.SelectedItem?.ToString() ?? "";

        if (_keyToDisplay != null)
        {
            foreach (var kv in _keyToDisplay)
            {
                if (SelectedValue.Contains(kv.Key, StringComparison.OrdinalIgnoreCase))
                {
                    SelectedKey = kv.Key;
                    break;
                }
            }
        }

        DialogResult = true;
    }

    private void OnCancelClick(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
    }
}
