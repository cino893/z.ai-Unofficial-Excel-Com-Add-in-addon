using System.Windows;
using System.Windows.Controls;

namespace ZaiExcelAddin.UI;

public partial class WpfLoginDialog : Window
{
    public string ApiKey => chkShow.IsChecked == true
        ? txtKeyVisible.Text
        : txtKey.Password;

    public WpfLoginDialog()
    {
        InitializeComponent();

        var i18n = AddIn.I18n;
        lblTitle.Text = i18n.T("auth.login_title");
        lblPrompt.Text = i18n.T("auth.prompt");
        chkShow.Content = i18n.T("auth.show_key");
        btnCancel.Content = i18n.T("select.cancel");
    }

    public void SetCurrentKey(string key)
    {
        txtKey.Password = key;
        txtKeyVisible.Text = key;
    }

    private void OnShowKeyChanged(object sender, RoutedEventArgs e)
    {
        if (chkShow.IsChecked == true)
        {
            txtKeyVisible.Text = txtKey.Password;
            txtKey.Visibility = Visibility.Collapsed;
            txtKeyVisible.Visibility = Visibility.Visible;
            txtKeyVisible.Focus();
        }
        else
        {
            txtKey.Password = txtKeyVisible.Text;
            txtKeyVisible.Visibility = Visibility.Collapsed;
            txtKey.Visibility = Visibility.Visible;
            txtKey.Focus();
        }
    }

    private void OnOkClick(object sender, RoutedEventArgs e)
    {
        DialogResult = true;
    }

    private void OnCancelClick(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
    }
}
