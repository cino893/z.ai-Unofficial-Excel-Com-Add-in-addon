using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ExcelDna.Integration;
using ZaiExcelAddin.Models;
using ZaiExcelAddin.Services;

namespace ZaiExcelAddin.UI;

public partial class ChatPanel : System.Windows.Controls.UserControl
{
    private readonly ObservableCollection<ChatMessage> _messages = new();
    private readonly System.Windows.Forms.Control _host;
    private bool _isProcessing;

    public ChatPanel(System.Windows.Forms.Control host)
    {
        InitializeComponent();
        _host = host;
        lstMessages.ItemsSource = _messages;

        txtInput.TextChanged += (s, e) =>
        {
            txtPlaceholder.Visibility = string.IsNullOrEmpty(txtInput.Text)
                ? Visibility.Visible : Visibility.Collapsed;
        };

        Loaded += OnLoaded;
    }

    // Design-time support
    public ChatPanel() : this(new System.Windows.Forms.Control()) { }

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        UpdateLabels();
        ShowWelcome();
    }

    public void RefreshLabels()
    {
        UpdateLabels();
    }

    private void UpdateLabels()
    {
        var i18n = AddIn.I18n;
        lblSubtitle.Text = i18n.T("chat.title");
        txtPlaceholder.Text = i18n.T("chat.send") + "...";
        lblStatus.Text = i18n.T("chat.ready");
        btnNewConv.Content = "🗨 " + i18n.T("chat.new");
        btnClear.Content = "🗑 " + i18n.T("chat.clear");
    }

    private void ShowWelcome()
    {
        _messages.Clear();
        AddIn.Conversation.Init();
        AddMessage("assistant", AddIn.I18n.T("chat.welcome"));
    }

    private void AddMessage(string role, string content)
    {
        _messages.Add(new ChatMessage { Role = role, Content = content });
        Dispatcher.BeginInvoke(() =>
        {
            scrollMessages.ScrollToEnd();
        }, System.Windows.Threading.DispatcherPriority.Background);
    }

    private void SetProcessing(bool processing)
    {
        _isProcessing = processing;
        btnSend.Visibility = processing ? Visibility.Collapsed : Visibility.Visible;
        btnStop.Visibility = processing ? Visibility.Visible : Visibility.Collapsed;
        txtInput.IsEnabled = !processing;
        typingIndicator.Visibility = processing ? Visibility.Visible : Visibility.Collapsed;
        lblStatus.Text = processing
            ? AddIn.I18n.T("chat.processing")
            : AddIn.I18n.TFormat("chat.ready_count", _messages.Count);

        // Show/hide Continue button based on stop reason
        if (!processing)
            UpdateContinueButton();
    }

    private void UpdateContinueButton()
    {
        var canContinue = AddIn.Conversation.CanContinue;
        btnContinue.Visibility = canContinue ? Visibility.Visible : Visibility.Collapsed;
        if (canContinue)
            btnContinue.Content = "▶ " + AddIn.I18n.T("chat.continue");
    }

    private async void OnSendClick(object sender, RoutedEventArgs e)
    {
        await SendMessage();
    }

    private void OnStopClick(object sender, RoutedEventArgs e)
    {
        AddIn.Conversation.Cancel();
    }

    private async void OnContinueClick(object sender, RoutedEventArgs e)
    {
        if (_isProcessing) return;

        btnContinue.Visibility = Visibility.Collapsed;
        AddMessage("info", AddIn.I18n.T("chat.continuing"));

        await RunAgentTask(() => AddIn.Conversation.Continue(ExecuteToolOnMainThread));
    }

    private void OnInputKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key == Key.Enter && !Keyboard.Modifiers.HasFlag(ModifierKeys.Shift))
        {
            e.Handled = true;
            _ = SendMessage();
        }
    }

    private async Task SendMessage()
    {
        var text = txtInput.Text?.Trim();
        if (string.IsNullOrEmpty(text) || _isProcessing) return;

        if (!AddIn.Auth.IsLoggedIn())
        {
            AddMessage("info", AddIn.I18n.T("auth.need_login"));
            AddIn.Auth.ShowLogin();
            return;
        }

        txtInput.Text = "";
        AddMessage("user", text);

        await RunAgentTask(() => AddIn.Conversation.SendMessage(text, ExecuteToolOnMainThread));
    }

    /// <summary>
    /// Wraps agent execution with Excel performance optimization.
    /// Disables ScreenUpdating, EnableEvents, and sets manual Calculation
    /// for the duration of tool execution (10-43x faster bulk operations).
    /// </summary>
    private async Task RunAgentTask(Func<string> agentAction)
    {
        SetProcessing(true);
        UndoService.CaptureSnapshot();

        bool oldScreenUpdating = true;
        bool oldEnableEvents = true;
        int oldCalculation = -4105; // xlCalculationAutomatic

        // Suppress Excel UI during agent execution for performance
        _host.Invoke(() =>
        {
            try
            {
                dynamic app = ExcelDnaUtil.Application;
                oldScreenUpdating = app.ScreenUpdating;
                oldEnableEvents = app.EnableEvents;
                oldCalculation = (int)app.Calculation;
                app.ScreenUpdating = false;
                app.EnableEvents = false;
                app.Calculation = -4135; // xlCalculationManual
            }
            catch { /* ignore if Excel not ready */ }
        });

        try
        {
            var response = await Task.Run(agentAction);

            if (!string.IsNullOrEmpty(response))
                AddMessage("assistant", response);

            UndoService.RegisterUndo();
        }
        catch (Exception ex)
        {
            AddIn.Logger.Error($"Agent error: {ex.Message}");
            AddMessage("info", $"Error: {ex.Message}");
        }
        finally
        {
            // Restore Excel state — always, even on error
            _host.Invoke(() =>
            {
                try
                {
                    dynamic app = ExcelDnaUtil.Application;
                    app.Calculation = oldCalculation;
                    app.EnableEvents = oldEnableEvents;
                    app.ScreenUpdating = oldScreenUpdating;
                }
                catch { /* ignore */ }
            });
            SetProcessing(false);
        }
    }

    private string ExecuteToolOnMainThread(string name, string args)
    {
        string result = "";
        _host.Invoke(() =>
        {
            result = AddIn.Skills.Execute(name, args);
        });
        return result;
    }

    private void OnNewChat(object sender, RoutedEventArgs e)
    {
        _messages.Clear();
        AddIn.Conversation.Init();
        btnContinue.Visibility = Visibility.Collapsed;
        AddMessage("info", AddIn.I18n.T("chat.new_started"));
        lblStatus.Text = AddIn.I18n.T("chat.ready");
    }

    private void OnClearChat(object sender, RoutedEventArgs e)
    {
        _messages.Clear();
        lblStatus.Text = AddIn.I18n.T("chat.ready");
    }
}

// Template selector for message bubbles
public class MessageTemplateSelector : DataTemplateSelector
{
    public DataTemplate? UserTemplate { get; set; }
    public DataTemplate? AssistantTemplate { get; set; }
    public DataTemplate? SystemTemplate { get; set; }

    public override DataTemplate? SelectTemplate(object item, DependencyObject container)
    {
        if (item is ChatMessage msg)
        {
            return msg.Role switch
            {
                "user" => UserTemplate,
                "assistant" => AssistantTemplate,
                _ => SystemTemplate
            };
        }
        return base.SelectTemplate(item, container);
    }
}
