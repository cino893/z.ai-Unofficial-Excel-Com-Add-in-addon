using ExcelDna.Integration;
using ZaiExcelAddin.Services;

namespace ZaiExcelAddin;

public class AddIn : IExcelAddIn
{
    public static DebugLogger Logger { get; private set; } = null!;
    public static AuthService Auth { get; private set; } = null!;
    public static I18nService I18n { get; private set; } = null!;
    public static ZaiApiService Api { get; private set; } = null!;
    public static ExcelSkillService Skills { get; private set; } = null!;
    public static ConversationService Conversation { get; private set; } = null!;

    public void AutoOpen()
    {
        try
        {
            // Initialize WPF infrastructure for hosting in CTP
            if (System.Windows.Application.Current == null)
            {
                var app = new System.Windows.Application
                {
                    ShutdownMode = System.Windows.ShutdownMode.OnExplicitShutdown
                };
            }

            Logger = new DebugLogger();
            I18n = new I18nService();
            Auth = new AuthService();
            Api = new ZaiApiService();
            Skills = new ExcelSkillService();
            Conversation = new ConversationService();

            Logger.Info("Z.AI Add-in v2.0 loaded (.NET COM)");
        }
        catch (Exception ex)
        {
            System.Windows.Forms.MessageBox.Show(
                $"Z.AI Add-in failed to load: {ex.Message}",
                "Z.AI Error", System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Error);
        }
    }

    public void AutoClose()
    {
        Logger?.Info("Z.AI Add-in unloaded");
    }

    public static dynamic ExcelApp => ExcelDnaUtil.Application;
}
