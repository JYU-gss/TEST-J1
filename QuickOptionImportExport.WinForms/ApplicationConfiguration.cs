using System.Windows.Forms;

namespace QuickOptionImportExport;

internal static class ApplicationConfiguration
{
    public static void Initialize()
    {
        Application.SetHighDpiMode(HighDpiMode.SystemAware);
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
    }
}
