using System.IO;
using System.Windows;
using System.Windows.Threading;
using WordTableProcessor.Services;

namespace WordTableProcessor;

public partial class App : Application
{
    private readonly SettingsService _settingsService = new();

    private void Application_Startup(object sender, StartupEventArgs e)
    {
        var settings = _settingsService.Load();
        settings.UseClassicUI = false;
        _settingsService.Save(settings);

        var mainWindow = new Views.MainWindow();
        mainWindow.Show();
    }

    private void Application_DispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        try
        {
            string logPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "WordTableProcessor", "crash.log");

            var logDir = Path.GetDirectoryName(logPath);
            if (!string.IsNullOrEmpty(logDir) && !Directory.Exists(logDir))
            {
                Directory.CreateDirectory(logDir);
            }

            var logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {e.Exception}\n\n";
            File.AppendAllText(logPath, logEntry);
        }
        catch
        {
            // Ignore logging errors
        }

        MessageBox.Show(
            $"Произошла непредвиденная ошибка:\n\n{e.Exception.Message}\n\nПриложение будет закрыто.",
            "Ошибка",
            MessageBoxButton.OK,
            MessageBoxImage.Error);

        e.Handled = true;
        Shutdown(1);
    }
}
