using System.Windows;
using WordTableProcessor.Services;

namespace WordTableProcessor.Views;

public partial class MainWindow : Window
{
    private readonly SettingsService _settingsService;

    public MainWindow()
    {
        InitializeComponent();
        _settingsService = new SettingsService();
    }

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
        var settings = _settingsService.Load();
        
        if (settings.WindowWidth > 0 && settings.WindowHeight > 0)
        {
            Width = settings.WindowWidth;
            Height = settings.WindowHeight;
        }

        if (settings.WindowLeft >= 0 && settings.WindowTop >= 0)
        {
            Left = settings.WindowLeft;
            Top = settings.WindowTop;
        }

        if (settings.IsMaximized)
        {
            WindowState = WindowState.Maximized;
        }
    }

    private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
    {
        _settingsService.SaveWindowPosition(
            Left, Top, Width, Height, 
            WindowState == WindowState.Maximized);
    }

    private void Navigation_Checked(object sender, RoutedEventArgs e)
    {
        if (!IsLoaded) return;

        ProcessorView.Visibility = NavProcessor.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
        CsvGeneratorView.Visibility = NavCsv.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
        HistoryView.Visibility = NavHistory.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;

        if (NavHistory.IsChecked == true)
        {
            HistoryView.LoadRecentFiles();
        }
    }

    private void ClassicUIButton_Click(object sender, RoutedEventArgs e)
    {
        MessageBox.Show(
            "Классический интерфейс находится в резервной копии (OldUI).\nДля его активации требуется дополнительная настройка.",
            "Информация",
            MessageBoxButton.OK,
            MessageBoxImage.Information);
    }
}
