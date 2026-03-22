using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WordTableProcessor.Models;
using WordTableProcessor.Services;

namespace WordTableProcessor.Views;

public partial class SettingsView : UserControl
{
    private readonly SettingsService _settingsService;

    public SettingsView()
    {
        InitializeComponent();
        _settingsService = new SettingsService();
        LoadRecentFiles();
    }

    public void LoadRecentFiles()
    {
        var settings = _settingsService.Load();
        RecentFilesList.ItemsSource = settings.RecentFiles;
        
        NoRecentFilesText.Visibility = settings.RecentFiles.Count == 0 
            ? Visibility.Visible 
            : Visibility.Collapsed;
    }

    private void RecentFile_Click(object sender, MouseButtonEventArgs e)
    {
        if (sender is Border border && border.DataContext is RecentFile file)
        {
            if (File.Exists(file.FilePath))
            {
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = file.FilePath,
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Не удалось открыть файл: {ex.Message}", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            else
            {
                MessageBox.Show("Файл не найден. Возможно, он был перемещен или удален.",
                    "Файл не найден", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
    }

    private void RemoveRecentFile_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button button && button.DataContext is RecentFile file)
        {
            var settings = _settingsService.Load();
            settings.RecentFiles.RemoveAll(f => f.FilePath == file.FilePath);
            _settingsService.Save(settings);
            LoadRecentFiles();
        }
        e.Handled = true;
    }

    private void ClearHistoryButton_Click(object sender, RoutedEventArgs e)
    {
        var result = MessageBox.Show("Очистить историю недавних файлов?", 
            "Подтверждение", MessageBoxButton.YesNo, MessageBoxImage.Question);
        
        if (result == MessageBoxResult.Yes)
        {
            _settingsService.ClearRecentFiles();
            LoadRecentFiles();
        }
    }
}
