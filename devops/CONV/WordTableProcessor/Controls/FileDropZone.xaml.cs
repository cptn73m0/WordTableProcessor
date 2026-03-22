using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Win32;

namespace WordTableProcessor.Controls;

public partial class FileDropZone : UserControl
{
    public static readonly DependencyProperty WordFilePathProperty =
        DependencyProperty.Register(nameof(WordFilePath), typeof(string), typeof(FileDropZone),
            new PropertyMetadata(string.Empty, OnFilePathChanged));

    public static readonly DependencyProperty CsvFilePathProperty =
        DependencyProperty.Register(nameof(CsvFilePath), typeof(string), typeof(FileDropZone),
            new PropertyMetadata(string.Empty, OnFilePathChanged));

    public static readonly DependencyProperty AllowCsvProperty =
        DependencyProperty.Register(nameof(AllowCsv), typeof(bool), typeof(FileDropZone),
            new PropertyMetadata(true, OnAllowCsvChanged));

    public string WordFilePath
    {
        get => (string)GetValue(WordFilePathProperty);
        set => SetValue(WordFilePathProperty, value);
    }

    public string CsvFilePath
    {
        get => (string)GetValue(CsvFilePathProperty);
        set => SetValue(CsvFilePathProperty, value);
    }

    public bool AllowCsv
    {
        get => (bool)GetValue(AllowCsvProperty);
        set => SetValue(AllowCsvProperty, value);
    }

    public event EventHandler<string>? FileSelected;
    public event EventHandler<string>? FileDropped;

    public FileDropZone()
    {
        InitializeComponent();
    }

    private static void OnFilePathChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is FileDropZone zone)
        {
            zone.UpdateDropZoneText();
        }
    }

    private static void OnAllowCsvChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is FileDropZone zone)
        {
            zone.SelectCsvButton.Visibility = (bool)e.NewValue ? Visibility.Visible : Visibility.Collapsed;
            zone.UpdateDropZoneText();
        }
    }

    private void UpdateDropZoneText()
    {
        bool hasFiles = !string.IsNullOrEmpty(WordFilePath) || !string.IsNullOrEmpty(CsvFilePath);
        
        if (hasFiles)
        {
            DropText.Text = "Перетащите дополнительные файлы";
        }
        else
        {
            DropText.Text = "Перетащите файлы сюда";
        }
    }

    private void OnDragEnter(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            var files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files != null && files.Any(IsValidFile))
            {
                e.Effects = DragDropEffects.Copy;
                DropZoneBorder.BorderBrush = (SolidColorBrush)FindResource("SystemControlForegroundAccentBrush");
                var gradientBrush = new LinearGradientBrush
                {
                    StartPoint = new System.Windows.Point(0, 0),
                    EndPoint = new System.Windows.Point(0, 1)
                };
                gradientBrush.GradientStops.Add(new GradientStop(Color.FromArgb(30, 0, 120, 215), 0));
                gradientBrush.GradientStops.Add(new GradientStop(Color.FromArgb(10, 0, 120, 215), 1));
                DropZoneBorder.Background = gradientBrush;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }
        e.Handled = true;
    }

    private void OnDragLeave(object sender, DragEventArgs e)
    {
        ResetDropZoneStyle();
    }

    private void OnDragOver(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            var files = (string[])e.Data.GetData(DataFormats.FileDrop);
            e.Effects = files != null && files.Any(IsValidFile) ? DragDropEffects.Copy : DragDropEffects.None;
        }
        e.Handled = true;
    }

    private void OnDrop(object sender, DragEventArgs e)
    {
        ResetDropZoneStyle();
        
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            var files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files != null)
            {
                foreach (var file in files.Where(IsValidFile))
                {
                    FileDropped?.Invoke(this, file);
                }
            }
        }
        e.Handled = true;
    }

    private void ResetDropZoneStyle()
    {
        DropZoneBorder.BorderBrush = (SolidColorBrush)FindResource("SystemControlForegroundBaseMediumBrush");
        var defaultGradient = new LinearGradientBrush
        {
            StartPoint = new System.Windows.Point(0, 0),
            EndPoint = new System.Windows.Point(0, 1)
        };
        defaultGradient.GradientStops.Add(new GradientStop(Color.FromArgb(255, 242, 242, 242), 0));
        defaultGradient.GradientStops.Add(new GradientStop(Color.FromArgb(255, 230, 230, 230), 1));
        DropZoneBorder.Background = defaultGradient;
    }

    private bool IsValidFile(string path)
    {
        if (string.IsNullOrEmpty(path)) return false;
        
        var ext = Path.GetExtension(path).ToLowerInvariant();
        
        if (ext == ".docx" && !string.IsNullOrEmpty(WordFilePath)) return false;
        if ((ext == ".csv" || ext == ".xlsx") && !string.IsNullOrEmpty(CsvFilePath)) return false;
        
        return ext == ".docx" || ext == ".csv" || ext == ".xlsx";
    }

    private void SelectWordButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Word Documents|*.docx",
            Title = "Выберите Word документ"
        };

        if (dialog.ShowDialog() == true)
        {
            WordFilePath = dialog.FileName;
            FileSelected?.Invoke(this, dialog.FileName);
        }
    }

    private void SelectCsvButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "CSV Files|*.csv|Excel Files|*.xlsx|All Files|*.*",
            Title = "Выберите файл данных"
        };

        if (dialog.ShowDialog() == true)
        {
            CsvFilePath = dialog.FileName;
            FileSelected?.Invoke(this, dialog.FileName);
        }
    }

    public void ClearFiles()
    {
        WordFilePath = string.Empty;
        CsvFilePath = string.Empty;
    }
}
