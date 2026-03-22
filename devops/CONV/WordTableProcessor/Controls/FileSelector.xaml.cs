using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;

namespace WordTableProcessor.Controls;

public partial class FileSelector : UserControl
{
    public static readonly DependencyProperty LabelProperty =
        DependencyProperty.Register(nameof(Label), typeof(string), typeof(FileSelector),
            new PropertyMetadata("Файл:"));

    public static readonly DependencyProperty FilePathProperty =
        DependencyProperty.Register(nameof(FilePath), typeof(string), typeof(FileSelector),
            new FrameworkPropertyMetadata(string.Empty, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault));

    public static readonly DependencyProperty FilterProperty =
        DependencyProperty.Register(nameof(Filter), typeof(string), typeof(FileSelector),
            new PropertyMetadata("All Files|*.*"));

    public static readonly DependencyProperty DialogTitleProperty =
        DependencyProperty.Register(nameof(DialogTitle), typeof(string), typeof(FileSelector),
            new PropertyMetadata("Выберите файл"));

    public static readonly DependencyProperty ShowClearButtonProperty =
        DependencyProperty.Register(nameof(ShowClearButton), typeof(bool), typeof(FileSelector),
            new PropertyMetadata(true, OnShowClearButtonChanged));

    public string Label
    {
        get => (string)GetValue(LabelProperty);
        set => SetValue(LabelProperty, value);
    }

    public string FilePath
    {
        get => (string)GetValue(FilePathProperty);
        set => SetValue(FilePathProperty, value);
    }

    public string Filter
    {
        get => (string)GetValue(FilterProperty);
        set => SetValue(FilterProperty, value);
    }

    public string DialogTitle
    {
        get => (string)GetValue(DialogTitleProperty);
        set => SetValue(DialogTitleProperty, value);
    }

    public bool ShowClearButton
    {
        get => (bool)GetValue(ShowClearButtonProperty);
        set => SetValue(ShowClearButtonProperty, value);
    }

    public event EventHandler<string>? FileSelected;
    public event EventHandler? FileCleared;

    public FileSelector()
    {
        InitializeComponent();
    }

    private static void OnShowClearButtonChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is FileSelector selector)
        {
            selector.ClearButton.Visibility = (bool)e.NewValue ? Visibility.Visible : Visibility.Collapsed;
        }
    }

    private void BrowseButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = Filter,
            Title = DialogTitle
        };

        if (!string.IsNullOrEmpty(FilePath) && System.IO.File.Exists(FilePath))
        {
            dialog.InitialDirectory = System.IO.Path.GetDirectoryName(FilePath);
        }

        if (dialog.ShowDialog() == true)
        {
            FilePath = dialog.FileName;
            FileSelected?.Invoke(this, dialog.FileName);
        }
    }

    private void ClearButton_Click(object sender, RoutedEventArgs e)
    {
        FilePath = string.Empty;
        FileCleared?.Invoke(this, EventArgs.Empty);
    }
}
