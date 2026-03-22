using System.Text;
using System.Windows;
using System.Windows.Controls;

namespace WordTableProcessor.Controls;

public partial class ProgressPanel : UserControl
{
    public static readonly DependencyProperty ProgressProperty =
        DependencyProperty.Register(nameof(Progress), typeof(double), typeof(ProgressPanel),
            new PropertyMetadata(0.0, OnProgressChanged));

    public static readonly DependencyProperty StatusTextProperty =
        DependencyProperty.Register(nameof(StatusText), typeof(string), typeof(ProgressPanel),
            new PropertyMetadata("Готово"));

    public static readonly DependencyProperty LogTextProperty =
        DependencyProperty.Register(nameof(LogText), typeof(string), typeof(ProgressPanel),
            new FrameworkPropertyMetadata(string.Empty, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, OnLogTextChanged));

    public static readonly DependencyProperty IsIndeterminateProperty =
        DependencyProperty.Register(nameof(IsIndeterminate), typeof(bool), typeof(ProgressPanel),
            new PropertyMetadata(false, OnIsIndeterminateChanged));

    public double Progress
    {
        get => (double)GetValue(ProgressProperty);
        set => SetValue(ProgressProperty, value);
    }

    public string StatusText
    {
        get => (string)GetValue(StatusTextProperty);
        set => SetValue(StatusTextProperty, value);
    }

    public string LogText
    {
        get => (string)GetValue(LogTextProperty);
        set => SetValue(LogTextProperty, value);
    }

    public bool IsIndeterminate
    {
        get => (bool)GetValue(IsIndeterminateProperty);
        set => SetValue(IsIndeterminateProperty, value);
    }

    public string PercentText => $"{Progress:F0}%";

    private readonly StringBuilder _logBuilder = new();
    private bool _autoScroll = true;

    public ProgressPanel()
    {
        InitializeComponent();
        LogTextBox.TextChanged += (s, e) =>
        {
            if (_autoScroll)
            {
                LogTextBox.ScrollToEnd();
            }
        };
    }

    private static void OnProgressChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is ProgressPanel panel)
        {
            panel.UpdatePercentText();
            panel.UpdateStatusFromProgress();
        }
    }

    private static void OnIsIndeterminateChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is ProgressPanel panel)
        {
            panel.ProgressBar.IsIndeterminate = (bool)e.NewValue;
        }
    }

    private static void OnLogTextChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        // Handled by binding
    }

    private void UpdatePercentText()
    {
        var textBlock = GetTemplateChild("PercentText") as TextBlock;
        if (textBlock != null)
        {
            textBlock.Text = PercentText;
        }
    }

    private void UpdateStatusFromProgress()
    {
        if (Progress > 0 && Progress < 100)
        {
            StatusText = "Обработка...";
        }
        else if (Progress >= 100)
        {
            StatusText = "Завершено";
        }
        else
        {
            StatusText = "Готово";
        }
    }

    public void AppendLog(string message)
    {
        Dispatcher.Invoke(() =>
        {
            var timestamp = DateTime.Now.ToString("HH:mm:ss");
            _logBuilder.AppendLine($"[{timestamp}] {message}");
            LogText = _logBuilder.ToString();
        });
    }

    public void ClearLog()
    {
        Dispatcher.Invoke(() =>
        {
            _logBuilder.Clear();
            LogText = string.Empty;
        });
    }

    public void SetProgress(double value, string? status = null)
    {
        Dispatcher.Invoke(() =>
        {
            Progress = Math.Max(0, Math.Min(100, value));
            if (!string.IsNullOrEmpty(status))
            {
                StatusText = status;
            }
        });
    }

    public void SetIndeterminate(bool value)
    {
        Dispatcher.Invoke(() =>
        {
            IsIndeterminate = value;
            if (value)
            {
                StatusText = "Обработка...";
            }
        });
    }

    public void SetCompleted(string? message = null)
    {
        Dispatcher.Invoke(() =>
        {
            Progress = 100;
            StatusText = message ?? "Завершено";
            if (!string.IsNullOrEmpty(message))
            {
                AppendLog(message);
            }
        });
    }

    public void SetError(string errorMessage)
    {
        Dispatcher.Invoke(() =>
        {
            Progress = 0;
            StatusText = "Ошибка";
            AppendLog($"ОШИБКА: {errorMessage}");
        });
    }
}
