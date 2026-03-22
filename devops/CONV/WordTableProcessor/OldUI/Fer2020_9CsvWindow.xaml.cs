using ClosedXML.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using WordTableProcessor.Models;

namespace WordTableProcessor
{
    public partial class Fer2020_9CsvWindow : Window
    {
        private IndexesSettings _indexesSettings = new IndexesSettings();

        public Fer2020_9CsvWindow()
        {
            InitializeComponent();
            DataContext = _indexesSettings;
        }

        private (Window window, ProgressBar progressBar, TextBox logTextBox, Button closeButton) CreateProgressWindow()
        {
            var window = new Window
            {
                Title = "Формирование CSV",
                Height = 300,
                Width = 400,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                ResizeMode = ResizeMode.NoResize
            };

            var grid = new Grid { Margin = new Thickness(10) };
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var progressBar = new ProgressBar { Height = 20, Margin = new Thickness(0, 0, 0, 10) };
            var logTextBox = new TextBox { IsReadOnly = true, VerticalScrollBarVisibility = ScrollBarVisibility.Auto, Margin = new Thickness(0, 0, 0, 10) };
            var closeButton = new Button { Name = "CloseButton", Content = "Закрыть", Width = 100, HorizontalAlignment = HorizontalAlignment.Right, IsEnabled = false };
            closeButton.Click += CloseProgressButton_Click;

            Grid.SetRow(progressBar, 0);
            Grid.SetRow(logTextBox, 1);
            Grid.SetRow(closeButton, 2);

            grid.Children.Add(progressBar);
            grid.Children.Add(logTextBox);
            grid.Children.Add(closeButton);

            window.Content = grid;
            return (window, progressBar, logTextBox, closeButton);
        }

        private void IndexesBrowseSourceButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog { Filter = "Excel Files|*.xlsx" };
            if (openFileDialog.ShowDialog() == true)
            {
                IndexesSourceFileTextBox.Text = openFileDialog.FileName;
            }
        }

        private void IndexesClearSourceButton_Click(object sender, RoutedEventArgs e)
        {
            IndexesSourceFileTextBox.Text = string.Empty;
        }

        private void IndexesBrowseSaveButton_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "CSV Files|*.csv",
                FileName = $"Индексы ФЕР-2020(0) {DateTime.Now.ToString("MM.yyyy")}.csv"
            };
            if (saveFileDialog.ShowDialog() == true)
            {
                IndexesSaveFileTextBox.Text = saveFileDialog.FileName;
            }
        }

        private void IndexesConfigureColumnsButton_Click(object sender, RoutedEventArgs e)
        {
            if (IndexesSettingsExpander != null)
            {
                IndexesSettingsExpander.IsExpanded = !IndexesSettingsExpander.IsExpanded;
            }
        }

        private async void IndexesGenerateButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(IndexesSourceFileTextBox.Text) || string.IsNullOrEmpty(IndexesSaveFileTextBox.Text))
            {
                MessageBox.Show("Выберите исходный и выходной файлы.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            string sourcePath = IndexesSourceFileTextBox.Text;
            string savePath = IndexesSaveFileTextBox.Text;

            try
            {
                var (progressWindow, progressBar, logTextBox, closeButton) = CreateProgressWindow();
                if (progressBar == null || logTextBox == null || closeButton == null)
                {
                    MessageBox.Show("Ошибка инициализации окна прогресса.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                progressWindow.Closing += (s, args) => { progressWindow = null; };
                progressWindow.Show();
                var startTime = DateTime.Now;
                await Task.Run(() => GenerateIndexesCsv(progressWindow, progressBar, logTextBox, sourcePath, savePath));
                var duration = DateTime.Now - startTime;
                if (progressWindow?.IsVisible == true)
                {
                    closeButton.IsEnabled = true;
                    UpdateProgress(progressBar, logTextBox, 100, $"Сформировано строк: {_indexesSettings.ProcessedRows}. Время выполнения: {duration.TotalSeconds:F2} сек.");
                }
                if (progressWindow?.IsVisible == true) MessageBox.Show("CSV успешно сформирован.", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                LogError(ex.ToString());
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void GenerateIndexesCsv(Window progressWindow, ProgressBar progressBar, TextBox logTextBox, string sourcePath, string savePath)
        {
            if (progressWindow == null) return;

            try
            {
                using var workbook = new XLWorkbook(sourcePath);
                var worksheet = workbook.Worksheet(1);
                var rows = worksheet.RowsUsed();
                _indexesSettings.ProcessedRows = 0;
                var csvLines = new List<string> { "ID_TABLE;CODE;EM;MR" }; // Заголовок

                // Первый проход: TABLE_ID 1 и 2 с региональными данными
                int tableId = 1;
                bool switchTable = false;
                foreach (var row in rows)
                {
                    if (progressWindow == null) break;
                    _indexesSettings.ProcessedRows++;

                    string code = row.Cell(_indexesSettings.Code).GetString() ?? "";
                    if (code == _indexesSettings.Table1End) { switchTable = true; }
                    else if (code == _indexesSettings.Table2End) { tableId = 2; switchTable = true; }

                    if (switchTable && code != _indexesSettings.Table1End && code != _indexesSettings.Table2End)
                    {
                        tableId = tableId == 1 ? 2 : tableId;
                        switchTable = false;
                    }

                    string emStr = row.Cell(_indexesSettings.Regional.Inem).GetString() ?? "1.000";
                    string mrStr = row.Cell(_indexesSettings.Regional.Mat).GetString() ?? "1.000";
                    double em = double.TryParse(emStr, out double emValue) ? emValue : 1.000;
                    double mr = double.TryParse(mrStr, out double mrValue) ? mrValue : 1.000;
                    csvLines.Add($"{tableId};{code};{FormatNumber(em)};{FormatNumber(mr)}");
                }

                // Второй проход: TABLE_ID 3 и 4 с отраслевыми данными
                tableId = 3;
                switchTable = false;
                foreach (var row in rows)
                {
                    if (progressWindow == null) break;
                    _indexesSettings.ProcessedRows++;

                    string code = row.Cell(_indexesSettings.Code).GetString() ?? "";
                    if (code == _indexesSettings.Table3End) { switchTable = true; }
                    else if (code == _indexesSettings.Table4End) { tableId = 4; switchTable = true; }

                    if (switchTable && code != _indexesSettings.Table3End && code != _indexesSettings.Table4End)
                    {
                        tableId = tableId == 3 ? 4 : tableId;
                        switchTable = false;
                    }

                    string emStr = row.Cell(_indexesSettings.Industry.Inem).GetString() ?? "1.000";
                    string mrStr = row.Cell(_indexesSettings.Industry.Mat).GetString() ?? "1.000";
                    double em = double.TryParse(emStr, out double emValue) ? emValue : 1.000;
                    double mr = double.TryParse(mrStr, out double mrValue) ? mrValue : 1.000;
                    csvLines.Add($"{tableId};{code};{FormatNumber(em)};{FormatNumber(mr)}");
                }

                if (File.Exists(savePath)) File.Delete(savePath);
                File.WriteAllLines(savePath, csvLines, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                LogError(ex.ToString());
                if (progressWindow != null && progressWindow.Dispatcher?.CheckAccess() == true)
                {
                    UpdateProgress(progressBar, logTextBox, 0, $"Ошибка: {ex.Message}");
                }
                else if (progressWindow?.Dispatcher != null)
                {
                    progressWindow.Dispatcher.Invoke(() => UpdateProgress(progressBar, logTextBox, 0, $"Ошибка: {ex.Message}"));
                }
            }
        }

        private void FsscBrowseSourceButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog { Filter = "Excel Files|*.xlsx" };
            if (openFileDialog.ShowDialog() == true)
            {
                FsscSourceFileTextBox.Text = openFileDialog.FileName;
            }
        }

        private void FsscClearSourceButton_Click(object sender, RoutedEventArgs e)
        {
            FsscSourceFileTextBox.Text = string.Empty;
        }

        private void FsscBrowseSaveButton_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "CSV Files|*.csv",
                FileName = $"Материалы ФЕР-2020(0) {DateTime.Now.ToString("MM.yyyy")}.csv"
            };
            if (saveFileDialog.ShowDialog() == true)
            {
                FsscSaveFileTextBox.Text = saveFileDialog.FileName;
            }
        }

        private void FsscConfigureColumnsButton_Click(object sender, RoutedEventArgs e)
        {
            // Заглушка. Реализуйте логику настройки колонок, если требуется.
        }

        private async void FsscGenerateButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(FsscSourceFileTextBox.Text) || string.IsNullOrEmpty(FsscSaveFileTextBox.Text))
            {
                MessageBox.Show("Выберите исходный и выходной файлы.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            string sourcePath = FsscSourceFileTextBox.Text;
            string savePath = FsscSaveFileTextBox.Text;

            try
            {
                var (progressWindow, progressBar, logTextBox, closeButton) = CreateProgressWindow();
                if (progressBar == null || logTextBox == null || closeButton == null)
                {
                    MessageBox.Show("Ошибка инициализации окна прогресса.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                progressWindow.Closing += (s, args) => { progressWindow = null; };
                progressWindow.Show();
                await Task.Run(() => GenerateFsscCsv(progressWindow, progressBar, logTextBox, sourcePath, savePath));
                if (progressWindow?.IsVisible == true) closeButton.IsEnabled = true;
                if (progressWindow?.IsVisible == true) MessageBox.Show("CSV успешно сформирован.", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                LogError(ex.ToString());
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void GenerateFsscCsv(Window progressWindow, ProgressBar progressBar, TextBox logTextBox, string sourcePath, string savePath)
        {
            if (progressWindow == null) return;

            try
            {
                using var workbook = new XLWorkbook(sourcePath);
                var worksheet = workbook.Worksheet(1);
                var rows = worksheet.RowsUsed();
                int totalRows = rows.Count();
                int processedRows = 0;
                var csvLines = new List<string> { "TABLE_ID;CODE;NAME;UNIT;PRICE_OPT;PRICE;INDX" };

                foreach (var row in rows)
                {
                    if (row.RowNumber() == 1) continue;
                    if (progressWindow == null) break;
                    processedRows++;
                    if (progressWindow.Dispatcher?.CheckAccess() == true)
                    {
                        UpdateProgress(progressBar, logTextBox, (double)processedRows / totalRows * 100, $"Обработано строк: {processedRows}/{totalRows}");
                    }
                    else if (progressWindow.Dispatcher != null)
                    {
                        progressWindow.Dispatcher.Invoke(() => UpdateProgress(progressBar, logTextBox, (double)processedRows / totalRows * 100, $"Обработано строк: {processedRows}/{totalRows}"));
                    }

                    string code = row.Cell(1).GetString() ?? "";
                    string name = row.Cell(2).GetString() ?? "";
                    string unit = row.Cell(3).GetString() ?? "";
                    double priceOpt = double.TryParse(row.Cell(4).GetString(), out double p) ? p : 0.0;
                    double price = double.TryParse(row.Cell(5).GetString(), out double pr) ? pr : 0.0;
                    double indx = double.TryParse(row.Cell(6).GetString(), out double i) ? i : 0.0;

                    string formattedPriceOpt = FormatNumber(priceOpt);
                    string formattedPrice = FormatNumber(price);
                    string formattedIndx = FormatNumber(indx);
                    csvLines.Add($"1;{code};{name};{unit};{formattedPriceOpt};{formattedPrice};{formattedIndx}");
                }

                if (File.Exists(savePath)) File.Delete(savePath);
                File.WriteAllLines(savePath, csvLines, Encoding.UTF8);
                if (progressWindow != null && progressWindow.Dispatcher?.CheckAccess() == true)
                {
                    UpdateProgress(progressBar, logTextBox, 100, $"Сформировано строк: {processedRows}");
                }
                else if (progressWindow?.Dispatcher != null)
                {
                    progressWindow.Dispatcher.Invoke(() => UpdateProgress(progressBar, logTextBox, 100, $"Сформировано строк: {processedRows}"));
                }
            }
            catch (Exception ex)
            {
                LogError(ex.ToString());
                if (progressWindow != null && progressWindow.Dispatcher?.CheckAccess() == true)
                {
                    UpdateProgress(progressBar, logTextBox, 0, $"Ошибка: {ex.Message}");
                }
                else if (progressWindow?.Dispatcher != null)
                {
                    progressWindow.Dispatcher.Invoke(() => UpdateProgress(progressBar, logTextBox, 0, $"Ошибка: {ex.Message}"));
                }
            }
        }

        private void FsemBrowseSourceButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog { Filter = "Excel Files|*.xlsx" };
            if (openFileDialog.ShowDialog() == true)
            {
                FsemSourceFileTextBox.Text = openFileDialog.FileName;
            }
        }

        private void FsemClearSourceButton_Click(object sender, RoutedEventArgs e)
        {
            FsemSourceFileTextBox.Text = string.Empty;
        }

        private void FsemBrowseSaveButton_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "CSV Files|*.csv",
                FileName = $"Механизмы ФЕР-2020(0) {DateTime.Now.ToString("MM.yyyy")}.csv"
            };
            if (saveFileDialog.ShowDialog() == true)
            {
                FsemSaveFileTextBox.Text = saveFileDialog.FileName;
            }
        }

        private void FsemConfigureColumnsButton_Click(object sender, RoutedEventArgs e)
        {
            // Заглушка. Реализуйте логику настройки колонок, если требуется.
        }

        private async void FsemGenerateButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(FsemSourceFileTextBox.Text) || string.IsNullOrEmpty(FsemSaveFileTextBox.Text))
            {
                MessageBox.Show("Выберите исходный и выходной файлы.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            string sourcePath = FsemSourceFileTextBox.Text;
            string savePath = FsemSaveFileTextBox.Text;

            try
            {
                var (progressWindow, progressBar, logTextBox, closeButton) = CreateProgressWindow();
                if (progressBar == null || logTextBox == null || closeButton == null)
                {
                    MessageBox.Show("Ошибка инициализации окна прогресса.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                progressWindow.Closing += (s, args) => { progressWindow = null; };
                progressWindow.Show();
                await Task.Run(() => GenerateFsemCsv(progressWindow, progressBar, logTextBox, sourcePath, savePath));
                if (progressWindow?.IsVisible == true) closeButton.IsEnabled = true;
                if (progressWindow?.IsVisible == true) MessageBox.Show("CSV успешно сформирован.", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                LogError(ex.ToString());
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void GenerateFsemCsv(Window progressWindow, ProgressBar progressBar, TextBox logTextBox, string sourcePath, string savePath)
        {
            if (progressWindow == null) return;

            try
            {
                using var workbook = new XLWorkbook(sourcePath);
                var worksheet = workbook.Worksheet(1);
                var rows = worksheet.RowsUsed();
                int totalRows = rows.Count();
                int processedRows = 0;
                var csvLines = new List<string> { "TABLE_ID;CODE;NAME;UNIT;PRICE;OTM;INDX" };

                foreach (var row in rows)
                {
                    if (row.RowNumber() == 1) continue;
                    if (progressWindow == null) break;
                    processedRows++;
                    if (progressWindow.Dispatcher?.CheckAccess() == true)
                    {
                        UpdateProgress(progressBar, logTextBox, (double)processedRows / totalRows * 100, $"Обработано строк: {processedRows}/{totalRows}");
                    }
                    else if (progressWindow.Dispatcher != null)
                    {
                        progressWindow.Dispatcher.Invoke(() => UpdateProgress(progressBar, logTextBox, (double)processedRows / totalRows * 100, $"Обработано строк: {processedRows}/{totalRows}"));
                    }

                    string code = row.Cell(1).GetString() ?? "";
                    string name = row.Cell(2).GetString() ?? "";
                    string unit = row.Cell(3).GetString() ?? "";
                    double price = double.TryParse(row.Cell(4).GetString(), out double p) ? p : 0.0;
                    double otm = double.TryParse(row.Cell(5).GetString(), out double o) ? o : 0.0;
                    double indx = double.TryParse(row.Cell(6).GetString(), out double i) ? i : 0.0;

                    string formattedPrice = FormatNumber(price);
                    string formattedOtm = FormatNumber(otm);
                    string formattedIndx = FormatNumber(indx);
                    csvLines.Add($"1;{code};{name};{unit};{formattedPrice};{formattedOtm};{formattedIndx}");
                }

                if (File.Exists(savePath)) File.Delete(savePath);
                File.WriteAllLines(savePath, csvLines, Encoding.UTF8);
                if (progressWindow != null && progressWindow.Dispatcher?.CheckAccess() == true)
                {
                    UpdateProgress(progressBar, logTextBox, 100, $"Сформировано строк: {processedRows}");
                }
                else if (progressWindow?.Dispatcher != null)
                {
                    progressWindow.Dispatcher.Invoke(() => UpdateProgress(progressBar, logTextBox, 100, $"Сформировано строк: {processedRows}"));
                }
            }
            catch (Exception ex)
            {
                LogError(ex.ToString());
                if (progressWindow != null && progressWindow.Dispatcher?.CheckAccess() == true)
                {
                    UpdateProgress(progressBar, logTextBox, 0, $"Ошибка: {ex.Message}");
                }
                else if (progressWindow?.Dispatcher != null)
                {
                    progressWindow.Dispatcher.Invoke(() => UpdateProgress(progressBar, logTextBox, 0, $"Ошибка: {ex.Message}"));
                }
            }
        }

        private void CloseProgressButton_Click(object sender, RoutedEventArgs e)
        {
            (sender as Button)?.FindAncestor<Window>()?.Close();
        }

        private void UpdateProgress(ProgressBar progressBar, TextBox logTextBox, double value, string logMessage)
        {
            if (progressBar != null)
            {
                progressBar.Value = value;
            }
            if (logTextBox != null)
            {
                logTextBox.AppendText(logMessage + "\n");
                logTextBox.ScrollToEnd();
            }
        }

        private string FormatNumber(double number)
        {
            return number.ToString("N3", new System.Globalization.CultureInfo("ru-RU")
            {
                NumberFormat = { NumberDecimalSeparator = ".", NumberGroupSeparator = " " }
            }).Replace(",", ".");
        }

        private void LogError(string errorMessage)
        {
            string logFilePath = "errors.log";
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            try
            {
                using (StreamWriter writer = new StreamWriter(logFilePath, true))
                {
                    writer.WriteLine($"{timestamp}: {errorMessage}");
                }
            }
            catch (Exception)
            {
                // Игнорируем ошибки записи лога
            }
        }
    }
}