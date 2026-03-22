using System.Collections.ObjectModel;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordTableProcessor.Controls;
using WordTableProcessor.Models;
using WordTableProcessor.Services;

namespace WordTableProcessor.Views;

public partial class DocumentProcessorView : UserControl
{
    private readonly SettingsService _settingsService;
    private string? _wordFilePath;
    private string? _csvFilePath;
    private SmetaDocumentType? _selectedDocumentType;

    public DocumentProcessorView()
    {
        InitializeComponent();
        _settingsService = new SettingsService();
        InitializeDocumentTypes();
        UpdateFilePanels();
    }

    private void InitializeDocumentTypes()
    {
        var items = new ObservableCollection<ComboBoxItem>();
        
        foreach (DocumentCategory category in Enum.GetValues<DocumentCategory>())
        {
            var headerItem = new ComboBoxItem
            {
                Content = SmetaDocumentType.GetCategoryDisplayName(category),
                IsEnabled = false,
                FontWeight = FontWeights.Bold,
                Padding = new Thickness(10, 8, 10, 8)
            };
            items.Add(headerItem);

            var types = SmetaDocumentType.GetTypesByCategory(category);
            foreach (var type in types)
            {
                var item = new ComboBoxItem
                {
                    Content = $"    {type.DisplayName}",
                    Tag = type,
                    Padding = new Thickness(10, 4, 10, 4)
                };
                items.Add(item);
            }
        }

        DocumentTypeComboBox.ItemsSource = items;
        if (items.Count > 1)
        {
            DocumentTypeComboBox.SelectedIndex = 1;
        }
    }

    private void DocumentTypeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (DocumentTypeComboBox.SelectedItem is ComboBoxItem item && item.Tag is SmetaDocumentType docType)
        {
            _selectedDocumentType = docType;
            ProgressControl.AppendLog($"Выбран тип документа: {docType.DisplayName}");
        }
    }

    private void DropZone_FileDropped(object sender, string filePath)
    {
        ProcessFile(filePath);
    }

    private void DropZone_FileSelected(object sender, string filePath)
    {
        ProcessFile(filePath);
    }

    private void ProcessFile(string filePath)
    {
        if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            return;

        var ext = Path.GetExtension(filePath).ToLowerInvariant();
        
        _settingsService.AddRecentFile(filePath, 
            ext == ".docx" ? RecentFileType.WordDocument : 
            ext == ".csv" ? RecentFileType.CsvFile : RecentFileType.ExcelFile);

        if (ext == ".docx")
        {
            _wordFilePath = filePath;
            ProgressControl.AppendLog($"Выбран Word документ: {Path.GetFileName(filePath)}");
        }
        else if (ext == ".csv")
        {
            _csvFilePath = filePath;
            ProgressControl.AppendLog($"Выбран CSV файл: {Path.GetFileName(filePath)}");
        }
        else
        {
            MessageBox.Show($"Поддерживаемые форматы: .docx, .csv", "Предупреждение", 
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        UpdateFilePanels();
        UpdateButtonStates();
    }

    private void UpdateFilePanels()
    {
        WordFileName.Text = string.IsNullOrEmpty(_wordFilePath) 
            ? "Word документ не выбран" 
            : Path.GetFileName(_wordFilePath);
        WordFileName.ToolTip = _wordFilePath;
        WordFilePanel.Opacity = string.IsNullOrEmpty(_wordFilePath) ? 0.5 : 1;

        CsvFileName.Text = string.IsNullOrEmpty(_csvFilePath) 
            ? "CSV файл не выбран" 
            : Path.GetFileName(_csvFilePath);
        CsvFileName.ToolTip = _csvFilePath;
        CsvFilePanel.Opacity = string.IsNullOrEmpty(_csvFilePath) ? 0.5 : 1;
    }

    private void UpdateButtonStates()
    {
        UpdatePricesButton.IsEnabled = !string.IsNullOrEmpty(_wordFilePath) && 
                                       !string.IsNullOrEmpty(_csvFilePath);
        CleanTableButton.IsEnabled = !string.IsNullOrEmpty(_wordFilePath);
    }

    private void ClearWordButton_Click(object sender, RoutedEventArgs e)
    {
        _wordFilePath = null;
        UpdateFilePanels();
        UpdateButtonStates();
        ProgressControl.AppendLog("Word документ удален");
    }

    private void ClearCsvButton_Click(object sender, RoutedEventArgs e)
    {
        _csvFilePath = null;
        UpdateFilePanels();
        UpdateButtonStates();
        ProgressControl.AppendLog("CSV файл удален");
    }

    private async void CleanTableButton_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_wordFilePath))
        {
            MessageBox.Show("Выберите Word документ!", "Ошибка", 
                MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        if (_selectedDocumentType == null)
        {
            MessageBox.Show("Выберите тип документа!", "Ошибка", 
                MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        ProgressControl.ClearLog();
        ProgressControl.SetIndeterminate(true);
        ProgressControl.AppendLog("Начало очистки таблицы...");

        try
        {
            await Task.Run(() => CleanTable(_wordFilePath, _selectedDocumentType));
            ProgressControl.SetCompleted("Очистка таблицы завершена");
            MessageBox.Show("Очистка таблицы завершена.", "Успех", 
                MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            ProgressControl.SetError(ex.Message);
            MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка", 
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            ProgressControl.SetIndeterminate(false);
        }
    }

    private void CleanTable(string filePath, SmetaDocumentType docType)
    {
        Dispatcher.Invoke(() => ProgressControl.AppendLog("Открытие документа..."));

        using var doc = WordprocessingDocument.Open(filePath, true);
        var tables = doc.MainDocumentPart?.Document?.Body?.Elements<Table>().ToList();

        if (tables == null || tables.Count < docType.TableCount)
        {
            Dispatcher.Invoke(() => ProgressControl.SetError("В документе недостаточно таблиц!"));
            return;
        }

        int totalRows = tables.Take(docType.TableCount).Sum(t => t.Elements<TableRow>().Count());
        int processedRows = 0;

        Dispatcher.Invoke(() => ProgressControl.AppendLog($"Найдено таблиц: {tables.Count}, строк: {totalRows}"));

        for (int t = 0; t < docType.TableCount; t++)
        {
            var table = tables[t];
            var rows = table.Elements<TableRow>().ToList();

            foreach (var row in rows)
            {
                var cells = row.Elements<TableCell>().ToList();
                if (cells.Count > 0)
                {
                    string code = cells[0].InnerText.Trim();
                    if (Regex.IsMatch(code, docType.CodePattern))
                    {
                        if (cells.Count > 2)
                        {
                            cells[1].RemoveAllChildren();
                            cells[1].Append(new Paragraph(new Run(new Text("!"))));
                            cells[2].RemoveAllChildren();
                            cells[2].Append(new Paragraph(new Run(new Text("!"))));
                        }
                    }
                }
                processedRows++;
                if (processedRows % 1000 == 0 || processedRows == totalRows)
                {
                    double progress = (double)processedRows / totalRows * 100;
                    Dispatcher.Invoke(() => ProgressControl.SetProgress(progress));
                }
            }
        }

        doc.Save();
        Dispatcher.Invoke(() => ProgressControl.AppendLog($"Обработано строк: {processedRows}"));
    }

    private async void UpdatePricesButton_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_wordFilePath) || string.IsNullOrEmpty(_csvFilePath))
        {
            MessageBox.Show("Выберите Word документ и CSV файл!", "Ошибка", 
                MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        if (_selectedDocumentType == null)
        {
            MessageBox.Show("Выберите тип документа!", "Ошибка", 
                MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        ProgressControl.ClearLog();
        ProgressControl.SetIndeterminate(true);
        ProgressControl.AppendLog("Чтение CSV файла...");

        try
        {
            var csvData = await Task.Run(() => ReadCsvFile(_csvFilePath!));
            ProgressControl.AppendLog($"Загружено записей из CSV: {csvData.Count}");

            ProgressControl.AppendLog("Начало обновления цен...");
            await Task.Run(() => UpdatePrices(_wordFilePath!, _csvFilePath!, csvData, _selectedDocumentType));

            ProgressControl.SetCompleted("Обновление цен завершено");
            MessageBox.Show("Обновление цен завершено.", "Успех", 
                MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            ProgressControl.SetError(ex.Message);
            MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка", 
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            ProgressControl.SetIndeterminate(false);
        }
    }

    private List<string[]> ReadCsvFile(string filePath)
    {
        var csvData = new List<string[]>();
        Encoding encoding;

        var firstLines = File.ReadLines(filePath, Encoding.UTF8).Take(3).ToList();
        if (firstLines.Any(l => l.Any(c => c > 127 && !IsValidUtf8(l))))
        {
            encoding = Encoding.GetEncoding(1251);
        }
        else
        {
            encoding = Encoding.UTF8;
        }

        using var reader = new StreamReader(filePath, encoding);
        while (!reader.EndOfStream)
        {
            var line = reader.ReadLine();
            if (line != null)
            {
                var cells = line.Split(';');
                if (cells.Length >= 4)
                {
                    csvData.Add(cells);
                }
            }
        }

        return csvData;
    }

    private bool IsValidUtf8(string line)
    {
        try
        {
            var bytes = Encoding.UTF8.GetBytes(line);
            Encoding.UTF8.GetString(bytes);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private void UpdatePrices(string wordPath, string csvPath, List<string[]> csvData, SmetaDocumentType docType)
    {
        Dispatcher.Invoke(() => ProgressControl.AppendLog("Открытие Word документа..."));

        using var doc = WordprocessingDocument.Open(wordPath, true);
        var tables = doc.MainDocumentPart?.Document?.Body?.Elements<Table>().ToList();

        if (tables == null || tables.Count < docType.TableCount)
        {
            Dispatcher.Invoke(() => ProgressControl.SetError("В документе недостаточно таблиц!"));
            return;
        }

        int totalRows = tables.Take(docType.TableCount).Sum(t => t.Elements<TableRow>().Count());
        int processedRows = 0;
        int updatedRows = 0;
        int csvIndex = 0;

        Dispatcher.Invoke(() => ProgressControl.AppendLog($"Обработано: {processedRows}/{totalRows}"));

        for (int t = 0; t < docType.TableCount; t++)
        {
            var table = tables[t];
            var rows = table.Elements<TableRow>().ToList();

            foreach (var row in rows)
            {
                var cells = row.Elements<TableCell>().ToList();
                if (cells.Count > 0)
                {
                    string code = cells[0].InnerText.Trim();
                    if (Regex.IsMatch(code, docType.CodePattern))
                    {
                        while (csvIndex < csvData.Count)
                        {
                            string csvCode = csvData[csvIndex][1].Trim();
                            string csvTableNum = csvData[csvIndex][0].Trim();

                            if (code.StartsWith(csvCode) && csvTableNum == (t + 1).ToString())
                            {
                                if (cells.Count > 2)
                                {
                                    try
                                    {
                                        double price1 = double.Parse(csvData[csvIndex][2], 
                                            System.Globalization.CultureInfo.InvariantCulture);
                                        double price2 = double.Parse(csvData[csvIndex][3], 
                                            System.Globalization.CultureInfo.InvariantCulture);

                                        string formattedPrice1 = price1.ToString("F2", 
                                            System.Globalization.CultureInfo.GetCultureInfo("ru-RU"));
                                        string formattedPrice2 = price2.ToString("F2", 
                                            System.Globalization.CultureInfo.GetCultureInfo("ru-RU"));

                                        cells[1].RemoveAllChildren();
                                        cells[1].Append(new Paragraph(new Run(new Text(formattedPrice1))));
                                        cells[2].RemoveAllChildren();
                                        cells[2].Append(new Paragraph(new Run(new Text(formattedPrice2))));
                                        
                                        updatedRows++;
                                    }
                                    catch (FormatException)
                                    {
                                        Dispatcher.Invoke(() => ProgressControl.AppendLog(
                                            $"Ошибка формата для кода {csvCode}"));
                                    }
                                }
                                csvIndex++;
                                break;
                            }
                            csvIndex++;
                        }
                    }
                }
                processedRows++;
                if (processedRows % 1000 == 0 || processedRows == totalRows)
                {
                    double progress = (double)processedRows / totalRows * 100;
                    Dispatcher.Invoke(() =>
                    {
                        ProgressControl.SetProgress(progress);
                        ProgressControl.AppendLog($"Обработано: {processedRows}/{totalRows}, обновлено: {updatedRows}");
                    });
                }
            }
        }

        doc.Save();
        Dispatcher.Invoke(() => ProgressControl.AppendLog($"Завершено. Обновлено строк: {updatedRows}"));
    }
}
