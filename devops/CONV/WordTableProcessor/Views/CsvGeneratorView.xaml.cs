using System.Collections.ObjectModel;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using ClosedXML.Excel;
using WordTableProcessor.Controls;
using WordTableProcessor.Models;
using WordTableProcessor.Services;

namespace WordTableProcessor.Views;

public partial class CsvGeneratorView : UserControl
{
    private readonly SettingsService _settingsService;
    private DocumentCategory _selectedCategory = DocumentCategory.FER2020;
    
    private string? _indexesSource;
    private string? _indexesSave;
    private string? _materialsSource;
    private string? _materialsSave;
    private string? _mechanismsSource;
    private string? _mechanismsSave;

    public CsvGeneratorView()
    {
        InitializeComponent();
        _settingsService = new SettingsService();
        InitializeDocumentTypes();
        LoadSettings();
    }

    private void InitializeDocumentTypes()
    {
        var items = new ObservableCollection<ComboBoxItem>();
        
        foreach (DocumentCategory category in Enum.GetValues<DocumentCategory>())
        {
            var item = new ComboBoxItem
            {
                Content = SmetaDocumentType.GetCategoryDisplayName(category),
                Tag = category
            };
            items.Add(item);
        }

        DocumentTypeComboBox.ItemsSource = items;
        DocumentTypeComboBox.SelectedIndex = 0;
    }

    private void LoadSettings()
    {
        var settings = _settingsService.Load();
        var s = settings.IndexesSettings;
        
        IndexesCodeColumn.Text = s.Code.ToString();
        IndexesTable1End.Text = s.Table1End;
        IndexesTable2End.Text = s.Table2End;
        IndexesTable3End.Text = s.Table3End;
        IndexesTable4End.Text = s.Table4End;
        IndexesRegionalInem.Text = s.Regional.Inem.ToString();
        IndexesRegionalMat.Text = s.Regional.Mat.ToString();
        IndexesIndustryInem.Text = s.Industry.Inem.ToString();
    }

    private void SaveSettings()
    {
        var settings = _settingsService.Load();
        settings.IndexesSettings = new Models.IndexesSettings
        {
            Code = int.TryParse(IndexesCodeColumn.Text, out var code) ? code : 1,
            Table1End = IndexesTable1End.Text,
            Table2End = IndexesTable2End.Text,
            Table3End = IndexesTable3End.Text,
            Table4End = IndexesTable4End.Text,
            Regional = new Models.RegionalSettings
            {
                Inem = int.TryParse(IndexesRegionalInem.Text, out var ri) ? ri : 5,
                Mat = int.TryParse(IndexesRegionalMat.Text, out var rm) ? rm : 6
            },
            Industry = new Models.IndustrySettings
            {
                Inem = int.TryParse(IndexesIndustryInem.Text, out var ii) ? ii : 8,
                Mat = 9
            }
        };
        _settingsService.Save(settings);
    }

    private void DocumentTypeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (DocumentTypeComboBox.SelectedItem is ComboBoxItem item && item.Tag is DocumentCategory category)
        {
            _selectedCategory = category;
            UpdateExpanderHeaders();
            ProgressControl.AppendLog($"Выбрана категория: {SmetaDocumentType.GetCategoryDisplayName(category)}");
        }
    }

    private void UpdateExpanderHeaders()
    {
        switch (_selectedCategory)
        {
            case DocumentCategory.FER2020:
                MaterialsExpander.Header = "Материалы (ФССЦ-2020)";
                MechanismsExpander.Header = "Механизмы (ФСЭМ-2020)";
                break;
            case DocumentCategory.TER2014:
                MaterialsExpander.Header = "Материалы (ТССЦ-2014)";
                MechanismsExpander.Header = "Механизмы (ТСЭМ-2014)";
                break;
            case DocumentCategory.TER2010:
                MaterialsExpander.Header = "Материалы (ТССЦ-2010)";
                MechanismsExpander.Header = "Механизмы (ТСЭМ-2010)";
                break;
        }
    }

    private void Settings_TextChanged(object sender, TextChangedEventArgs e)
    {
        SaveSettings();
    }

    private void IndexesSource_FileSelected(object sender, string filePath)
    {
        _indexesSource = filePath;
        _settingsService.AddRecentFile(filePath, RecentFileType.ExcelFile);
    }

    private void IndexesSave_FileSelected(object sender, string filePath)
    {
        _indexesSave = filePath;
    }

    private void MaterialsSource_FileSelected(object sender, string filePath)
    {
        _materialsSource = filePath;
        _settingsService.AddRecentFile(filePath, RecentFileType.ExcelFile);
    }

    private void MaterialsSave_FileSelected(object sender, string filePath)
    {
        _materialsSave = filePath;
    }

    private void MechanismsSource_FileSelected(object sender, string filePath)
    {
        _mechanismsSource = filePath;
        _settingsService.AddRecentFile(filePath, RecentFileType.ExcelFile);
    }

    private void MechanismsSave_FileSelected(object sender, string filePath)
    {
        _mechanismsSave = filePath;
    }

    private async void IndexesGenerateButton_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_indexesSource) || string.IsNullOrEmpty(_indexesSave))
        {
            MessageBox.Show("Укажите исходный и выходной файлы!", "Ошибка", 
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        ProgressControl.ClearLog();
        ProgressControl.SetIndeterminate(true);
        ProgressControl.AppendLog("Генерация индексов...");

        try
        {
            await Task.Run(() => GenerateIndexesCsv(_indexesSource, _indexesSave));
            ProgressControl.SetCompleted($"CSV успешно создан: {Path.GetFileName(_indexesSave)}");
            MessageBox.Show("CSV успешно сформирован.", "Успех", 
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

    private void GenerateIndexesCsv(string sourcePath, string savePath)
    {
        Dispatcher.Invoke(() => ProgressControl.AppendLog("Открытие Excel файла..."));

        using var workbook = new XLWorkbook(sourcePath);
        var worksheet = workbook.Worksheet(1);
        var rows = worksheet.RowsUsed().ToList();
        
        Dispatcher.Invoke(() => ProgressControl.AppendLog($"Найдено строк: {rows.Count}"));

        var csvLines = new List<string> { "ID_TABLE;CODE;EM;MR" };
        int codeColumn = int.TryParse(IndexesCodeColumn.Text, out var c) ? c : 1;

        int tableId = 1;
        bool switchTable = false;
        int processedRows = 0;

        foreach (var row in rows)
        {
            string code = row.Cell(codeColumn).GetString() ?? "";
            
            if (code == IndexesTable1End.Text) { switchTable = true; }
            else if (code == IndexesTable2End.Text) { tableId = 2; switchTable = true; }
            else if (code == IndexesTable3End.Text) { tableId = 3; switchTable = true; }
            else if (code == IndexesTable4End.Text) { tableId = 4; switchTable = true; }

            if (switchTable && code != IndexesTable1End.Text && code != IndexesTable2End.Text 
                && code != IndexesTable3End.Text && code != IndexesTable4End.Text)
            {
                tableId = tableId == 1 ? 2 : tableId == 2 ? 3 : tableId == 3 ? 4 : 4;
                switchTable = false;
            }

            int inemCol = tableId <= 2 
                ? (int.TryParse(IndexesRegionalInem.Text, out var ri) ? ri : 5)
                : (int.TryParse(IndexesIndustryInem.Text, out var ii) ? ii : 8);
            int matCol = tableId <= 2 
                ? (int.TryParse(IndexesRegionalMat.Text, out var rm) ? rm : 6)
                : 9;

            string emStr = row.Cell(inemCol).GetString() ?? "1.000";
            string mrStr = row.Cell(matCol).GetString() ?? "1.000";
            
            double em = double.TryParse(emStr, out double emValue) ? emValue : 1.000;
            double mr = double.TryParse(mrStr, out double mrValue) ? mrValue : 1.000;
            
            csvLines.Add($"{tableId};{code};{FormatNumber(em)};{FormatNumber(mr)}");
            processedRows++;

            if (processedRows % 1000 == 0)
            {
                Dispatcher.Invoke(() => ProgressControl.AppendLog($"Обработано: {processedRows}"));
            }
        }

        if (File.Exists(savePath)) File.Delete(savePath);
        File.WriteAllLines(savePath, csvLines, Encoding.UTF8);
        
        Dispatcher.Invoke(() => ProgressControl.AppendLog($"Сохранено строк: {csvLines.Count - 1}"));
    }

    private async void MaterialsGenerateButton_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_materialsSource) || string.IsNullOrEmpty(_materialsSave))
        {
            MessageBox.Show("Укажите исходный и выходной файлы!", "Ошибка", 
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        ProgressControl.ClearLog();
        ProgressControl.SetIndeterminate(true);
        ProgressControl.AppendLog("Генерация материалов...");

        try
        {
            await Task.Run(() => GenerateMaterialsCsv(_materialsSource, _materialsSave));
            ProgressControl.SetCompleted($"CSV успешно создан: {Path.GetFileName(_materialsSave)}");
            MessageBox.Show("CSV успешно сформирован.", "Успех", 
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

    private void GenerateMaterialsCsv(string sourcePath, string savePath)
    {
        Dispatcher.Invoke(() => ProgressControl.AppendLog("Открытие Excel файла..."));

        using var workbook = new XLWorkbook(sourcePath);
        var worksheet = workbook.Worksheet(1);
        var rows = worksheet.RowsUsed().ToList();
        
        Dispatcher.Invoke(() => ProgressControl.AppendLog($"Найдено строк: {rows.Count}"));

        var csvLines = new List<string> { "TABLE_ID;CODE;NAME;UNIT;PRICE_OPT;PRICE;INDX" };
        int processedRows = 0;

        foreach (var row in rows)
        {
            if (row.RowNumber() == 1) continue;

            string code = row.Cell(1).GetString() ?? "";
            string name = row.Cell(2).GetString() ?? "";
            string unit = row.Cell(3).GetString() ?? "";
            double priceOpt = double.TryParse(row.Cell(4).GetString(), out double p) ? p : 0.0;
            double price = double.TryParse(row.Cell(5).GetString(), out double pr) ? pr : 0.0;
            double indx = double.TryParse(row.Cell(6).GetString(), out double i) ? i : 0.0;

            csvLines.Add($"1;{code};{name};{unit};{FormatNumber(priceOpt)};{FormatNumber(price)};{FormatNumber(indx)}");
            processedRows++;

            if (processedRows % 1000 == 0)
            {
                Dispatcher.Invoke(() => ProgressControl.SetProgress((double)processedRows / rows.Count * 100));
            }
        }

        if (File.Exists(savePath)) File.Delete(savePath);
        File.WriteAllLines(savePath, csvLines, Encoding.UTF8);
        
        Dispatcher.Invoke(() => ProgressControl.AppendLog($"Сохранено строк: {csvLines.Count - 1}"));
    }

    private async void MechanismsGenerateButton_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_mechanismsSource) || string.IsNullOrEmpty(_mechanismsSave))
        {
            MessageBox.Show("Укажите исходный и выходной файлы!", "Ошибка", 
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        ProgressControl.ClearLog();
        ProgressControl.SetIndeterminate(true);
        ProgressControl.AppendLog("Генерация механизмов...");

        try
        {
            await Task.Run(() => GenerateMechanismsCsv(_mechanismsSource, _mechanismsSave));
            ProgressControl.SetCompleted($"CSV успешно создан: {Path.GetFileName(_mechanismsSave)}");
            MessageBox.Show("CSV успешно сформирован.", "Успех", 
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

    private void GenerateMechanismsCsv(string sourcePath, string savePath)
    {
        Dispatcher.Invoke(() => ProgressControl.AppendLog("Открытие Excel файла..."));

        using var workbook = new XLWorkbook(sourcePath);
        var worksheet = workbook.Worksheet(1);
        var rows = worksheet.RowsUsed().ToList();
        
        Dispatcher.Invoke(() => ProgressControl.AppendLog($"Найдено строк: {rows.Count}"));

        var csvLines = new List<string> { "TABLE_ID;CODE;NAME;UNIT;PRICE;OTM;INDX" };
        int processedRows = 0;

        foreach (var row in rows)
        {
            if (row.RowNumber() == 1) continue;

            string code = row.Cell(1).GetString() ?? "";
            string name = row.Cell(2).GetString() ?? "";
            string unit = row.Cell(3).GetString() ?? "";
            double price = double.TryParse(row.Cell(4).GetString(), out double p) ? p : 0.0;
            double otm = double.TryParse(row.Cell(5).GetString(), out double o) ? o : 0.0;
            double indx = double.TryParse(row.Cell(6).GetString(), out double i) ? i : 0.0;

            csvLines.Add($"1;{code};{name};{unit};{FormatNumber(price)};{FormatNumber(otm)};{FormatNumber(indx)}");
            processedRows++;

            if (processedRows % 1000 == 0)
            {
                Dispatcher.Invoke(() => ProgressControl.SetProgress((double)processedRows / rows.Count * 100));
            }
        }

        if (File.Exists(savePath)) File.Delete(savePath);
        File.WriteAllLines(savePath, csvLines, Encoding.UTF8);
        
        Dispatcher.Invoke(() => ProgressControl.AppendLog($"Сохранено строк: {csvLines.Count - 1}"));
    }

    private string FormatNumber(double number)
    {
        return number.ToString("N3", new System.Globalization.CultureInfo("ru-RU")
        {
            NumberFormat = { NumberDecimalSeparator = ".", NumberGroupSeparator = " " }
        }).Replace(",", ".");
    }
}
