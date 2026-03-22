using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Win32;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Collections.Generic;
using System.Globalization;

namespace WordTableProcessor
{
    public partial class TSSC2014 : UserControl
    {
        private string? wordFilePath;
        private string? csvFilePath;
        private const string LogFilePath = "log.txt";

        public TSSC2014()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            InitializeComponent();
        }

        private void SelectWordButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Word Documents|*.docx"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                wordFilePath = openFileDialog.FileName;
                WordFilePathTextBlock.Text = wordFilePath;
                StatusTextBox.Text += $"[{DateTime.Now}] Выбран Word документ: {wordFilePath}\n";
                File.AppendAllText(LogFilePath, $"[{DateTime.Now}] Выбран Word документ: {wordFilePath}\n");
            }
        }

        private void SelectCsvButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "CSV Files|*.csv"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                csvFilePath = openFileDialog.FileName;
                CsvFilePathTextBlock.Text = csvFilePath;
                StatusTextBox.Text += $"[{DateTime.Now}] Выбран CSV файл: {csvFilePath}\n";
                File.AppendAllText(LogFilePath, $"[{DateTime.Now}] Выбран CSV файл: {csvFilePath}\n");
            }
        }

        private async void CleanTableButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(wordFilePath))
            {
                MessageBox.Show("Сначала выберите Word документ!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var stopwatch = Stopwatch.StartNew();
            StringBuilder logBuilder = new StringBuilder();

            try
            {
                await Task.Run(() =>
                {
                    logBuilder.AppendLine($"[{DateTime.Now}] Начало очистки таблицы...");
                    logBuilder.AppendLine($"Проверка файла: {wordFilePath}");
                    if (!File.Exists(wordFilePath))
                    {
                        throw new FileNotFoundException($"Файл {wordFilePath} не найден.");
                    }
                    var fileInfo = new FileInfo(wordFilePath);
                    logBuilder.AppendLine($"Размер файла: {fileInfo.Length / 1024.0 / 1024.0:F2} МБ");
                    logBuilder.AppendLine($"[{DateTime.Now}] Открытие Word документа...");

                    using (WordprocessingDocument doc = WordprocessingDocument.Open(wordFilePath, true))
                    {
                        var tables = doc.MainDocumentPart!.Document.Body!.Elements<Table>();
                        int tableCount = tables.Count();
                        if (tableCount == 0)
                        {
                            logBuilder.AppendLine($"[{DateTime.Now}] Таблица отсутствует!");
                            Dispatcher.Invoke(() =>
                            {
                                StatusTextBox.Text += $"[{DateTime.Now}] Таблица отсутствует!\n";
                                MessageBox.Show("Таблица отсутствует!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                            });
                            return;
                        }

                        int totalRows = tables.Sum(table => table.Elements<TableRow>().Count());
                        logBuilder.AppendLine($"[{DateTime.Now}] Найдено {tableCount} таблиц, {totalRows} строк.");
                        int processedRows = 0;

                        foreach (var table in tables)
                        {
                            foreach (var row in table.Elements<TableRow>())
                            {
                                var cells = row.Elements<TableCell>().ToList();
                                if (cells.Count >= 7)
                                {
                                    string code = cells[1].InnerText.Trim();
                                    if (Regex.IsMatch(code, @"^\d{3}-\d{4}.*$"))
                                    {
                                        cells[4].RemoveAllChildren();
                                        cells[4].Append(new Paragraph(new Run(new Text("!"))));
                                        cells[5].RemoveAllChildren();
                                        cells[5].Append(new Paragraph(new Run(new Text("!"))));
                                        cells[6].RemoveAllChildren();
                                        cells[6].Append(new Paragraph(new Run(new Text("!"))));
                                    }
                                }
                                processedRows++;
                                if (processedRows % 100 == 0 || processedRows == totalRows)
                                {
                                    Dispatcher.Invoke(() => ProgressBar.Value = (double)processedRows / totalRows * 100);
                                }
                            }
                        }
                        logBuilder.AppendLine($"[{DateTime.Now}] Очищено {processedRows} строк.");
                        doc.Save();
                    }
                });

                stopwatch.Stop();
                var timeSpan = stopwatch.Elapsed;
                string formattedTime = $"{(int)timeSpan.TotalMinutes}:{timeSpan.Seconds:D2}";
                logBuilder.AppendLine($"[{DateTime.Now}] Очистка таблицы завершена за {formattedTime}.");
                await File.AppendAllTextAsync(LogFilePath, logBuilder.ToString());
                Dispatcher.Invoke(() =>
                {
                    StatusTextBox.Text += logBuilder.ToString();
                    MessageBox.Show("Очистка таблицы завершена.", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                });
            }
            catch (Exception ex)
            {
                logBuilder.AppendLine($"[{DateTime.Now}] Ошибка при очистке таблицы: {ex.Message}");
                await File.AppendAllTextAsync(LogFilePath, logBuilder.ToString());
                Dispatcher.Invoke(() =>
                {
                    StatusTextBox.Text += logBuilder.ToString();
                    MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                });
            }
            finally
            {
                Dispatcher.Invoke(() => ProgressBar.Value = 0);
            }
        }

        private async void UpdatePricesButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(wordFilePath) || string.IsNullOrEmpty(csvFilePath))
            {
                MessageBox.Show("Выберите Word документ и CSV файл!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var stopwatch = Stopwatch.StartNew();
            StringBuilder logBuilder = new StringBuilder();
            int updatedRows = 0;
            int skippedRows = 0;

            try
            {
                logBuilder.AppendLine($"[{DateTime.Now}] Начало обновления цен...");
                logBuilder.AppendLine($"Проверка файлов: Word={wordFilePath}, CSV={csvFilePath}");
                if (!File.Exists(wordFilePath))
                    throw new FileNotFoundException($"Файл {wordFilePath} не найден.");
                if (!File.Exists(csvFilePath))
                    throw new FileNotFoundException($"Файл {csvFilePath} не найден.");
                var wordFileInfo = new FileInfo(wordFilePath);
                var csvFileInfo = new FileInfo(csvFilePath);
                logBuilder.AppendLine($"Размер Word файла: {wordFileInfo.Length / 1024.0 / 1024.0:F2} МБ");
                logBuilder.AppendLine($"Размер CSV файла: {csvFileInfo.Length / 1024.0 / 1024.0:F2} МБ");
                logBuilder.AppendLine($"Версия .NET: {Environment.Version}");
                logBuilder.AppendLine($"Доступная память: {GC.GetTotalMemory(false) / 1024.0 / 1024.0:F2} МБ");

                // Кэширование CSV
                Dictionary<string, (string val5, string val6, string val7)> csvCache = new Dictionary<string, (string, string, string)>(50000);
                logBuilder.AppendLine($"[{DateTime.Now}] Загрузка CSV в кэш...");
                Encoding encoding = Encoding.UTF8;

                // Проверка кодировки
                try
                {
                    using (var reader = new StreamReader(csvFilePath, Encoding.UTF8, true))
                    {
                        string header = reader.ReadLine();
                        if (header != null)
                        {
                            var headerCells = header.Split(';');
                            logBuilder.AppendLine($"[{DateTime.Now}] Структура CSV: {headerCells.Length} столбцов, заголовок: {header}");
                        }
                        string sample = reader.ReadLine();
                        if (sample != null && sample.Contains("�"))
                        {
                            encoding = Encoding.GetEncoding(1251);
                            logBuilder.AppendLine($"[{DateTime.Now}] Обнаружены нечитаемые символы, переключение на кодировку Windows-1251.");
                        }
                    }
                }
                catch (Exception ex)
                {
                    logBuilder.AppendLine($"[{DateTime.Now}] Ошибка при определении кодировки CSV: {ex.Message}");
                    encoding = Encoding.UTF8; // Fallback на UTF-8
                }

                using (var reader = new StreamReader(csvFilePath, encoding))
                {
                    reader.ReadLine(); // Пропуск заголовочной строки
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        if (line != null)
                        {
                            var cells = line.Split(';');
                            if (cells.Length >= 7)
                            {
                                if (!string.IsNullOrWhiteSpace(cells[1]))
                                {
                                    string csvCode = cells[1].Trim();
                                    string price = cells.Length > 4 ? cells[4].Trim() : "";
                                    string otm = cells.Length > 5 ? cells[5].Trim() : "";
                                    string indx = cells.Length > 6 ? cells[6].Trim() : "";
                                    // Форматирование чисел
                                    price = FormatNumber(price);
                                    otm = FormatNumber(otm);
                                    indx = FormatNumber(indx);
                                    csvCache[csvCode] = (price, otm, indx);
                                }
                            }
                        }
                    }
                }
                logBuilder.AppendLine($"[{DateTime.Now}] CSV загружен в кэш, найдено {csvCache.Count} записей.");
                logBuilder.AppendLine($"[{DateTime.Now}] Первые 5 кодов из CSV: {string.Join(", ", csvCache.Keys.Take(5))}");

                await Task.Run(() =>
                {
                    logBuilder.AppendLine($"[{DateTime.Now}] Открытие Word документа...");

                    using (WordprocessingDocument doc = WordprocessingDocument.Open(wordFilePath, true))
                    {
                        var tables = doc.MainDocumentPart!.Document.Body!.Elements<Table>();
                        int tableCount = tables.Count();
                        if (tableCount == 0)
                        {
                            logBuilder.AppendLine($"[{DateTime.Now}] Таблица отсутствует!");
                            Dispatcher.Invoke(() =>
                            {
                                StatusTextBox.Text += $"[{DateTime.Now}] Таблица отсутствует!\n";
                                MessageBox.Show("Таблица отсутствует!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                            });
                            return;
                        }

                        int totalRows = tables.Sum(table => table.Elements<TableRow>().Count());
                        logBuilder.AppendLine($"[{DateTime.Now}] Найдено {tableCount} таблиц, {totalRows} строк.");
                        int processedRows = 0;

                        foreach (var table in tables)
                        {
                            foreach (var row in table.Elements<TableRow>())
                            {
                                var cells = row.Elements<TableCell>().ToList();
                                if (cells.Count >= 7)
                                {
                                    string code = cells[1].InnerText.Trim();
                                    if (Regex.IsMatch(code, @"^\d{3}-\d{4}.*$"))
                                    {
                                        // Упрощенная нормализация: удаляем только xe и кавычки
                                        string normalizedCode = Regex.Replace(code, @"xe\s*""[^""]*""", "").Trim();
                                        if (csvCache.TryGetValue(normalizedCode, out var csvRow))
                                        {
                                            cells[4].RemoveAllChildren();
                                            cells[4].Append(new Paragraph(new Run(new Text(csvRow.val5))));
                                            cells[5].RemoveAllChildren();
                                            cells[5].Append(new Paragraph(new Run(new Text(csvRow.val6))));
                                            cells[6].RemoveAllChildren();
                                            cells[6].Append(new Paragraph(new Run(new Text(csvRow.val7))));
                                            updatedRows++;
                                        }
                                        else
                                        {
                                            skippedRows++;
                                        }
                                    }
                                    else
                                    {
                                        skippedRows++;
                                    }
                                }
                                else
                                {
                                    skippedRows++;
                                }
                                processedRows++;
                                if (processedRows % 100 == 0 || processedRows == totalRows)
                                {
                                    Dispatcher.Invoke(() => ProgressBar.Value = (double)processedRows / totalRows * 100);
                                }
                            }
                        }
                        logBuilder.AppendLine($"[{DateTime.Now}] Обработано {processedRows} строк, обновлено {updatedRows} строк, пропущено {skippedRows} строк.");
                        doc.Save();
                    }
                });

                stopwatch.Stop();
                var timeSpan = stopwatch.Elapsed;
                string formattedTime = $"{(int)timeSpan.TotalMinutes}:{timeSpan.Seconds:D2}";
                logBuilder.AppendLine($"[{DateTime.Now}] Обновление цен завершено за {formattedTime}.");
                await File.AppendAllTextAsync(LogFilePath, logBuilder.ToString());
                Dispatcher.Invoke(() =>
                {
                    StatusTextBox.Text += logBuilder.ToString();
                    MessageBox.Show($"Обновление цен завершено. Обновлено {updatedRows} строк, пропущено {skippedRows} строк.", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                });
            }
            catch (Exception ex)
            {
                logBuilder.AppendLine($"[{DateTime.Now}] Ошибка при обновлении цен: {ex.Message}");
                await File.AppendAllTextAsync(LogFilePath, logBuilder.ToString());
                Dispatcher.Invoke(() =>
                {
                    StatusTextBox.Text += logBuilder.ToString();
                    MessageBox.Show($"Ошибка: {ex.Message}. Обновлено {updatedRows} строк, пропущено {skippedRows} строк.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                });
            }
            finally
            {
                Dispatcher.Invoke(() => ProgressBar.Value = 0);
            }
        }

        private string FormatNumber(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return "0,00";

            // Удаляем пробелы и неразрывные пробелы
            string cleanedInput = input.Trim().Replace(" ", "").Replace("\u00A0", "");
            try
            {
                // Парсим число, заменяя запятую на точку
                if (double.TryParse(cleanedInput.Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double number))
                {
                    // Форматируем с двумя знаками после запятой и запятой как разделителем
                    string formatted = number.ToString("N2", CultureInfo.InvariantCulture).Replace(",", " ").Replace(".", ",");
                    return formatted;
                }
                else
                {
                    throw new FormatException("Не удалось распарсить число.");
                }
            }
            catch (FormatException ex)
            {
                // Дополнительная диагностика: выводим ASCII-коды
                string asciiDebug = string.Join(" ", cleanedInput.Select(c => ((int)c).ToString()));
                Dispatcher.InvokeAsync(() => StatusTextBox.Text += $"[{DateTime.Now}] Ошибка формата числа: '{input}' (очищено: '{cleanedInput}', ASCII: {asciiDebug}). Детали: {ex.Message}\n");
                File.AppendAllText(LogFilePath, $"[{DateTime.Now}] Ошибка формата числа: '{input}' (очищено: '{cleanedInput}', ASCII: {asciiDebug}). Детали: {ex.Message}\n");
                return "0,00";
            }
        }
    }
}