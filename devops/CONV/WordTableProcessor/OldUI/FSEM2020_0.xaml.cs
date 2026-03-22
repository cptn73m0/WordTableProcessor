using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Win32;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Globalization;
using System.Text;

namespace WordTableProcessor
{
    public partial class FSEM2020_0 : UserControl
    {
        private string wordFilePath;
        private string csvFilePath;
        private const string LogFilePath = "log.txt";

        public FSEM2020_0()
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
            ProgressBar.Value = 0;
            try
            {
                logBuilder.AppendLine($"[{DateTime.Now}] Начало очистки таблицы...");
                logBuilder.AppendLine($"Проверка файла: Word={wordFilePath}");
                if (!File.Exists(wordFilePath))
                    throw new FileNotFoundException($"Файл {wordFilePath} не найден.");

                await Task.Run(() =>
                {
                    using (WordprocessingDocument doc = WordprocessingDocument.Open(wordFilePath, true))
                    {
                        var tables = doc.MainDocumentPart.Document.Body.Elements<Table>().ToList();
                        if (tables.Count < 4)
                        {
                            Dispatcher.Invoke(() =>
                            {
                                StatusTextBox.Text += $"[{DateTime.Now}] В документе недостаточно таблиц!\n";
                                MessageBox.Show("В документе недостаточно таблиц!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                            });
                            return;
                        }

                        var table = tables[3]; // Таблица 4 (индекс 0-based)
                        var rows = table.Elements<TableRow>().ToList();
                        int totalRows = rows.Count;
                        int processedRows = 0;

                        foreach (var row in rows)
                        {
                            var cells = row.Elements<TableCell>().ToList();
                            if (cells.Count > 1)
                            {
                                string code = cells[1].InnerText.Trim(); // 2-й столбец
                                if (Regex.IsMatch(code, @"^\d{2}\.\d{2}\.\d{2}-\d{3}.*$"))
                                {
                                    if (cells.Count > 6)
                                    {
                                        cells[4].RemoveAllChildren(); // 5-й столбец
                                        cells[4].Append(new Paragraph(new Run(new Text("!"))));
                                        cells[5].RemoveAllChildren(); // 6-й столбец
                                        cells[5].Append(new Paragraph(new Run(new Text("!"))));
                                        cells[6].RemoveAllChildren(); // 7-й столбец
                                        cells[6].Append(new Paragraph(new Run(new Text("!"))));
                                    }
                                }
                            }
                            processedRows++;
                            if (processedRows % 1000 == 0 || processedRows == totalRows)
                            {
                                double progress = (double)processedRows / totalRows * 100;
                                Dispatcher.Invoke(() => ProgressBar.Value = progress);
                            }
                        }
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
                    StatusTextBox.Text += $"[{DateTime.Now}] Очистка таблицы завершена за {formattedTime}.\n";
                    MessageBox.Show("Очистка таблицы завершена.", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                });
            }
            catch (Exception ex)
            {
                logBuilder.AppendLine($"[{DateTime.Now}] Ошибка при очистке таблицы: {ex.Message}");
                await File.AppendAllTextAsync(LogFilePath, logBuilder.ToString());
                Dispatcher.Invoke(() =>
                {
                    StatusTextBox.Text += $"[{DateTime.Now}] Ошибка при очистке таблицы: {ex.Message}\n";
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

                // Чтение CSV файла
                List<string[]> csvData = new List<string[]>();
                Encoding encoding = Encoding.UTF8;
                try
                {
                    using (var reader = new StreamReader(csvFilePath, Encoding.UTF8))
                    {
                        reader.ReadLine(); // Попытка чтения для проверки кодировки
                    }
                }
                catch
                {
                    encoding = Encoding.GetEncoding(1251); // Windows-1251
                }
                using (var reader = new StreamReader(csvFilePath, encoding))
                {
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        if (line != null)
                        {
                            var cells = line.Split(';');
                            if (cells.Length >= 7)
                            {
                                csvData.Add(cells);
                            }
                        }
                    }
                }
                logBuilder.AppendLine($"[{DateTime.Now}] CSV загружен, найдено {csvData.Count} записей, кодировка: {encoding.EncodingName}");

                await Task.Run(() =>
                {
                    using (WordprocessingDocument doc = WordprocessingDocument.Open(wordFilePath, true))
                    {
                        var tables = doc.MainDocumentPart.Document.Body.Elements<Table>().ToList();
                        if (tables.Count < 4)
                        {
                            Dispatcher.Invoke(() =>
                            {
                                StatusTextBox.Text += $"[{DateTime.Now}] В документе недостаточно таблиц!\n";
                                MessageBox.Show("В документе недостаточно таблиц!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                            });
                            return;
                        }

                        int[] tableIndices = { 1, 3 }; // Таблицы 2 и 4 (индекс 0-based)
                        int totalRows = tableIndices.Sum(t => tables[t].Elements<TableRow>().Count());
                        int processedRows = 0;
                        int csvIndex = 0;

                        foreach (int t in tableIndices)
                        {
                            var table = tables[t];
                            var rows = table.Elements<TableRow>().ToList();
                            foreach (var row in rows)
                            {
                                var cells = row.Elements<TableCell>().ToList();
                                if (cells.Count > 1)
                                {
                                    string code = cells[1].InnerText.Trim(); // 2-й столбец
                                    if (Regex.IsMatch(code, @"^\d{2}\.\d{2}\.\d{2}-\d{3}.*$"))
                                    {
                                        bool found = false;
                                        while (csvIndex < csvData.Count)
                                        {
                                            string csvCode = csvData[csvIndex][1].Trim();
                                            string csvTableNum = csvData[csvIndex][0].Trim();
                                            if (code.StartsWith(csvCode) && csvTableNum == (t == 1 ? "1" : "2"))
                                            {
                                                found = true;
                                                if (cells.Count > 6)
                                                {
                                                    cells[4].RemoveAllChildren(); // 5-й столбец (PRICE)
                                                    cells[4].Append(new Paragraph(new Run(new Text(FormatNumber(csvData[csvIndex][4])))));
                                                    cells[5].RemoveAllChildren(); // 6-й столбец (OTM)
                                                    cells[5].Append(new Paragraph(new Run(new Text(FormatNumber(csvData[csvIndex][5])))));
                                                    cells[6].RemoveAllChildren(); // 7-й столбец (INDX)
                                                    cells[6].Append(new Paragraph(new Run(new Text(FormatNumber(csvData[csvIndex][6])))));
                                                    updatedRows++;
                                                }
                                                csvIndex++;
                                                break;
                                            }
                                            csvIndex++;
                                        }
                                        if (!found)
                                        {
                                            skippedRows++;
                                            Dispatcher.InvokeAsync(() => StatusTextBox.Text += $"[{DateTime.Now}] Код {code} не найден в CSV для таблицы {(t == 1 ? 2 : 4)}\n");
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
                                if (processedRows % 1000 == 0 || processedRows == totalRows)
                                {
                                    double progress = (double)processedRows / totalRows * 100;
                                    Dispatcher.Invoke(() => ProgressBar.Value = progress);
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
                    StatusTextBox.Text += $"[{DateTime.Now}] Обновление цен завершено. Обновлено {updatedRows} строк, пропущено {skippedRows} строк за {formattedTime}.\n";
                    MessageBox.Show($"Обновление цен завершено. Обновлено {updatedRows} строк, пропущено {skippedRows} строк.", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                });
            }
            catch (Exception ex)
            {
                logBuilder.AppendLine($"[{DateTime.Now}] Ошибка при обновлении цен: {ex.Message}");
                await File.AppendAllTextAsync(LogFilePath, logBuilder.ToString());
                Dispatcher.Invoke(() =>
                {
                    StatusTextBox.Text += $"[{DateTime.Now}] Ошибка при обновлении цен: {ex.Message}. Обновлено {updatedRows} строк, пропущено {skippedRows} строк.\n";
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
                Dispatcher.InvokeAsync(() =>
                {
                    StatusTextBox.Text += $"[{DateTime.Now}] Ошибка формата числа: '{input}' (очищено: '{cleanedInput}', ASCII: {asciiDebug}). Детали: {ex.Message}\n";
                });
                File.AppendAllText(LogFilePath, $"[{DateTime.Now}] Ошибка формата числа: '{input}' (очищено: '{cleanedInput}', ASCII: {asciiDebug}). Детали: {ex.Message}\n");
                return "0,00";
            }
        }
    }
}