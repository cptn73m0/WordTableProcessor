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

namespace WordTableProcessor
{
    public partial class TER2010 : UserControl
    {
        private string? wordFilePath;
        private string? csvFilePath;

        public TER2010()
        {
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
                StatusTextBox.Text += $"Выбран Word документ: {wordFilePath}\n";
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
                StatusTextBox.Text += $"Выбран CSV файл: {csvFilePath}\n";
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
            ProgressBar.Value = 0;
            try
            {
                await Task.Run(() =>
                {
                    using (WordprocessingDocument doc = WordprocessingDocument.Open(wordFilePath!, true))
                    {
                        var tables = doc.MainDocumentPart!.Document.Body!.Elements<Table>().ToList();
                        if (tables.Count < 4)
                        {
                            Dispatcher.Invoke(() => MessageBox.Show("В документе недостаточно таблиц!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error));
                            return;
                        }

                        int totalRows = tables.Take(4).Sum(t => t.Elements<TableRow>().Count());
                        int processedRows = 0;

                        for (int t = 0; t < 4; t++) // Таблицы 1-4 (индекс 0-based)
                        {
                            var table = tables[t];
                            var rows = table.Elements<TableRow>().ToList();
                            foreach (var row in rows)
                            {
                                var cells = row.Elements<TableCell>().ToList();
                                if (cells.Count > 0)
                                {
                                    string code = cells[0].InnerText.Trim(); // 1-й столбец
                                    if (Regex.IsMatch(code, @"^\d{2}-\d{2}-\d{3}-\d{2}.*$"))
                                    {
                                        if (cells.Count > 2)
                                        {
                                            cells[1].RemoveAllChildren(); // 2-й столбец
                                            cells[1].Append(new Paragraph(new Run(new Text("!"))));
                                            cells[2].RemoveAllChildren(); // 3-й столбец
                                            cells[2].Append(new Paragraph(new Run(new Text("!"))));
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
                        }
                        doc.Save();
                    }
                });

                stopwatch.Stop();
                var timeSpan = stopwatch.Elapsed;
                string formattedTime = $"{(int)timeSpan.TotalMinutes}:{timeSpan.Seconds:D2}";
                StatusTextBox.Text += $"Очистка таблицы завершена за {formattedTime}.\n";
                MessageBox.Show("Очистка таблицы завершена.", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                StatusTextBox.Text += $"Ошибка при очистке таблицы: {ex.Message}\n";
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                ProgressBar.Value = 0;
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
            ProgressBar.Value = 0;
            try
            {
                // Чтение CSV файла
                List<string[]> csvData = new List<string[]>();
                using (var reader = new StreamReader(csvFilePath))
                {
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
                }

                await Task.Run(() =>
                {
                    using (WordprocessingDocument doc = WordprocessingDocument.Open(wordFilePath!, true))
                    {
                        var tables = doc.MainDocumentPart!.Document.Body!.Elements<Table>().ToList();
                        if (tables.Count < 4)
                        {
                            Dispatcher.Invoke(() => MessageBox.Show("В документе недостаточно таблиц!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error));
                            return;
                        }

                        int totalRows = tables.Take(4).Sum(t => t.Elements<TableRow>().Count());
                        int processedRows = 0;
                        int csvIndex = 0;

                        for (int t = 0; t < 4; t++) // Таблицы 1-4 (индекс 0-based)
                        {
                            var table = tables[t];
                            var rows = table.Elements<TableRow>().ToList();
                            foreach (var row in rows)
                            {
                                var cells = row.Elements<TableCell>().ToList();
                                if (cells.Count > 0)
                                {
                                    string code = cells[0].InnerText.Trim(); // 1-й столбец
                                    if (Regex.IsMatch(code, @"^\d{2}-\d{2}-\d{3}-\d{2}.*$"))
                                    {
                                        bool found = false;
                                        while (csvIndex < csvData.Count)
                                        {
                                            string csvCode = csvData[csvIndex][1].Trim();
                                            string csvTableNum = csvData[csvIndex][0].Trim();
                                            if (code.StartsWith(csvCode) && csvTableNum == (t + 1).ToString())
                                            {
                                                found = true;
                                                if (cells.Count > 2)
                                                {
                                                    string formattedPrice1, formattedPrice2;
                                                    try
                                                    {
                                                        // Парсим значения из CSV, где точка — десятичный разделитель
                                                        double price1 = double.Parse(csvData[csvIndex][2], System.Globalization.CultureInfo.InvariantCulture);
                                                        double price2 = double.Parse(csvData[csvIndex][3], System.Globalization.CultureInfo.InvariantCulture);
                                                        // Форматируем до двух знаков после запятой с запятой (ru-RU)
                                                        formattedPrice1 = price1.ToString("F2", System.Globalization.CultureInfo.GetCultureInfo("ru-RU")); // 25.020 -> 25,02
                                                        formattedPrice2 = price2.ToString("F2", System.Globalization.CultureInfo.GetCultureInfo("ru-RU"));
                                                    }
                                                    catch (FormatException ex)
                                                    {
                                                        Dispatcher.Invoke(() => StatusTextBox.Text += $"Ошибка формата числа в CSV для кода {csvData[csvIndex][1]}: {ex.Message}\n");
                                                        csvIndex++;
                                                        continue;
                                                    }

                                                    cells[1].RemoveAllChildren(); // 2-й столбец
                                                    cells[1].Append(new Paragraph(new Run(new Text(formattedPrice1))));
                                                    cells[2].RemoveAllChildren(); // 3-й столбец
                                                    cells[2].Append(new Paragraph(new Run(new Text(formattedPrice2))));
                                                }
                                                csvIndex++;
                                                break;
                                            }
                                            csvIndex++;
                                        }
                                        if (!found)
                                        {
                                            Dispatcher.Invoke(() => 
                                            {
                                                StatusTextBox.Text += $"Код {code} не найден в CSV для таблицы {t + 1}\n";
                                                MessageBox.Show($"Код {code} не найден в CSV для таблицы {t + 1}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                                            });
                                            csvIndex = 0; // Сброс для следующей итерации
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
                        }
                        doc.Save();
                    }
                });

                stopwatch.Stop();
                var timeSpan = stopwatch.Elapsed;
                string formattedTime = $"{(int)timeSpan.TotalMinutes}:{timeSpan.Seconds:D2}";
                StatusTextBox.Text += $"Обновление цен завершено за {formattedTime}.\n";
                MessageBox.Show("Обновление цен завершено.", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                StatusTextBox.Text += $"Ошибка при обновлении цен: {ex.Message}\n";
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                ProgressBar.Value = 0;
            }
        }
    }
}