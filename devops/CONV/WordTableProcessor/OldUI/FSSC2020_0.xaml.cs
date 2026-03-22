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

namespace WordTableProcessor
{
    public partial class FSSC2020_0 : UserControl
    {
        private string wordFilePath;
        private string csvFilePath;

        public FSSC2020_0()
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
                MessageBox.Show("Сначала выберите Word!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var stopwatch = Stopwatch.StartNew();
            ProgressBar.Value = 0;
            try
            {
                await Task.Run(() =>
                {
                    using (WordprocessingDocument doc = WordprocessingDocument.Open(wordFilePath, true))
                    {
                        var tables = doc.MainDocumentPart.Document.Body.Elements<Table>().ToList();
                        if (tables.Count == 0)
                        {
                            Dispatcher.Invoke(() => MessageBox.Show("Не найдены таблицы в документе!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error));
                            return;
                        }

                        int totalRows = tables.Sum(t => t.Elements<TableRow>().Count());
                        int processedRows = 0;

                        foreach (var table in tables)
                        {
                            var rows = table.Elements<TableRow>().ToList();
                            foreach (var row in rows)
                            {
                                var cells = row.Elements<TableCell>().ToList();
                                if (cells.Count > 1)
                                {
                                    string code = cells[1].InnerText.Trim(); // 2nd column (1-based index in VBA)
                                    if (Regex.IsMatch(code, @"^\d{2}\.\d\.\d{2}\.\d{2}-.*$"))
                                    {
                                        if (cells.Count > 6)
                                        {
                                            cells[4].RemoveAllChildren(); // 5th column
                                            cells[4].Append(new Paragraph(new Run(new Text("!"))));
                                            cells[5].RemoveAllChildren(); // 6th column
                                            cells[5].Append(new Paragraph(new Run(new Text("!"))));
                                            cells[6].RemoveAllChildren(); // 7th column
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
                        }
                        doc.Save();
                    }
                });

                stopwatch.Stop();
                var timeSpan = stopwatch.Elapsed;
                string formattedTime = $"{(int)timeSpan.TotalMinutes}:{timeSpan.Seconds:D2}";
                StatusTextBox.Text += $"Очистка таблицы завершена за {formattedTime}.\n";
                MessageBox.Show("Таблица очищена.", "Успешно!", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                StatusTextBox.Text += $"Ошибка очистки таблицы: {ex.Message}\n";
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
                MessageBox.Show("Пожалуйста, выберите Word и CSV файл!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var stopwatch = Stopwatch.StartNew();
            ProgressBar.Value = 0;
            try
            {
                // Read CSV file
                List<string[]> csvData = new List<string[]>();
                using (var reader = new StreamReader(csvFilePath, System.Text.Encoding.GetEncoding("windows-1251")))
                {
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        var cells = line.Split(';');
                        if (cells.Length >= 7)
                        {
                            csvData.Add(cells);
                        }
                    }
                }

                await Task.Run(() =>
                {
                    using (WordprocessingDocument doc = WordprocessingDocument.Open(wordFilePath, true))
                    {
                        var tables = doc.MainDocumentPart.Document.Body.Elements<Table>().ToList();
                        if (tables.Count == 0)
                        {
                            Dispatcher.Invoke(() => MessageBox.Show("Не найдены таблицы в документе!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error));
                            return;
                        }

                        int totalRows = tables.Sum(t => t.Elements<TableRow>().Count());
                        int processedRows = 0;
                        int csvIndex = 0;

                        foreach (var table in tables)
                        {
                            var rows = table.Elements<TableRow>().ToList();
                            foreach (var row in rows)
                            {
                                var cells = row.Elements<TableCell>().ToList();
                                if (cells.Count > 1)
                                {
                                    string code = cells[1].InnerText.Trim(); // 2nd column
                                    if (Regex.IsMatch(code, @"^\d{2}\.\d\.\d{2}\.\d{2}-.*$"))
                                    {
                                        bool found = false;
                                        while (csvIndex < csvData.Count)
                                        {
                                            string csvCode = csvData[csvIndex][1].Trim();
                                            if (code.StartsWith(csvCode))
                                            {
                                                found = true;
                                                if (cells.Count > 6)
                                                {
                                                    // Проверяем и форматируем числа
                                                    string formatValue(string input, bool isIndexColumn)
                                                    {
                                                        if (string.IsNullOrWhiteSpace(input))
                                                        {
                                                            Dispatcher.InvokeAsync(() => StatusTextBox.Text += $"Пустое значение в строке с кодом {csvData[csvIndex][1]}\n");
                                                            return "0,00";
                                                        }

                                                        // Удаляем пробелы и неразрывные пробелы
                                                        string cleanedInput = input.Trim().Replace(" ", "").Replace("\u00A0", "");
                                                        try
                                                        {
                                                            // Проверяем, является ли строка числом с точкой
                                                            if (!Regex.IsMatch(cleanedInput, @"^-?\d*\.?\d+$"))
                                                            {
                                                                throw new FormatException($"Недопустимые символы в числе: '{cleanedInput}'");
                                                            }
                                                            // Парсим с использованием InvariantCulture
                                                            if (!double.TryParse(cleanedInput, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double value))
                                                            {
                                                                throw new FormatException("Не удалось распарсить число.");
                                                            }
                                                            // Форматируем с двумя знаками после запятой
                                                            string formatted = value.ToString("F2", System.Globalization.CultureInfo.InvariantCulture).Replace(".", ",");
                                                            // Для столбцов PRICE_OPT и PRICE добавляем пробелы
                                                            if (!isIndexColumn)
                                                            {
                                                                string[] parts = formatted.Split(',');
                                                                string integerPart = parts[0];
                                                                string decimalPart = parts.Length > 1 ? "," + parts[1] : ",00";
                                                                // Добавляем пробелы каждые три цифры справа налево
                                                                string formattedInteger = "";
                                                                for (int i = integerPart.Length - 1, j = 0; i >= 0; i--, j++)
                                                                {
                                                                    if (j > 0 && j % 3 == 0 && integerPart[i] != '-')
                                                                        formattedInteger = " " + formattedInteger;
                                                                    formattedInteger = integerPart[i] + formattedInteger;
                                                                }
                                                                return formattedInteger + decimalPart;
                                                            }
                                                            return formatted; // Для INDX без пробелов
                                                        }
                                                        catch (FormatException ex)
                                                        {
                                                            // Дополнительная диагностика: выводим ASCII-коды символов
                                                            string asciiDebug = string.Join(" ", cleanedInput.Select(c => ((int)c).ToString()));
                                                            Dispatcher.InvokeAsync(() => StatusTextBox.Text += $"Ошибка формата числа: '{input}' (очищено: '{cleanedInput}', ASCII: {asciiDebug}) в строке с кодом {csvData[csvIndex][1]}. Детали: {ex.Message}\n");
                                                            return "0,00";
                                                        }
                                                    }

                                                    cells[4].RemoveAllChildren(); // 5th column (PRICE_OPT)
                                                    cells[4].Append(new Paragraph(new Run(new Text(formatValue(csvData[csvIndex][4], false)))));
                                                    cells[5].RemoveAllChildren(); // 6th column (PRICE)
                                                    cells[5].Append(new Paragraph(new Run(new Text(formatValue(csvData[csvIndex][5], false)))));
                                                    cells[6].RemoveAllChildren(); // 7th column (INDX)
                                                    cells[6].Append(new Paragraph(new Run(new Text(formatValue(csvData[csvIndex][6], true)))));
                                                }
                                                csvIndex++;
                                                break;
                                            }
                                            csvIndex++;
                                        }
                                        if (!found)
                                        {
                                            Dispatcher.InvokeAsync(() => StatusTextBox.Text += $"Код {code} не найден в CSV для {tables.IndexOf(table) + 1}\n");
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
                StatusTextBox.Text += $"Цены обновлены за {formattedTime}.\n";
                MessageBox.Show("Цены обновлены.", "Успех!", MessageBoxButton.OK, MessageBoxImage.Information);
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