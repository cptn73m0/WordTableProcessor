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
    public partial class FER2020_9 : UserControl
    {
        private string wordFilePath;
        private string csvFilePath;

        public FER2020_9()
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
                StatusTextBox.Text += $"Selected Word document: {wordFilePath}\n";
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
                StatusTextBox.Text += $"Selected CSV file: {csvFilePath}\n";
            }
        }

        private async void CleanTableButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(wordFilePath))
            {
                MessageBox.Show("Please select a Word document first!", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
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
                        if (tables.Count < 4)
                        {
                            Dispatcher.Invoke(() => MessageBox.Show("Document does not contain enough tables!", "Error", MessageBoxButton.OK, MessageBoxImage.Error));
                            return;
                        }

                        var table = tables[3]; // Table 4 (0-based index)
                        var rows = table.Elements<TableRow>().ToList();
                        int totalRows = rows.Count;
                        int processedRows = 0;

                        foreach (var row in rows)
                        {
                            var cells = row.Elements<TableCell>().ToList();
                            if (cells.Count > 0)
                            {
                                string code = cells[0].InnerText.Trim(); // 1st column (0-based index)
                                if (Regex.IsMatch(code, @"^\d{2}-\d{2}-\d{3}-\d{2}.*$"))
                                {
                                    if (cells.Count > 2)
                                    {
                                        cells[1].RemoveAllChildren(); // 2nd column
                                        cells[1].Append(new Paragraph(new Run(new Text("!"))));
                                        cells[2].RemoveAllChildren(); // 3rd column
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
                        doc.Save();
                    }
                });

                stopwatch.Stop();
                var timeSpan = stopwatch.Elapsed;
                string formattedTime = $"{(int)timeSpan.TotalMinutes}:{timeSpan.Seconds:D2}";
                StatusTextBox.Text += $"Table cleaning completed in {formattedTime}.\n";
                MessageBox.Show("Table cleaning completed.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                StatusTextBox.Text += $"Error during table cleaning: {ex.Message}\n";
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
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
                MessageBox.Show("Please select both Word document and CSV file!", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var stopwatch = Stopwatch.StartNew();
            ProgressBar.Value = 0;
            try
            {
                // Read CSV file
                List<string[]> csvData = new List<string[]>();
                using (var reader = new StreamReader(csvFilePath))
                {
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        var cells = line.Split(';');
                        if (cells.Length >= 4)
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
                        if (tables.Count < 4)
                        {
                            Dispatcher.Invoke(() => MessageBox.Show("Document does not contain enough tables!", "Error", MessageBoxButton.OK, MessageBoxImage.Error));
                            return;
                        }

                        int totalRows = tables.Take(4).Sum(t => t.Elements<TableRow>().Count());
                        int processedRows = 0;
                        int csvIndex = 0;

                        for (int t = 0; t < 4; t++) // Tables 1-4 (0-based index)
                        {
                            var table = tables[t];
                            var rows = table.Elements<TableRow>().ToList();
                            foreach (var row in rows)
                            {
                                var cells = row.Elements<TableCell>().ToList();
                                if (cells.Count > 0)
                                {
                                    string code = cells[0].InnerText.Trim(); // 1st column
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
                                                        formattedPrice1 = price1.ToString("F2", System.Globalization.CultureInfo.GetCultureInfo("ru-RU"));
                                                        formattedPrice2 = price2.ToString("F2", System.Globalization.CultureInfo.GetCultureInfo("ru-RU"));
                                                    }
                                                    catch (FormatException ex)
                                                    {
                                                        Dispatcher.Invoke(() => StatusTextBox.Text += $"Ошибка формата числа в CSV для кода {csvData[csvIndex][1]}: {ex.Message}\n");
                                                        csvIndex++;
                                                        continue;
                                                    }

                                                    cells[1].RemoveAllChildren(); // 2-я колонка
                                                    cells[1].Append(new Paragraph(new Run(new Text(formattedPrice1))));
                                                    cells[2].RemoveAllChildren(); // 3-я колонка
                                                    cells[2].Append(new Paragraph(new Run(new Text(formattedPrice2))));
                                                }
                                                csvIndex++;
                                                break;
                                            }
                                            csvIndex++;
                                        }
                                        if (!found)
                                        {
                                            Dispatcher.Invoke(() => StatusTextBox.Text += $"Code {code} not found in CSV for table {t + 1}\n");
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
                StatusTextBox.Text += $"Price update completed in {formattedTime}.\n";
                MessageBox.Show("Price update completed.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                StatusTextBox.Text += $"Error during price update: {ex.Message}\n";
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                ProgressBar.Value = 0;
            }
        }
    }
}