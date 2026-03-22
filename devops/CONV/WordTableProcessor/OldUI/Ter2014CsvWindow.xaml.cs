using Microsoft.Win32;
using System;
using System.Windows;

namespace WordTableProcessor
{
    public partial class Ter2014CsvWindow : Window
    {
        public Ter2014CsvWindow()
        {
            InitializeComponent();
        }

        // Индексы
        private void IndexesBrowseSourceButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx"
            };
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
                Filter = "Excel Files|*.xlsx",
                FileName = $"Индексы ТЕР-2014 {DateTime.Now.ToString("MM.yyyy")}.xlsx"
            };
            if (saveFileDialog.ShowDialog() == true)
            {
                IndexesSaveFileTextBox.Text = saveFileDialog.FileName;
            }
        }

        // ФССЦ
        private void FsscBrowseSourceButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx"
            };
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
                Filter = "Excel Files|*.xlsx",
                FileName = $"Материалы ТЕР-2014 {DateTime.Now.ToString("MM.yyyy")}.xlsx"
            };
            if (saveFileDialog.ShowDialog() == true)
            {
                FsscSaveFileTextBox.Text = saveFileDialog.FileName;
            }
        }

        // ФСЭМ
        private void FsemBrowseSourceButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx"
            };
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
                Filter = "Excel Files|*.xlsx",
                FileName = $"Механизмы ТЕР-2014 {DateTime.Now.ToString("MM.yyyy")}.xlsx"
            };
            if (saveFileDialog.ShowDialog() == true)
            {
                FsemSaveFileTextBox.Text = saveFileDialog.FileName;
            }
        }
    }
}