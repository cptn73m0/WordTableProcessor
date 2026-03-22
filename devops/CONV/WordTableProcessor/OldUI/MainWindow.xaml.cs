using System.Windows;
using System.Windows.Controls;

namespace WordTableProcessor
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void GenerateCsvButton_Click(object sender, RoutedEventArgs e)
        {
            if (CsvButtonsPanel.Visibility == Visibility.Collapsed)
            {
                MainTabControl.Visibility = Visibility.Collapsed;
                CsvButtonsPanel.Visibility = Visibility.Visible;
                GenerateCsvButton.Content = "Вернуться к вкладкам";
            }
            else
            {
                MainTabControl.Visibility = Visibility.Visible;
                CsvButtonsPanel.Visibility = Visibility.Collapsed;
                GenerateCsvButton.Content = "Формирование CSV файлов";
            }
        }

        private void CsvButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button)
            {
                switch (button.Content.ToString())
                {
                    case "ФЕР-2020(9)":
                        var fer2020_9Window = new Fer2020_9CsvWindow();
                        fer2020_9Window.ShowDialog();
                        break;
                    case "ФЕР-2020(0)":
                        var fer2020_0Window = new Fer2020_0CsvWindow();
                        fer2020_0Window.ShowDialog();
                        break;
                    case "ТЕР-2014":
                        var ter2014Window = new Ter2014CsvWindow();
                        ter2014Window.ShowDialog();
                        break;
                    case "ТЕР-2010":
                        var ter2010Window = new Ter2010CsvWindow();
                        ter2010Window.ShowDialog();
                        break;
                }
            }
        }
    }
}