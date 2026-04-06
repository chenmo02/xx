using System.Windows;
using System.Windows.Controls;

namespace WpfApp1
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            MainFrame.Navigate(new Views.HomePage());
        }

        private void NavButton_Checked(object sender, RoutedEventArgs e)
        {
            if (MainFrame == null) return;
            var rb = sender as RadioButton;
            if (rb == null) return;

            if (rb == NavHome) MainFrame.Navigate(new Views.HomePage());
            else if (rb == NavImport) MainFrame.Navigate(new Views.DataImportPage());
            else if (rb == NavCsvViewer) MainFrame.Navigate(new Views.CsvViewerPage());
            else if (rb == NavJsonTool) MainFrame.Navigate(new Views.JsonToolPage());
            else if (rb == NavJsonDiff) MainFrame.Navigate(new Views.JsonDiffPage());
            else if (rb == NavSettings) MainFrame.Navigate(new Views.SettingsPage());
        }
    }
}
