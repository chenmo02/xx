using System.Windows;
using System.Windows.Controls;
using WpfApp1.Views;

namespace WpfApp1
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            MainFrame.Navigate(new HomePage());
        }

        private void NavButton_Checked(object sender, RoutedEventArgs e)
        {
            if (MainFrame == null) return;
            var rb = sender as RadioButton;
            if (rb == null) return;

            if (rb == NavHome) MainFrame.Navigate(new HomePage());
            else if (rb == NavImport) MainFrame.Navigate(new DataImportPage());
            else if (rb == NavCsvViewer) MainFrame.Navigate(new CsvViewerPage());
            else if (rb == NavCsvCompare) MainFrame.Navigate(new CsvComparePage());
            else if (rb == NavJsonTool) MainFrame.Navigate(new JsonToolPage());
            else if (rb == NavJsonDiff) MainFrame.Navigate(new JsonDiffPage());
            else if (rb == NavDrawBoard) MainFrame.Navigate(new DrawBoardPage());
            else if (rb == NavInvoice) MainFrame.Navigate(new InvoicePrintPage());
            else if (rb == NavSettings) MainFrame.Navigate(new SettingsPage());
        }
    }
}
