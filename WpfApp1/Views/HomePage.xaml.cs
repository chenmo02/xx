using System.Windows;
using System.Windows.Controls;

namespace WpfApp1.Views
{
    public partial class HomePage : Page
    {
        public HomePage()
        {
            InitializeComponent();
            TxtDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
            TxtRuntime.Text = $".NET {Environment.Version.Major}";
        }

        private void GoToImport_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (Application.Current.MainWindow is MainWindow mw)
            { mw.MainFrame.Navigate(new DataImportPage()); mw.NavImport.IsChecked = true; }
        }

        private void GoToCsvViewer_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (Application.Current.MainWindow is MainWindow mw)
            { mw.MainFrame.Navigate(new CsvViewerPage()); mw.NavCsvViewer.IsChecked = true; }
        }

        private void GoToJsonTool_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (Application.Current.MainWindow is MainWindow mw)
            { mw.MainFrame.Navigate(new JsonToolPage()); mw.NavJsonTool.IsChecked = true; }
        }
    }
}
