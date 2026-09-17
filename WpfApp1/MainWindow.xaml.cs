using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WpfApp1.Views;

namespace WpfApp1
{
    public partial class MainWindow : Window
    {
        private readonly HomePage _homePage = new();
        private QuickNotesWindow? _quickNotesWindow;
        private Point _notesButtonStartPoint;
        private Thickness _notesButtonStartMargin;
        private bool _notesButtonDragging;

        public MainWindow()
        {
            InitializeComponent();
            MainFrame.Navigate(_homePage);
        }

        public void NavigateToHome()
        {
            MainFrame.Navigate(_homePage);
            NavHome.IsChecked = true;
        }

        public void NavigateToCsvCompare()
        {
            NavSettings.IsChecked = true;
            MainFrame.Navigate(new CsvComparePage());
        }

        public void NavigateToQuickTable()
        {
            NavQuickTable.IsChecked = true;
        }

        private void NavButton_Checked(object sender, RoutedEventArgs e)
        {
            if (MainFrame == null) return;
            var rb = sender as RadioButton;
            if (rb == null) return;

            if (rb == NavHome) MainFrame.Navigate(_homePage);
            else if (rb == NavImport) MainFrame.Navigate(new DataImportPage());
            else if (rb == NavCsvViewer) MainFrame.Navigate(new CsvViewerPage());
            else if (rb == NavQuickTable) MainFrame.Navigate(new QuickTablePage());
            else if (rb == NavDataValidation) MainFrame.Navigate(new DataValidationPage());
            else if (rb == NavJsonTool) MainFrame.Navigate(new JsonToolPage());
            else if (rb == NavJsonDiff) MainFrame.Navigate(new JsonDiffPage());
            else if (rb == NavDrawBoard) MainFrame.Navigate(new DrawBoardPage());
            else if (rb == NavInvoice) MainFrame.Navigate(new InvoicePrintPage());
        }

        private void NavSettings_Click(object sender, RoutedEventArgs e)
        {
            MainFrame.Navigate(new SettingsPage());
        }

        private void NotesButton_Click(object sender, RoutedEventArgs e)
        {
            if (_notesButtonDragging) return;

            if (_quickNotesWindow == null)
            {
                _quickNotesWindow = new QuickNotesWindow { Owner = this };
                _quickNotesWindow.Closed += (_, _) => _quickNotesWindow = null;
            }

            if (_quickNotesWindow.IsVisible)
            {
                _quickNotesWindow.Activate();
            }
            else
            {
                _quickNotesWindow.Show();
            }
        }

        private void NotesButton_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _notesButtonStartPoint = e.GetPosition(this);
            _notesButtonStartMargin = NotesButton.Margin;
            _notesButtonDragging = false;
            NotesButton.CaptureMouse();
        }

        private void NotesButton_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton != MouseButtonState.Pressed || !NotesButton.IsMouseCaptured) return;

            var currentPoint = e.GetPosition(this);
            var delta = currentPoint - _notesButtonStartPoint;
            if (!_notesButtonDragging && (Math.Abs(delta.X) < 3 && Math.Abs(delta.Y) < 3)) return;

            _notesButtonDragging = true;
            var parentGrid = NotesButton.Parent as Grid;
            var contentWidth = (parentGrid?.ActualWidth ?? ActualWidth) - NotesButton.ActualWidth;
            var contentHeight = (parentGrid?.ActualHeight ?? ActualHeight) - NotesButton.ActualHeight;
            var right = Math.Clamp(_notesButtonStartMargin.Right - delta.X, 0, Math.Max(0, contentWidth));
            var bottom = Math.Clamp(_notesButtonStartMargin.Bottom - delta.Y, 0, Math.Max(0, contentHeight));
            NotesButton.Margin = new Thickness(0, 0, right, bottom);
        }

        private void NotesButton_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (NotesButton.IsMouseCaptured) NotesButton.ReleaseMouseCapture();
            if (!_notesButtonDragging) NotesButton_Click(sender, new RoutedEventArgs());
            e.Handled = true;
        }
    }
}
