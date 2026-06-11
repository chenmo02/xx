using Microsoft.Win32;
using System.Data;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using WpfApp1.Services;

namespace WpfApp1.Views
{
    public partial class CsvViewerPage : Page
    {
        private static readonly Brush MatchBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#4E6EF2")!);
        private static readonly Brush ErrorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#DC2626")!);
        private static readonly Brush MutedBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#64748B")!);

        private DataTable? _data;
        private readonly List<(int row, int col)> _matchPositions = [];
        private int _currentMatchIndex = -1;

        public CsvViewerPage()
        {
            InitializeComponent();
            TxtMatchInfo.Foreground = MutedBrush;
        }

        private async void BtnOpen_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "支持的文件|*.csv;*.txt;*.tsv;*.dbf;*.bdf|CSV / 文本文件|*.csv;*.txt;*.tsv|DBF / BDF 文件|*.dbf;*.bdf|所有文件|*.*",
                Title = "打开数据文件"
            };

            if (dialog.ShowDialog() != true)
            {
                return;
            }

            await LoadFileAsync(dialog.FileName);
        }

        private async Task LoadFileAsync(string filePath)
        {
            if (!IsSupportedFile(filePath))
            {
                MessageBox.Show("仅支持 CSV、TXT、TSV、DBF、BDF 文件。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            BtnOpen.IsEnabled = false;
            TxtStatus.Text = "正在加载文件...";

            try
            {
                CsvViewerLoadResult result = await Task.Run(() => LoadPreviewFile(filePath));
                _data = result.Table;

                DgCsv.ItemsSource = _data.DefaultView;
                TxtFileInfo.Text = $"{result.FileName}  |  {FormatFileSize(result.FileSize)}  |  {_data.Rows.Count} 行 × {_data.Columns.Count} 列  |  {result.Detail}";
                TxtEncoding.Text = result.EncodingLabel;
                TxtStatus.Text = $"已加载 {_data.Rows.Count} 行，{_data.Columns.Count} 列";

                ClearSearch(clearKeyword: true);
            }
            catch (Exception ex)
            {
                TxtStatus.Text = "文件加载失败";
                MessageBox.Show($"文件加载失败：\n{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                BtnOpen.IsEnabled = true;
            }
        }

        private void Page_DragEnter(object sender, DragEventArgs e)
        {
            e.Effects = GetDroppedSupportedFile(e) == null ? DragDropEffects.None : DragDropEffects.Copy;
            e.Handled = true;
        }

        private async void Page_Drop(object sender, DragEventArgs e)
        {
            string? filePath = GetDroppedSupportedFile(e);
            if (filePath == null)
            {
                return;
            }

            await LoadFileAsync(filePath);
        }

        private void TxtSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                DoSearch();
            }
        }

        private void BtnSearch_Click(object sender, RoutedEventArgs e) => DoSearch();

        private void BtnNextMatch_Click(object sender, RoutedEventArgs e)
        {
            if (_matchPositions.Count == 0)
            {
                return;
            }

            _currentMatchIndex = (_currentMatchIndex + 1) % _matchPositions.Count;
            NavigateToMatch(_currentMatchIndex);
        }

        private void BtnClearSearch_Click(object sender, RoutedEventArgs e) => ClearSearch(clearKeyword: true);

        private void DoSearch()
        {
            if (_data == null || _data.Rows.Count == 0)
            {
                TxtMatchInfo.Text = "请先打开文件";
                TxtMatchInfo.Foreground = ErrorBrush;
                return;
            }

            string keyword = TxtSearch.Text.Trim();
            if (string.IsNullOrWhiteSpace(keyword))
            {
                TxtMatchInfo.Text = "请输入查找内容";
                TxtMatchInfo.Foreground = ErrorBrush;
                return;
            }

            _matchPositions.Clear();
            _currentMatchIndex = -1;

            for (int rowIndex = 0; rowIndex < _data.Rows.Count; rowIndex++)
            {
                for (int columnIndex = 0; columnIndex < _data.Columns.Count; columnIndex++)
                {
                    string cellText = _data.Rows[rowIndex][columnIndex]?.ToString() ?? string.Empty;
                    if (cellText.Contains(keyword, StringComparison.OrdinalIgnoreCase))
                    {
                        _matchPositions.Add((rowIndex, columnIndex));
                    }
                }
            }

            if (_matchPositions.Count == 0)
            {
                TxtMatchInfo.Text = $"未找到 “{keyword}”";
                TxtMatchInfo.Foreground = ErrorBrush;
                TxtStatus.Text = "未找到匹配项";
                ResetSelection();
                TxtRowCol.Text = $"共 {_data.Rows.Count} 行，{_data.Columns.Count} 列";
                return;
            }

            TxtMatchInfo.Foreground = MatchBrush;
            _currentMatchIndex = 0;
            NavigateToMatch(_currentMatchIndex);
        }

        private void NavigateToMatch(int index)
        {
            if (_data == null || index < 0 || index >= _matchPositions.Count)
            {
                return;
            }

            var (rowIndex, columnIndex) = _matchPositions[index];
            if (rowIndex >= DgCsv.Items.Count || columnIndex >= DgCsv.Columns.Count)
            {
                return;
            }

            object item = DgCsv.Items[rowIndex];
            var column = DgCsv.Columns[columnIndex];
            var cellInfo = new DataGridCellInfo(item, column);

            DgCsv.UnselectAllCells();
            DgCsv.SelectedCells.Clear();
            DgCsv.CurrentCell = cellInfo;
            DgCsv.SelectedCells.Add(cellInfo);
            DgCsv.ScrollIntoView(item, column);
            DgCsv.UpdateLayout();
            DgCsv.Focus();

            var cell = GetCell(rowIndex, columnIndex);
            if (cell != null)
            {
                cell.Focus();
                Keyboard.Focus(cell);
            }

            TxtMatchInfo.Text = $"匹配 {index + 1}/{_matchPositions.Count}";
            TxtStatus.Text = $"定位到第 {rowIndex + 1} 行，第 {columnIndex + 1} 列 ({_data.Columns[columnIndex].ColumnName})";
            TxtRowCol.Text = $"第 {rowIndex + 1} 行，第 {columnIndex + 1} 列 ({_data.Columns[columnIndex].ColumnName})";
        }

        private void ClearSearch(bool clearKeyword)
        {
            if (clearKeyword)
            {
                TxtSearch.Clear();
            }

            _matchPositions.Clear();
            _currentMatchIndex = -1;
            TxtMatchInfo.Text = string.Empty;
            TxtMatchInfo.Foreground = MutedBrush;
            ResetSelection();

            if (_data == null)
            {
                TxtRowCol.Text = "未选择单元格";
                return;
            }

            TxtStatus.Text = $"已加载 {_data.Rows.Count} 行，{_data.Columns.Count} 列";
            TxtRowCol.Text = $"共 {_data.Rows.Count} 行，{_data.Columns.Count} 列";
        }

        private void ResetSelection()
        {
            DgCsv.UnselectAllCells();
            DgCsv.SelectedCells.Clear();
            DgCsv.CurrentCell = new DataGridCellInfo();
        }

        private void DgCsv_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = (e.Row.GetIndex() + 1).ToString(CultureInfo.InvariantCulture);
        }

        private void DgCsv_SelectedCellsChanged(object sender, SelectedCellsChangedEventArgs e)
        {
            if (_data == null || DgCsv.CurrentCell.Column == null)
            {
                return;
            }

            int rowIndex = DgCsv.Items.IndexOf(DgCsv.CurrentCell.Item);
            int columnIndex = DgCsv.Columns.IndexOf(DgCsv.CurrentCell.Column);

            if (rowIndex >= 0 && columnIndex >= 0 && columnIndex < _data.Columns.Count)
            {
                TxtRowCol.Text = $"第 {rowIndex + 1} 行，第 {columnIndex + 1} 列 ({_data.Columns[columnIndex].ColumnName})";
            }
        }

        private DataGridCell? GetCell(int rowIndex, int columnIndex)
        {
            if (rowIndex < 0 || columnIndex < 0 || rowIndex >= DgCsv.Items.Count || columnIndex >= DgCsv.Columns.Count)
            {
                return null;
            }

            object item = DgCsv.Items[rowIndex];
            var column = DgCsv.Columns[columnIndex];

            DgCsv.ScrollIntoView(item, column);
            DgCsv.UpdateLayout();

            var rowContainer = DgCsv.ItemContainerGenerator.ContainerFromIndex(rowIndex) as DataGridRow;
            if (rowContainer == null)
            {
                return null;
            }

            rowContainer.ApplyTemplate();
            var presenter = FindVisualChild<DataGridCellsPresenter>(rowContainer);
            if (presenter == null)
            {
                return column.GetCellContent(rowContainer)?.Parent as DataGridCell;
            }

            return presenter.ItemContainerGenerator.ContainerFromIndex(columnIndex) as DataGridCell
                ?? column.GetCellContent(rowContainer)?.Parent as DataGridCell;
        }

        private static T? FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
        {
            for (int index = 0; index < VisualTreeHelper.GetChildrenCount(parent); index++)
            {
                var child = VisualTreeHelper.GetChild(parent, index);
                if (child is T typedChild)
                {
                    return typedChild;
                }

                var childMatch = FindVisualChild<T>(child);
                if (childMatch != null)
                {
                    return childMatch;
                }
            }

            return null;
        }

        private static string FormatFileSize(long fileSize)
        {
            if (fileSize >= 1024 * 1024)
            {
                return $"{fileSize / 1024.0 / 1024.0:F1} MB";
            }

            return $"{fileSize / 1024.0:F1} KB";
        }

        private static CsvViewerLoadResult LoadPreviewFile(string filePath)
        {
            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            if (ext is ".dbf" or ".bdf")
            {
                DataTable table = FileParserService.ParseFile(filePath, dbfEncoding: "GBK");
                var fileInfo = new FileInfo(filePath);
                return new CsvViewerLoadResult
                {
                    FileName = fileInfo.Name,
                    FileSize = fileInfo.Length,
                    Detail = $"类型: {(ext == ".bdf" ? "BDF" : "DBF")}",
                    EncodingLabel = "编码: GBK",
                    Table = table
                };
            }

            DelimitedTextLoadResult result = DelimitedTextFileService.LoadFile(filePath);
            return new CsvViewerLoadResult
            {
                FileName = result.FileName,
                FileSize = result.FileSize,
                Detail = $"分隔符: {DelimitedTextFileService.GetDelimiterName(result.Delimiter)}",
                EncodingLabel = $"编码: {result.Encoding.EncodingName}",
                Table = result.Table
            };
        }

        private static string? GetDroppedSupportedFile(DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                return null;
            }

            string[]? files = e.Data.GetData(DataFormats.FileDrop) as string[];
            string? filePath = files?.FirstOrDefault();
            return filePath != null && IsSupportedFile(filePath) ? filePath : null;
        }

        private static bool IsSupportedFile(string filePath)
        {
            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            return ext is ".csv" or ".txt" or ".tsv" or ".dbf" or ".bdf";
        }

        private sealed class CsvViewerLoadResult
        {
            public required string FileName { get; init; }

            public required long FileSize { get; init; }

            public required string Detail { get; init; }

            public required string EncodingLabel { get; init; }

            public required DataTable Table { get; init; }
        }
    }
}
