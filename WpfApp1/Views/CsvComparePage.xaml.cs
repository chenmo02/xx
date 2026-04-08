using Microsoft.Win32;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using WpfApp1.Services;

namespace WpfApp1.Views
{
    public partial class CsvComparePage : Page
    {
        [DllImport("user32.dll")] private static extern bool OpenClipboard(IntPtr hWndNewOwner);
        [DllImport("user32.dll")] private static extern bool CloseClipboard();
        [DllImport("user32.dll")] private static extern bool EmptyClipboard();
        [DllImport("user32.dll")] private static extern IntPtr SetClipboardData(uint uFormat, IntPtr hMem);
        private const uint CF_UNICODETEXT = 13;

        private sealed class LoadedFileContext
        {
            public required DelimitedTextLoadResult LoadResult { get; init; }

            public DataTable Table => LoadResult.Table;
        }

        private LoadedFileContext? _leftFile;
        private LoadedFileContext? _rightFile;
        private CsvCompareResult? _lastResult;

        public CsvComparePage()
        {
            InitializeComponent();
            UpdateModeUi();
            ResetResults();
        }

        private async void BtnLoadLeft_Click(object sender, RoutedEventArgs e)
            => await LoadFileAsync(isLeft: true);

        private async void BtnLoadRight_Click(object sender, RoutedEventArgs e)
            => await LoadFileAsync(isLeft: false);

        private async Task LoadFileAsync(bool isLeft)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "分隔文本|*.csv;*.txt;*.tsv|CSV 文件|*.csv|文本文件|*.txt;*.tsv|所有文件|*.*",
                Title = isLeft ? "选择文件 A" : "选择文件 B"
            };

            if (dialog.ShowDialog() != true)
            {
                return;
            }

            try
            {
                DelimitedTextLoadResult result = await Task.Run(() => DelimitedTextFileService.LoadFile(dialog.FileName));
                var context = new LoadedFileContext { LoadResult = result };

                if (isLeft)
                {
                    _leftFile = context;
                }
                else
                {
                    _rightFile = context;
                }

                UpdateFilePanel(isLeft, context);
                RefreshKeyColumns();
                ResetResults();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"文件加载失败：\n{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnSwap_Click(object sender, RoutedEventArgs e)
        {
            (_leftFile, _rightFile) = (_rightFile, _leftFile);
            RefreshFilePanels();
            RefreshKeyColumns();
            ResetResults();
        }

        private void BtnClear_Click(object sender, RoutedEventArgs e)
        {
            _leftFile = null;
            _rightFile = null;
            RefreshFilePanels();
            RefreshKeyColumns();
            ResetResults();
        }

        private void CompareMode_Checked(object sender, RoutedEventArgs e)
        {
            UpdateModeUi();
            UpdateCompareButtonState();
        }

        private void BtnCompare_Click(object sender, RoutedEventArgs e)
        {
            if (_leftFile == null || _rightFile == null)
            {
                MessageBox.Show("请先导入左右两个文件。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            CsvCompareMode mode = RbCompareByKey.IsChecked == true ? CsvCompareMode.ByKeyColumns : CsvCompareMode.ByRowNumber;
            List<string> selectedKeyColumns = GetSelectedKeyColumns();

            if (mode == CsvCompareMode.ByKeyColumns && selectedKeyColumns.Count == 0)
            {
                MessageBox.Show("主键模式至少需要选择一个公共列。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _lastResult = CsvCompareService.Compare(_leftFile.Table, _rightFile.Table, mode, selectedKeyColumns);
            BtnCopyReport.IsEnabled = true;
            BtnExportResult.IsEnabled = true;
            RenderResult();
        }

        private void FilterChanged(object sender, RoutedEventArgs e)
        {
            if (_lastResult == null)
            {
                return;
            }

            RenderResult();
        }

        private void BtnCopyReport_Click(object sender, RoutedEventArgs e)
        {
            if (_lastResult == null)
            {
                return;
            }

            SafeCopyToClipboard(BuildReport());
            MessageBox.Show("报告已复制到剪贴板。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void BtnExportResult_Click(object sender, RoutedEventArgs e)
        {
            if (_lastResult == null)
            {
                return;
            }

            var dialog = new SaveFileDialog
            {
                Filter = "CSV 文件|*.csv",
                FileName = $"csv_compare_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
            };

            if (dialog.ShowDialog() != true)
            {
                return;
            }

            DataTable exportTable = BuildExportTable(GetVisibleDiffItems());
            ExportService.ExportCsv(dialog.FileName, exportTable);
            MessageBox.Show($"结果已导出到：\n{dialog.FileName}", "导出成功", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void RefreshFilePanels()
        {
            if (_leftFile == null)
            {
                TxtLeftFileName.Text = "未选择文件";
                TxtLeftMeta.Text = "支持 .csv / .txt / .tsv";
            }
            else
            {
                UpdateFilePanel(true, _leftFile);
            }

            if (_rightFile == null)
            {
                TxtRightFileName.Text = "未选择文件";
                TxtRightMeta.Text = "支持 .csv / .txt / .tsv";
            }
            else
            {
                UpdateFilePanel(false, _rightFile);
            }
        }

        private void UpdateFilePanel(bool isLeft, LoadedFileContext context)
        {
            string metaText = $"{FormatFileSize(context.LoadResult.FileSize)}  |  {context.Table.Rows.Count} 行 × {context.Table.Columns.Count} 列  |  分隔符: {DelimitedTextFileService.GetDelimiterName(context.LoadResult.Delimiter)}  |  编码: {context.LoadResult.Encoding.EncodingName}";

            if (isLeft)
            {
                TxtLeftFileName.Text = context.LoadResult.FileName;
                TxtLeftMeta.Text = metaText;
            }
            else
            {
                TxtRightFileName.Text = context.LoadResult.FileName;
                TxtRightMeta.Text = metaText;
            }
        }

        private void UpdateModeUi()
        {
            if (RbCompareByKey == null || KeyColumnsContainer == null || TxtModeHint == null)
            {
                return;
            }

            bool byKey = RbCompareByKey.IsChecked == true;
            KeyColumnsContainer.Visibility = byKey ? Visibility.Visible : Visibility.Collapsed;
            TxtModeHint.Text = byKey
                ? "主键模式会按选中的公共列组成复合主键进行匹配。"
                : "行号模式会按第 N 行对第 N 行进行比较。";
        }

        private void RefreshKeyColumns()
        {
            if (KeyColumnsPanel == null || TxtKeyColumnsHint == null)
            {
                return;
            }

            var selectedBeforeRefresh = GetSelectedKeyColumns().ToHashSet(StringComparer.OrdinalIgnoreCase);
            KeyColumnsPanel.Children.Clear();

            List<string> commonColumns = GetCommonColumns();
            if (commonColumns.Count == 0)
            {
                TxtKeyColumnsHint.Text = "两边没有公共列，主键模式不可用。";
                UpdateCompareButtonState();
                return;
            }

            TxtKeyColumnsHint.Text = $"可选公共列共 {commonColumns.Count} 个，请选择一个或多个列作为复合主键。";
            foreach (string column in commonColumns)
            {
                var checkBox = new CheckBox
                {
                    Content = column,
                    Tag = column,
                    Margin = new Thickness(0, 0, 12, 8),
                    IsChecked = selectedBeforeRefresh.Contains(column)
                };
                checkBox.Checked += KeyColumnSelectionChanged;
                checkBox.Unchecked += KeyColumnSelectionChanged;
                KeyColumnsPanel.Children.Add(checkBox);
            }

            UpdateCompareButtonState();
        }

        private void KeyColumnSelectionChanged(object sender, RoutedEventArgs e)
        {
            int selectedCount = GetSelectedKeyColumns().Count;
            List<string> commonColumns = GetCommonColumns();
            TxtKeyColumnsHint.Text = commonColumns.Count == 0
                ? "两边没有公共列，主键模式不可用。"
                : selectedCount == 0
                    ? $"可选公共列共 {commonColumns.Count} 个，请至少选择一个主键列。"
                    : $"已选择 {selectedCount} 个主键列，可执行复合主键匹配。";
            UpdateCompareButtonState();
        }

        private List<string> GetSelectedKeyColumns()
        {
            if (KeyColumnsPanel == null)
            {
                return [];
            }

            return KeyColumnsPanel.Children
                .OfType<CheckBox>()
                .Where(checkBox => checkBox.IsChecked == true)
                .Select(checkBox => checkBox.Tag?.ToString() ?? string.Empty)
                .Where(column => !string.IsNullOrWhiteSpace(column))
                .ToList();
        }

        private List<string> GetCommonColumns()
        {
            if (_leftFile == null || _rightFile == null)
            {
                return [];
            }

            var rightColumns = _rightFile.Table.Columns.Cast<DataColumn>()
                .Select(column => column.ColumnName)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            return _leftFile.Table.Columns.Cast<DataColumn>()
                .Select(column => column.ColumnName)
                .Where(column => rightColumns.Contains(column))
                .ToList();
        }

        private void UpdateCompareButtonState()
        {
            if (BtnCompare == null || RbCompareByKey == null)
            {
                return;
            }

            if (_leftFile == null || _rightFile == null)
            {
                BtnCompare.IsEnabled = false;
                return;
            }

            if (RbCompareByKey.IsChecked == true)
            {
                BtnCompare.IsEnabled = GetSelectedKeyColumns().Count > 0;
                return;
            }

            BtnCompare.IsEnabled = true;
        }

        private void ResetResults()
        {
            _lastResult = null;
            ValidationBanner.Visibility = Visibility.Collapsed;
            TxtValidationMessage.Text = string.Empty;
            TxtSummary.Text = "请先导入两个文件并执行对比。";
            TxtEmptyState.Text = "当前没有可显示的差异结果。";
            TxtEmptyState.Visibility = Visibility.Visible;
            DgDiffs.Visibility = Visibility.Collapsed;
            DgDiffs.ItemsSource = null;
            BtnCopyReport.IsEnabled = false;
            BtnExportResult.IsEnabled = false;
            ChkShowColumnChanges.IsEnabled = true;
            ChkShowRowChanges.IsEnabled = true;
            ChkShowCellChanges.IsEnabled = true;
            UpdateCompareButtonState();
        }

        private void RenderResult()
        {
            if (_lastResult == null)
            {
                ResetResults();
                return;
            }

            if (_lastResult.HasValidationErrors)
            {
                ValidationBanner.Visibility = Visibility.Visible;
                TxtValidationMessage.Text = string.Join(Environment.NewLine, _lastResult.ValidationErrors);
                ChkShowColumnChanges.IsEnabled = false;
                ChkShowRowChanges.IsEnabled = false;
                ChkShowCellChanges.IsEnabled = false;

                List<CsvDiffItem> duplicateItems = _lastResult.DiffItems
                    .Where(item => item.DiffType == CsvDiffType.DuplicateKey)
                    .ToList();

                TxtSummary.Text = $"发现 {_lastResult.DuplicateKeyCount} 个重复主键，已停止正常对比，请先修复数据。";
                BindDiffItems(duplicateItems, "当前没有可显示的重复主键明细。");
                return;
            }

            ValidationBanner.Visibility = Visibility.Collapsed;
            TxtValidationMessage.Text = string.Empty;
            ChkShowColumnChanges.IsEnabled = true;
            ChkShowRowChanges.IsEnabled = true;
            ChkShowCellChanges.IsEnabled = true;

            List<CsvDiffItem> visibleItems = GetVisibleDiffItems();
            if (_lastResult.DiffItems.Count == 0)
            {
                TxtSummary.Text = "✅ 两个文件内容一致，没有发现任何差异。";
                BindDiffItems([], "两个文件内容一致，没有发现任何差异。");
                return;
            }

            TxtSummary.Text = $"总差异 {_lastResult.DiffItems.Count} 处  |  行新增 {_lastResult.RowAddedCount}  |  行删除 {_lastResult.RowRemovedCount}  |  单元格修改 {_lastResult.CellModifiedCount}  |  列新增 {_lastResult.ColumnAddedCount}  |  列删除 {_lastResult.ColumnRemovedCount}"
                + (visibleItems.Count != _lastResult.DiffItems.Count ? $"  |  当前显示 {visibleItems.Count} 条" : string.Empty);

            BindDiffItems(visibleItems, "当前筛选条件下没有可显示的差异结果。");
        }

        private List<CsvDiffItem> GetVisibleDiffItems()
        {
            if (_lastResult == null)
            {
                return [];
            }

            if (_lastResult.HasValidationErrors)
            {
                return _lastResult.DiffItems.Where(item => item.DiffType == CsvDiffType.DuplicateKey).ToList();
            }

            bool showColumnChanges = ChkShowColumnChanges.IsChecked == true;
            bool showRowChanges = ChkShowRowChanges.IsChecked == true;
            bool showCellChanges = ChkShowCellChanges.IsChecked == true;

            return _lastResult.DiffItems.Where(item =>
                (showColumnChanges && item.DiffType is CsvDiffType.ColumnAdded or CsvDiffType.ColumnRemoved) ||
                (showRowChanges && item.DiffType is CsvDiffType.RowAdded or CsvDiffType.RowRemoved) ||
                (showCellChanges && item.DiffType == CsvDiffType.CellModified)
            ).ToList();
        }

        private void BindDiffItems(IReadOnlyList<CsvDiffItem> items, string emptyMessage)
        {
            if (items.Count == 0)
            {
                DgDiffs.Visibility = Visibility.Collapsed;
                DgDiffs.ItemsSource = null;
                TxtEmptyState.Visibility = Visibility.Visible;
                TxtEmptyState.Text = emptyMessage;
                return;
            }

            TxtEmptyState.Visibility = Visibility.Collapsed;
            DgDiffs.Visibility = Visibility.Visible;
            DgDiffs.ItemsSource = items;
        }

        private string BuildReport()
        {
            var visibleItems = GetVisibleDiffItems();
            var builder = new StringBuilder();

            builder.AppendLine("CSV 对比报告");
            builder.AppendLine($"生成时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            builder.AppendLine($"文件 A: {_leftFile?.LoadResult.FileName ?? "未选择"}");
            builder.AppendLine($"文件 B: {_rightFile?.LoadResult.FileName ?? "未选择"}");
            builder.AppendLine();

            if (_lastResult == null)
            {
                builder.AppendLine("尚未执行对比。");
                return builder.ToString();
            }

            if (_lastResult.HasValidationErrors)
            {
                builder.AppendLine("校验失败：");
                foreach (string error in _lastResult.ValidationErrors)
                {
                    builder.AppendLine($"- {error}");
                }
            }
            else if (_lastResult.DiffItems.Count == 0)
            {
                builder.AppendLine("结果：两个文件内容一致，没有发现任何差异。");
            }
            else
            {
                builder.AppendLine(TxtSummary.Text);
            }

            builder.AppendLine();

            if (visibleItems.Count == 0)
            {
                builder.AppendLine("当前没有可导出的差异明细。");
                return builder.ToString();
            }

            foreach (CsvDiffItem item in visibleItems)
            {
                builder.AppendLine($"[{item.DiffTypeText}] {item.Locator}");
                if (!string.IsNullOrWhiteSpace(item.ColumnName))
                {
                    builder.AppendLine($"列名: {item.ColumnName}");
                }
                if (!string.IsNullOrEmpty(item.LeftValue))
                {
                    builder.AppendLine($"A 值: {item.LeftValue}");
                }
                if (!string.IsNullOrEmpty(item.RightValue))
                {
                    builder.AppendLine($"B 值: {item.RightValue}");
                }
                builder.AppendLine($"说明: {item.Message}");
                builder.AppendLine();
            }

            return builder.ToString().TrimEnd();
        }

        private static DataTable BuildExportTable(IReadOnlyList<CsvDiffItem> items)
        {
            var table = new DataTable("CsvCompareResult");
            table.Columns.Add("差异类型", typeof(string));
            table.Columns.Add("定位", typeof(string));
            table.Columns.Add("列名", typeof(string));
            table.Columns.Add("A 值", typeof(string));
            table.Columns.Add("B 值", typeof(string));
            table.Columns.Add("说明", typeof(string));

            foreach (CsvDiffItem item in items)
            {
                DataRow row = table.NewRow();
                row[0] = item.DiffTypeText;
                row[1] = item.Locator;
                row[2] = item.ColumnName;
                row[3] = item.LeftValue;
                row[4] = item.RightValue;
                row[5] = item.Message;
                table.Rows.Add(row);
            }

            return table;
        }

        private void SafeCopyToClipboard(string text)
        {
            try
            {
                if (OpenClipboard(IntPtr.Zero))
                {
                    try
                    {
                        EmptyClipboard();
                        IntPtr handle = Marshal.StringToHGlobalUni(text);
                        SetClipboardData(CF_UNICODETEXT, handle);
                    }
                    finally
                    {
                        CloseClipboard();
                    }
                }
                else
                {
                    Clipboard.SetText(text);
                }
            }
            catch
            {
                try
                {
                    Clipboard.SetText(text);
                }
                catch
                {
                }
            }
        }

        private static string FormatFileSize(long fileSize)
        {
            if (fileSize >= 1024 * 1024)
            {
                return $"{fileSize / 1024.0 / 1024.0:F1} MB";
            }

            return $"{fileSize / 1024.0:F1} KB";
        }
    }
}
