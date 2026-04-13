using Microsoft.Win32;
using System.Data;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
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

        private sealed class HeaderComparisonInfo
        {
            public int LeftColumnCount { get; init; }
            public int RightColumnCount { get; init; }
            public int MatchedCount { get; init; }
            public required List<string> LeftOnlyColumns { get; init; }
            public required List<string> RightOnlyColumns { get; init; }
        }

        private sealed class CsvCompareDetailRow
        {
            public string ColumnName { get; init; } = string.Empty;
            public string LeftValue { get; init; } = string.Empty;
            public string RightValue { get; init; } = string.Empty;
            public string Message { get; init; } = string.Empty;
        }

        private sealed class CsvCompareDisplayRow
        {
            public CsvDiffType DiffGroupType { get; init; }
            public string DiffTypeText { get; init; } = string.Empty;
            public string Locator { get; init; } = string.Empty;
            public string SummaryText { get; init; } = string.Empty;
            public int DiffCount { get; init; }
            public string LeftPreview { get; init; } = string.Empty;
            public string RightPreview { get; init; } = string.Empty;
            public int GroupSortOrder { get; init; }
            public string LocatorKey { get; init; } = string.Empty;
            public int? RowNumber { get; init; }
            public required IReadOnlyList<CsvDiffItem> Details { get; init; }
        }

        private readonly List<int> _pageSizes = [10, 20, 50, 100, 200];
        private readonly List<CsvCompareDisplayRow> _allDisplayRows = [];
        private readonly List<CsvCompareDisplayRow> _filteredDisplayRows = [];
        private readonly List<CsvCompareDisplayRow> _currentPageRows = [];
        private readonly HashSet<string> _restoredSelectedKeyColumns = new(StringComparer.OrdinalIgnoreCase);

        private LoadedFileContext? _leftFile;
        private LoadedFileContext? _rightFile;
        private CsvCompareResult? _lastResult;
        private HeaderComparisonInfo? _headerComparison;
        private int _pageIndex;
        private int _pageSize = 100;
        private CsvCompareMode _currentMode = CsvCompareMode.ByRowNumber;
        private bool _isKeyColumnsExpanded = true;
        private bool _isRefreshingKeyColumns;

        public CsvComparePage()
        {
            InitializeComponent();

            CmbPageSize.ItemsSource = _pageSizes;
            CmbPageSize.SelectedItem = _pageSize;
            CmbPageSize.SelectionChanged += CmbPageSize_SelectionChanged;

            RbCompareByRow.Checked += CompareMode_Checked;
            RbCompareByKey.Checked += CompareMode_Checked;
            ChkShowColumnChanges.Checked += Filter_CheckedChanged;
            ChkShowColumnChanges.Unchecked += Filter_CheckedChanged;
            ChkShowRowChanges.Checked += Filter_CheckedChanged;
            ChkShowRowChanges.Unchecked += Filter_CheckedChanged;
            ChkShowCellChanges.Checked += Filter_CheckedChanged;
            ChkShowCellChanges.Unchecked += Filter_CheckedChanged;

            UpdateResponsiveLayout();
            UpdateModeUi();
            UpdateKeyColumnsExpanderUi();
            ResetResults();
            RestoreState();
        }

        // ── 事件处理 ──────────────────────────────────────────────

        private void Page_SizeChanged(object sender, SizeChangedEventArgs e) => UpdateResponsiveLayout();

        private async void BtnLoadLeft_Click(object sender, RoutedEventArgs e) => await LoadFileAsync(isLeft: true);

        private async void BtnLoadRight_Click(object sender, RoutedEventArgs e) => await LoadFileAsync(isLeft: false);

        private void BtnSwap_Click(object sender, RoutedEventArgs e)
        {
            (_leftFile, _rightFile) = (_rightFile, _leftFile);
            RefreshFilePanels();
            RefreshKeyColumns();
            ResetResults();
            SaveState();
        }

        private void BtnClear_Click(object sender, RoutedEventArgs e)
        {
            _leftFile = null;
            _rightFile = null;
            RefreshFilePanels();
            RefreshKeyColumns();
            ResetResults();
            SaveState();
        }

        private void BtnCompare_Click(object sender, RoutedEventArgs e)
        {
            if (_leftFile == null || _rightFile == null)
            {
                MessageBox.Show("请先导入左右两个文件。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _currentMode = RbCompareByKey.IsChecked == true ? CsvCompareMode.ByKeyColumns : CsvCompareMode.ByRowNumber;
            List<string> selectedKeyColumns = GetSelectedKeyColumns();
            if (_currentMode == CsvCompareMode.ByKeyColumns && selectedKeyColumns.Count == 0)
            {
                MessageBox.Show("主键模式至少需要选择一个公共列。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            RunCompare();
        }

        private void BtnToggleKeyColumns_Click(object sender, RoutedEventArgs e)
        {
            _isKeyColumnsExpanded = !_isKeyColumnsExpanded;
            UpdateKeyColumnsExpanderUi();
        }

        private void CompareMode_Checked(object sender, RoutedEventArgs e)
        {
            if (RbCompareByKey == null) return;
            _currentMode = RbCompareByKey.IsChecked == true ? CsvCompareMode.ByKeyColumns : CsvCompareMode.ByRowNumber;
            UpdateModeUi();
            UpdateCompareButtonState();
            SaveState();
        }

        private void Filter_CheckedChanged(object sender, RoutedEventArgs e)
        {
            if (_lastResult == null) return;
            RenderResult(resetPage: true);
            SaveState();
        }

        private void CmbPageSize_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CmbPageSize?.SelectedItem is not int selected || selected == _pageSize) return;
            _pageSize = selected;
            _pageIndex = 0;
            BindCurrentPage();
            SaveState();
        }

        private void BtnFirstPage_Click(object sender, RoutedEventArgs e)
        {
            if (_pageIndex == 0) return;
            _pageIndex = 0;
            BindCurrentPage();
        }

        private void BtnPrevPage_Click(object sender, RoutedEventArgs e)
        {
            if (_pageIndex <= 0) return;
            _pageIndex--;
            BindCurrentPage();
        }

        private void BtnNextPage_Click(object sender, RoutedEventArgs e)
        {
            if (_filteredDisplayRows.Count == 0) return;
            int totalPages = (int)Math.Ceiling(_filteredDisplayRows.Count / (double)_pageSize);
            if (_pageIndex >= totalPages - 1) return;
            _pageIndex++;
            BindCurrentPage();
        }

        private void BtnLastPage_Click(object sender, RoutedEventArgs e)
        {
            if (_filteredDisplayRows.Count == 0) return;
            _pageIndex = Math.Max(0, (int)Math.Ceiling(_filteredDisplayRows.Count / (double)_pageSize) - 1);
            BindCurrentPage();
        }

        private void DgSummaryRows_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = (_pageIndex * _pageSize + e.Row.GetIndex() + 1).ToString();
        }

        private void DgSummaryRows_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            BindDetail((CsvCompareDisplayRow?)DgSummaryRows.SelectedItem);
        }

        private void ResultGrid_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (PageScrollViewer == null) return;
            bool canScrollUp = PageScrollViewer.VerticalOffset > 0;
            bool canScrollDown = PageScrollViewer.VerticalOffset < PageScrollViewer.ScrollableHeight;
            if ((e.Delta > 0 && canScrollUp) || (e.Delta < 0 && canScrollDown))
            {
                double nextOffset = Math.Max(0, Math.Min(
                    PageScrollViewer.VerticalOffset - e.Delta,
                    PageScrollViewer.ScrollableHeight));
                PageScrollViewer.ScrollToVerticalOffset(nextOffset);
                e.Handled = true;
            }
        }

        private void BtnCopyReport_Click(object sender, RoutedEventArgs e)
        {
            if (_lastResult == null) return;
            SafeCopyToClipboard(BuildReport());
            MessageBox.Show("报告已复制到剪贴板。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void BtnExportResult_Click(object sender, RoutedEventArgs e)
        {
            if (_lastResult == null) return;
            var dialog = new SaveFileDialog
            {
                Filter = "CSV 文件|*.csv",
                FileName = $"csv_compare_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
            };
            if (dialog.ShowDialog() != true) return;
            ExportService.ExportCsv(dialog.FileName, BuildExportTable(GetExportDiffItems()));
            MessageBox.Show($"结果已导出到：\n{dialog.FileName}", "导出成功", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        // ── 文件加载 ──────────────────────────────────────────────

        private async Task LoadFileAsync(bool isLeft)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "分隔文本|*.csv;*.txt;*.tsv|CSV 文件|*.csv|文本文件|*.txt;*.tsv|所有文件|*.*",
                Title = isLeft ? "选择文件 A" : "选择文件 B"
            };
            if (dialog.ShowDialog() != true) return;

            try
            {
                DelimitedTextLoadResult result = await Task.Run(() => DelimitedTextFileService.LoadFile(dialog.FileName));
                var context = new LoadedFileContext { LoadResult = result };
                if (isLeft) _leftFile = context;
                else _rightFile = context;

                UpdateFilePanel(isLeft, context);
                RefreshKeyColumns();
                ResetResults();
                SaveState();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"文件加载失败：\n{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ── 比对 ──────────────────────────────────────────────────

        private void TryAutoCompare()
        {
            if (_leftFile == null || _rightFile == null) return;
            if (_currentMode == CsvCompareMode.ByKeyColumns && GetSelectedKeyColumns().Count == 0) return;
            RunCompare();
        }

        private void RunCompare()
        {
            if (_leftFile == null || _rightFile == null) return;
            List<string> keyColumns = GetSelectedKeyColumns();
            _lastResult = CsvCompareService.Compare(_leftFile.Table, _rightFile.Table, _currentMode, keyColumns);
            _headerComparison = BuildHeaderComparison(_leftFile.Table, _rightFile.Table);
            BtnCopyReport.IsEnabled = true;
            BtnExportResult.IsEnabled = true;
            _pageIndex = 0;
            RenderResult(resetPage: true);
            SaveState();
        }

        // ── UI 更新 ───────────────────────────────────────────────

        private void UpdateResponsiveLayout()
        {
            // 当前布局为固定两列，无需动态切换
        }

        private void RefreshFilePanels()
        {
            if (_leftFile == null)
            {
                TxtLeftFileName.Text = "尚未加载文件";
                TxtLeftMeta.Text = "请选择左侧文件后开始自动比对。";
            }
            else
            {
                UpdateFilePanel(true, _leftFile);
            }

            if (_rightFile == null)
            {
                TxtRightFileName.Text = "尚未加载文件";
                TxtRightMeta.Text = "请选择右侧文件后开始自动比对。";
            }
            else
            {
                UpdateFilePanel(false, _rightFile);
            }
        }

        private void UpdateFilePanel(bool isLeft, LoadedFileContext context)
        {
            string meta = $"{FormatFileSize(context.LoadResult.FileSize)} | {context.Table.Rows.Count} 行 × {context.Table.Columns.Count} 列 | 分隔符：{DelimitedTextFileService.GetDelimiterName(context.LoadResult.Delimiter)} | 编码：{context.LoadResult.Encoding.EncodingName}";
            if (isLeft)
            {
                TxtLeftFileName.Text = context.LoadResult.FileName;
                TxtLeftMeta.Text = meta;
            }
            else
            {
                TxtRightFileName.Text = context.LoadResult.FileName;
                TxtRightMeta.Text = meta;
            }
        }

        private void UpdateModeUi()
        {
            if (RbCompareByKey == null || KeyColumnsContainer == null || TxtModeHint == null) return;
            bool byKey = RbCompareByKey.IsChecked == true;
            KeyColumnsContainer.Visibility = byKey ? Visibility.Visible : Visibility.Collapsed;
            TxtModeHint.Text = byKey
                ? "主键模式会按选中的公共列组成复合主键进行匹配。"
                : "行号模式会按第 N 行对第 N 行进行比较，先汇总显示发生变化的行。";
        }

        private void UpdateKeyColumnsExpanderUi()
        {
            if (BtnToggleKeyColumns == null || KeyColumnsContentPanel == null) return;
            BtnToggleKeyColumns.Content = _isKeyColumnsExpanded ? "收起" : "展开";
            KeyColumnsContentPanel.Visibility = _isKeyColumnsExpanded ? Visibility.Visible : Visibility.Collapsed;
        }

        private void RefreshKeyColumns()
        {
            if (KeyColumnsPanel == null || TxtKeyColumnsHint == null) return;

            _isRefreshingKeyColumns = true;
            try
            {
                var prevSelected = GetSelectedKeyColumns().ToHashSet(StringComparer.OrdinalIgnoreCase);
                foreach (string col in _restoredSelectedKeyColumns) prevSelected.Add(col);

                KeyColumnsPanel.Children.Clear();
                List<string> common = GetCommonColumns();

                if (common.Count == 0)
                {
                    TxtKeyColumnsHint.Text = _leftFile == null || _rightFile == null
                        ? "请选择两侧文件后查看可选主键列。"
                        : "两边没有公共列，主键模式不可用。";
                    TxtKeyColumnsSummary.Text = "当前未选择主键列";
                    UpdateCompareButtonState();
                    return;
                }

                foreach (string col in common)
                {
                    var cb = new CheckBox
                    {
                        Content = col,
                        Tag = col,
                        IsChecked = prevSelected.Contains(col)
                    };
                    cb.Checked += KeyColumnSelectionChanged;
                    cb.Unchecked += KeyColumnSelectionChanged;
                    KeyColumnsPanel.Children.Add(cb);
                }
            }
            finally
            {
                _isRefreshingKeyColumns = false;
            }

            UpdateKeyColumnsSummary();
            UpdateCompareButtonState();
        }

        private void KeyColumnSelectionChanged(object sender, RoutedEventArgs e)
        {
            if (_isRefreshingKeyColumns) return;
            UpdateKeyColumnsSummary();
            UpdateCompareButtonState();
            SaveState();
        }

        private void UpdateKeyColumnsSummary()
        {
            if (TxtKeyColumnsHint == null || TxtKeyColumnsSummary == null) return;
            List<string> selected = GetSelectedKeyColumns();
            List<string> common = GetCommonColumns();
            TxtKeyColumnsHint.Text = common.Count == 0
                ? "两边没有公共列，主键模式不可用。"
                : selected.Count == 0
                    ? $"可选公共列共 {common.Count} 个，请至少选择一个主键列。"
                    : $"已选择 {selected.Count} 个主键列，可执行复合主键匹配。";
            TxtKeyColumnsSummary.Text = selected.Count == 0
                ? "当前未选择主键列"
                : $"已选 {selected.Count} 个：{string.Join("、", selected.Take(3))}{(selected.Count > 3 ? " 等" : "")}";
        }

        private void UpdateCompareButtonState()
        {
            if (BtnCompare == null || RbCompareByKey == null) return;
            if (_leftFile == null || _rightFile == null) { BtnCompare.IsEnabled = false; return; }
            BtnCompare.IsEnabled = RbCompareByKey.IsChecked != true || GetSelectedKeyColumns().Count > 0;
        }

        // ── 结果渲染 ──────────────────────────────────────────────

        private void ResetResults()
        {
            _lastResult = null;
            _headerComparison = null;
            _allDisplayRows.Clear();
            _filteredDisplayRows.Clear();
            _currentPageRows.Clear();
            _pageIndex = 0;

            ValidationBanner.Visibility = Visibility.Collapsed;
            TxtValidationMessage.Text = string.Empty;
            TxtSummary.Text = "请先导入两个文件并执行对比。";
            TxtScopeSummary.Text = string.Empty;
            HeaderChangesBanner.Visibility = Visibility.Collapsed;
            TxtHeaderChangesSummary.Text = string.Empty;
            TxtResultEmptyState.Text = "当前没有可显示的差异结果。";
            TxtResultEmptyState.Visibility = Visibility.Visible;
            DgSummaryRows.Visibility = Visibility.Collapsed;
            DgSummaryRows.ItemsSource = null;
            DgDetailRows.ItemsSource = null;
            DgDetailRows.Visibility = Visibility.Collapsed;
            TxtDetailTitle.Text = "请选择上方一条汇总结果以查看具体明细。";
            TxtDetailLeftPreview.Text = "—";
            TxtDetailRightPreview.Text = "—";
            TxtDetailEmptyState.Text = "请选择上方一条汇总结果以查看具体明细。";
            TxtDetailEmptyState.Visibility = Visibility.Visible;
            TxtPageInfo.Text = "第 0 / 0 页";
            BtnCopyReport.IsEnabled = false;
            BtnExportResult.IsEnabled = false;
            SetFilterControlsEnabled(true);
            UpdateCompareButtonState();
        }

        private void RenderResult(bool resetPage)
        {
            if (_lastResult == null) { ResetResults(); return; }
            if (resetPage) _pageIndex = 0;

            if (_lastResult.HasValidationErrors)
            {
                RenderValidationResult();
                return;
            }

            ValidationBanner.Visibility = Visibility.Collapsed;
            TxtValidationMessage.Text = string.Empty;
            SetFilterControlsEnabled(true);

            _allDisplayRows.Clear();
            _allDisplayRows.AddRange(BuildDisplayRows(_lastResult.DiffItems));

            _filteredDisplayRows.Clear();
            _filteredDisplayRows.AddRange(_allDisplayRows.Where(ShouldShowRow));

            TxtSummary.Text = _lastResult.DiffItems.Count == 0
                ? "两个文件内容一致，没有发现任何差异。"
                : $"差异 {_lastResult.DiffItems.Count} 项 | 行新增 {_lastResult.RowAddedCount} | 行删除 {_lastResult.RowRemovedCount} | 单元格修改 {_lastResult.CellModifiedCount} | 字段变更 {_lastResult.ColumnAddedCount + _lastResult.ColumnRemovedCount}";

            UpdateHeaderChangesBanner();
            BindCurrentPage();
        }

        private void RenderValidationResult()
        {
            SetFilterControlsEnabled(false);
            ValidationBanner.Visibility = Visibility.Visible;
            TxtValidationMessage.Text = string.Join(Environment.NewLine, _lastResult!.ValidationErrors);
            HeaderChangesBanner.Visibility = Visibility.Collapsed;
            TxtHeaderChangesSummary.Text = string.Empty;
            _allDisplayRows.Clear();
            _allDisplayRows.AddRange(BuildDuplicateDisplayRows(
                _lastResult.DiffItems.Where(i => i.DiffType == CsvDiffType.DuplicateKey).ToList()));
            _filteredDisplayRows.Clear();
            _filteredDisplayRows.AddRange(_allDisplayRows);
            TxtSummary.Text = $"发现 {_lastResult.DuplicateKeyCount} 条重复主键，已停止正常对比，请先修复数据。";
            BindCurrentPage();
        }

        private void SetFilterControlsEnabled(bool enabled)
        {
            ChkShowColumnChanges.IsEnabled = enabled;
            ChkShowRowChanges.IsEnabled = enabled;
            ChkShowCellChanges.IsEnabled = enabled;
            CmbPageSize.IsEnabled = enabled;
        }

        private void UpdateHeaderChangesBanner()
        {
            // 只要两边文件都加载了，就显示字段对比摘要
            if (_headerComparison == null)
            {
                HeaderChangesBanner.Visibility = Visibility.Collapsed;
                TxtHeaderChangesSummary.Text = string.Empty;
                return;
            }

            var sb = new StringBuilder();
            sb.Append($"字段对比：A 表 {_headerComparison.LeftColumnCount} 个字段 | B 表 {_headerComparison.RightColumnCount} 个字段 | 共同字段 {_headerComparison.MatchedCount} 个");

            if (_headerComparison.LeftOnlyColumns.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine();
                sb.Append($"A 表独有（B 表缺失，共 {_headerComparison.LeftOnlyColumns.Count} 个）：");
                sb.Append(string.Join("、", _headerComparison.LeftOnlyColumns));
            }

            if (_headerComparison.RightOnlyColumns.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine();
                sb.Append($"B 表独有（A 表缺失，共 {_headerComparison.RightOnlyColumns.Count} 个）：");
                sb.Append(string.Join("、", _headerComparison.RightOnlyColumns));
            }

            TxtHeaderChangesSummary.Text = sb.ToString();
            HeaderChangesBanner.Visibility = Visibility.Visible;
        }

        private void BindCurrentPage()
        {
            int total = _filteredDisplayRows.Count;
            int totalPages = total == 0 ? 0 : (int)Math.Ceiling(total / (double)_pageSize);
            _pageIndex = totalPages == 0 ? 0 : Math.Clamp(_pageIndex, 0, totalPages - 1);

            // 先断开 DataGrid 的数据源，避免 Clear 时触发 SelectionChanged 异常
            DgSummaryRows.ItemsSource = null;
            DgDetailRows.ItemsSource = null;

            _currentPageRows.Clear();
            if (total > 0)
                _currentPageRows.AddRange(_filteredDisplayRows.Skip(_pageIndex * _pageSize).Take(_pageSize));

            TxtPageInfo.Text = totalPages == 0 ? "第 0 / 0 页" : $"第 {_pageIndex + 1} / {totalPages} 页";
            TxtScopeSummary.Text = _lastResult == null
                ? string.Empty
                : $"汇总结果 {_allDisplayRows.Count} | 当前筛选 {_filteredDisplayRows.Count} | 当前页 {_currentPageRows.Count}";

            if (_currentPageRows.Count == 0)
            {
                DgSummaryRows.Visibility = Visibility.Collapsed;
                TxtResultEmptyState.Visibility = Visibility.Visible;
                TxtResultEmptyState.Text = _lastResult?.DiffItems.Count == 0
                    ? "两个文件内容一致，没有发现任何差异。"
                    : "当前筛选下没有可显示的行级差异结果。";
                BindDetail(null);
                return;
            }

            TxtResultEmptyState.Visibility = Visibility.Collapsed;
            DgSummaryRows.Visibility = Visibility.Visible;
            DgSummaryRows.ItemsSource = _currentPageRows;
            DgSummaryRows.SelectedItem = _currentPageRows[0];
            BindDetail(_currentPageRows[0]);
        }

        private void BindDetail(CsvCompareDisplayRow? row)
        {
            if (row == null)
            {
                DgDetailRows.Visibility = Visibility.Collapsed;
                DgDetailRows.ItemsSource = null;
                TxtDetailTitle.Text = "请选择上方一条汇总结果以查看具体明细。";
                TxtDetailLeftPreview.Text = "—";
                TxtDetailRightPreview.Text = "—";
                TxtDetailEmptyState.Text = "请选择上方一条汇总结果以查看具体明细。";
                TxtDetailEmptyState.Visibility = Visibility.Visible;
                return;
            }

            TxtDetailTitle.Text = $"{row.Locator} · {row.DiffTypeText} · {row.SummaryText}";
            TxtDetailLeftPreview.Text = string.IsNullOrWhiteSpace(row.LeftPreview) ? "—" : row.LeftPreview;
            TxtDetailRightPreview.Text = string.IsNullOrWhiteSpace(row.RightPreview) ? "—" : row.RightPreview;

            List<CsvCompareDetailRow> details = BuildDetailRows(row);
            if (details.Count == 0)
            {
                DgDetailRows.Visibility = Visibility.Collapsed;
                DgDetailRows.ItemsSource = null;
                TxtDetailEmptyState.Text = "当前项没有可显示的明细。";
                TxtDetailEmptyState.Visibility = Visibility.Visible;
                return;
            }

            TxtDetailEmptyState.Visibility = Visibility.Collapsed;
            DgDetailRows.Visibility = Visibility.Visible;
            DgDetailRows.ItemsSource = details;
        }

        // ── 数据构建 ──────────────────────────────────────────────

        private static HeaderComparisonInfo BuildHeaderComparison(DataTable left, DataTable right)
        {
            var leftCols = left.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToList();
            var rightCols = right.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToList();
            var rightSet = new HashSet<string>(rightCols, StringComparer.OrdinalIgnoreCase);
            var leftSet = new HashSet<string>(leftCols, StringComparer.OrdinalIgnoreCase);

            return new HeaderComparisonInfo
            {
                LeftColumnCount = leftCols.Count,
                RightColumnCount = rightCols.Count,
                MatchedCount = leftCols.Count(c => rightSet.Contains(c)),
                LeftOnlyColumns = leftCols.Where(c => !rightSet.Contains(c)).ToList(),
                RightOnlyColumns = rightCols.Where(c => !leftSet.Contains(c)).ToList()
            };
        }

        private static List<CsvCompareDisplayRow> BuildDisplayRows(IReadOnlyList<CsvDiffItem> diffItems)
        {
            return diffItems
                .Where(i => i.DiffType is CsvDiffType.RowAdded or CsvDiffType.RowRemoved or CsvDiffType.CellModified)
                .GroupBy(i => i.LocatorKey, StringComparer.Ordinal)
                .Select(group =>
                {
                    var details = group
                        .OrderBy(i => string.IsNullOrWhiteSpace(i.ColumnName) ? 1 : 0)
                        .ThenBy(i => i.ColumnName, StringComparer.OrdinalIgnoreCase)
                        .ToList();
                    CsvDiffItem first = details[0];
                    int changedCols = details
                        .Where(i => !string.IsNullOrWhiteSpace(i.ColumnName))
                        .Select(i => i.ColumnName)
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .Count();

                    return new CsvCompareDisplayRow
                    {
                        DiffGroupType = first.DiffType,
                        DiffTypeText = first.DiffTypeText,
                        Locator = first.Locator,
                        SummaryText = BuildSummaryText(first.DiffType, details, changedCols),
                        DiffCount = details.Count,
                        LeftPreview = FirstNonEmpty(details.Select(i => i.LeftRowPreview)),
                        RightPreview = FirstNonEmpty(details.Select(i => i.RightRowPreview)),
                        GroupSortOrder = first.GroupSortOrder,
                        LocatorKey = first.LocatorKey,
                        RowNumber = details.Select(i => i.RowNumber).FirstOrDefault(n => n.HasValue),
                        Details = details
                    };
                })
                .OrderBy(r => r.GroupSortOrder)
                .ThenBy(r => r.RowNumber ?? int.MaxValue)
                .ThenBy(r => r.Locator, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static List<CsvCompareDisplayRow> BuildDuplicateDisplayRows(IReadOnlyList<CsvDiffItem> diffItems)
        {
            return diffItems
                .Select(i => new CsvCompareDisplayRow
                {
                    DiffGroupType = CsvDiffType.DuplicateKey,
                    DiffTypeText = i.DiffTypeText,
                    Locator = i.Locator,
                    SummaryText = i.Message,
                    DiffCount = 1,
                    LeftPreview = i.LeftRowPreview,
                    RightPreview = i.RightRowPreview,
                    GroupSortOrder = i.GroupSortOrder,
                    LocatorKey = i.LocatorKey,
                    RowNumber = i.RowNumber,
                    Details = [i]
                })
                .OrderBy(r => r.Locator, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static List<CsvCompareDetailRow> BuildDetailRows(CsvCompareDisplayRow row)
        {
            if (row.Details.Count == 0) return [];

            if (row.DiffGroupType == CsvDiffType.CellModified)
            {
                return row.Details.Select(i => new CsvCompareDetailRow
                {
                    ColumnName = i.ColumnName,
                    LeftValue = i.LeftValue,
                    RightValue = i.RightValue,
                    Message = i.Message
                }).ToList();
            }

            CsvDiffItem first = row.Details[0];
            return
            [
                new CsvCompareDetailRow
                {
                    ColumnName = "整行预览",
                    LeftValue = row.LeftPreview == "—" ? string.Empty : row.LeftPreview,
                    RightValue = row.RightPreview == "—" ? string.Empty : row.RightPreview,
                    Message = first.Message
                }
            ];
        }

        private static string BuildSummaryText(CsvDiffType diffType, IReadOnlyList<CsvDiffItem> details, int changedCols)
        {
            return diffType switch
            {
                CsvDiffType.RowAdded => "该记录仅存在于文件 B",
                CsvDiffType.RowRemoved => "该记录仅存在于文件 A",
                CsvDiffType.CellModified => BuildChangedColsSummary(details, changedCols),
                _ => details[0].Message
            };
        }

        private static string BuildChangedColsSummary(IReadOnlyList<CsvDiffItem> details, int changedCols)
        {
            if (changedCols == 0) return "该行存在内容差异";
            var cols = details.Where(i => !string.IsNullOrWhiteSpace(i.ColumnName))
                .Select(i => i.ColumnName).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            string preview = string.Join("、", cols.Take(3));
            return cols.Count > 3 ? $"{changedCols} 列不同：{preview} 等" : $"{changedCols} 列不同：{preview}";
        }

        private static string FirstNonEmpty(IEnumerable<string> values)
        {
            string? v = values.FirstOrDefault(s => !string.IsNullOrWhiteSpace(s));
            return string.IsNullOrWhiteSpace(v) ? "—" : v;
        }

        private bool ShouldShowRow(CsvCompareDisplayRow row) => row.DiffGroupType switch
        {
            CsvDiffType.RowAdded or CsvDiffType.RowRemoved => ChkShowRowChanges?.IsChecked == true,
            CsvDiffType.CellModified => ChkShowCellChanges?.IsChecked == true,
            _ => true
        };

        // ── 辅助方法 ──────────────────────────────────────────────

        private List<string> GetSelectedKeyColumns()
        {
            if (KeyColumnsPanel == null) return [];
            return KeyColumnsPanel.Children.OfType<CheckBox>()
                .Where(cb => cb.IsChecked == true)
                .Select(cb => cb.Tag?.ToString() ?? string.Empty)
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .ToList();
        }

        private List<string> GetCommonColumns()
        {
            if (_leftFile == null || _rightFile == null) return [];
            var rightSet = _rightFile.Table.Columns.Cast<DataColumn>()
                .Select(c => c.ColumnName)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            return _leftFile.Table.Columns.Cast<DataColumn>()
                .Select(c => c.ColumnName)
                .Where(rightSet.Contains)
                .ToList();
        }

        // ── 状态缓存 ──────────────────────────────────────────────

        private void RestoreState()
        {
            CsvCompareCachedState state = CsvCompareStateService.State;
            if (state.LeftFile == null && state.RightFile == null && state.LastResult == null) return;

            _leftFile = state.LeftFile == null ? null : FromCachedFile(state.LeftFile);
            _rightFile = state.RightFile == null ? null : FromCachedFile(state.RightFile);
            _lastResult = state.LastResult;
            _pageSize = _pageSizes.Contains(state.PageSize) ? state.PageSize : 100;
            _pageIndex = Math.Max(0, state.PageIndex);
            _currentMode = state.CurrentMode;

            _restoredSelectedKeyColumns.Clear();
            foreach (string col in state.SelectedKeyColumns) _restoredSelectedKeyColumns.Add(col);

            if (RbCompareByKey != null && RbCompareByRow != null)
            {
                RbCompareByKey.IsChecked = _currentMode == CsvCompareMode.ByKeyColumns;
                RbCompareByRow.IsChecked = _currentMode != CsvCompareMode.ByKeyColumns;
            }

            if (ChkShowColumnChanges != null) ChkShowColumnChanges.IsChecked = state.ShowColumnChanges;
            if (ChkShowRowChanges != null) ChkShowRowChanges.IsChecked = state.ShowRowChanges;
            if (ChkShowCellChanges != null) ChkShowCellChanges.IsChecked = state.ShowCellChanges;
            if (ChkExportFilteredOnly != null) ChkExportFilteredOnly.IsChecked = state.ExportFilteredOnly;
            if (CmbPageSize != null) CmbPageSize.SelectedItem = _pageSize;

            RefreshFilePanels();
            RefreshKeyColumns();
            UpdateModeUi();
            UpdateCompareButtonState();

            if (_lastResult != null)
            {
                if (_leftFile != null && _rightFile != null)
                    _headerComparison = BuildHeaderComparison(_leftFile.Table, _rightFile.Table);
                BtnCopyReport.IsEnabled = true;
                BtnExportResult.IsEnabled = true;
                RenderResult(resetPage: false);
            }
        }

        private void SaveState()
        {
            CsvCompareCachedState state = CsvCompareStateService.State;
            state.LeftFile = _leftFile == null ? null : ToCachedFile(_leftFile);
            state.RightFile = _rightFile == null ? null : ToCachedFile(_rightFile);
            state.LastResult = _lastResult;
            state.CurrentMode = _currentMode;
            state.SelectedKeyColumns = GetSelectedKeyColumns();
            state.ShowColumnChanges = ChkShowColumnChanges?.IsChecked == true;
            state.ShowRowChanges = ChkShowRowChanges?.IsChecked == true;
            state.ShowCellChanges = ChkShowCellChanges?.IsChecked == true;
            state.ExportFilteredOnly = ChkExportFilteredOnly?.IsChecked == true;
            state.PageSize = _pageSize;
            state.PageIndex = _pageIndex;
        }

        private static CsvCompareCachedFile ToCachedFile(LoadedFileContext ctx) => new()
        {
            FilePath = ctx.LoadResult.FilePath,
            FileName = ctx.LoadResult.FileName,
            FileSize = ctx.LoadResult.FileSize,
            Delimiter = ctx.LoadResult.Delimiter,
            Encoding = ctx.LoadResult.Encoding,
            Table = ctx.Table.Copy()
        };

        private static LoadedFileContext FromCachedFile(CsvCompareCachedFile f) => new()
        {
            LoadResult = new DelimitedTextLoadResult
            {
                FilePath = f.FilePath,
                FileName = f.FileName,
                FileSize = f.FileSize,
                Encoding = f.Encoding,
                Delimiter = f.Delimiter,
                Table = f.Table.Copy()
            }
        };

        // ── 导出 & 报告 ───────────────────────────────────────────

        private IReadOnlyList<CsvDiffItem> GetExportDiffItems()
        {
            if (_lastResult == null) return [];
            if (_lastResult.HasValidationErrors)
                return _lastResult.DiffItems.Where(i => i.DiffType == CsvDiffType.DuplicateKey).ToList();
            if (ChkExportFilteredOnly?.IsChecked != true)
                return _lastResult.DiffItems;

            var headerItems = _lastResult.DiffItems
                .Where(i => i.DiffType is CsvDiffType.ColumnAdded or CsvDiffType.ColumnRemoved);
            var rowItems = _filteredDisplayRows.SelectMany(r => r.Details);
            return headerItems.Concat(rowItems)
                .OrderBy(i => i.GroupSortOrder)
                .ThenBy(i => i.RowNumber ?? int.MaxValue)
                .ThenBy(i => i.Locator, StringComparer.OrdinalIgnoreCase)
                .ThenBy(i => i.ColumnName, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private string BuildReport()
        {
            var sb = new StringBuilder();
            sb.AppendLine("CSV 对比报告");
            sb.AppendLine($"生成时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine($"对比模式: {(_currentMode == CsvCompareMode.ByKeyColumns ? "主键模式" : "行号模式")}");
            sb.AppendLine($"文件 A: {_leftFile?.LoadResult.FileName ?? "未选择"}");
            sb.AppendLine($"文件 B: {_rightFile?.LoadResult.FileName ?? "未选择"}");

            if (_headerComparison != null)
            {
                sb.AppendLine();
                sb.AppendLine($"字段对比：A 表 {_headerComparison.LeftColumnCount} 列 | B 表 {_headerComparison.RightColumnCount} 列 | 共同 {_headerComparison.MatchedCount} 列");
                if (_headerComparison.LeftOnlyColumns.Count > 0)
                    sb.AppendLine($"A 表独有（B 表缺失，{_headerComparison.LeftOnlyColumns.Count} 个）：{string.Join("、", _headerComparison.LeftOnlyColumns)}");
                if (_headerComparison.RightOnlyColumns.Count > 0)
                    sb.AppendLine($"B 表独有（A 表缺失，{_headerComparison.RightOnlyColumns.Count} 个）：{string.Join("、", _headerComparison.RightOnlyColumns)}");
            }

            sb.AppendLine();
            if (_lastResult == null) { sb.AppendLine("尚未执行对比。"); return sb.ToString(); }

            if (_lastResult.HasValidationErrors)
            {
                sb.AppendLine("校验失败：");
                foreach (string err in _lastResult.ValidationErrors) sb.AppendLine($"- {err}");
                sb.AppendLine();
                foreach (CsvCompareDisplayRow row in _allDisplayRows)
                {
                    sb.AppendLine($"[{row.DiffTypeText}] {row.Locator}");
                    sb.AppendLine($"摘要: {row.SummaryText}");
                    if (row.LeftPreview != "—") sb.AppendLine($"A 预览: {row.LeftPreview}");
                    if (row.RightPreview != "—") sb.AppendLine($"B 预览: {row.RightPreview}");
                    sb.AppendLine();
                }
                return sb.ToString().TrimEnd();
            }

            sb.AppendLine(TxtSummary.Text);
            if (!string.IsNullOrWhiteSpace(TxtScopeSummary.Text)) sb.AppendLine(TxtScopeSummary.Text);

            if (_filteredDisplayRows.Count == 0)
            {
                sb.AppendLine(); sb.AppendLine("当前筛选下没有行级差异。");
                return sb.ToString().TrimEnd();
            }

            sb.AppendLine();
            foreach (CsvCompareDisplayRow row in _filteredDisplayRows)
            {
                sb.AppendLine($"[{row.DiffTypeText}] {row.Locator}");
                sb.AppendLine($"摘要: {row.SummaryText}");
                if (row.LeftPreview != "—") sb.AppendLine($"A 预览: {row.LeftPreview}");
                if (row.RightPreview != "—") sb.AppendLine($"B 预览: {row.RightPreview}");
                foreach (CsvCompareDetailRow d in BuildDetailRows(row))
                    sb.AppendLine($"- {d.ColumnName}: A={FormatVal(d.LeftValue)} | B={FormatVal(d.RightValue)} | {d.Message}");
                sb.AppendLine();
            }
            return sb.ToString().TrimEnd();
        }

        private static string FormatVal(string v) => string.IsNullOrEmpty(v) ? "(空)" : v;

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
                    finally { CloseClipboard(); }
                }
                else { Clipboard.SetText(text); }
            }
            catch { try { Clipboard.SetText(text); } catch { } }
        }

        private static string FormatFileSize(long bytes) =>
            bytes >= 1024 * 1024 ? $"{bytes / 1024.0 / 1024.0:F1} MB" : $"{bytes / 1024.0:F1} KB";
    }
}
