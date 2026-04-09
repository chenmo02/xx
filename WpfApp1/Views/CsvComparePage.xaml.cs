using Microsoft.Win32;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
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

        private sealed class CsvCompareDisplayRow
        {
            public required CsvDiffType DiffGroupType { get; init; }
            public required string DiffTypeText { get; init; }
            public required string Locator { get; init; }
            public required string SummaryText { get; init; }
            public required int DiffCount { get; init; }
            public required int ChangedColumnCount { get; init; }
            public required string LeftPreview { get; init; }
            public required string RightPreview { get; init; }
            public required int GroupSortOrder { get; init; }
            public required string LocatorKey { get; init; }
            public int? RowNumber { get; init; }
            public required IReadOnlyList<CsvDiffItem> Details { get; init; }
        }

        private sealed class CsvCompareDetailRow
        {
            public required string ColumnName { get; init; }
            public required string LeftValue { get; init; }
            public required string RightValue { get; init; }
            public required string Message { get; init; }
        }

        private LoadedFileContext? _leftFile;
        private LoadedFileContext? _rightFile;
        private CsvCompareResult? _lastResult;
        private readonly List<int> _pageSizes = [50, 100, 200];
        private List<CsvCompareDisplayRow> _allDisplayRows = [];
        private List<CsvCompareDisplayRow> _filteredDisplayRows = [];
        private List<CsvCompareDisplayRow> _currentPageRows = [];
        private List<CsvDiffItem> _visibleHeaderChanges = [];
        private int _pageIndex;
        private int _pageSize = 100;
        private CsvCompareMode _currentMode = CsvCompareMode.ByRowNumber;

        public CsvComparePage()
        {
            InitializeComponent();
            CmbPageSize.ItemsSource = _pageSizes;
            CmbPageSize.SelectedItem = _pageSize;
            UpdateResponsiveLayout();
            UpdateModeUi();
            ResetResults();
            RestoreState();
        }

        private void Page_SizeChanged(object sender, SizeChangedEventArgs e) => UpdateResponsiveLayout();

        private async void BtnLoadLeft_Click(object sender, RoutedEventArgs e) => await LoadFileAsync(isLeft: true);

        private async void BtnLoadRight_Click(object sender, RoutedEventArgs e) => await LoadFileAsync(isLeft: false);

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
                SaveState();
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

        private void CompareMode_Checked(object sender, RoutedEventArgs e)
        {
            UpdateModeUi();
            UpdateCompareButtonState();
            if (RbCompareByKey == null)
            {
                return;
            }

            _currentMode = RbCompareByKey.IsChecked == true ? CsvCompareMode.ByKeyColumns : CsvCompareMode.ByRowNumber;
            SaveState();
        }

        private void FilterChanged(object sender, RoutedEventArgs e)
        {
            if (_lastResult == null)
            {
                return;
            }

            RenderResult(resetPage: true);
            SaveState();
        }

        private void CmbPageSize_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CmbPageSize?.SelectedItem is not int selectedPageSize || selectedPageSize == _pageSize)
            {
                return;
            }

            _pageSize = selectedPageSize;
            _pageIndex = 0;
            BindCurrentPage();
            SaveState();
        }

        private void BtnFirstPage_Click(object sender, RoutedEventArgs e)
        {
            if (_pageIndex == 0)
            {
                return;
            }

            _pageIndex = 0;
            BindCurrentPage();
            SaveState();
        }

        private void BtnPrevPage_Click(object sender, RoutedEventArgs e)
        {
            if (_pageIndex <= 0)
            {
                return;
            }

            _pageIndex--;
            BindCurrentPage();
            SaveState();
        }

        private void BtnNextPage_Click(object sender, RoutedEventArgs e)
        {
            if (_filteredDisplayRows.Count == 0)
            {
                return;
            }

            int totalPages = (int)Math.Ceiling(_filteredDisplayRows.Count / (double)_pageSize);
            if (_pageIndex >= totalPages - 1)
            {
                return;
            }

            _pageIndex++;
            BindCurrentPage();
            SaveState();
        }

        private void BtnLastPage_Click(object sender, RoutedEventArgs e)
        {
            if (_filteredDisplayRows.Count == 0)
            {
                return;
            }

            _pageIndex = Math.Max(0, (int)Math.Ceiling(_filteredDisplayRows.Count / (double)_pageSize) - 1);
            BindCurrentPage();
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

            _lastResult = CsvCompareService.Compare(_leftFile.Table, _rightFile.Table, _currentMode, selectedKeyColumns);
            BtnCopyReport.IsEnabled = true;
            BtnExportResult.IsEnabled = true;
            _pageIndex = 0;
            RenderResult(resetPage: true);
            SaveState();
        }

        private void DgSummaryRows_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            BindDetail((CsvCompareDisplayRow?)DgSummaryRows.SelectedItem);
        }

        private void UpdateResponsiveLayout()
        {
            if (FileCardsGrid == null || LeftFileCard == null || RightFileCard == null)
            {
                return;
            }

            bool useStackLayout = ActualWidth < 1120;
            if (useStackLayout)
            {
                FileCardsGrid.ColumnDefinitions[0].Width = new GridLength(1, GridUnitType.Star);
                FileCardsGrid.ColumnDefinitions[1].Width = new GridLength(0);
                FileCardsGrid.ColumnDefinitions[2].Width = new GridLength(0);
                FileCardsGrid.RowDefinitions[0].Height = GridLength.Auto;
                FileCardsGrid.RowDefinitions[1].Height = new GridLength(12);
                FileCardsGrid.RowDefinitions[2].Height = GridLength.Auto;
                Grid.SetRow(LeftFileCard, 0);
                Grid.SetColumn(LeftFileCard, 0);
                Grid.SetColumnSpan(LeftFileCard, 3);
                Grid.SetRow(RightFileCard, 2);
                Grid.SetColumn(RightFileCard, 0);
                Grid.SetColumnSpan(RightFileCard, 3);
            }
            else
            {
                FileCardsGrid.ColumnDefinitions[0].Width = new GridLength(1, GridUnitType.Star);
                FileCardsGrid.ColumnDefinitions[1].Width = new GridLength(12);
                FileCardsGrid.ColumnDefinitions[2].Width = new GridLength(1, GridUnitType.Star);
                FileCardsGrid.RowDefinitions[0].Height = GridLength.Auto;
                FileCardsGrid.RowDefinitions[1].Height = new GridLength(0);
                FileCardsGrid.RowDefinitions[2].Height = new GridLength(0);
                Grid.SetRow(LeftFileCard, 0);
                Grid.SetColumn(LeftFileCard, 0);
                Grid.SetColumnSpan(LeftFileCard, 1);
                Grid.SetRow(RightFileCard, 0);
                Grid.SetColumn(RightFileCard, 2);
                Grid.SetColumnSpan(RightFileCard, 1);
            }
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
            string metaText = $"{FormatFileSize(context.LoadResult.FileSize)} | {context.Table.Rows.Count} 行 × {context.Table.Columns.Count} 列 | 分隔符：{DelimitedTextFileService.GetDelimiterName(context.LoadResult.Delimiter)} | 编码：{context.LoadResult.Encoding.EncodingName}";
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
                : "行号模式会按第 N 行对第 N 行进行比较，并先汇总显示发生变化的行。";
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
            SaveState();
        }

        private List<string> GetSelectedKeyColumns()
        {
            if (KeyColumnsPanel == null)
            {
                return [];
            }

            return KeyColumnsPanel.Children.OfType<CheckBox>()
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

            BtnCompare.IsEnabled = RbCompareByKey.IsChecked == true
                ? GetSelectedKeyColumns().Count > 0
                : true;
        }

        private void RestoreState()
        {
            CsvCompareCachedState state = CsvCompareStateService.State;
            if (state.LeftFile == null && state.RightFile == null && state.LastResult == null)
            {
                return;
            }

            _leftFile = state.LeftFile == null ? null : FromCachedFile(state.LeftFile);
            _rightFile = state.RightFile == null ? null : FromCachedFile(state.RightFile);
            _lastResult = state.LastResult;
            _pageSize = _pageSizes.Contains(state.PageSize) ? state.PageSize : 100;
            _pageIndex = Math.Max(0, state.PageIndex);
            _currentMode = state.CurrentMode;

            RefreshFilePanels();
            RefreshKeyColumns();

            if (RbCompareByKey != null && RbCompareByRow != null)
            {
                RbCompareByKey.IsChecked = _currentMode == CsvCompareMode.ByKeyColumns;
                RbCompareByRow.IsChecked = _currentMode != CsvCompareMode.ByKeyColumns;
            }

            ApplySelectedKeyColumns(state.SelectedKeyColumns);
            UpdateModeUi();

            if (ChkShowColumnChanges != null) ChkShowColumnChanges.IsChecked = state.ShowColumnChanges;
            if (ChkShowRowChanges != null) ChkShowRowChanges.IsChecked = state.ShowRowChanges;
            if (ChkShowCellChanges != null) ChkShowCellChanges.IsChecked = state.ShowCellChanges;
            if (ChkExportFilteredOnly != null) ChkExportFilteredOnly.IsChecked = state.ExportFilteredOnly;
            if (CmbPageSize != null) CmbPageSize.SelectedItem = _pageSize;

            UpdateCompareButtonState();

            if (_lastResult != null)
            {
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

        private void ApplySelectedKeyColumns(IReadOnlyCollection<string> selectedColumns)
        {
            if (KeyColumnsPanel == null)
            {
                return;
            }

            foreach (CheckBox checkBox in KeyColumnsPanel.Children.OfType<CheckBox>())
            {
                string column = checkBox.Tag?.ToString() ?? string.Empty;
                checkBox.IsChecked = selectedColumns.Contains(column, StringComparer.OrdinalIgnoreCase);
            }
        }

        private static CsvCompareCachedFile ToCachedFile(LoadedFileContext context)
        {
            return new CsvCompareCachedFile
            {
                FileName = context.LoadResult.FileName,
                FileSize = context.LoadResult.FileSize,
                Delimiter = context.LoadResult.Delimiter,
                Encoding = context.LoadResult.Encoding,
                Table = context.Table.Copy()
            };
        }

        private static LoadedFileContext FromCachedFile(CsvCompareCachedFile cachedFile)
        {
            return new LoadedFileContext
            {
                LoadResult = new DelimitedTextLoadResult
                {
                    FilePath = cachedFile.FileName,
                    FileName = cachedFile.FileName,
                    FileSize = cachedFile.FileSize,
                    Encoding = cachedFile.Encoding,
                    Delimiter = cachedFile.Delimiter,
                    Table = cachedFile.Table.Copy()
                }
            };
        }

        private void ResetResults()
        {
            _lastResult = null;
            _allDisplayRows = [];
            _filteredDisplayRows = [];
            _currentPageRows = [];
            _visibleHeaderChanges = [];
            _pageIndex = 0;

            ValidationBanner.Visibility = Visibility.Collapsed;
            TxtValidationMessage.Text = string.Empty;
            TxtSummary.Text = "请先导入两个文件并执行对比。";
            TxtScopeSummary.Text = string.Empty;
            TxtHeaderChangesSummary.Text = string.Empty;
            HeaderChangesBanner.Visibility = Visibility.Collapsed;
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
            if (_lastResult == null)
            {
                ResetResults();
                return;
            }

            if (resetPage)
            {
                _pageIndex = 0;
            }

            if (_lastResult.HasValidationErrors)
            {
                RenderValidationResult();
                return;
            }

            ValidationBanner.Visibility = Visibility.Collapsed;
            TxtValidationMessage.Text = string.Empty;
            SetFilterControlsEnabled(true);

            _visibleHeaderChanges = ChkShowColumnChanges.IsChecked == true
                ? _lastResult.DiffItems.Where(item => item.DiffType is CsvDiffType.ColumnAdded or CsvDiffType.ColumnRemoved)
                    .OrderBy(item => item.GroupSortOrder)
                    .ThenBy(item => item.ColumnName, StringComparer.OrdinalIgnoreCase)
                    .ToList()
                : [];

            _allDisplayRows = BuildDisplayRows(_lastResult.DiffItems);
            _filteredDisplayRows = _allDisplayRows.Where(ShouldDisplayRowBeVisible).ToList();
            TxtSummary.Text = _lastResult.DiffItems.Count == 0
                ? "两个文件内容一致，没有发现任何差异。"
                : $"原始差异 {_lastResult.DiffItems.Count} 项 | 行新增 {_lastResult.RowAddedCount} | 行删除 {_lastResult.RowRemovedCount} | 单元格修改 {_lastResult.CellModifiedCount} | 表头变更 {_lastResult.ColumnAddedCount + _lastResult.ColumnRemovedCount}";

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
            _visibleHeaderChanges = [];
            _allDisplayRows = BuildDuplicateDisplayRows(_lastResult.DiffItems.Where(item => item.DiffType == CsvDiffType.DuplicateKey).ToList());
            _filteredDisplayRows = _allDisplayRows;
            TxtSummary.Text = $"发现 {_lastResult.DuplicateKeyCount} 条重复主键，已停止正常对比，请先修复数据。";
            BindCurrentPage();
        }

        private void SetFilterControlsEnabled(bool isEnabled)
        {
            ChkShowColumnChanges.IsEnabled = isEnabled;
            ChkShowRowChanges.IsEnabled = isEnabled;
            ChkShowCellChanges.IsEnabled = isEnabled;
            CmbPageSize.IsEnabled = isEnabled;
        }

        private void UpdateHeaderChangesBanner()
        {
            if (_visibleHeaderChanges.Count == 0)
            {
                HeaderChangesBanner.Visibility = Visibility.Collapsed;
                TxtHeaderChangesSummary.Text = string.Empty;
                return;
            }

            List<string> addedColumns = _visibleHeaderChanges.Where(item => item.DiffType == CsvDiffType.ColumnAdded).Select(item => item.ColumnName).ToList();
            List<string> removedColumns = _visibleHeaderChanges.Where(item => item.DiffType == CsvDiffType.ColumnRemoved).Select(item => item.ColumnName).ToList();
            var sections = new List<string>();
            if (addedColumns.Count > 0)
            {
                sections.Add($"新增列：{string.Join("、", addedColumns)}");
            }
            if (removedColumns.Count > 0)
            {
                sections.Add($"删除列：{string.Join("、", removedColumns)}");
            }

            TxtHeaderChangesSummary.Text = $"表头变更 {_visibleHeaderChanges.Count} 项。{string.Join("；", sections)}";
            HeaderChangesBanner.Visibility = Visibility.Visible;
        }

        private bool ShouldDisplayRowBeVisible(CsvCompareDisplayRow row)
        {
            return row.DiffGroupType switch
            {
                CsvDiffType.RowAdded or CsvDiffType.RowRemoved => ChkShowRowChanges.IsChecked == true,
                CsvDiffType.CellModified => ChkShowCellChanges.IsChecked == true,
                _ => true
            };
        }

        private void BindCurrentPage()
        {
            int totalItems = _filteredDisplayRows.Count;
            int totalPages = totalItems == 0 ? 0 : (int)Math.Ceiling(totalItems / (double)_pageSize);
            _pageIndex = totalPages == 0 ? 0 : Math.Clamp(_pageIndex, 0, totalPages - 1);
            _currentPageRows = totalItems == 0 ? [] : _filteredDisplayRows.Skip(_pageIndex * _pageSize).Take(_pageSize).ToList();

            TxtPageInfo.Text = totalPages == 0 ? "第 0 / 0 页" : $"第 {_pageIndex + 1} / {totalPages} 页";
            TxtScopeSummary.Text = _lastResult == null ? string.Empty : $"汇总结果 {_allDisplayRows.Count} | 当前筛选 {_filteredDisplayRows.Count} | 当前页 {_currentPageRows.Count}";

            if (_currentPageRows.Count == 0)
            {
                DgSummaryRows.Visibility = Visibility.Collapsed;
                DgSummaryRows.ItemsSource = null;
                TxtResultEmptyState.Visibility = Visibility.Visible;
                TxtResultEmptyState.Text = _lastResult?.DiffItems.Count == 0
                    ? "两个文件内容一致，没有发现任何差异。"
                    : _visibleHeaderChanges.Count > 0
                        ? "当前筛选下没有行级差异，表头变更请查看上方提示。"
                        : "当前筛选下没有可显示的差异结果。";
                BindDetail(null);
                return;
            }

            TxtResultEmptyState.Visibility = Visibility.Collapsed;
            DgSummaryRows.Visibility = Visibility.Visible;
            DgSummaryRows.ItemsSource = _currentPageRows;
            DgSummaryRows.SelectedItem = _currentPageRows[0];
            BindDetail(_currentPageRows[0]);
            SaveState();
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

            List<CsvCompareDetailRow> detailRows = BuildDetailRows(row);
            if (detailRows.Count == 0)
            {
                DgDetailRows.Visibility = Visibility.Collapsed;
                DgDetailRows.ItemsSource = null;
                TxtDetailEmptyState.Text = "当前项没有可显示的明细。";
                TxtDetailEmptyState.Visibility = Visibility.Visible;
                return;
            }

            TxtDetailEmptyState.Visibility = Visibility.Collapsed;
            DgDetailRows.Visibility = Visibility.Visible;
            DgDetailRows.ItemsSource = detailRows;
        }

        private static List<CsvCompareDisplayRow> BuildDisplayRows(IReadOnlyList<CsvDiffItem> diffItems)
        {
            return diffItems.Where(item => item.DiffType is CsvDiffType.RowAdded or CsvDiffType.RowRemoved or CsvDiffType.CellModified)
                .GroupBy(item => item.LocatorKey, StringComparer.Ordinal)
                .Select(group =>
                {
                    List<CsvDiffItem> details = group.OrderBy(item => string.IsNullOrWhiteSpace(item.ColumnName) ? 1 : 0)
                        .ThenBy(item => item.ColumnName, StringComparer.OrdinalIgnoreCase).ToList();
                    CsvDiffItem first = details[0];
                    int changedColumnCount = details.Where(item => !string.IsNullOrWhiteSpace(item.ColumnName))
                        .Select(item => item.ColumnName).Distinct(StringComparer.OrdinalIgnoreCase).Count();

                    return new CsvCompareDisplayRow
                    {
                        DiffGroupType = first.DiffType == CsvDiffType.CellModified ? CsvDiffType.CellModified : first.DiffType,
                        DiffTypeText = first.DiffTypeText,
                        Locator = first.Locator,
                        SummaryText = BuildSummaryText(first.DiffType, details, changedColumnCount),
                        DiffCount = details.Count,
                        ChangedColumnCount = changedColumnCount,
                        LeftPreview = FirstNonEmpty(details.Select(item => item.LeftRowPreview)),
                        RightPreview = FirstNonEmpty(details.Select(item => item.RightRowPreview)),
                        GroupSortOrder = first.GroupSortOrder,
                        LocatorKey = first.LocatorKey,
                        RowNumber = details.Select(item => item.RowNumber).FirstOrDefault(rowNumber => rowNumber.HasValue),
                        Details = details
                    };
                })
                .OrderBy(row => row.GroupSortOrder)
                .ThenBy(row => row.RowNumber ?? int.MaxValue)
                .ThenBy(row => row.Locator, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static List<CsvCompareDisplayRow> BuildDuplicateDisplayRows(IReadOnlyList<CsvDiffItem> diffItems)
        {
            return diffItems.Select(item => new CsvCompareDisplayRow
                {
                    DiffGroupType = CsvDiffType.DuplicateKey,
                    DiffTypeText = item.DiffTypeText,
                    Locator = item.Locator,
                    SummaryText = item.Message,
                    DiffCount = 1,
                    ChangedColumnCount = 0,
                    LeftPreview = item.LeftRowPreview,
                    RightPreview = item.RightRowPreview,
                    GroupSortOrder = item.GroupSortOrder,
                    LocatorKey = item.LocatorKey,
                    RowNumber = item.RowNumber,
                    Details = [item]
                })
                .OrderBy(row => row.Locator, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static string BuildSummaryText(CsvDiffType diffType, IReadOnlyList<CsvDiffItem> details, int changedColumnCount)
        {
            return diffType switch
            {
                CsvDiffType.RowAdded => "该记录仅存在于文件 B",
                CsvDiffType.RowRemoved => "该记录仅存在于文件 A",
                CsvDiffType.CellModified => BuildChangedColumnsSummary(details, changedColumnCount),
                _ => details[0].Message
            };
        }

        private static string BuildChangedColumnsSummary(IReadOnlyList<CsvDiffItem> details, int changedColumnCount)
        {
            if (changedColumnCount == 0)
            {
                return "该行存在内容差异";
            }

            List<string> columnNames = details.Where(item => !string.IsNullOrWhiteSpace(item.ColumnName))
                .Select(item => item.ColumnName).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            string previewColumns = string.Join("、", columnNames.Take(3));
            return columnNames.Count > 3 ? $"{changedColumnCount} 列不同：{previewColumns} 等" : $"{changedColumnCount} 列不同：{previewColumns}";
        }

        private static string FirstNonEmpty(IEnumerable<string> values)
        {
            string? value = values.FirstOrDefault(item => !string.IsNullOrWhiteSpace(item));
            return string.IsNullOrWhiteSpace(value) ? "—" : value;
        }

        private static List<CsvCompareDetailRow> BuildDetailRows(CsvCompareDisplayRow row)
        {
            if (row.Details.Count == 0)
            {
                return [];
            }

            if (row.DiffGroupType == CsvDiffType.CellModified)
            {
                return row.Details.Select(item => new CsvCompareDetailRow
                {
                    ColumnName = item.ColumnName,
                    LeftValue = item.LeftValue,
                    RightValue = item.RightValue,
                    Message = item.Message
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

        private IReadOnlyList<CsvDiffItem> GetExportDiffItems()
        {
            if (_lastResult == null)
            {
                return [];
            }

            if (_lastResult.HasValidationErrors)
            {
                return _lastResult.DiffItems.Where(item => item.DiffType == CsvDiffType.DuplicateKey).ToList();
            }

            if (ChkExportFilteredOnly.IsChecked != true)
            {
                return _lastResult.DiffItems;
            }

            return _visibleHeaderChanges.Concat(_filteredDisplayRows.SelectMany(row => row.Details))
                .OrderBy(item => item.GroupSortOrder)
                .ThenBy(item => item.RowNumber ?? int.MaxValue)
                .ThenBy(item => item.Locator, StringComparer.OrdinalIgnoreCase)
                .ThenBy(item => item.ColumnName, StringComparer.OrdinalIgnoreCase)
                .ToList();
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

            ExportService.ExportCsv(dialog.FileName, BuildExportTable(GetExportDiffItems()));
            MessageBox.Show($"结果已导出到：\n{dialog.FileName}", "导出成功", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private string BuildReport()
        {
            var builder = new StringBuilder();
            builder.AppendLine("CSV 对比报告");
            builder.AppendLine($"生成时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            builder.AppendLine($"对比模式: {(_currentMode == CsvCompareMode.ByKeyColumns ? "主键模式" : "行号模式")}");
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

                builder.AppendLine();
                foreach (CsvCompareDisplayRow row in _allDisplayRows)
                {
                    builder.AppendLine($"[{row.DiffTypeText}] {row.Locator}");
                    builder.AppendLine($"摘要: {row.SummaryText}");
                    if (row.LeftPreview != "—")
                    {
                        builder.AppendLine($"A 预览: {row.LeftPreview}");
                    }
                    if (row.RightPreview != "—")
                    {
                        builder.AppendLine($"B 预览: {row.RightPreview}");
                    }
                    builder.AppendLine();
                }

                return builder.ToString().TrimEnd();
            }

            builder.AppendLine(TxtSummary.Text);
            if (!string.IsNullOrWhiteSpace(TxtScopeSummary.Text))
            {
                builder.AppendLine(TxtScopeSummary.Text);
            }

            if (_visibleHeaderChanges.Count > 0)
            {
                builder.AppendLine();
                builder.AppendLine("表头变更：");
                builder.AppendLine(TxtHeaderChangesSummary.Text);
            }

            if (_filteredDisplayRows.Count == 0)
            {
                builder.AppendLine();
                builder.AppendLine("当前筛选下没有行级差异。");
                return builder.ToString().TrimEnd();
            }

            builder.AppendLine();
            foreach (CsvCompareDisplayRow row in _filteredDisplayRows)
            {
                builder.AppendLine($"[{row.DiffTypeText}] {row.Locator}");
                builder.AppendLine($"摘要: {row.SummaryText}");
                if (row.LeftPreview != "—")
                {
                    builder.AppendLine($"A 预览: {row.LeftPreview}");
                }
                if (row.RightPreview != "—")
                {
                    builder.AppendLine($"B 预览: {row.RightPreview}");
                }

                foreach (CsvCompareDetailRow detail in BuildDetailRows(row))
                {
                    builder.AppendLine($"- {detail.ColumnName}: A={FormatReportValue(detail.LeftValue)} | B={FormatReportValue(detail.RightValue)} | {detail.Message}");
                }

                builder.AppendLine();
            }

            return builder.ToString().TrimEnd();
        }

        private static string FormatReportValue(string value) => string.IsNullOrEmpty(value) ? "(空)" : value;

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
