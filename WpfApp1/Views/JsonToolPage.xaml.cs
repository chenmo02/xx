using Microsoft.Win32;
using System.Collections.ObjectModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using WpfApp1.Services;

namespace WpfApp1.Views
{
    public partial class JsonToolPage : Page
    {
        private readonly DispatcherTimer _debounceTimer = new();
        private readonly DispatcherTimer _gridSearchDebounceTimer = new();
        private static readonly Brush BorderColor = MakeBrush("#E5E7EB");
        private static readonly Brush HeaderBg = MakeBrush("#0EA5E9");
        private static readonly Brush HeaderBorderColor = MakeBrush("#0284C7");
        private static readonly Brush RowNumBg = MakeBrush("#F1F5F9");
        private static readonly Brush EvenRowBg = MakeBrush("#FFFDE7");
        private static readonly Brush LinkColor = MakeBrush("#2563EB");
        private static readonly Brush KeyColor = MakeBrush("#374151");
        private static readonly Brush HighlightBg = MakeBrush("#FBBF24");
        private static readonly Brush HighlightFg = MakeBrush("#1E293B");
        private static readonly Brush CurrentHighlightBg = MakeBrush("#F97316");
        private static readonly Brush NodeHighlightBg = MakeBrush("#FEF3C7");
        private static readonly Brush NormalJsonFg = MakeBrush("#A5F3FC");
        private static readonly Brush GridSearchHighlightBg = MakeBrush("#C4B5FD");
        private static readonly Brush GridSearchCurrentBg = MakeBrush("#7C3AED");
        private static readonly Brush GridSearchCurrentFg = Brushes.White;
        private static readonly Brush ClickHighlightBg = MakeBrush("#34D399");

        // 左侧搜索
        private List<TextRange> _searchMatches = new();
        private int _currentMatchIndex = -1;
        private bool _isUpdatingText = false;

        // GRID 节点追踪
        private ObservableCollection<JsonGridNode>? _currentNodes;
        private readonly Dictionary<string, (TextBlock header, FrameworkElement content, JsonGridNode node)> _gridSections = new();

        // 右侧 GRID 搜索
        private List<TextBlock> _gridSearchMatches = new();
        private int _gridSearchCurrentIndex = -1;
        private readonly List<(TextBlock tb, Brush originalBg, Brush originalFg)> _gridSearchHighlighted = new();

        // 点击联动
        private List<TextRange> _clickHighlights = new();

        // 所有 GRID 值 TextBlock
        private readonly List<TextBlock> _allGridValueTextBlocks = new();

        // 搜索结果限制（防止大 JSON 卡死）
        private const int MaxGridSearchMatches = 500;
        private const int MaxJsonHighlights = 200;
        private const int MaxGridAutoExpand = 50;
        // 每批展开的最大节点数
        private const int ExpandBatchSize = 8;
        // GRID 搜索取消令牌（用于中断上一次搜索）
        private CancellationTokenSource? _gridSearchCts;

        private sealed class EditorTextSegment
        {
            public required int Start { get; init; }
            public required int Length { get; init; }
            public required TextPointer Pointer { get; init; }
        }

        private sealed class EditorTextSnapshot
        {
            public required string Text { get; init; }
            public required List<EditorTextSegment> Segments { get; init; }
        }

        public JsonToolPage()
        {
            InitializeComponent();
            _debounceTimer.Interval = TimeSpan.FromMilliseconds(600);
            _debounceTimer.Tick += (s, e) => { _debounceTimer.Stop(); RebuildGrid(); };

            // GRID 搜索防抖 300ms
            _gridSearchDebounceTimer.Interval = TimeSpan.FromMilliseconds(300);
            _gridSearchDebounceTimer.Tick += (s, e) => { _gridSearchDebounceTimer.Stop(); PerformGridSearchAsync(); };
        }

        private void Page_Loaded(object sender, RoutedEventArgs e)
        {
            Focusable = true;
            Focus();
        }

        // ==================== RichTextBox 文本辅助 ====================

        private string GetEditorText()
        {
            var range = new TextRange(TxtJsonEditor.Document.ContentStart, TxtJsonEditor.Document.ContentEnd);
            return range.Text.TrimEnd();
        }

        private void SetEditorText(string text)
        {
            _isUpdatingText = true;
            TxtJsonEditor.Document.Blocks.Clear();
            var paragraph = new Paragraph();
            paragraph.Inlines.Add(new Run(text) { Foreground = NormalJsonFg });
            TxtJsonEditor.Document.Blocks.Add(paragraph);
            _isUpdatingText = false;
        }

        private void TxtJsonEditor_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_isUpdatingText) return;
            _debounceTimer.Stop();
            _debounceTimer.Start();
        }

        // ==================== Ctrl+F 智能判断 ====================

        private void FindCommand_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            if (IsGridFocused())
                ToggleGridSearchPanel(true);
            else
                ToggleSearchPanel(true);
        }

        private void EscCommand_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            if (GridSearchPanel.Visibility == Visibility.Visible)
                ToggleGridSearchPanel(false);
            else if (SearchPanel.Visibility == Visibility.Visible)
                ToggleSearchPanel(false);
        }

        private bool IsGridFocused()
        {
            var focused = Keyboard.FocusedElement as DependencyObject;
            while (focused != null)
            {
                if (focused == GridContainer || focused == GridScrollViewer || focused == TxtGridSearchKeyword)
                    return true;
                focused = VisualTreeHelper.GetParent(focused);
            }
            return false;
        }

        // ==================== 左侧 JSON 搜索 ====================

        private void BtnToggleSearch_Click(object sender, RoutedEventArgs e)
            => ToggleSearchPanel(SearchPanel.Visibility != Visibility.Visible);

        private void BtnCloseSearch_Click(object sender, RoutedEventArgs e)
            => ToggleSearchPanel(false);

        private void ToggleSearchPanel(bool show)
        {
            SearchPanel.Visibility = show ? Visibility.Visible : Visibility.Collapsed;
            if (show)
            {
                TxtSearchKeyword.Focus();
                TxtSearchKeyword.SelectAll();
            }
            else
            {
                ClearSearchHighlights();
                TxtSearchKeyword.Text = "";
                TxtSearchCount.Text = "";
                ClearGridHighlights();
            }
        }

        private void TxtSearchKeyword_TextChanged(object sender, TextChangedEventArgs e)
            => PerformSearch();

        private void SearchOptionChanged(object sender, RoutedEventArgs e)
            => PerformSearch();

        private void TxtSearchKeyword_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                NavigateMatch(Keyboard.Modifiers == ModifierKeys.Shift ? -1 : 1);
                e.Handled = true;
            }
            else if (e.Key == Key.Escape)
            {
                ToggleSearchPanel(false);
                e.Handled = true;
            }
        }

        private void BtnSearchPrev_Click(object sender, RoutedEventArgs e) => NavigateMatch(-1);
        private void BtnSearchNext_Click(object sender, RoutedEventArgs e) => NavigateMatch(1);

        private void PerformSearch()
        {
            _isUpdatingText = true;
            try
            {
                ClearSearchHighlights();
                ClearGridHighlights();
                _searchMatches.Clear();
                _currentMatchIndex = -1;

                string keyword = TxtSearchKeyword.Text?.Trim() ?? "";
                if (string.IsNullOrEmpty(keyword)) { TxtSearchCount.Text = ""; return; }
                StringComparison comparison = GetLeftSearchComparison();

                foreach (var range in FindEditorMatches(keyword, comparison, MaxJsonHighlights))
                {
                    _searchMatches.Add(range);
                    range.ApplyPropertyValue(TextElement.BackgroundProperty, HighlightBg);
                    range.ApplyPropertyValue(TextElement.ForegroundProperty, HighlightFg);
                }

                if (_searchMatches.Count > 0)
                {
                    _currentMatchIndex = 0;
                    HighlightCurrentMatch();
                    string suffix = _searchMatches.Count >= MaxJsonHighlights ? "+" : "";
                    TxtSearchCount.Text = $"1/{_searchMatches.Count}{suffix}";
                    HighlightGridMatches(keyword, comparison);
                }
                else
                {
                    TxtSearchCount.Text = "无匹配";
                }
            }
            finally { _isUpdatingText = false; }
        }

        private StringComparison GetLeftSearchComparison()
            => ChkSearchCaseSensitive?.IsChecked == true ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

        private List<TextRange> FindEditorMatches(string keyword, StringComparison comparison, int maxMatches)
        {
            var matches = new List<TextRange>();
            if (string.IsNullOrEmpty(keyword) || maxMatches <= 0)
            {
                return matches;
            }

            var snapshot = CaptureEditorTextSnapshot();
            if (string.IsNullOrEmpty(snapshot.Text))
            {
                return matches;
            }

            int startIndex = 0;
            while (matches.Count < maxMatches)
            {
                int matchIndex = snapshot.Text.IndexOf(keyword, startIndex, comparison);
                if (matchIndex < 0)
                {
                    break;
                }

                var range = CreateEditorTextRange(snapshot, matchIndex, keyword.Length);
                if (range != null)
                {
                    matches.Add(range);
                }

                startIndex = matchIndex + Math.Max(keyword.Length, 1);
            }

            return matches;
        }

        private EditorTextSnapshot CaptureEditorTextSnapshot()
        {
            var segments = new List<EditorTextSegment>();
            var textBuilder = new StringBuilder();
            var current = TxtJsonEditor.Document.ContentStart;
            var contentEnd = TxtJsonEditor.Document.ContentEnd;

            while (current != null && current.CompareTo(contentEnd) < 0)
            {
                if (current.GetPointerContext(LogicalDirection.Forward) == TextPointerContext.Text)
                {
                    string textRun = current.GetTextInRun(LogicalDirection.Forward);
                    if (!string.IsNullOrEmpty(textRun))
                    {
                        segments.Add(new EditorTextSegment
                        {
                            Start = textBuilder.Length,
                            Length = textRun.Length,
                            Pointer = current
                        });
                        textBuilder.Append(textRun);
                    }
                }

                current = current.GetNextContextPosition(LogicalDirection.Forward);
            }

            return new EditorTextSnapshot
            {
                Text = textBuilder.ToString(),
                Segments = segments
            };
        }

        private TextRange? CreateEditorTextRange(EditorTextSnapshot snapshot, int startOffset, int length)
        {
            if (startOffset < 0 || length <= 0)
            {
                return null;
            }

            var start = GetEditorTextPointerAtOffset(snapshot, startOffset);
            var end = GetEditorTextPointerAtOffset(snapshot, startOffset + length);
            if (start == null || end == null)
            {
                return null;
            }

            return new TextRange(start, end);
        }

        private TextPointer? GetEditorTextPointerAtOffset(EditorTextSnapshot snapshot, int offset)
        {
            if (offset < 0)
            {
                return null;
            }

            foreach (var segment in snapshot.Segments)
            {
                if (offset < segment.Start || offset > segment.Start + segment.Length)
                {
                    continue;
                }

                return segment.Pointer.GetPositionAtOffset(offset - segment.Start);
            }

            if (offset == snapshot.Text.Length)
            {
                var lastSegment = snapshot.Segments.LastOrDefault();
                if (lastSegment != null)
                {
                    return lastSegment.Pointer.GetPositionAtOffset(lastSegment.Length);
                }
            }

            return offset == 0 ? TxtJsonEditor.Document.ContentStart : null;
        }

        private void HighlightCurrentMatch()
        {
            if (_currentMatchIndex < 0 || _currentMatchIndex >= _searchMatches.Count) return;
            bool wasUpdating = _isUpdatingText;
            _isUpdatingText = true;
            try
            {
                foreach (var match in _searchMatches)
                    match.ApplyPropertyValue(TextElement.BackgroundProperty, HighlightBg);
                var current = _searchMatches[_currentMatchIndex];
                current.ApplyPropertyValue(TextElement.BackgroundProperty, CurrentHighlightBg);
                var rect = current.Start.GetCharacterRect(LogicalDirection.Forward);
                TxtJsonEditor.ScrollToVerticalOffset(TxtJsonEditor.VerticalOffset + rect.Top - TxtJsonEditor.ActualHeight / 3);
            }
            finally { _isUpdatingText = wasUpdating; }
        }

        private void NavigateMatch(int direction)
        {
            if (_searchMatches.Count == 0) return;
            _currentMatchIndex += direction;
            if (_currentMatchIndex >= _searchMatches.Count) _currentMatchIndex = 0;
            if (_currentMatchIndex < 0) _currentMatchIndex = _searchMatches.Count - 1;
            HighlightCurrentMatch();
            TxtSearchCount.Text = $"{_currentMatchIndex + 1}/{_searchMatches.Count}";
        }

        private void ClearSearchHighlights()
        {
            if (_searchMatches.Count == 0) { _currentMatchIndex = -1; return; }
            bool wasUpdating = _isUpdatingText;
            _isUpdatingText = true;
            try
            {
                foreach (var match in _searchMatches)
                {
                    match.ApplyPropertyValue(TextElement.BackgroundProperty, Brushes.Transparent);
                    match.ApplyPropertyValue(TextElement.ForegroundProperty, NormalJsonFg);
                }
                _searchMatches.Clear();
                _currentMatchIndex = -1;
            }
            finally { _isUpdatingText = wasUpdating; }
        }

        // ==================== 左侧搜索 → 右侧 GRID 联动（只高亮已渲染的，不强制展开） ====================

        private void HighlightGridMatches(string keyword, StringComparison comparison)
        {
            if (_currentNodes == null) return;
            ExpandMatchingGridSections(keyword, comparison, MaxGridAutoExpand);
            // 只高亮已展开的节点标题，不强制展开
            foreach (var kvp in _gridSections.ToList())
            {
                var (header, content, node) = kvp.Value;
                if (NodeContainsKeyword(node, keyword, comparison))
                {
                    header.Background = NodeHighlightBg;
                    header.FontWeight = FontWeights.Bold;
                }
            }
            // 只高亮已渲染的单元格
            HighlightGridCells(GridContainer, keyword, comparison);
        }

        private bool NodeContainsKeyword(JsonGridNode node, string keyword, StringComparison comparison = StringComparison.OrdinalIgnoreCase)
        {
            if (ContainsKeyword(node.Key, keyword, comparison) ||
                ContainsKeyword(node.Value, keyword, comparison))
                return true;
            foreach (var child in node.Children)
                if (NodeContainsKeyword(child, keyword, comparison)) return true;
            foreach (var row in node.TableRows)
                foreach (var cell in row.Cells)
                    if (CellContainsKeyword(cell, keyword, comparison))
                        return true;
            return false;
        }

        private bool CellContainsKeyword(JsonGridCell cell, string keyword, StringComparison comparison = StringComparison.OrdinalIgnoreCase)
        {
            if (ContainsKeyword(cell.ColumnName, keyword, comparison) ||
                ContainsKeyword(cell.Value, keyword, comparison) ||
                ContainsKeyword(cell.NestedSummary, keyword, comparison))
            {
                return true;
            }

            foreach (var nested in cell.NestedChildren)
            {
                if (NodeContainsKeyword(nested, keyword, comparison))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool ContainsKeyword(string? text, string keyword, StringComparison comparison)
            => !string.IsNullOrEmpty(text) && text.Contains(keyword, comparison);

        private void ExpandMatchingGridSections(string keyword, StringComparison comparison, int maxTotalExpand)
        {
            int totalExpanded = 0;
            bool expandedInPass;

            do
            {
                expandedInPass = false;
                foreach (var kvp in _gridSections.ToList())
                {
                    if (totalExpanded >= maxTotalExpand)
                    {
                        return;
                    }

                    var (header, content, node) = kvp.Value;
                    if (content.Visibility != Visibility.Collapsed || !NodeContainsKeyword(node, keyword, comparison))
                    {
                        continue;
                    }

                    ExpandGridSection(header, content, node);
                    totalExpanded++;
                    expandedInPass = true;
                }
            }
            while (expandedInPass && totalExpanded < maxTotalExpand);
        }

        private void ExpandGridSection(TextBlock header, FrameworkElement content, JsonGridNode node)
        {
            content.Visibility = Visibility.Visible;
            header.Text = header.Text.Replace("[+]", "[-]");

            if (content is StackPanel stackPanel && stackPanel.Children.Count == 0)
            {
                if (node.HasTable)
                {
                    RenderTable(node, stackPanel, 0);
                }
                else if (node.IsContainer)
                {
                    foreach (var child in node.Children)
                    {
                        RenderNode(child, stackPanel, 1);
                    }
                }
            }
        }

        private void HighlightGridCells(Panel container, string keyword, StringComparison comparison)
        {
            foreach (var child in container.Children)
            {
                if (child is Border border) HighlightInElement(border, keyword, comparison);
                else if (child is Grid grid)
                    foreach (var gc in grid.Children)
                        if (gc is Border b) HighlightInElement(b, keyword, comparison);
                else if (child is StackPanel sp) HighlightGridCells(sp, keyword, comparison);
            }
        }

        private void HighlightInElement(Border border, string keyword, StringComparison comparison)
        {
            if (border.Child is TextBlock tb)
            {
                if (ContainsKeyword(tb.Text, keyword, comparison))
                    tb.Background = NodeHighlightBg;
            }
            else if (border.Child is StackPanel sp)
            {
                foreach (var c in sp.Children)
                    if (c is TextBlock stb && ContainsKeyword(stb.Text, keyword, comparison))
                        stb.Background = NodeHighlightBg;
            }
            else if (border.Child is Grid g)
            {
                foreach (var gc in g.Children)
                    if (gc is Border gb) HighlightInElement(gb, keyword, comparison);
            }
        }

        private void ClearGridHighlights()
        {
            foreach (var kvp in _gridSections.ToList())
            {
                var (header, _, _) = kvp.Value;
                header.Background = Brushes.Transparent;
                header.FontWeight = FontWeights.SemiBold;
            }
            ClearCellHighlights(GridContainer);
        }

        private void ClearCellHighlights(Panel container)
        {
            foreach (var child in container.Children)
            {
                if (child is Border border) ClearBorderHighlight(border);
                else if (child is Grid grid)
                    foreach (var gc in grid.Children)
                        if (gc is Border b) ClearBorderHighlight(b);
                else if (child is StackPanel sp) ClearCellHighlights(sp);
            }
        }

        private void ClearBorderHighlight(Border border)
        {
            if (border.Child is TextBlock tb && tb.Background == NodeHighlightBg)
                tb.Background = Brushes.Transparent;
            else if (border.Child is StackPanel sp)
                foreach (var c in sp.Children)
                    if (c is TextBlock stb && stb.Background == NodeHighlightBg)
                        stb.Background = Brushes.Transparent;
        }

        // ==================== 右侧 GRID 搜索（防抖 + 分批展开 + 限量） ====================

        private void BtnToggleGridSearch_Click(object sender, RoutedEventArgs e)
            => ToggleGridSearchPanel(GridSearchPanel.Visibility != Visibility.Visible);

        private void BtnCloseGridSearch_Click(object sender, RoutedEventArgs e)
            => ToggleGridSearchPanel(false);

        private void ToggleGridSearchPanel(bool show)
        {
            GridSearchPanel.Visibility = show ? Visibility.Visible : Visibility.Collapsed;
            if (show)
            {
                TxtGridSearchKeyword.Focus();
                TxtGridSearchKeyword.SelectAll();
            }
            else
            {
                CancelGridSearch();
                ClearGridSearchHighlights();
                ClearClickHighlights();
                TxtGridSearchKeyword.Text = "";
                TxtGridSearchCount.Text = "";
            }
        }

        private void TxtGridSearchKeyword_TextChanged(object sender, TextChangedEventArgs e)
        {
            CancelGridSearch();
            _gridSearchDebounceTimer.Stop();
            _gridSearchDebounceTimer.Start();
        }

        private void GridSearchOptionChanged(object sender, RoutedEventArgs e)
        {
            CancelGridSearch();
            _gridSearchDebounceTimer.Stop();
            _gridSearchDebounceTimer.Start();
        }

        private void TxtGridSearchKeyword_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                _gridSearchDebounceTimer.Stop();
                if (_gridSearchMatches.Count == 0)
                {
                    CancelGridSearch();
                    PerformGridSearchAsync();
                }
                else
                    NavigateGridMatch(Keyboard.Modifiers == ModifierKeys.Shift ? -1 : 1);
                e.Handled = true;
            }
            else if (e.Key == Key.Escape)
            {
                ToggleGridSearchPanel(false);
                e.Handled = true;
            }
        }

        private void BtnGridSearchPrev_Click(object sender, RoutedEventArgs e) => NavigateGridMatch(-1);
        private void BtnGridSearchNext_Click(object sender, RoutedEventArgs e) => NavigateGridMatch(1);

        /// <summary>取消正在进行的 GRID 搜索</summary>
        private void CancelGridSearch()
        {
            _gridSearchCts?.Cancel();
            _gridSearchCts?.Dispose();
            _gridSearchCts = null;
        }

        /// <summary>GRID 搜索入口（防抖 Timer 调用）</summary>
        private void PerformGridSearchAsync()
        {
            ClearGridSearchHighlights();
            ClearClickHighlights();
            _gridSearchMatches.Clear();
            _gridSearchCurrentIndex = -1;

            string keyword = TxtGridSearchKeyword.Text?.Trim() ?? "";
            if (string.IsNullOrEmpty(keyword)) { TxtGridSearchCount.Text = ""; return; }
            StringComparison comparison = GetGridSearchComparison();

            TxtGridSearchCount.Text = "搜索中...";

            _gridSearchCts = new CancellationTokenSource();
            var token = _gridSearchCts.Token;

            // 启动分批展开搜索
            _ = ExpandAndSearchAsync(keyword, comparison, token);
        }

        /// <summary>
        /// 分批展开包含关键字的折叠节点，每批展开后搜索新渲染的 TextBlock。
        /// 使用 Dispatcher 让出 UI 线程，避免卡死。
        /// </summary>
        private StringComparison GetGridSearchComparison()
            => ChkGridSearchCaseSensitive?.IsChecked == true ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

        private async Task ExpandAndSearchAsync(string keyword, StringComparison comparison, CancellationToken token)
        {
            try
            {
                int totalExpanded = 0;
                int maxTotalExpand = 50; // 最多展开 50 个节点

                // 循环：每轮找出需要展开的折叠节点，分批展开
                while (totalExpanded < maxTotalExpand)
                {
                    if (token.IsCancellationRequested) return;

                    // 找出当前所有折叠且包含关键字的节点
                    var toExpand = new List<(TextBlock header, FrameworkElement content, JsonGridNode node)>();
                    foreach (var kvp in _gridSections.ToList())
                    {
                        if (token.IsCancellationRequested) return;
                        var (header, content, node) = kvp.Value;
                        if (content.Visibility == Visibility.Collapsed && NodeContainsKeyword(node, keyword, comparison))
                            toExpand.Add((header, content, node));
                    }

                    if (toExpand.Count == 0) break; // 没有更多需要展开的

                    // 分批展开
                    int batchCount = 0;
                    foreach (var (header, content, node) in toExpand)
                    {
                        if (token.IsCancellationRequested) return;
                        if (batchCount >= ExpandBatchSize)
                        {
                            // 让出 UI 线程，让界面有时间渲染
                            await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.Background);
                            batchCount = 0;
                        }

                        ExpandGridSection(header, content, node);

                        totalExpanded++;
                        batchCount++;

                        if (totalExpanded >= maxTotalExpand) break;
                    }

                    // 每轮展开后让 UI 喘息
                    await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.Background);
                }

                if (token.IsCancellationRequested) return;

                // 展开完成后，搜索所有可见的 TextBlock
                CollectGridSearchMatches(keyword, comparison, token);

                if (token.IsCancellationRequested) return;

                // 联动左侧 JSON
                if (_gridSearchMatches.Count > 0)
                    HighlightJsonForKeyword(keyword, comparison);
            }
            catch (OperationCanceledException) { }
        }

        /// <summary>在所有已渲染且可见的 TextBlock 中收集匹配项并高亮</summary>
        private void CollectGridSearchMatches(string keyword, StringComparison comparison, CancellationToken token)
        {
            int matchCount = 0;
            foreach (var tb in EnumerateSearchableGridTextBlocks(GridContainer))
            {
                if (token.IsCancellationRequested) return;
                if (matchCount >= MaxGridSearchMatches) break;
                if (string.IsNullOrEmpty(tb.Text)) continue;
                if (!ContainsKeyword(tb.Text, keyword, comparison)) continue;
                if (!IsTextBlockVisible(tb)) continue;

                _gridSearchMatches.Add(tb);
                _gridSearchHighlighted.Add((tb, tb.Background, tb.Foreground));
                tb.Background = GridSearchHighlightBg;
                matchCount++;
            }

            // 标记仍然折叠的节点（如果还有超出展开限制的）
            int collapsedMatchCount = 0;
            foreach (var kvp in _gridSections.ToList())
            {
                if (token.IsCancellationRequested) return;
                var (header, content, node) = kvp.Value;
                if (content.Visibility == Visibility.Collapsed && NodeContainsKeyword(node, keyword, comparison))
                {
                    header.Background = NodeHighlightBg;
                    header.FontWeight = FontWeights.Bold;
                    collapsedMatchCount++;
                }
            }

            if (_gridSearchMatches.Count > 0)
            {
                _gridSearchCurrentIndex = 0;
                HighlightCurrentGridMatch();
                string suffix = matchCount >= MaxGridSearchMatches ? "+" : "";
                string extra = collapsedMatchCount > 0 ? $" (还有 {collapsedMatchCount} 个折叠节点)" : "";
                TxtGridSearchCount.Text = $"1/{_gridSearchMatches.Count}{suffix}{extra}";
            }
            else if (collapsedMatchCount > 0)
            {
                TxtGridSearchCount.Text = $"展开高亮节点可查看 ({collapsedMatchCount} 个)";
            }
            else
            {
                TxtGridSearchCount.Text = "无匹配";
            }
        }

        /// <summary>检查 TextBlock 是否在可见的父容器链中</summary>
        private static IEnumerable<TextBlock> EnumerateSearchableGridTextBlocks(DependencyObject root)
        {
            var queue = new Queue<DependencyObject>();
            queue.Enqueue(root);

            while (queue.Count > 0)
            {
                DependencyObject current = queue.Dequeue();
                if (current is TextBlock textBlock)
                {
                    yield return textBlock;
                }

                int childCount = VisualTreeHelper.GetChildrenCount(current);
                for (int i = 0; i < childCount; i++)
                {
                    queue.Enqueue(VisualTreeHelper.GetChild(current, i));
                }
            }
        }

        private static bool IsTextBlockVisible(TextBlock tb)
        {
            DependencyObject? current = tb;
            while (current != null)
            {
                if (current is UIElement ui && ui.Visibility != Visibility.Visible)
                    return false;
                current = VisualTreeHelper.GetParent(current);
            }
            return true;
        }

        private void HighlightCurrentGridMatch()
        {
            if (_gridSearchCurrentIndex < 0 || _gridSearchCurrentIndex >= _gridSearchMatches.Count) return;

            // 恢复所有为浅紫色
            foreach (var tb in _gridSearchMatches)
            {
                tb.Background = GridSearchHighlightBg;
                var entry = _gridSearchHighlighted.FirstOrDefault(x => x.tb == tb);
                if (entry.tb != null) tb.Foreground = entry.originalFg;
            }

            // 当前匹配深紫色白字
            var current = _gridSearchMatches[_gridSearchCurrentIndex];
            current.Background = GridSearchCurrentBg;
            current.Foreground = GridSearchCurrentFg;
            current.BringIntoView();
        }

        private void NavigateGridMatch(int direction)
        {
            if (_gridSearchMatches.Count == 0) return;
            _gridSearchCurrentIndex += direction;
            if (_gridSearchCurrentIndex >= _gridSearchMatches.Count) _gridSearchCurrentIndex = 0;
            if (_gridSearchCurrentIndex < 0) _gridSearchCurrentIndex = _gridSearchMatches.Count - 1;
            HighlightCurrentGridMatch();
            TxtGridSearchCount.Text = $"{_gridSearchCurrentIndex + 1}/{_gridSearchMatches.Count}";
        }

        private void ClearGridSearchHighlights()
        {
            foreach (var (tb, originalBg, originalFg) in _gridSearchHighlighted)
            {
                tb.Background = originalBg;
                tb.Foreground = originalFg;
            }
            _gridSearchHighlighted.Clear();
            _gridSearchMatches.Clear();
            _gridSearchCurrentIndex = -1;

            // 清除折叠节点标题高亮
            foreach (var kvp in _gridSections.ToList())
            {
                var (header, _, _) = kvp.Value;
                if (header.Background == NodeHighlightBg)
                {
                    header.Background = Brushes.Transparent;
                    header.FontWeight = FontWeights.SemiBold;
                }
            }
        }

        /// <summary>GRID 搜索联动左侧 JSON（限量高亮，防止大文档卡死）</summary>
        private void HighlightJsonForKeyword(string keyword, StringComparison comparison)
        {
            _isUpdatingText = true;
            try
            {
                ClearClickHighlights();
                foreach (var range in FindEditorMatches(keyword, comparison, MaxJsonHighlights))
                {
                    _clickHighlights.Add(range);
                    range.ApplyPropertyValue(TextElement.BackgroundProperty, HighlightBg);
                    range.ApplyPropertyValue(TextElement.ForegroundProperty, HighlightFg);
                }
                if (_clickHighlights.Count > 0)
                {
                    var rect = _clickHighlights[0].Start.GetCharacterRect(LogicalDirection.Forward);
                    TxtJsonEditor.ScrollToVerticalOffset(TxtJsonEditor.VerticalOffset + rect.Top - TxtJsonEditor.ActualHeight / 3);
                }
            }
            finally { _isUpdatingText = false; }
        }

        // ==================== 右侧 GRID 点击值 → 左侧 JSON 联动 ====================

        private void OnGridValueClicked(string clickedValue)
        {
            if (string.IsNullOrEmpty(clickedValue) || clickedValue == "null") return;

            _isUpdatingText = true;
            try
            {
                ClearClickHighlights();

                foreach (var range in FindEditorMatches(clickedValue, StringComparison.Ordinal, MaxJsonHighlights))
                {
                    _clickHighlights.Add(range);
                    range.ApplyPropertyValue(TextElement.BackgroundProperty, ClickHighlightBg);
                    range.ApplyPropertyValue(TextElement.ForegroundProperty, HighlightFg);
                }

                if (_clickHighlights.Count > 0)
                {
                    var rect = _clickHighlights[0].Start.GetCharacterRect(LogicalDirection.Forward);
                    TxtJsonEditor.ScrollToVerticalOffset(TxtJsonEditor.VerticalOffset + rect.Top - TxtJsonEditor.ActualHeight / 3);
                    string suffix = _clickHighlights.Count >= MaxJsonHighlights ? "+" : "";
                    SetStatus($"🔗 已定位: \"{clickedValue}\" ({_clickHighlights.Count}{suffix} 处)");
                }
                else
                {
                    SetStatus($"⚠️ 未找到: \"{clickedValue}\"");
                }
            }
            finally { _isUpdatingText = false; }
        }

        private void ClearClickHighlights()
        {
            if (_clickHighlights.Count == 0) return;
            bool wasUpdating = _isUpdatingText;
            _isUpdatingText = true;
            try
            {
                foreach (var range in _clickHighlights)
                {
                    range.ApplyPropertyValue(TextElement.BackgroundProperty, Brushes.Transparent);
                    range.ApplyPropertyValue(TextElement.ForegroundProperty, NormalJsonFg);
                }
                _clickHighlights.Clear();
            }
            finally { _isUpdatingText = wasUpdating; }
        }

        // ==================== 核心：构建嵌套网格 ====================

        private void RebuildGrid()
        {
            GridContainer.Children.Clear();
            _gridSections.Clear();
            _allGridValueTextBlocks.Clear();
            string json = GetEditorText();
            if (string.IsNullOrEmpty(json)) { SetStatus("就绪"); return; }

            try
            {
                _currentNodes = JsonGridParser.Parse(json);
                if (_currentNodes.Count > 0)
                    RenderNode(_currentNodes[0], GridContainer, 0);
                SetStatus("✅ 已同步");
            }
            catch { SetStatus("⚠️ JSON 格式错误"); }
        }

        private void RenderNode(JsonGridNode node, Panel container, int depth)
        {
            if (node.Key == "root" && node.IsContainer)
            {
                foreach (var child in node.Children)
                    RenderNode(child, container, depth);
                if (node.HasTable)
                    RenderCollapsibleTable(node, container, depth);
                return;
            }

            if (node.HasTable)
                RenderCollapsibleTable(node, container, depth);
            else if (node.IsContainer)
                RenderCollapsibleObject(node, container, depth);
            else
                RenderLeafNode(node, container, depth);
        }

        private void RenderCollapsibleObject(JsonGridNode node, Panel container, int depth)
        {
            string collapsedLabel = node.ExpandLabel.Replace("[-]", "[+]");
            string expandedLabel = node.ExpandLabel;

            var headerText = new TextBlock
            {
                Text = collapsedLabel,
                FontWeight = FontWeights.SemiBold,
                Foreground = MakeBrush("#1E40AF"),
                FontSize = 13,
                Margin = new Thickness(depth * 20, 6, 0, 3),
                Cursor = Cursors.Hand
            };

            var contentPanel = new StackPanel { Visibility = Visibility.Collapsed };
            string sectionKey = $"obj_{node.JsonPath}_{_gridSections.Count}";
            _gridSections[sectionKey] = (headerText, contentPanel, node);
            _allGridValueTextBlocks.Add(headerText);

            headerText.MouseLeftButtonUp += (s, e) =>
            {
                if (contentPanel.Visibility == Visibility.Collapsed)
                {
                    contentPanel.Visibility = Visibility.Visible;
                    headerText.Text = expandedLabel;
                    if (contentPanel.Children.Count == 0)
                        foreach (var child in node.Children)
                            RenderNode(child, contentPanel, depth + 1);
                }
                else
                {
                    contentPanel.Visibility = Visibility.Collapsed;
                    headerText.Text = collapsedLabel;
                }
                e.Handled = true;
            };

            container.Children.Add(headerText);
            container.Children.Add(contentPanel);
        }

        private void RenderCollapsibleTable(JsonGridNode node, Panel container, int depth)
        {
            string collapsedLabel = node.ExpandLabel.Replace("[-]", "[+]");
            string expandedLabel = node.ExpandLabel;
            string displayLabel = node.Key == "root" ? "[+] 数据表格" : collapsedLabel;
            string displayExpandedLabel = node.Key == "root" ? "[-] 数据表格" : expandedLabel;

            var headerText = new TextBlock
            {
                Text = displayLabel,
                FontWeight = FontWeights.SemiBold,
                Foreground = MakeBrush("#1E40AF"),
                FontSize = 13,
                Margin = new Thickness(depth * 20, 6, 0, 3),
                Cursor = Cursors.Hand
            };

            var contentPanel = new StackPanel { Visibility = Visibility.Collapsed };
            string sectionKey = $"tbl_{node.JsonPath}_{_gridSections.Count}";
            _gridSections[sectionKey] = (headerText, contentPanel, node);
            _allGridValueTextBlocks.Add(headerText);

            headerText.MouseLeftButtonUp += (s, e) =>
            {
                if (contentPanel.Visibility == Visibility.Collapsed)
                {
                    contentPanel.Visibility = Visibility.Visible;
                    headerText.Text = displayExpandedLabel;
                    if (contentPanel.Children.Count == 0)
                        RenderTable(node, contentPanel, depth);
                }
                else
                {
                    contentPanel.Visibility = Visibility.Collapsed;
                    headerText.Text = displayLabel;
                }
                e.Handled = true;
            };

            container.Children.Add(headerText);
            container.Children.Add(contentPanel);
        }

        // ==================== 全部展开 / 全部折叠 ====================

        private void BtnExpandAll_Click(object sender, RoutedEventArgs e)
        {
            bool hasCollapsed = true;
            int maxIterations = 50;
            while (hasCollapsed && maxIterations-- > 0)
            {
                hasCollapsed = false;
                var snapshot = _gridSections.ToList();
                foreach (var kvp in snapshot)
                {
                    var (header, content, node) = kvp.Value;
                    if (content.Visibility == Visibility.Collapsed)
                    {
                        hasCollapsed = true;
                        content.Visibility = Visibility.Visible;
                        header.Text = header.Text.Replace("[+]", "[-]");
                        if (content is StackPanel sp && sp.Children.Count == 0)
                        {
                            if (node.HasTable) RenderTable(node, sp, 0);
                            else if (node.IsContainer)
                                foreach (var child in node.Children) RenderNode(child, sp, 1);
                        }
                    }
                }
            }
            SetStatus("📂 已全部展开");
        }

        private void BtnCollapseAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var kvp in _gridSections.ToList())
            {
                var (header, content, _) = kvp.Value;
                content.Visibility = Visibility.Collapsed;
                header.Text = header.Text.Replace("[-]", "[+]");
            }
            SetStatus("📁 已全部折叠");
        }

        // ==================== 叶子节点渲染 ====================

        private void RenderLeafNode(JsonGridNode node, Panel container, int depth)
        {
            var row = new Grid { Margin = new Thickness(depth * 20, 0, 0, 0) };
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(160) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            var keyTextBlock = new TextBlock
            {
                Text = node.Key, FontWeight = FontWeights.SemiBold,
                Foreground = KeyColor, FontSize = 13,
                VerticalAlignment = VerticalAlignment.Center,
                Cursor = Cursors.Hand,
                ToolTip = "点击定位到左侧 JSON"
            };
            string capturedKey = node.Key;
            keyTextBlock.MouseLeftButtonUp += (s, e) =>
            {
                OnGridValueClicked(capturedKey);
                e.Handled = true;
            };
            _allGridValueTextBlocks.Add(keyTextBlock);

            var keyBorder = new Border
            {
                BorderBrush = BorderColor,
                BorderThickness = new Thickness(0, 0, 0, 1),
                Padding = new Thickness(6, 4, 6, 4),
                Child = keyTextBlock
            };
            Grid.SetColumn(keyBorder, 0);
            row.Children.Add(keyBorder);

            var valTextBlock = new TextBlock
            {
                Text = node.Value,
                Foreground = MakeBrush(node.ValueColor),
                FontSize = 13,
                VerticalAlignment = VerticalAlignment.Center,
                Cursor = Cursors.Hand,
                ToolTip = "点击定位到左侧 JSON"
            };
            string capturedValue = node.Value;
            valTextBlock.MouseLeftButtonUp += (s, e) =>
            {
                OnGridValueClicked(capturedValue);
                e.Handled = true;
            };
            _allGridValueTextBlocks.Add(valTextBlock);

            var valBorder = new Border
            {
                BorderBrush = BorderColor,
                BorderThickness = new Thickness(0, 0, 0, 1),
                Padding = new Thickness(10, 4, 6, 4),
                Child = valTextBlock
            };
            Grid.SetColumn(valBorder, 1);
            row.Children.Add(valBorder);

            container.Children.Add(row);
        }

        private void AddLabel(string text, Panel container, int depth, bool isBold)
        {
            container.Children.Add(new TextBlock
            {
                Text = text,
                FontWeight = isBold ? FontWeights.SemiBold : FontWeights.Normal,
                Foreground = MakeBrush("#1E40AF"),
                FontSize = 13,
                Margin = new Thickness(depth * 20, 6, 0, 3)
            });
        }

        // ==================== 表格渲染 ====================

        private void RenderTable(JsonGridNode node, Panel container, int depth)
        {
            if (node.TableColumns.Count == 0 || node.TableRows.Count == 0) return;

            var tableBorder = new Border
            {
                BorderBrush = MakeBrush("#CBD5E1"),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(4),
                Margin = new Thickness(depth * 20, 2, 4, 8),
                HorizontalAlignment = HorizontalAlignment.Left
            };

            var tableGrid = new Grid();
            tableGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            foreach (var _ in node.TableColumns)
                tableGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto, MinWidth = 100 });
            tableGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            foreach (var _ in node.TableRows)
                tableGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var menuBtn = CreateHeaderCell("···");
            menuBtn.Cursor = Cursors.Hand;
            menuBtn.MouseLeftButtonUp += (s, e) => ExportNodeToCsv(node);
            menuBtn.ToolTip = "导出此表格为 CSV";
            Grid.SetRow(menuBtn, 0); Grid.SetColumn(menuBtn, 0);
            tableGrid.Children.Add(menuBtn);

            for (int c = 0; c < node.TableColumns.Count; c++)
            {
                var headerCell = CreateHeaderCell(node.TableColumns[c]);
                Grid.SetRow(headerCell, 0); Grid.SetColumn(headerCell, c + 1);
                tableGrid.Children.Add(headerCell);

                // 注册表头列名 TextBlock 到搜索列表
                if (headerCell.Child is TextBlock headerTb)
                    _allGridValueTextBlocks.Add(headerTb);
            }

            for (int r = 0; r < node.TableRows.Count; r++)
            {
                var row = node.TableRows[r];
                int gridRow = r + 1;

                var rowNum = new Border
                {
                    Background = RowNumBg, BorderBrush = BorderColor,
                    BorderThickness = new Thickness(0, 0, 1, 1),
                    Padding = new Thickness(8, 4, 8, 4),
                    Child = new TextBlock { Text = row.Index.ToString(), Foreground = Brushes.Gray, FontSize = 12, HorizontalAlignment = HorizontalAlignment.Center }
                };
                Grid.SetRow(rowNum, gridRow); Grid.SetColumn(rowNum, 0);
                tableGrid.Children.Add(rowNum);

                for (int c = 0; c < row.Cells.Count; c++)
                {
                    var cell = row.Cells[c];
                    var cellBorder = new Border
                    {
                        BorderBrush = BorderColor,
                        BorderThickness = new Thickness(0, 0, 1, 1),
                        Padding = new Thickness(8, 4, 8, 4),
                        MinWidth = 100
                    };
                    if (r % 2 == 0) cellBorder.Background = EvenRowBg;

                    if (cell.IsNested)
                    {
                        var nestedPanel = new StackPanel();
                        var summaryBtn = new TextBlock
                        {
                            Text = cell.NestedSummary,
                            Foreground = LinkColor,
                            FontWeight = FontWeights.SemiBold,
                            FontSize = 12,
                            Cursor = Cursors.Hand
                        };
                        var nestedContainer = new StackPanel { Visibility = Visibility.Collapsed };
                        summaryBtn.MouseLeftButtonUp += (s, e) =>
                        {
                            if (nestedContainer.Visibility == Visibility.Collapsed)
                            {
                                nestedContainer.Visibility = Visibility.Visible;
                                summaryBtn.Text = summaryBtn.Text.Replace("[+]", "[-]");
                                if (nestedContainer.Children.Count == 0)
                                    foreach (var nested in cell.NestedChildren)
                                        RenderNestedContent(nested, nestedContainer);
                            }
                            else
                            {
                                nestedContainer.Visibility = Visibility.Collapsed;
                                summaryBtn.Text = summaryBtn.Text.Replace("[-]", "[+]");
                            }
                            e.Handled = true;
                        };
                        nestedPanel.Children.Add(summaryBtn);
                        nestedPanel.Children.Add(nestedContainer);
                        cellBorder.Child = nestedPanel;
                    }
                    else
                    {
                        var valTb = new TextBlock
                        {
                            Text = cell.Value,
                            Foreground = MakeBrush(cell.ValueColor),
                            FontSize = 12,
                            VerticalAlignment = VerticalAlignment.Center,
                            TextTrimming = TextTrimming.CharacterEllipsis,
                            MaxWidth = 280,
                            Cursor = Cursors.Hand,
                            ToolTip = "点击定位到左侧 JSON"
                        };
                        string cellValue = cell.Value;
                        valTb.MouseLeftButtonUp += (s, e) =>
                        {
                            OnGridValueClicked(cellValue);
                            e.Handled = true;
                        };
                        _allGridValueTextBlocks.Add(valTb);
                        cellBorder.Child = valTb;
                    }

                    Grid.SetRow(cellBorder, gridRow); Grid.SetColumn(cellBorder, c + 1);
                    tableGrid.Children.Add(cellBorder);
                }
            }

            tableBorder.Child = tableGrid;
            container.Children.Add(tableBorder);
        }

        private void RenderNestedContent(JsonGridNode nested, Panel container)
        {
            if (nested.HasTable)
                RenderTable(nested, container, 0);
            else if (nested.NodeType == "Object")
                foreach (var child in nested.Children) RenderNode(child, container, 0);
            else
                RenderLeafNode(nested, container, 0);
        }

        // ==================== 导出 CSV ====================

        private void ExportNodeToCsv(JsonGridNode node)
        {
            if (node.TableColumns.Count == 0 || node.TableRows.Count == 0)
            { MessageBox.Show("没有可导出的表格数据", "提示", MessageBoxButton.OK, MessageBoxImage.Warning); return; }

            var sb = new StringBuilder();
            sb.AppendLine(string.Join(",", node.TableColumns.Select(EscapeCsvField)));
            foreach (var row in node.TableRows)
            {
                var fields = new List<string>();
                foreach (var cell in row.Cells)
                    fields.Add(EscapeCsvField(cell.IsNested ? cell.NestedSummary : cell.Value));
                sb.AppendLine(string.Join(",", fields));
            }

            var dlg = new SaveFileDialog
            {
                Filter = "CSV 文件|*.csv",
                FileName = $"{node.Key}_{DateTime.Now:yyyyMMdd_HHmmss}.csv",
                Title = $"导出 {node.Key} 表格为 CSV"
            };
            if (dlg.ShowDialog() == true)
            {
                File.WriteAllText(dlg.FileName, sb.ToString(), new UTF8Encoding(true));
                SetStatus($"✅ 已导出: {Path.GetFileName(dlg.FileName)}");
                MessageBox.Show($"CSV 已导出到：\n{dlg.FileName}", "导出成功", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private static string EscapeCsvField(string field)
        {
            if (string.IsNullOrEmpty(field)) return "";
            if (field.Contains(',') || field.Contains('"') || field.Contains('\n'))
                return $"\"{field.Replace("\"", "\"\"")}\"";
            return field;
        }

        // ==================== 辅助方法 ====================

        private Border CreateHeaderCell(string text)
        {
            return new Border
            {
                Background = HeaderBg,
                Padding = new Thickness(10, 5, 10, 5),
                BorderBrush = HeaderBorderColor,
                BorderThickness = new Thickness(0, 0, 1, 0),
                Child = new TextBlock { Text = text, Foreground = Brushes.White, FontWeight = FontWeights.SemiBold, FontSize = 12 }
            };
        }

        private static SolidColorBrush MakeBrush(string hex)
            => new((Color)ColorConverter.ConvertFromString(hex)!);

        // ==================== Win32 剪贴板 ====================

        [DllImport("user32.dll")]
        private static extern bool OpenClipboard(IntPtr hWndNewOwner);
        [DllImport("user32.dll")]
        private static extern bool CloseClipboard();
        [DllImport("user32.dll")]
        private static extern bool EmptyClipboard();
        [DllImport("user32.dll")]
        private static extern IntPtr SetClipboardData(uint uFormat, IntPtr hMem);
        [DllImport("kernel32.dll")]
        private static extern IntPtr GlobalAlloc(uint uFlags, UIntPtr dwBytes);
        [DllImport("kernel32.dll")]
        private static extern IntPtr GlobalLock(IntPtr hMem);
        [DllImport("kernel32.dll")]
        private static extern bool GlobalUnlock(IntPtr hMem);

        private static bool NativeCopy(string text)
        {
            if (!OpenClipboard(IntPtr.Zero)) return false;
            try
            {
                EmptyClipboard();
                var hGlobal = GlobalAlloc(0x0002, (UIntPtr)((text.Length + 1) * 2));
                if (hGlobal == IntPtr.Zero) return false;
                var target = GlobalLock(hGlobal);
                Marshal.Copy(text.ToCharArray(), 0, target, text.Length);
                Marshal.WriteInt16(target, text.Length * 2, 0);
                GlobalUnlock(hGlobal);
                SetClipboardData(13, hGlobal);
                return true;
            }
            finally { CloseClipboard(); }
        }

        // ==================== 格式化 / 压缩 / 校验 ====================

        private void BtnBeautify_Click(object sender, RoutedEventArgs e)
        {
            try { SetEditorText(JsonToolService.Beautify(GetEditorText())); SetStatus("✅ 格式化完成"); }
            catch (JsonException ex) { SetStatus($"❌ {ex.Message}"); }
        }

        private void BtnMinify_Click(object sender, RoutedEventArgs e)
        {
            try { SetEditorText(JsonToolService.Minify(GetEditorText())); SetStatus("✅ 压缩完成"); }
            catch (JsonException ex) { SetStatus($"❌ {ex.Message}"); }
        }

        private void BtnValidate_Click(object sender, RoutedEventArgs e)
        {
            var (isValid, message, _) = JsonToolService.Validate(GetEditorText());
            SetStatus(isValid ? "✅ JSON 格式正确" : "❌ 格式错误");
            MessageBox.Show(message, isValid ? "校验通过" : "校验失败", MessageBoxButton.OK,
                isValid ? MessageBoxImage.Information : MessageBoxImage.Warning);
        }

        // ==================== 文件导入导出 ====================

        private void BtnImportJson_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog { Filter = "JSON 文件|*.json|所有文件|*.*" };
            if (dlg.ShowDialog() == true)
            {
                string content = File.ReadAllText(dlg.FileName);
                try { content = JsonToolService.Beautify(content); } catch { }
                SetEditorText(content);
                SetStatus($"✅ 已导入: {Path.GetFileName(dlg.FileName)}");
            }
        }

        private void BtnExportJson_Click(object sender, RoutedEventArgs e)
        {
            string text = GetEditorText();
            if (string.IsNullOrWhiteSpace(text)) return;
            var dlg = new SaveFileDialog { Filter = "JSON 文件|*.json", FileName = $"export_{DateTime.Now:yyyyMMdd_HHmmss}.json" };
            if (dlg.ShowDialog() == true)
            {
                File.WriteAllText(dlg.FileName, text, Encoding.UTF8);
                SetStatus("✅ 已导出");
            }
        }

        private void BtnExportCsv_Click(object sender, RoutedEventArgs e)
        {
            string text = GetEditorText();
            if (string.IsNullOrWhiteSpace(text)) return;
            try
            {
                string csv = JsonToolService.JsonToCsv(text);
                var dlg = new SaveFileDialog { Filter = "CSV 文件|*.csv", FileName = $"export_{DateTime.Now:yyyyMMdd_HHmmss}.csv" };
                if (dlg.ShowDialog() == true)
                {
                    File.WriteAllText(dlg.FileName, csv, new UTF8Encoding(true));
                    SetStatus("✅ CSV 已导出");
                }
            }
            catch (Exception ex) { MessageBox.Show($"CSV 转换失败: {ex.Message}"); }
        }

        private void SetStatus(string msg) => TxtStatus.Text = msg;
    }
}
