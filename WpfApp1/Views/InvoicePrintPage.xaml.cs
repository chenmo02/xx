using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using System.Printing;
using WpfApp1.Services;

namespace WpfApp1.Views
{
    public partial class InvoicePrintPage : Page
    {
        [DllImport("user32.dll")] private static extern bool OpenClipboard(IntPtr hWndNewOwner);
        [DllImport("user32.dll")] private static extern bool CloseClipboard();
        [DllImport("user32.dll")] private static extern bool EmptyClipboard();
        [DllImport("user32.dll")] private static extern IntPtr SetClipboardData(uint uFormat, IntPtr hMem);

        private readonly InvoicePrintService _service = new();
        private readonly ObservableCollection<InvoiceFileItem> _fileItems = new();
        private readonly ObservableCollection<PrinterOption> _printers = new();
        private List<PrintTemplate> _templates = new();
        private double _zoomLevel = 1.0;
        private bool _isInitialized = false;
        private int _selectedLayout = 1; // 1, 2, 4

        private sealed class PrinterOption
        {
            public string Name { get; init; } = string.Empty;
            public string FullName { get; init; } = string.Empty;
        }

        public InvoicePrintPage()
        {
            InitializeComponent();
            FileListBox.ItemsSource = _fileItems;
            CmbPrinter.ItemsSource = _printers;
            LoadPrinters();
            LoadTemplates();
            UpdateLayoutCardSelection();
            _isInitialized = true;
        }

        private void LoadPrinters()
        {
            string? selectedName = (CmbPrinter.SelectedItem as PrinterOption)?.FullName;

            _printers.Clear();

            try
            {
                var server = new LocalPrintServer();
                var queues = server
                    .GetPrintQueues(new[]
                    {
                        EnumeratedPrintQueueTypes.Local,
                        EnumeratedPrintQueueTypes.Connections
                    })
                    .OrderBy(q => q.Name)
                    .ToList();

                foreach (var queue in queues)
                {
                    _printers.Add(new PrinterOption
                    {
                        Name = queue.Name,
                        FullName = queue.FullName
                    });
                }

                string? defaultName = server.DefaultPrintQueue?.FullName;
                var preferred = _printers.FirstOrDefault(p => string.Equals(p.FullName, selectedName, StringComparison.OrdinalIgnoreCase))
                    ?? _printers.FirstOrDefault(p => string.Equals(p.FullName, defaultName, StringComparison.OrdinalIgnoreCase))
                    ?? _printers.FirstOrDefault();

                if (preferred != null)
                    CmbPrinter.SelectedItem = preferred;
            }
            catch (Exception ex)
            {
                SetStatus($"打印机列表加载失败: {ex.Message}");
            }
        }

        private PrintQueue? GetSelectedPrinterQueue()
        {
            var selected = CmbPrinter.SelectedItem as PrinterOption;
            if (selected == null)
                return null;

            var server = new LocalPrintServer();
            return server
                .GetPrintQueues(new[]
                {
                    EnumeratedPrintQueueTypes.Local,
                    EnumeratedPrintQueueTypes.Connections
                })
                .FirstOrDefault(q =>
                    string.Equals(q.FullName, selected.FullName, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(q.Name, selected.Name, StringComparison.OrdinalIgnoreCase));
        }

        private void BtnRefreshPrinters_Click(object sender, RoutedEventArgs e)
        {
            LoadPrinters();

            if (_printers.Count > 0)
                SetStatus($"已刷新打印机列表，当前打印机：{((PrinterOption?)CmbPrinter.SelectedItem)?.Name}");
            else
                SetStatus("未找到可用的打印机。");
        }

        // ═══════════════════════════════════════
        // Ctrl + 鼠标滚轮缩放
        // ═══════════════════════════════════════

        private void PreviewScroller_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (Keyboard.Modifiers == ModifierKeys.Control)
            {
                e.Handled = true;
                _zoomLevel = e.Delta > 0
                    ? Math.Min(_zoomLevel + 0.1, 5.0)
                    : Math.Max(_zoomLevel - 0.1, 0.1);
                ApplyZoom();
            }
        }

        // ═══════════════════════════════════════
        // 排版方式卡片选择
        // ═══════════════════════════════════════

        private void LayoutCard1_Click(object sender, MouseButtonEventArgs e) { _selectedLayout = 1; UpdateLayoutCardSelection(); UpdatePreview(); }
        private void LayoutCard2_Click(object sender, MouseButtonEventArgs e) { _selectedLayout = 2; UpdateLayoutCardSelection(); UpdatePreview(); }
        private void LayoutCard4_Click(object sender, MouseButtonEventArgs e) { _selectedLayout = 4; UpdateLayoutCardSelection(); UpdatePreview(); }

        private void UpdateLayoutCardSelection()
        {
            var active = new SolidColorBrush(Color.FromRgb(78, 110, 242));   // #4E6EF2
            var inactive = new SolidColorBrush(Color.FromRgb(209, 213, 219)); // #D1D5DB
            var activeBg = new SolidColorBrush(Color.FromRgb(237, 240, 255)); // #EDF0FF
            var inactiveBg = new SolidColorBrush(Color.FromRgb(247, 248, 250)); // #F7F8FA

            LayoutCard1.BorderBrush = _selectedLayout == 1 ? active : inactive;
            LayoutCard1.Background = _selectedLayout == 1 ? activeBg : inactiveBg;
            LayoutCard2.BorderBrush = _selectedLayout == 2 ? active : inactive;
            LayoutCard2.Background = _selectedLayout == 2 ? activeBg : inactiveBg;
            LayoutCard4.BorderBrush = _selectedLayout == 4 ? active : inactive;
            LayoutCard4.Background = _selectedLayout == 4 ? activeBg : inactiveBg;
        }

        // ═══════════════════════════════════════
        // 纸张方向 & 裁剪线
        // ═══════════════════════════════════════

        private bool IsLandscape => RbLandscape.IsChecked == true;
        private bool ShowCutLine => ChkCutLine.IsChecked == true;

        private void Orientation_Changed(object sender, RoutedEventArgs e)
        {
            if (!_isInitialized) return;
            UpdatePreview();
        }

        private void CutLine_Changed(object sender, RoutedEventArgs e)
        {
            if (!_isInitialized) return;
            UpdatePreview();
        }

        // ═══════════════════════════════════════
        // 模板管理
        // ═══════════════════════════════════════

        private void LoadTemplates()
        {
            _templates = _service.LoadTemplates();
            var templateFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "invoice_templates.json");
            if (!File.Exists(templateFile))
            {
                foreach (var template in _templates)
                {
                    template.MarginTop = 0;
                    template.MarginBottom = 0;
                    template.MarginLeft = 0;
                    template.MarginRight = 0;
                }
            }

            CmbTemplate.Items.Clear();
            foreach (var t in _templates)
                CmbTemplate.Items.Add(new ComboBoxItem { Content = t.Name, Tag = t });
            if (CmbTemplate.Items.Count > 0)
                CmbTemplate.SelectedIndex = 0;
        }

        private PrintTemplate GetCurrentTemplate()
        {
            var t = new PrintTemplate();
            if (CmbPaperMode.SelectedItem is ComboBoxItem pi)
                t.PaperMode = pi.Tag?.ToString() ?? "A4";
            t.LayoutCount = _selectedLayout;
            double.TryParse(TxtMarginTop.Text, out var mt); t.MarginTop = mt;
            double.TryParse(TxtMarginBottom.Text, out var mb); t.MarginBottom = mb;
            double.TryParse(TxtMarginLeft.Text, out var ml); t.MarginLeft = ml;
            double.TryParse(TxtMarginRight.Text, out var mr); t.MarginRight = mr;
            double.TryParse(TxtOffsetX.Text, out var ox); t.OffsetX = ox;
            double.TryParse(TxtOffsetY.Text, out var oy); t.OffsetY = oy;
            t.PrintQuality = CmbQuality.SelectedIndex switch { 0 => "草稿", 2 => "高画质", _ => "标准" };
            return t;
        }

        private void ApplyTemplate(PrintTemplate t)
        {
            _isInitialized = false;
            CmbPaperMode.SelectedIndex = t.PaperMode == "Invoice" ? 1 : 0;
            UpdatePaperModeUI(t.PaperMode);
            _selectedLayout = t.LayoutCount;
            UpdateLayoutCardSelection();
            TxtMarginTop.Text = t.MarginTop.ToString();
            TxtMarginBottom.Text = t.MarginBottom.ToString();
            TxtMarginLeft.Text = t.MarginLeft.ToString();
            TxtMarginRight.Text = t.MarginRight.ToString();
            TxtOffsetX.Text = t.OffsetX.ToString();
            TxtOffsetY.Text = t.OffsetY.ToString();
            CmbQuality.SelectedIndex = t.PrintQuality switch { "草稿" => 0, "高画质" => 2, _ => 1 };
            _isInitialized = true;
            UpdatePreview();
        }

        private void CmbTemplate_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CmbTemplate.SelectedItem is ComboBoxItem item && item.Tag is PrintTemplate t)
                ApplyTemplate(t);
        }

        private void BtnSaveTemplate_Click(object sender, RoutedEventArgs e)
        {
            var template = GetCurrentTemplate();
            var name = ShowInputDialog("保存模板", "请输入模板名称：", template.Name);
            if (string.IsNullOrWhiteSpace(name)) return;
            template.Name = name;
            var idx = _templates.FindIndex(t => t.Name == name);
            if (idx >= 0) _templates[idx] = template; else _templates.Add(template);
            _service.SaveTemplates(_templates);
            LoadTemplates();
            for (int i = 0; i < CmbTemplate.Items.Count; i++)
                if (CmbTemplate.Items[i] is ComboBoxItem ci && ci.Content?.ToString() == name)
                { CmbTemplate.SelectedIndex = i; break; }
            SetStatus("✅ 模板已保存");
        }

        private void BtnDeleteTemplate_Click(object sender, RoutedEventArgs e)
        {
            if (CmbTemplate.SelectedItem is not ComboBoxItem item) return;
            var name = item.Content?.ToString();
            if (string.IsNullOrEmpty(name)) return;
            if (MessageBox.Show($"确定删除模板 \"{name}\"？", "确认", MessageBoxButton.YesNo) != MessageBoxResult.Yes) return;
            _templates.RemoveAll(t => t.Name == name);
            _service.SaveTemplates(_templates);
            LoadTemplates();
            SetStatus($"🗑️ 模板 \"{name}\" 已删除");
        }

        // ═══════════════════════════════════════
        // 文件导入 + 去重
        // ═══════════════════════════════════════

        private void BtnImportFiles_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Title = "选择发票文件",
                Filter = "发票文件|*.pdf;*.ofd;*.jpg;*.jpeg;*.png;*.bmp;*.tif;*.tiff|PDF|*.pdf|OFD|*.ofd|图片|*.jpg;*.jpeg;*.png;*.bmp",
                Multiselect = true
            };
            if (dlg.ShowDialog() != true) return;
            AddFiles(dlg.FileNames);
        }

        private void BtnImportFolder_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFolderDialog { Title = "选择发票文件夹" };
            if (dlg.ShowDialog() != true) return;
            var files = Directory.GetFiles(dlg.FolderName, "*.*", SearchOption.AllDirectories)
                .Where(f => InvoicePrintService.SupportedExtensions.Contains(Path.GetExtension(f).ToLowerInvariant())).ToArray();
            AddFiles(files);
        }

        private void AddFiles(string[] paths)
        {
            var items = _service.ImportFiles(paths);
            int added = 0;
            var dupNames = new List<string>();
            foreach (var item in items)
            {
                if (_fileItems.Any(f => f.FilePath == item.FilePath)) { dupNames.Add(item.FileName); continue; }
                var newHash = ComputeFileHash(item.FilePath);
                bool isDup = false;
                if (newHash != null)
                    foreach (var ex in _fileItems)
                    {
                        var exHash = ComputeFileHash(ex.FilePath);
                        if (exHash != null && newHash == exHash)
                        { dupNames.Add($"{item.FileName}（与 {ex.FileName} 内容相同）"); isDup = true; break; }
                    }
                if (!isDup) { _fileItems.Add(item); added++; }
            }
            UpdateFileCount();
            if (dupNames.Count > 0)
                MessageBox.Show($"以下 {dupNames.Count} 个文件已存在，已自动跳过：\n\n{string.Join("\n", dupNames.Select(n => $"  • {n}"))}", "⚠️ 重复文件提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            if (added > 0)
            {
                SetStatus($"📥 已导入 {added} 个文件" + (dupNames.Count > 0 ? $"，跳过 {dupNames.Count} 个重复" : ""));
                DropHintPanel.Visibility = Visibility.Collapsed;
                PreviewScroller.Visibility = Visibility.Visible;
                if (FileListBox.SelectedIndex < 0) FileListBox.SelectedIndex = 0;
            }
            else if (dupNames.Count > 0) SetStatus("⚠️ 所有文件均已存在");
            else SetStatus("⚠️ 没有找到支持的文件格式");
        }

        private static string? ComputeFileHash(string path)
        {
            try { using var s = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite); return Convert.ToHexString(SHA256.HashData(s)); }
            catch { return null; }
        }

        private void Page_DragOver(object sender, DragEventArgs e) { if (e.Data.GetDataPresent(DataFormats.FileDrop)) { e.Effects = DragDropEffects.Copy; e.Handled = true; } }
        private void Page_Drop(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;
            var paths = (string[])e.Data.GetData(DataFormats.FileDrop)!;
            var all = new List<string>();
            foreach (var p in paths)
            {
                if (Directory.Exists(p)) all.AddRange(Directory.GetFiles(p, "*.*", SearchOption.AllDirectories).Where(f => InvoicePrintService.SupportedExtensions.Contains(Path.GetExtension(f).ToLowerInvariant())));
                else if (File.Exists(p)) all.Add(p);
            }
            if (all.Count > 0) AddFiles(all.ToArray());
        }

        private void BtnClearList_Click(object sender, RoutedEventArgs e)
        {
            _fileItems.Clear(); UpdateFileCount(); LayoutPreviewGrid.Children.Clear();
            DropHintPanel.Visibility = Visibility.Visible; PreviewScroller.Visibility = Visibility.Collapsed;
            PanelPageNav.Visibility = Visibility.Collapsed; SetStatus("🗑️ 列表已清空");
        }
        private void BtnSelectAll_Click(object sender, RoutedEventArgs e) => FileListBox.SelectAll();
        private void UpdateFileCount() => TxtFileCount.Text = $"已导入 {_fileItems.Count} 个文件";

        // ═══════════════════════════════════════
        // 预览渲染（核心）
        // ═══════════════════════════════════════

        private void FileListBox_SelectionChanged(object sender, SelectionChangedEventArgs e) => UpdatePreview();

        private void FileListBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (FileListBox.SelectedItem is InvoiceFileItem item)
            {
                _fileItems.Remove(item);
                UpdateFileCount();
                UpdatePreview();
                if (_fileItems.Count == 0)
                {
                    LayoutPreviewGrid.Children.Clear();
                    DropHintPanel.Visibility = Visibility.Visible;
                    PreviewScroller.Visibility = Visibility.Collapsed;
                    PanelPageNav.Visibility = Visibility.Collapsed;
                    SetStatus("🗑️ 列表已清空");
                }
                else
                {
                    SetStatus($"🗑️ 已移除 {item.FileName}");
                }
            }
        }

        private void UpdatePreview()
        {
            if (!_isInitialized) return;
            if (_fileItems.Count == 0) { DropHintPanel.Visibility = Visibility.Visible; PreviewScroller.Visibility = Visibility.Collapsed; return; }
            DropHintPanel.Visibility = Visibility.Collapsed;
            PreviewScroller.Visibility = Visibility.Visible;

            try
            {
                var template = GetCurrentTemplate();
                // 预览始终显示所有文件的排版效果（与打印行为一致）
                var previewItems = _fileItems.ToList();

                // PDF 分页
                if (FileListBox.SelectedItems.Count == 1 && FileListBox.SelectedItem is InvoiceFileItem si)
                {
                    var ext = Path.GetExtension(si.FilePath).ToLowerInvariant();
                    if (ext == ".pdf") { int pc = InvoicePrintService.GetPdfPageCount(si.FilePath); si.PageCount = pc; PanelPageNav.Visibility = pc > 1 ? Visibility.Visible : Visibility.Collapsed; TxtPageInfo.Text = $"第 {si.SelectedPage + 1} / {pc} 页"; }
                    else PanelPageNav.Visibility = Visibility.Collapsed;
                }
                else PanelPageNav.Visibility = Visibility.Collapsed;

                var pages = RenderLayoutPreviewPages(previewItems, template);
                LayoutPreviewGrid.Children.Clear();
                LayoutPreviewGrid.RowDefinitions.Clear();

                if (pages.Count == 0) return;

                for (int pi = 0; pi < pages.Count; pi++)
                {
                    // 每页一行，页间留 20px 间距
                    LayoutPreviewGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

                    var pagePanel = new StackPanel();

                    // 页码标签
                    if (pages.Count > 1)
                    {
                        var pageLabel = new TextBlock
                        {
                            Text = $"第 {pi + 1} / {pages.Count} 页",
                            FontSize = 12,
                            Foreground = new SolidColorBrush(Color.FromRgb(120, 120, 120)),
                            HorizontalAlignment = HorizontalAlignment.Center,
                            Margin = new Thickness(0, pi == 0 ? 0 : 16, 0, 6)
                        };
                        pagePanel.Children.Add(pageLabel);
                    }

                    // 页面图片（带阴影边框）
                    var pageBorder = new Border
                    {
                        Background = Brushes.White,
                        Margin = new Thickness(0, 0, 0, 8),
                        Effect = new System.Windows.Media.Effects.DropShadowEffect
                        {
                            BlurRadius = 8, ShadowDepth = 2, Opacity = 0.12
                        }
                    };
                    var img = new System.Windows.Controls.Image { Source = pages[pi], Stretch = Stretch.Uniform };
                    RenderOptions.SetBitmapScalingMode(img, BitmapScalingMode.HighQuality);
                    pageBorder.Child = img;
                    pagePanel.Children.Add(pageBorder);

                    Grid.SetRow(pagePanel, pi);
                    LayoutPreviewGrid.Children.Add(pagePanel);
                }

                PreviewContainer.LayoutTransform = new ScaleTransform(_zoomLevel, _zoomLevel);
                TxtZoomLevel.Text = $"{(int)(_zoomLevel * 100)}%";
                if (pages.Count > 1)
                    SetStatus($"📄 共 {pages.Count} 页预览 · {previewItems.Count} 个文件");
            }
            catch (Exception ex) { SetStatus($"❌ 预览失败: {ex.Message}"); }
        }

        /// <summary>
        /// 渲染排版预览：支持多页分页、纸张方向、裁剪线
        /// 返回每页一张 BitmapSource 的列表
        /// </summary>
        private List<BitmapSource> RenderLayoutPreviewPages(List<InvoiceFileItem> items, PrintTemplate template)
        {
            var result = new List<BitmapSource>();
            double mmToWpf = 96.0 / 25.4;

            // 纸张尺寸
            double paperW, paperH;
            if (template.PaperMode == "Invoice")
            { paperW = 241 * mmToWpf; paperH = 140 * mmToWpf; }
            else
            { paperW = 210 * mmToWpf; paperH = 297 * mmToWpf; }

            // 横向时交换宽高
            if (IsLandscape) { (paperW, paperH) = (paperH, paperW); }

            double ml = template.MarginLeft * mmToWpf, mr = template.MarginRight * mmToWpf;
            double mt = template.MarginTop * mmToWpf, mb = template.MarginBottom * mmToWpf;
            double ox = template.OffsetX * mmToWpf, oy = template.OffsetY * mmToWpf;
            double contentW = paperW - ml - mr, contentH = paperH - mt - mb;

            int perPage = template.PaperMode == "Invoice" ? 1 : _selectedLayout;
            int cols = perPage == 4 ? 2 : 1;
            int rows = perPage >= 2 ? 2 : 1;
            double gap = 3 * mmToWpf;
            double totalGapW = gap * Math.Max(0, cols - 1);
            double totalGapH = gap * Math.Max(0, rows - 1);
            double cellW = Math.Max(0, (contentW - totalGapW) / cols);
            double cellH = Math.Max(0, (contentH - totalGapH) / rows);

            // 如果没有文件，渲染一页空白模板
            int totalPages = items.Count == 0 ? 1 : (int)Math.Ceiling((double)items.Count / perPage);

            for (int pageIdx = 0; pageIdx < totalPages; pageIdx++)
            {
                var dv = new DrawingVisual();
                using (var dc = dv.RenderOpen())
                {
                    // 纸张背景
                    dc.DrawRectangle(Brushes.White, new Pen(new SolidColorBrush(Color.FromRgb(200, 200, 200)), 1),
                        new Rect(0, 0, paperW, paperH));

                    // 边距参考线
                    var marginPen = new Pen(new SolidColorBrush(Color.FromArgb(50, 78, 110, 242)), 1) { DashStyle = DashStyles.Dash };
                    dc.DrawRectangle(null, marginPen, new Rect(ml, mt, contentW, contentH));

                // 绘制每个发票槽位
                for (int j = 0; j < perPage; j++)
                {
                    int globalIdx = pageIdx * perPage + j; // 全局文件索引
                    int col = j % cols, row = j / cols;
                    double x = ml + col * (cellW + gap) + ox;
                    double y = mt + row * (cellH + gap) + oy;
                    double w = cellW;
                    double h = cellH;

                    // 槽位背景
                    var slotPen = new Pen(new SolidColorBrush(Color.FromArgb(60, 150, 150, 150)), 1) { DashStyle = DashStyles.Dot };
                    dc.DrawRectangle(new SolidColorBrush(Color.FromArgb(6, 0, 0, 0)), slotPen, new Rect(x, y, w, h));

                    if (globalIdx < items.Count)
                    {
                        var bmp = GetPreviewBitmap(items[globalIdx]);
                        if (bmp != null)
                        {
                            double pad = 4, aw = w - pad * 2, ah = h - pad * 2;
                            double sc = Math.Min(aw / bmp.PixelWidth, ah / bmp.PixelHeight);
                            double dw = bmp.PixelWidth * sc, dh = bmp.PixelHeight * sc;
                            dc.DrawImage(bmp, new Rect(x + (w - dw) / 2, y + (h - dh) / 2, dw, dh));
                        }
                    }
                    else
                    {
                        var ft = new FormattedText($"发票 {globalIdx + 1}", System.Globalization.CultureInfo.CurrentCulture,
                            FlowDirection.LeftToRight, new Typeface("Microsoft YaHei"), 14,
                            new SolidColorBrush(Color.FromRgb(180, 180, 180)),
                            VisualTreeHelper.GetDpi(dv).PixelsPerDip);
                        dc.DrawText(ft, new Point(x + (w - ft.Width) / 2, y + (h - ft.Height) / 2));
                    }
                }

                // ═══ 裁剪线 ═══
                if (ShowCutLine && perPage > 1)
                {
                    var cutPen = new Pen(new SolidColorBrush(Color.FromArgb(180, 150, 150, 150)), 1) { DashStyle = DashStyles.Dash };
                    double scissorSize = 10;
                    var scissorBrush = new SolidColorBrush(Color.FromRgb(150, 150, 150));

                    // 水平裁剪线（2张或4张时，中间横线）
                    if (rows == 2)
                    {
                        double cy = mt + cellH + gap / 2;
                        dc.DrawLine(cutPen, new Point(0, cy), new Point(paperW, cy));
                        // 剪刀符号
                        var ft = new FormattedText("✂", System.Globalization.CultureInfo.CurrentCulture,
                            FlowDirection.LeftToRight, new Typeface("Segoe UI Symbol"), scissorSize, scissorBrush,
                            VisualTreeHelper.GetDpi(dv).PixelsPerDip);
                        dc.DrawText(ft, new Point(4, cy - ft.Height / 2));
                    }

                    // 垂直裁剪线（4张时，中间竖线）
                    if (cols == 2)
                    {
                        double cx = ml + cellW + gap / 2;
                        dc.DrawLine(cutPen, new Point(cx, 0), new Point(cx, paperH));
                        var ft = new FormattedText("✂", System.Globalization.CultureInfo.CurrentCulture,
                            FlowDirection.LeftToRight, new Typeface("Segoe UI Symbol"), scissorSize, scissorBrush,
                            VisualTreeHelper.GetDpi(dv).PixelsPerDip);
                        dc.DrawText(ft, new Point(cx - ft.Width / 2, 4));
                    }
                }

                // 底部信息
                var orient = IsLandscape ? "横向" : "纵向";
                var pageInfo = totalPages > 1 ? $" · 第{pageIdx + 1}/{totalPages}页" : "";
                var info = new FormattedText(
                    $"{template.PaperMode} · {orient} · {perPage}张/页{pageInfo}" + (ShowCutLine ? " · 裁剪线" : ""),
                    System.Globalization.CultureInfo.CurrentCulture, FlowDirection.LeftToRight,
                    new Typeface("Microsoft YaHei"), 10, new SolidColorBrush(Color.FromRgb(160, 160, 160)),
                    VisualTreeHelper.GetDpi(dv).PixelsPerDip);
                dc.DrawText(info, new Point(ml, paperH - mb + 4));
                }

                var rtb = new RenderTargetBitmap((int)paperW, (int)paperH, 96, 96, PixelFormats.Pbgra32);
                rtb.Render(dv);
                rtb.Freeze();
                result.Add(rtb);
            }

            return result;
        }

        private BitmapSource? GetPreviewBitmap(InvoiceFileItem item)
        {
            try
            {
                var ext = Path.GetExtension(item.FilePath).ToLowerInvariant();
                BitmapSource? img = ext == ".pdf"
                    ? InvoicePrintService.RenderPdfPage(item.FilePath, item.SelectedPage, 150)
                    : (item.PreviewImage ?? (item.PreviewImage = InvoicePrintService.LoadPreviewImage(item.FilePath, ext)));
                if (img != null && item.RotationAngle != 0) img = InvoicePrintService.RotateImage(img, item.RotationAngle);
                return img;
            }
            catch { return null; }
        }

        // ── 旋转 & 缩放 ──
        private void BtnRotate_Click(object sender, RoutedEventArgs e) { if (FileListBox.SelectedItem is InvoiceFileItem item) { item.RotationAngle = (item.RotationAngle + 90) % 360; item.PreviewImage = null; UpdatePreview(); } }
        private void BtnZoomIn_Click(object sender, RoutedEventArgs e) { _zoomLevel = Math.Min(_zoomLevel + 0.2, 5.0); ApplyZoom(); }
        private void BtnZoomOut_Click(object sender, RoutedEventArgs e) { _zoomLevel = Math.Max(_zoomLevel - 0.2, 0.1); ApplyZoom(); }
        private void BtnZoomReset_Click(object sender, RoutedEventArgs e) { _zoomLevel = 1.0; ApplyZoom(); }
        private void ApplyZoom() { PreviewContainer.LayoutTransform = new ScaleTransform(_zoomLevel, _zoomLevel); TxtZoomLevel.Text = $"{(int)(_zoomLevel * 100)}%"; }

        // ── PDF 分页 ──
        private void BtnPrevPage_Click(object sender, RoutedEventArgs e) { if (FileListBox.SelectedItem is InvoiceFileItem item && item.SelectedPage > 0) { item.SelectedPage--; UpdatePreview(); } }
        private void BtnNextPage_Click(object sender, RoutedEventArgs e) { if (FileListBox.SelectedItem is InvoiceFileItem item && item.SelectedPage < item.PageCount - 1) { item.SelectedPage++; UpdatePreview(); } }

        // ═══════════════════════════════════════
        // 设置面板事件
        // ═══════════════════════════════════════

        private void CmbPaperMode_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (!_isInitialized) return;
            var mode = CmbPaperMode.SelectedIndex == 1 ? "Invoice" : "A4";
            UpdatePaperModeUI(mode);
            UpdatePreview();
        }

        private void UpdatePaperModeUI(string mode)
        {
            if (PanelLayoutCount == null || PanelOffset == null) return;
            PanelLayoutCount.Visibility = mode == "Invoice" ? Visibility.Collapsed : Visibility.Visible;
            PanelOffset.Visibility = mode == "Invoice" ? Visibility.Visible : Visibility.Collapsed;
        }

        private void LayoutChanged(object sender, SelectionChangedEventArgs e) { } // 不再使用 ComboBox

        private DispatcherTimer? _marginDebounceTimer;
        private void MarginChanged(object sender, TextChangedEventArgs e)
        {
            if (!_isInitialized) return;
            _marginDebounceTimer?.Stop();
            _marginDebounceTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(500) };
            _marginDebounceTimer.Tick += (_, _) => { _marginDebounceTimer.Stop(); UpdatePreview(); };
            _marginDebounceTimer.Start();
        }

        // ═══════════════════════════════════════
        // 打印
        // ═══════════════════════════════════════

        private void BtnPrint_Click(object sender, RoutedEventArgs e)
        {
            // 始终打印所有导入的文件（与预览一致）
            var printItems = _fileItems.ToList();
            if (printItems.Count == 0) { MessageBox.Show("请先导入发票文件", "提示", MessageBoxButton.OK, MessageBoxImage.Information); return; }
            if (!int.TryParse(TxtCopies.Text.Trim(), out var copies) || copies < 1) copies = 1;
            if (copies > 99) copies = 99;
            var template = GetCurrentTemplate();
            try
            {
                SetStatus("⏳ 正在准备打印...");
                var dlg = new PrintDialog();

                // 设置纸张方向
                if (IsLandscape)
                    dlg.PrintTicket.PageOrientation = System.Printing.PageOrientation.Landscape;
                else
                    dlg.PrintTicket.PageOrientation = System.Printing.PageOrientation.Portrait;

                if (dlg.ShowDialog() != true) { SetStatus("❌ 打印已取消"); return; }
                var pageSize = new Size(dlg.PrintableAreaWidth, dlg.PrintableAreaHeight);
                var pages = InvoicePrintService.BuildPrintPages(printItems, template, pageSize);
                if (pages.Count == 0) { SetStatus("❌ 没有可打印的内容"); return; }
                bool ok = InvoicePrintService.PrintPages(pages, copies, dlg);
                if (ok) { foreach (var it in printItems) it.IsPrinted = true; _service.RecordPrintHistory(printItems); SetStatus($"✅ 打印完成！共 {pages.Count} 页 × {copies} 份"); }
                else SetStatus("❌ 打印失败");
            }
            catch (Exception ex) { SetStatus($"❌ 打印出错: {ex.Message}"); MessageBox.Show($"打印失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error); }
        }

        // ═══════════════════════════════════════
        // 工具方法
        // ═══════════════════════════════════════

        private void BtnPrint_Click2(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!TryPreparePrintJob(showDialog: false, out var printItems, out var pages, out var copies, out var context))
                    return;

                var printerName = context.PrintQueue?.Name ?? "当前打印机";
                var confirm = MessageBox.Show(
                    $"将按当前页面设置发送到打印机“{printerName}”，共 {pages.Count} 页 × {copies} 份。是否继续？",
                    "确认打印",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);

                if (confirm != MessageBoxResult.Yes)
                {
                    SetStatus("已取消当前打印。");
                    return;
                }

                bool ok = InvoicePrintService.PrintPages(pages, copies, context);
                if (ok)
                {
                    foreach (var it in printItems) it.IsPrinted = true;
                    _service.RecordPrintHistory(printItems);
                    SetStatus($"打印完成，共 {pages.Count} 页 x {copies} 份。");
                }
                else
                {
                    SetStatus("打印失败。");
                }
            }
            catch (Exception ex)
            {
                SetStatus($"打印出错: {ex.Message}");
                MessageBox.Show($"打印失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnPrinterPreview_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!TryPreparePrintJob(showDialog: false, out _, out var pages, out _, out var context))
                    return;

                var doc = InvoicePrintService.BuildFixedDocument(pages, context);
                ShowPrinterPreviewWindow(doc, context, pages.Count);
                SetStatus($"已生成打印机预览，共 {pages.Count} 页。");
            }
            catch (Exception ex)
            {
                SetStatus($"预览出错: {ex.Message}");
                MessageBox.Show($"打印机预览失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private bool TryPreparePrintJob(
            bool showDialog,
            out List<InvoiceFileItem> printItems,
            out List<DrawingVisual> pages,
            out int copies,
            out InvoicePrintService.PrintLayoutContext context)
        {
            printItems = _fileItems.ToList();
            pages = new List<DrawingVisual>();
            context = new InvoicePrintService.PrintLayoutContext();

            if (printItems.Count == 0)
            {
                MessageBox.Show("请先导入发票文件", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                copies = 0;
                return false;
            }

            if (!int.TryParse(TxtCopies.Text.Trim(), out copies) || copies < 1)
                copies = 1;
            if (copies > 99)
                copies = 99;

            var template = GetCurrentTemplate();
            SetStatus(showDialog ? "正在准备打印..." : "正在准备打印机预览...");

            var dlg = CreateConfiguredPrintDialog();
            if (showDialog && dlg.ShowDialog() != true)
            {
                SetStatus("已取消当前操作。");
                return false;
            }

            context = InvoicePrintService.CreatePrintLayoutContext(dlg, template, IsLandscape);
            pages = InvoicePrintService.BuildPrintPages(printItems, template, context.ContentSize);

            if (pages.Count == 0)
            {
                SetStatus("没有可输出的页面内容。");
                return false;
            }

            return true;
        }

        private PrintDialog CreateConfiguredPrintDialog()
        {
            var dlg = new PrintDialog();
            var selectedQueue = GetSelectedPrinterQueue();
            if (selectedQueue == null)
                throw new InvalidOperationException("请先选择可用的打印机。");

            dlg.PrintQueue = selectedQueue;
            dlg.PrintTicket = selectedQueue.DefaultPrintTicket ?? new PrintTicket();

            dlg.PrintTicket.PageOrientation = IsLandscape
                ? PageOrientation.Landscape
                : PageOrientation.Portrait;

            return dlg;
        }

        private void ShowPrinterPreviewWindow(FixedDocument document, InvoicePrintService.PrintLayoutContext context, int pageCount)
        {
            var toolbar = new DockPanel
            {
                Margin = new Thickness(16, 12, 16, 12),
                LastChildFill = false
            };

            var info = new TextBlock
            {
                Text = $"{context.PrintQueue?.Name ?? "当前打印机"} · {pageCount} 页 · {(IsLandscape ? "横向" : "纵向")}",
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(Color.FromRgb(51, 51, 51)),
                VerticalAlignment = VerticalAlignment.Center
            };
            DockPanel.SetDock(info, Dock.Left);
            toolbar.Children.Add(info);

            var closeButton = new Button
            {
                Content = "关闭",
                Padding = new Thickness(18, 6, 18, 6),
                Margin = new Thickness(8, 0, 0, 0),
                MinWidth = 88
            };

            var viewer = new DocumentViewer
            {
                Document = document,
                Margin = new Thickness(16, 0, 16, 16),
                Background = new SolidColorBrush(Color.FromRgb(241, 245, 249))
            };

            var layout = new DockPanel();
            DockPanel.SetDock(toolbar, Dock.Top);
            layout.Children.Add(toolbar);
            layout.Children.Add(viewer);

            var win = new Window
            {
                Title = "打印机预览",
                Width = 1100,
                Height = 780,
                MinWidth = 860,
                MinHeight = 620,
                Content = layout,
                Owner = Application.Current.MainWindow,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Background = Brushes.White
            };

            closeButton.Click += (_, _) => win.Close();
            toolbar.Children.Add(closeButton);

            win.ShowDialog();
        }

        private void SetStatus(string text) => TxtStatus.Text = text;

        private static string? ShowInputDialog(string title, string prompt, string defaultValue = "")
        {
            var win = new Window { Title = title, Width = 380, Height = 170, WindowStartupLocation = WindowStartupLocation.CenterOwner, Owner = Application.Current.MainWindow, ResizeMode = ResizeMode.NoResize, WindowStyle = WindowStyle.ToolWindow };
            var sp = new StackPanel { Margin = new Thickness(16) };
            sp.Children.Add(new TextBlock { Text = prompt, FontSize = 13, Margin = new Thickness(0, 0, 0, 8) });
            var tb = new TextBox { Text = defaultValue, FontSize = 13, Padding = new Thickness(8, 6, 8, 6) };
            sp.Children.Add(tb);
            var bp = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 12, 0, 0) };
            string? result = null;
            var ok = new Button { Content = "确定", Padding = new Thickness(20, 6, 20, 6), IsDefault = true };
            ok.Click += (_, _) => { result = tb.Text; win.Close(); };
            var cancel = new Button { Content = "取消", Padding = new Thickness(20, 6, 20, 6), Margin = new Thickness(8, 0, 0, 0), IsCancel = true };
            cancel.Click += (_, _) => win.Close();
            bp.Children.Add(ok); bp.Children.Add(cancel); sp.Children.Add(bp);
            win.Content = sp; tb.Focus(); tb.SelectAll(); win.ShowDialog();
            return result;
        }
    }
}
