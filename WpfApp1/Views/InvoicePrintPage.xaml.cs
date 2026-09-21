using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
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
        private bool _isLeftPanelHidden = false;
        private int _selectedLayout = 1; // 1, 2, 4
        private readonly Dictionary<string, string?> _fileHashes = new(StringComparer.OrdinalIgnoreCase);
        private CancellationTokenSource? _previewCancellation;
        private int _previewVersion;
        private bool _isImporting;
        private bool _isPrintBusy;
        private bool _isLoadingPrinters;
        private int _importVersion;
        private CancellationTokenSource? _importCancellation;
        private readonly SemaphoreSlim _previewGate = new(1, 1);
        private readonly SemaphoreSlim _printerPreviewGate = new(1, 1);
        private readonly Dictionary<InvoiceFileItem, Task<int>> _pdfPageCounts = new();
        private int _navigationVersion;

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
            _ = LoadPrintersAsync();
            LoadTemplates();
            UpdateLayoutCardSelection();
            _isInitialized = true;
            Unloaded += (_, _) =>
            {
                _previewCancellation?.Cancel();
                _marginDebounceTimer?.Stop();
                _importVersion++;
                _importCancellation?.Cancel();
                _navigationVersion++;
            };
            Loaded += (_, _) => UpdatePreview();
        }

        private async Task LoadPrintersAsync()
        {
            if (_isLoadingPrinters) return;
            _isLoadingPrinters = true;
            string? selectedName = (CmbPrinter.SelectedItem as PrinterOption)?.FullName;

            try
            {
                SetStatus("正在加载打印机列表...");
                var result = await InvoicePrintService.RunPrintWorkerAsync(() =>
                {
                    using var server = new LocalPrintServer();
                    using var queues = server.GetPrintQueues(new[]
                    { EnumeratedPrintQueueTypes.Local, EnumeratedPrintQueueTypes.Connections });
                    var options = new List<PrinterOption>();
                    foreach (var queue in queues)
                    {
                        using (queue) options.Add(new PrinterOption { Name = queue.Name, FullName = queue.FullName });
                    }
                    using var defaultQueue = server.DefaultPrintQueue;
                    return (options, defaultName: defaultQueue?.FullName);
                });
                selectedName = (CmbPrinter.SelectedItem as PrinterOption)?.FullName ?? selectedName;
                _printers.Clear();
                foreach (var option in result.options.OrderBy(p => p.Name)) _printers.Add(option);
                string? defaultName = result.defaultName;
                var preferred = _printers.FirstOrDefault(p => string.Equals(p.FullName, selectedName, StringComparison.OrdinalIgnoreCase))
                    ?? _printers.FirstOrDefault(p => string.Equals(p.FullName, defaultName, StringComparison.OrdinalIgnoreCase))
                    ?? _printers.FirstOrDefault();

                if (preferred != null)
                    CmbPrinter.SelectedItem = preferred;
                SetStatus(preferred == null ? "未找到可用的打印机。" : $"已刷新打印机列表，当前打印机：{preferred.Name}");
            }
            catch (Exception ex)
            {
                SetStatus($"打印机列表加载失败: {ex.Message}");
            }
            finally { _isLoadingPrinters = false; }
        }

        private async void BtnRefreshPrinters_Click(object sender, RoutedEventArgs e)
        {
            await LoadPrintersAsync();
        }

        private void BtnToggleLeftPanel_Click(object sender, RoutedEventArgs e)
        {
            _isLeftPanelHidden = !_isLeftPanelHidden;
            LeftFilePanel.Visibility = _isLeftPanelHidden ? Visibility.Collapsed : Visibility.Visible;
            LeftPanelColumn.Width = _isLeftPanelHidden ? new GridLength(0) : new GridLength(240);
            BtnToggleLeftPanel.Content = _isLeftPanelHidden ? "显示左侧" : "隐藏左侧";
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
            t.IsLandscape = IsLandscape;
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
            RbPortrait.IsChecked = !t.IsLandscape;
            RbLandscape.IsChecked = t.IsLandscape;
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
            ToastService.Show(this, "模板已保存");
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
            ToastService.Show(this, "模板已删除");
        }

        // ═══════════════════════════════════════
        // 文件导入 + 去重
        // ═══════════════════════════════════════

        private async void BtnImportFiles_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Title = "选择发票文件",
                Filter = "发票文件|*.pdf;*.ofd;*.jpg;*.jpeg;*.png;*.bmp;*.tif;*.tiff|PDF|*.pdf|OFD|*.ofd|图片|*.jpg;*.jpeg;*.png;*.bmp",
                Multiselect = true
            };
            if (dlg.ShowDialog() != true) return;
            await AddFilesAsync(dlg.FileNames);
        }

        private async void BtnImportFolder_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFolderDialog { Title = "选择发票文件夹" };
            if (dlg.ShowDialog() != true) return;
            await AddFilesAsync(new[] { dlg.FolderName });
        }

        private async Task AddFilesAsync(string[] paths)
        {
            if (_isImporting)
            {
                SetStatus("正在导入文件，请稍候...");
                return;
            }

            _isImporting = true;
            int version = ++_importVersion;
            using var cancellation = new CancellationTokenSource();
            _importCancellation = cancellation;
            try
            {
                SetStatus("正在导入文件...");
                var existingPaths = _fileItems.Select(f => f.FilePath).ToHashSet(StringComparer.OrdinalIgnoreCase);
                var existingHashes = _fileHashes.Values.Where(h => h != null).Cast<string>().ToHashSet(StringComparer.Ordinal);

                var result = await Task.Run(() =>
                {
                    var addedItems = new List<(InvoiceFileItem Item, string? Hash)>();
                    var duplicates = new List<string>();
                    var files = paths.SelectMany(path => Directory.Exists(path)
                        ? Directory.EnumerateFiles(path, "*", new EnumerationOptions
                        { RecurseSubdirectories = true, IgnoreInaccessible = true, AttributesToSkip = FileAttributes.ReparsePoint })
                        : new[] { path });
                    foreach (var item in _service.ImportFiles(files, cancellation.Token))
                    {
                        cancellation.Token.ThrowIfCancellationRequested();
                        if (!existingPaths.Add(item.FilePath))
                        {
                            duplicates.Add(item.FileName);
                            continue;
                        }

                        var hash = ComputeFileHash(item.FilePath);
                        if (hash != null && !existingHashes.Add(hash))
                        {
                            duplicates.Add(item.FileName);
                            continue;
                        }

                        addedItems.Add((item, hash));
                    }
                    return (addedItems, duplicates);
                });

                if (version != _importVersion) return;
                int applied = 0;
                foreach (var (item, hash) in result.addedItems)
                {
                    if (version != _importVersion) return;
                    _fileItems.Add(item);
                    _fileHashes[item.FilePath] = hash;
                    if (++applied % 50 == 0)
                    {
                        UpdateFileCount();
                        await Dispatcher.Yield(DispatcherPriority.Background);
                    }
                }

                if (version != _importVersion) return;
                UpdateFileCount();
                if (result.duplicates.Count > 0)
                    MessageBox.Show($"以下 {result.duplicates.Count} 个文件已存在，已自动跳过（最多展示 20 项）：\n\n{string.Join("\n", result.duplicates.Take(20).Select(n => $"  • {n}"))}", "⚠️ 重复文件提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                if (result.addedItems.Count > 0)
                {
                    SetStatus($"📥 已导入 {result.addedItems.Count} 个文件" + (result.duplicates.Count > 0 ? $"，跳过 {result.duplicates.Count} 个重复" : ""));
                    ToastService.Show(this, $"已导入 {result.addedItems.Count} 个文件");
                    DropHintPanel.Visibility = Visibility.Collapsed;
                    PreviewScroller.Visibility = Visibility.Visible;
                    if (FileListBox.SelectedIndex < 0) FileListBox.SelectedIndex = 0;
                    UpdatePreview();
                }
                else if (result.duplicates.Count > 0) SetStatus("⚠️ 所有文件均已存在");
                else SetStatus("⚠️ 没有找到支持的文件格式");
            }
            catch (OperationCanceledException) { }
            catch (Exception ex)
            {
                if (version == _importVersion) SetStatus($"❌ 导入失败: {ex.Message}");
            }
            finally { _isImporting = false; _importCancellation = null; }
        }

        private static string? ComputeFileHash(string path)
        {
            try { using var s = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite); return Convert.ToHexString(SHA256.HashData(s)); }
            catch { return null; }
        }

        private void Page_DragOver(object sender, DragEventArgs e) { if (e.Data.GetDataPresent(DataFormats.FileDrop)) { e.Effects = DragDropEffects.Copy; e.Handled = true; } }
        private async void Page_Drop(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;
            var paths = (string[])e.Data.GetData(DataFormats.FileDrop)!;
            await AddFilesAsync(paths);
        }

        private void BtnClearList_Click(object sender, RoutedEventArgs e)
        {
            _importVersion++;
            _importCancellation?.Cancel();
            _previewCancellation?.Cancel();
            _fileItems.Clear(); _fileHashes.Clear(); _pdfPageCounts.Clear(); _navigationVersion++;
            UpdateFileCount(); LayoutPreviewGrid.Children.Clear(); LayoutPreviewGrid.RowDefinitions.Clear();
            DropHintPanel.Visibility = Visibility.Visible; PreviewScroller.Visibility = Visibility.Collapsed;
            PanelPageNav.Visibility = Visibility.Collapsed; SetStatus("🗑️ 列表已清空");
            ToastService.Show(this, "列表已清空");
        }
        private void BtnSelectAll_Click(object sender, RoutedEventArgs e) => FileListBox.SelectAll();
        private void UpdateFileCount() => TxtFileCount.Text = $"已导入 {_fileItems.Count} 个文件";

        // ═══════════════════════════════════════
        // 预览渲染（核心）
        // ═══════════════════════════════════════

        private void FileListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_isImporting) UpdatePageNavigation();
        }

        private void FileListBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (_isImporting) return;
            if (FileListBox.SelectedItem is InvoiceFileItem item)
            {
                _fileItems.Remove(item);
                _fileHashes.Remove(item.FilePath);
                _pdfPageCounts.Remove(item);
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
            UpdatePageNavigation();
            _previewCancellation?.Cancel();
            _previewCancellation?.Dispose();
            _previewCancellation = new CancellationTokenSource();
            _ = UpdatePreviewAsync(++_previewVersion, _previewCancellation.Token);
        }

        private async void UpdatePageNavigation()
        {
            int version = ++_navigationVersion;
            PanelPageNav.Visibility = Visibility.Collapsed;
            if (FileListBox.SelectedItems.Count != 1 || FileListBox.SelectedItem is not InvoiceFileItem item ||
                !Path.GetExtension(item.FilePath).Equals(".pdf", StringComparison.OrdinalIgnoreCase)) return;
            try
            {
                if (!_pdfPageCounts.TryGetValue(item, out var task))
                {
                    var path = item.FilePath;
                    task = Task.Run(() => InvoicePrintService.GetPdfPageCount(path));
                    _pdfPageCounts[item] = task;
                }
                int count = await task;
                if (version != _navigationVersion) return;
                item.PageCount = Math.Max(1, count);
                PanelPageNav.Visibility = count > 1 ? Visibility.Visible : Visibility.Collapsed;
                TxtPageInfo.Text = $"第 {item.SelectedPage + 1} / {item.PageCount} 页";
            }
            catch (Exception ex) { if (version == _navigationVersion) SetStatus($"分页读取失败: {ex.Message}"); }
        }

        private async Task UpdatePreviewAsync(int version, CancellationToken cancellationToken)
        {
            if (_fileItems.Count == 0)
            {
                DropHintPanel.Visibility = Visibility.Visible;
                PreviewScroller.Visibility = Visibility.Collapsed;
                return;
            }

            DropHintPanel.Visibility = Visibility.Collapsed;
            PreviewScroller.Visibility = Visibility.Visible;

            bool entered = false;
            try
            {
                await Task.Delay(120, cancellationToken);
                await _previewGate.WaitAsync(cancellationToken);
                entered = true;
                cancellationToken.ThrowIfCancellationRequested();
                var template = GetCurrentTemplate();
                var previewItems = _fileItems.ToList();
                var isLandscape = IsLandscape;
                var showCutLine = ShowCutLine;

                var previewImages = new Dictionary<InvoiceFileItem, BitmapSource?>();
                var imagesToLoad = new List<(InvoiceFileItem Item, int PageIndex)>();
                foreach (var item in previewItems)
                {
                    if (item.PreviewPageIndex == item.SelectedPage)
                        previewImages[item] = item.PreviewImage;
                    else
                        imagesToLoad.Add((item, item.SelectedPage));
                }

                if (imagesToLoad.Count > 0)
                {
                    var loadedImages = await Task.Run(
                        () => LoadPreviewImages(imagesToLoad, cancellationToken), cancellationToken);
                    if (cancellationToken.IsCancellationRequested || version != _previewVersion) return;

                    foreach (var (item, pageIndex) in imagesToLoad)
                    {
                        loadedImages.TryGetValue(item, out var image);
                        item.PreviewImage = image;
                        item.PreviewPageIndex = pageIndex;
                        previewImages[item] = image;
                    }
                }

                var rotations = previewItems.Select(item => item.RotationAngle).ToArray();
                var pages = await InvoicePrintService.RunPrintWorkerAsync(() => RenderLayoutPreviewPages(
                    previewItems, template, previewImages, rotations, isLandscape, showCutLine, cancellationToken));
                if (cancellationToken.IsCancellationRequested || version != _previewVersion) return;

                LayoutPreviewGrid.Children.Clear();
                LayoutPreviewGrid.RowDefinitions.Clear();

                if (pages.Count == 0) return;

                for (int pi = 0; pi < pages.Count; pi++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
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
                    if (pi % 10 == 9) await Dispatcher.Yield(DispatcherPriority.Background);
                }

                cancellationToken.ThrowIfCancellationRequested();
                PreviewContainer.LayoutTransform = new ScaleTransform(_zoomLevel, _zoomLevel);
                TxtZoomLevel.Text = $"{(int)(_zoomLevel * 100)}%";
                if (pages.Count > 1)
                    SetStatus($"📄 共 {pages.Count} 页预览 · {previewItems.Count} 个文件");
            }
            catch (OperationCanceledException) { }
            catch (Exception ex)
            {
                if (version == _previewVersion)
                    SetStatus($"❌ 预览失败: {ex.Message}");
            }
            finally { if (entered) _previewGate.Release(); }
        }

        private static Dictionary<InvoiceFileItem, BitmapSource?> LoadPreviewImages(
            List<(InvoiceFileItem Item, int PageIndex)> items, CancellationToken cancellationToken)
        {
            var images = new Dictionary<InvoiceFileItem, BitmapSource?>();
            foreach (var (item, pageIndex) in items)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var ext = Path.GetExtension(item.FilePath).ToLowerInvariant();
                images[item] = ext == ".pdf"
                    ? InvoicePrintService.RenderPdfPage(item.FilePath, pageIndex, 150)
                    : InvoicePrintService.LoadPreviewImage(item.FilePath, ext);
            }
            return images;
        }

        /// <summary>
        /// 渲染排版预览：支持多页分页、纸张方向、裁剪线
        /// 返回每页一张 BitmapSource 的列表
        /// </summary>
        private static List<BitmapSource> RenderLayoutPreviewPages(
            List<InvoiceFileItem> items,
            PrintTemplate template,
            IReadOnlyDictionary<InvoiceFileItem, BitmapSource?> previewImages,
            double[] rotations,
            bool isLandscape,
            bool showCutLine,
            CancellationToken cancellationToken)
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
            if (isLandscape) { (paperW, paperH) = (paperH, paperW); }

            double ml = template.MarginLeft * mmToWpf, mr = template.MarginRight * mmToWpf;
            double mt = template.MarginTop * mmToWpf, mb = template.MarginBottom * mmToWpf;
            double ox = template.OffsetX * mmToWpf, oy = template.OffsetY * mmToWpf;
            double contentW = paperW - ml - mr, contentH = paperH - mt - mb;

            int perPage = template.PaperMode == "Invoice" ? 1 : template.LayoutCount;
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
                cancellationToken.ThrowIfCancellationRequested();
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
                        previewImages.TryGetValue(items[globalIdx], out var bmp);
                        if (bmp != null)
                        {
                            if (rotations[globalIdx] != 0)
                                bmp = InvoicePrintService.RotateImage(bmp, rotations[globalIdx]);
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
                if (showCutLine && perPage > 1)
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
                var orient = isLandscape ? "横向" : "纵向";
                var pageInfo = totalPages > 1 ? $" · 第{pageIdx + 1}/{totalPages}页" : "";
                var info = new FormattedText(
                    $"{template.PaperMode} · {orient} · {perPage}张/页{pageInfo}" + (showCutLine ? " · 裁剪线" : ""),
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

        // ── 旋转 & 缩放 ──
        private void BtnRotate_Click(object sender, RoutedEventArgs e) { if (FileListBox.SelectedItem is InvoiceFileItem item) { item.RotationAngle = (item.RotationAngle + 90) % 360; UpdatePreview(); } }
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

        private async void BtnPrint_Click2(object sender, RoutedEventArgs e)
        {
            if (_isPrintBusy) return;
            DispatcherTimer? progressTimer = null;
            PrintProgressPanel.Visibility = Visibility.Collapsed;
            SetPrintBusy(true);
            try
            {
                var request = CapturePrintRequest();
                if (request == null) return;
                var originalItems = _fileItems.ToList();
                var pageCount = (request.Items.Count + request.Template.LayoutCount - 1) / request.Template.LayoutCount;
                var confirm = MessageBox.Show(
                    $"将按当前页面设置发送到打印机“{request.PrinterName}”，共 {pageCount} 页 × {request.Copies} 份。是否继续？",
                    "确认打印",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);

                if (confirm != MessageBoxResult.Yes)
                {
                    SetStatus("已取消当前打印。");
                    return;
                }

                PrintProgressPanel.Visibility = Visibility.Visible;
                PrintProgressBar.Visibility = Visibility.Visible;
                PrintProgressBar.IsIndeterminate = true;
                BtnStartPrint.Content = "正在打印...";
                SetPrintProgress("正在连接打印机...");
                string? pendingProgress = null;
                Action<string> reportProgress = message => Interlocked.Exchange(ref pendingProgress, message);
                progressTimer = new DispatcherTimer(DispatcherPriority.Background)
                { Interval = TimeSpan.FromMilliseconds(100) };
                progressTimer.Tick += (_, _) =>
                {
                    var message = Interlocked.Exchange(ref pendingProgress, null);
                    if (message != null) SetPrintProgress(message);
                };
                progressTimer.Start();
                var result = await InvoicePrintService.RunPrintWorkerAsync(() =>
                    InvoicePrintService.PreparePrintJob(request.PrinterName, request.Items, request.Template,
                        request.Landscape, request.CutLine,
                        (pages, context) => InvoicePrintService.PrintPages(pages, request.Copies, context, reportProgress),
                        reportProgress: reportProgress));
                progressTimer.Stop();
                await Task.Run(() => _service.RecordPrintHistory(request.Items));
                if (result.Completed)
                {
                    foreach (var it in originalItems) it.IsPrinted = true;
                    SetPrintProgress($"{result.Message} 共 {pageCount} 页 x {request.Copies} 份。");
                    ToastService.Show(this, result.Message);
                }
                else
                {
                    SetPrintProgress(result.Message);
                }
            }
            catch (Exception ex)
            {
                progressTimer?.Stop();
                SetPrintProgress($"打印出错: {ex.Message}");
                PrintProgressBar.IsIndeterminate = false;
                PrintProgressBar.Visibility = Visibility.Collapsed;
                MessageBox.Show($"打印失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                progressTimer?.Stop();
                PrintProgressBar.IsIndeterminate = false;
                PrintProgressBar.Visibility = Visibility.Collapsed;
                BtnStartPrint.Content = "🖨️ 开始打印";
                SetPrintBusy(false);
            }
        }

        private void SetPrintProgress(string message)
        {
            PrintProgressPanel.Visibility = Visibility.Visible;
            TxtPrintProgress.Text = message;
            SetStatus(message);
        }

        private void BtnPrinterPreview_Click(object sender, RoutedEventArgs e)
        {
            if (_isPrintBusy) return;
            SetPrintBusy(true);
            try
            {
                var request = CapturePrintRequest();
                if (request != null) ShowPrinterPreviewWindow(request);
            }
            catch (Exception ex)
            {
                SetStatus($"预览出错: {ex.Message}");
                MessageBox.Show($"打印机预览失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally { SetPrintBusy(false); }
        }

        private void SetPrintBusy(bool busy)
        {
            _isPrintBusy = busy;
            BtnStartPrint.IsEnabled = !busy;
            BtnPrinterPreview.IsEnabled = !busy;
        }

        private sealed record PrintRequest(string PrinterName, List<InvoiceFileItem> Items,
            PrintTemplate Template, int Copies, bool Landscape, bool CutLine);

        private PrintRequest? CapturePrintRequest()
        {
            if (_isImporting) { SetStatus("请等待文件导入完成后再打印。"); return null; }
            if (_fileItems.Count == 0) { MessageBox.Show("请先导入发票文件", "提示"); return null; }
            if (CmbPrinter.SelectedItem is not PrinterOption printer)
            { MessageBox.Show("请先选择可用的打印机。", "提示"); return null; }
            var template = GetCurrentTemplate();
            if (template.PaperMode == "Invoice") template.LayoutCount = 1;
            if (!int.TryParse(TxtCopies.Text.Trim(), out var copies)) copies = 1;
            var items = _fileItems.Select(item => new InvoiceFileItem
            {
                FilePath = item.FilePath, FileName = item.FileName, FileType = item.FileType,
                FileSize = item.FileSize, SelectedPage = item.SelectedPage,
                RotationAngle = item.RotationAngle, CropRect = item.CropRect
            }).ToList();
            return new PrintRequest(printer.FullName, items, template, Math.Clamp(copies, 1, 99), IsLandscape, ShowCutLine);
        }

        private void ShowPrinterPreviewWindow(PrintRequest request)
        {
            var toolbar = new DockPanel
            {
                Margin = new Thickness(16, 12, 16, 12),
                LastChildFill = false
            };

            var info = new TextBlock
            {
                Text = "正在准备打印机预览...",
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

            bool closed = false;
            bool started = false;
            using var cancellation = new CancellationTokenSource();
            var cancellationToken = cancellation.Token;
            win.Closed += (_, _) => { closed = true; cancellation.Cancel(); };
            win.ContentRendered += async (_, _) =>
            {
                if (started) return;
                started = true;
                bool entered = false;
                try
                {
                    await _printerPreviewGate.WaitAsync(cancellationToken);
                    entered = true;
                    cancellationToken.ThrowIfCancellationRequested();
                    var data = await InvoicePrintService.RunPrintWorkerAsync(() =>
                        InvoicePrintService.PreparePrintJob(request.PrinterName, request.Items, request.Template,
                            request.Landscape, request.CutLine,
                            (pages, context) => InvoicePrintService.RenderPreviewDocument(pages, context, cancellationToken),
                            cancellationToken));
                    if (closed) return;
                    var document = new FixedDocument();
                    document.DocumentPaginator.PageSize = data.MediaSize;
                    foreach (var bitmap in data.Images)
                    {
                        if (closed) return;
                        var page = new FixedPage { Width = data.MediaSize.Width, Height = data.MediaSize.Height };
                        var image = new System.Windows.Controls.Image
                        {
                            Source = bitmap, Width = data.ContentSize.Width, Height = data.ContentSize.Height,
                            Stretch = Stretch.Fill
                        };
                        FixedPage.SetLeft(image, data.ContentOrigin.X);
                        FixedPage.SetTop(image, data.ContentOrigin.Y);
                        page.Children.Add(image);
                        var content = new PageContent();
                        ((System.Windows.Markup.IAddChild)content).AddChild(page);
                        document.Pages.Add(content);
                        await Dispatcher.Yield(DispatcherPriority.Background);
                    }
                    if (closed) return;
                    viewer.Document = document;
                    info.Text = $"{request.PrinterName} · {data.Images.Count} 页 · {(request.Landscape ? "横向" : "纵向")}";
                    SetStatus($"已生成打印机预览，共 {data.Images.Count} 页。");
                }
                catch (OperationCanceledException) { }
                catch (Exception ex)
                {
                    if (!closed) info.Text = $"预览失败：{ex.Message}";
                }
                finally { if (entered) _printerPreviewGate.Release(); }
            };

            win.ShowDialog();
        }

        private void SetStatus(string text) => TxtStatus.Text = text;

        private static string? ShowInputDialog(string title, string prompt, string defaultValue = "")
        {
            var dialog = new AppInputDialog(title, prompt, defaultValue)
            {
                Owner = Application.Current.MainWindow
            };

            return dialog.ShowDialog() == true ? dialog.InputValue : null;
        }
    }
}
