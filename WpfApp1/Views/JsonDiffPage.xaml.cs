using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using WpfApp1.Services;

namespace WpfApp1.Views
{
    public partial class JsonDiffPage : Page
    {
        // ── Win32 剪贴板 ──
        [DllImport("user32.dll")] private static extern bool OpenClipboard(IntPtr hWndNewOwner);
        [DllImport("user32.dll")] private static extern bool CloseClipboard();
        [DllImport("user32.dll")] private static extern bool EmptyClipboard();
        [DllImport("user32.dll")] private static extern IntPtr SetClipboardData(uint uFormat, IntPtr hMem);
        private const uint CF_UNICODETEXT = 13;

        private List<DiffEntry> _allDiffs = [];

        public JsonDiffPage()
        {
            InitializeComponent();
            ShowWelcome();
        }

        // ═══════════════════════════════════════
        // 工具栏按钮
        // ═══════════════════════════════════════

        private void BtnImportA_Click(object sender, RoutedEventArgs e) => ImportJsonTo(TxtJsonA);
        private void BtnImportB_Click(object sender, RoutedEventArgs e) => ImportJsonTo(TxtJsonB);

        private void ImportJsonTo(TextBox target)
        {
            var dlg = new OpenFileDialog { Filter = "JSON 文件|*.json|所有文件|*.*" };
            if (dlg.ShowDialog() == true)
            {
                var content = File.ReadAllText(dlg.FileName);
                try { content = JsonToolService.Beautify(content); } catch { }
                target.Text = content;
            }
        }

        private void BtnBeautifyA_Click(object sender, RoutedEventArgs e) => BeautifyBox(TxtJsonA);
        private void BtnBeautifyB_Click(object sender, RoutedEventArgs e) => BeautifyBox(TxtJsonB);

        private void BeautifyBox(TextBox box)
        {
            try { box.Text = JsonToolService.Beautify(box.Text); }
            catch { MessageBox.Show("JSON 格式错误，无法美化", "提示", MessageBoxButton.OK, MessageBoxImage.Warning); }
        }

        private void BtnClearA_Click(object sender, RoutedEventArgs e) => TxtJsonA.Text = "";
        private void BtnClearB_Click(object sender, RoutedEventArgs e) => TxtJsonB.Text = "";

        private void BtnClearAll_Click(object sender, RoutedEventArgs e)
        {
            TxtJsonA.Text = "";
            TxtJsonB.Text = "";
            _allDiffs.Clear();
            DiffResultContainer.Children.Clear();
            TxtDiffSummary.Text = "";
            ShowWelcome();
        }

        private void BtnSwap_Click(object sender, RoutedEventArgs e)
        {
            var temp = TxtJsonA.Text;
            TxtJsonA.Text = TxtJsonB.Text;
            TxtJsonB.Text = temp;

            if (_allDiffs.Count > 0)
                BtnCompare_Click(sender, e);
        }

        // ═══════════════════════════════════════
        // 核心：开始对比
        // ═══════════════════════════════════════

        private void BtnCompare_Click(object sender, RoutedEventArgs e)
        {
            var jsonA = TxtJsonA.Text?.Trim() ?? "";
            var jsonB = TxtJsonB.Text?.Trim() ?? "";

            if (string.IsNullOrEmpty(jsonA) && string.IsNullOrEmpty(jsonB))
            {
                MessageBox.Show("请在左右两侧输入 JSON", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            JsonElement? elA = null, elB = null;

            if (!string.IsNullOrEmpty(jsonA))
            {
                try { elA = JsonDocument.Parse(jsonA).RootElement.Clone(); }
                catch (JsonException ex)
                {
                    MessageBox.Show($"JSON A 格式错误：\n{ex.Message}", "解析失败", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
            }

            if (!string.IsNullOrEmpty(jsonB))
            {
                try { elB = JsonDocument.Parse(jsonB).RootElement.Clone(); }
                catch (JsonException ex)
                {
                    MessageBox.Show($"JSON B 格式错误：\n{ex.Message}", "解析失败", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
            }

            _allDiffs = [];
            CompareElements(elA, elB, "$");

            if (_allDiffs.Count == 0)
            {
                TxtDiffSummary.Text = "✅ 两个 JSON 完全相同，无差异";
                DiffResultContainer.Children.Clear();
                ShowNoDiffMessage();
                return;
            }

            RenderDiffResults();
        }

        // ═══════════════════════════════════════
        // 递归深度对比
        // ═══════════════════════════════════════

        private void CompareElements(JsonElement? a, JsonElement? b, string path)
        {
            // 一侧为 null（整体新增/删除）
            if (a == null && b != null)
            {
                CollectAll(b.Value, path, DiffType.Added);
                return;
            }
            if (a != null && b == null)
            {
                CollectAll(a.Value, path, DiffType.Removed);
                return;
            }
            if (a == null || b == null) return;

            var elA = a.Value;
            var elB = b.Value;

            // 类型不同
            if (elA.ValueKind != elB.ValueKind)
            {
                _allDiffs.Add(new DiffEntry
                {
                    Path = path,
                    Type = DiffType.TypeChanged,
                    OldValue = FormatValue(elA),
                    NewValue = FormatValue(elB),
                    OldType = elA.ValueKind.ToString(),
                    NewType = elB.ValueKind.ToString()
                });
                return;
            }

            switch (elA.ValueKind)
            {
                case JsonValueKind.Object:
                    CompareObjects(elA, elB, path);
                    break;

                case JsonValueKind.Array:
                    CompareArrays(elA, elB, path);
                    break;

                default:
                    // 基本类型值对比
                    var valA = FormatValue(elA);
                    var valB = FormatValue(elB);
                    if (valA != valB)
                    {
                        _allDiffs.Add(new DiffEntry
                        {
                            Path = path,
                            Type = DiffType.Modified,
                            OldValue = valA,
                            NewValue = valB
                        });
                    }
                    break;
            }
        }

        private void CompareObjects(JsonElement a, JsonElement b, string path)
        {
            var keysA = a.EnumerateObject().Select(p => p.Name).ToHashSet();
            var keysB = b.EnumerateObject().Select(p => p.Name).ToHashSet();

            // A 有 B 没有 → 删除
            foreach (var key in keysA.Except(keysB))
            {
                var childPath = $"{path}.{key}";
                CollectAll(a.GetProperty(key), childPath, DiffType.Removed);
            }

            // B 有 A 没有 → 新增
            foreach (var key in keysB.Except(keysA))
            {
                var childPath = $"{path}.{key}";
                CollectAll(b.GetProperty(key), childPath, DiffType.Added);
            }

            // 两者都有 → 递归对比
            foreach (var key in keysA.Intersect(keysB))
            {
                var childPath = $"{path}.{key}";
                CompareElements(a.GetProperty(key), b.GetProperty(key), childPath);
            }
        }

        private void CompareArrays(JsonElement a, JsonElement b, string path)
        {
            int lenA = a.GetArrayLength();
            int lenB = b.GetArrayLength();
            int maxLen = Math.Max(lenA, lenB);

            for (int i = 0; i < maxLen; i++)
            {
                var childPath = $"{path}[{i}]";
                if (i >= lenA)
                {
                    CollectAll(b[i], childPath, DiffType.Added);
                }
                else if (i >= lenB)
                {
                    CollectAll(a[i], childPath, DiffType.Removed);
                }
                else
                {
                    CompareElements(a[i], b[i], childPath);
                }
            }
        }

        /// <summary>递归展开节点，每个叶子值单独记录为一条差异（新增/删除）</summary>
        private void CollectAll(JsonElement el, string path, DiffType type)
        {
            switch (el.ValueKind)
            {
                case JsonValueKind.Object:
                    foreach (var prop in el.EnumerateObject())
                        CollectAll(prop.Value, $"{path}.{prop.Name}", type);
                    break;

                case JsonValueKind.Array:
                    for (int i = 0; i < el.GetArrayLength(); i++)
                        CollectAll(el[i], $"{path}[{i}]", type);
                    break;

                default:
                    var val = FormatValue(el);
                    _allDiffs.Add(new DiffEntry
                    {
                        Path = path,
                        Type = type,
                        OldValue = type == DiffType.Removed ? val : "",
                        NewValue = type == DiffType.Added ? val : "",
                        OldType = type == DiffType.Removed ? el.ValueKind.ToString() : "",
                        NewType = type == DiffType.Added ? el.ValueKind.ToString() : ""
                    });
                    break;
            }
        }

        private static string FormatValue(JsonElement el)
        {
            return el.ValueKind switch
            {
                JsonValueKind.String => $"\"{el.GetString()}\"",
                JsonValueKind.True => "true",
                JsonValueKind.False => "false",
                JsonValueKind.Null => "null",
                JsonValueKind.Number => el.GetRawText(),
                JsonValueKind.Object => $"{{{el.EnumerateObject().Count()} 个字段}}",
                JsonValueKind.Array => $"[{el.GetArrayLength()} 项]",
                _ => el.GetRawText()
            };
        }

        // ═══════════════════════════════════════
        // 渲染差异结果
        // ═══════════════════════════════════════

        private void RenderDiffResults()
        {
            var filtered = GetFilteredDiffs();

            int added = _allDiffs.Count(d => d.Type == DiffType.Added);
            int removed = _allDiffs.Count(d => d.Type == DiffType.Removed);
            int modified = _allDiffs.Count(d => d.Type == DiffType.Modified);
            int typeChanged = _allDiffs.Count(d => d.Type == DiffType.TypeChanged);

            TxtDiffSummary.Text = $"共 {_allDiffs.Count} 处差异  |  " +
                                  $"🟢 新增 {added}  🔴 删除 {removed}  🟠 修改 {modified}  🟣 类型变化 {typeChanged}" +
                                  (filtered.Count != _allDiffs.Count ? $"  （当前显示 {filtered.Count} 条）" : "");

            DiffResultContainer.Children.Clear();

            if (filtered.Count == 0)
            {
                ShowFilterEmptyMessage();
                return;
            }

            // 表头
            var headerGrid = CreateDiffRow(
                "路径", "差异类型", "JSON A（原始值）", "JSON B（新值）",
                Br("#F1F5F9"), Br("#555"), Br("#555"), Br("#555"), FontWeights.SemiBold);
            headerGrid.Margin = new Thickness(0);
            DiffResultContainer.Children.Add(headerGrid);

            // 数据行
            for (int i = 0; i < filtered.Count; i++)
            {
                var d = filtered[i];
                var (typeLabel, typeColor) = GetTypeDisplay(d.Type);
                var bgBrush = GetRowBackground(d.Type, i);

                string oldDisplay = d.Type == DiffType.TypeChanged
                    ? $"({d.OldType}) {d.OldValue}"
                    : d.OldValue;
                string newDisplay = d.Type == DiffType.TypeChanged
                    ? $"({d.NewType}) {d.NewValue}"
                    : d.NewValue;

                var row = CreateDiffRow(d.Path, typeLabel, oldDisplay, newDisplay,
                    bgBrush, typeColor, Br("#DC2626"), Br("#16A34A"), FontWeights.Normal);
                DiffResultContainer.Children.Add(row);
            }
        }

        private Grid CreateDiffRow(string col1, string col2, string col3, string col4,
            Brush bg, Brush fg2, Brush fg3, Brush fg4, FontWeight weight)
        {
            var grid = new Grid
            {
                Background = bg,
                MinHeight = 32
            };
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(2.5, GridUnitType.Star), MinWidth = 180 });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star), MinWidth = 80 });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(2, GridUnitType.Star), MinWidth = 140 });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(2, GridUnitType.Star), MinWidth = 140 });

            AddCell(grid, 0, col1, Br("#333"), weight);
            AddCell(grid, 1, col2, fg2, weight);
            AddCell(grid, 2, col3, fg3, weight);
            AddCell(grid, 3, col4, fg4, weight);

            return grid;
        }

        private void AddCell(Grid grid, int col, string text, Brush fg, FontWeight weight)
        {
            var border = new Border
            {
                BorderBrush = Br("#E5E7EB"),
                BorderThickness = new Thickness(0, 0, col < 3 ? 1 : 0, 1),
                Padding = new Thickness(10, 6, 10, 6)
            };

            var tb = new TextBlock
            {
                Text = text,
                Foreground = fg,
                FontWeight = weight,
                FontSize = 12.5,
                FontFamily = col == 0 ? new FontFamily("Consolas, 'Courier New'") : new FontFamily("Segoe UI"),
                VerticalAlignment = VerticalAlignment.Center,
                TextWrapping = TextWrapping.Wrap
            };

            border.Child = tb;
            Grid.SetColumn(border, col);
            grid.Children.Add(border);
        }

        private static (string label, Brush color) GetTypeDisplay(DiffType type) => type switch
        {
            DiffType.Added => ("➕ 新增", Br("#16A34A")),
            DiffType.Removed => ("➖ 删除", Br("#DC2626")),
            DiffType.Modified => ("✏️ 修改", Br("#D97706")),
            DiffType.TypeChanged => ("🔄 类型变化", Br("#9333EA")),
            _ => ("—", Br("#999"))
        };

        private static Brush GetRowBackground(DiffType type, int index)
        {
            return type switch
            {
                DiffType.Added => Br(index % 2 == 0 ? "#F0FDF4" : "#ECFDF5"),
                DiffType.Removed => Br(index % 2 == 0 ? "#FEF2F2" : "#FFF5F5"),
                DiffType.Modified => Br(index % 2 == 0 ? "#FFFBEB" : "#FEF3C7"),
                DiffType.TypeChanged => Br(index % 2 == 0 ? "#FAF5FF" : "#F3E8FF"),
                _ => Br(index % 2 == 0 ? "#FFF" : "#FAFAFA")
            };
        }

        // ═══════════════════════════════════════
        // 筛选
        // ═══════════════════════════════════════

        private void Filter_Changed(object sender, RoutedEventArgs e)
        {
            if (_allDiffs.Count > 0)
                RenderDiffResults();
        }

        private List<DiffEntry> GetFilteredDiffs()
        {
            return _allDiffs.Where(d =>
                (d.Type == DiffType.Added && ChkShowAdded.IsChecked == true) ||
                (d.Type == DiffType.Removed && ChkShowRemoved.IsChecked == true) ||
                (d.Type == DiffType.Modified && ChkShowModified.IsChecked == true) ||
                (d.Type == DiffType.TypeChanged && ChkShowTypeChanged.IsChecked == true)
            ).ToList();
        }

        // ═══════════════════════════════════════
        // 复制 / 导出报告
        // ═══════════════════════════════════════

        private void BtnCopyReport_Click(object sender, RoutedEventArgs e)
        {
            if (_allDiffs.Count == 0)
            {
                MessageBox.Show("请先执行对比", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var report = BuildReport();
            SafeCopyToClipboard(report);
            ShowToast("✅ 差异报告已复制到剪贴板");
        }

        private void BtnExportReport_Click(object sender, RoutedEventArgs e)
        {
            if (_allDiffs.Count == 0)
            {
                MessageBox.Show("请先执行对比", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var dlg = new SaveFileDialog
            {
                Filter = "文本文件|*.txt|Markdown|*.md|所有文件|*.*",
                FileName = $"json_diff_{DateTime.Now:yyyyMMdd_HHmmss}.txt"
            };

            if (dlg.ShowDialog() == true)
            {
                File.WriteAllText(dlg.FileName, BuildReport(), Encoding.UTF8);
                MessageBox.Show($"报告已导出到：\n{dlg.FileName}", "导出成功", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private string BuildReport()
        {
            var sb = new StringBuilder();
            sb.AppendLine("═══════════════════════════════════════");
            sb.AppendLine("  JSON 对比报告");
            sb.AppendLine($"  生成时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine("═══════════════════════════════════════");
            sb.AppendLine();

            int added = _allDiffs.Count(d => d.Type == DiffType.Added);
            int removed = _allDiffs.Count(d => d.Type == DiffType.Removed);
            int modified = _allDiffs.Count(d => d.Type == DiffType.Modified);
            int typeChanged = _allDiffs.Count(d => d.Type == DiffType.TypeChanged);

            sb.AppendLine($"  总差异数: {_allDiffs.Count}");
            sb.AppendLine($"  ● 新增: {added}  ● 删除: {removed}  ● 修改: {modified}  ● 类型变化: {typeChanged}");
            sb.AppendLine();
            sb.AppendLine("───────────────────────────────────────");

            foreach (var d in _allDiffs)
            {
                var typeStr = d.Type switch
                {
                    DiffType.Added => "[+新增]",
                    DiffType.Removed => "[-删除]",
                    DiffType.Modified => "[~修改]",
                    DiffType.TypeChanged => "[!类型变化]",
                    _ => "[?]"
                };

                sb.AppendLine();
                sb.AppendLine($"  {typeStr}  {d.Path}");

                if (d.Type == DiffType.Modified)
                {
                    sb.AppendLine($"    旧值: {d.OldValue}");
                    sb.AppendLine($"    新值: {d.NewValue}");
                }
                else if (d.Type == DiffType.TypeChanged)
                {
                    sb.AppendLine($"    旧: ({d.OldType}) {d.OldValue}");
                    sb.AppendLine($"    新: ({d.NewType}) {d.NewValue}");
                }
                else if (d.Type == DiffType.Added)
                {
                    sb.AppendLine($"    值: {d.NewValue}");
                }
                else if (d.Type == DiffType.Removed)
                {
                    sb.AppendLine($"    值: {d.OldValue}");
                }
            }

            sb.AppendLine();
            sb.AppendLine("═══════════════════════════════════════");
            return sb.ToString();
        }

        // ═══════════════════════════════════════
        // 占位提示
        // ═══════════════════════════════════════

        private void ShowWelcome()
        {
            DiffResultContainer.Children.Clear();
            DiffResultContainer.Children.Add(new TextBlock
            {
                Text = "💡 在左右两侧粘贴或导入 JSON，点击「⚡ 开始对比」查看差异\n\n" +
                       "支持功能：\n" +
                       "  • 深度递归对比，精确到每个字段和数组元素\n" +
                       "  • 识别新增、删除、值修改、类型变化四种差异\n" +
                       "  • 忽略 JSON 对象字段顺序（只比较内容）\n" +
                       "  • 按差异类型筛选显示\n" +
                       "  • 一键复制或导出差异报告",
                FontSize = 13,
                Foreground = Br("#999"),
                Margin = new Thickness(20, 20, 20, 20),
                LineHeight = 24
            });
        }

        private void ShowNoDiffMessage()
        {
            DiffResultContainer.Children.Clear();
            DiffResultContainer.Children.Add(new Border
            {
                Background = Br("#F0FDF4"),
                CornerRadius = new CornerRadius(8),
                Padding = new Thickness(20, 16, 20, 16),
                Margin = new Thickness(12),
                Child = new TextBlock
                {
                    Text = "✅ 两个 JSON 完全相同，没有任何差异！",
                    FontSize = 15,
                    FontWeight = FontWeights.SemiBold,
                    Foreground = Br("#16A34A"),
                    HorizontalAlignment = HorizontalAlignment.Center
                }
            });
        }

        private void ShowFilterEmptyMessage()
        {
            DiffResultContainer.Children.Clear();
            DiffResultContainer.Children.Add(new TextBlock
            {
                Text = "当前筛选条件下没有差异记录，请调整上方的筛选选项",
                FontSize = 13,
                Foreground = Br("#999"),
                Margin = new Thickness(20),
                HorizontalAlignment = HorizontalAlignment.Center
            });
        }

        // ═══════════════════════════════════════
        // 剪贴板 + Toast
        // ═══════════════════════════════════════

        private void SafeCopyToClipboard(string text)
        {
            try
            {
                if (OpenClipboard(IntPtr.Zero))
                {
                    try
                    {
                        EmptyClipboard();
                        var hGlobal = Marshal.StringToHGlobalUni(text);
                        SetClipboardData(CF_UNICODETEXT, hGlobal);
                    }
                    finally { CloseClipboard(); }
                }
                else { Clipboard.SetText(text); }
            }
            catch { try { Clipboard.SetText(text); } catch { } }
        }

        private void ShowToast(string message)
        {
            var popup = new Border
            {
                Background = Br("#333"),
                CornerRadius = new CornerRadius(6),
                Padding = new Thickness(16, 8, 16, 8),
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Top,
                Margin = new Thickness(0, 10, 0, 0),
                Opacity = 0.92,
                Child = new TextBlock
                {
                    Text = message,
                    Foreground = Brushes.White,
                    FontSize = 13
                }
            };

            if (this.Content is Grid rootGrid)
            {
                rootGrid.Children.Add(popup);
                var timer = new System.Windows.Threading.DispatcherTimer { Interval = TimeSpan.FromSeconds(1.5) };
                timer.Tick += (_, _) => { rootGrid.Children.Remove(popup); timer.Stop(); };
                timer.Start();
            }
        }

        // ═══════════════════════════════════════
        // 工具
        // ═══════════════════════════════════════

        private static SolidColorBrush Br(string hex)
            => new((Color)ColorConverter.ConvertFromString(hex)!);

        // ═══════════════════════════════════════
        // 数据模型
        // ═══════════════════════════════════════

        private enum DiffType { Added, Removed, Modified, TypeChanged }

        private class DiffEntry
        {
            public string Path { get; set; } = "";
            public DiffType Type { get; set; }
            public string OldValue { get; set; } = "";
            public string NewValue { get; set; } = "";
            public string OldType { get; set; } = "";
            public string NewType { get; set; } = "";
        }
    }
}
