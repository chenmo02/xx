using Microsoft.Win32;
using System.Collections.ObjectModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using WpfApp1.Services;

namespace WpfApp1.Views
{
    public partial class JsonToolPage : Page
    {
        private readonly DispatcherTimer _debounceTimer = new();
        private static readonly Brush BorderColor = MakeBrush("#E5E7EB");
        private static readonly Brush HeaderBg = MakeBrush("#0EA5E9");
        private static readonly Brush HeaderBorderColor = MakeBrush("#0284C7");
        private static readonly Brush RowNumBg = MakeBrush("#F1F5F9");
        private static readonly Brush EvenRowBg = MakeBrush("#FFFDE7");
        private static readonly Brush LinkColor = MakeBrush("#2563EB");
        private static readonly Brush KeyColor = MakeBrush("#374151");

        public JsonToolPage()
        {
            InitializeComponent();
            _debounceTimer.Interval = TimeSpan.FromMilliseconds(600);
            _debounceTimer.Tick += (s, e) => { _debounceTimer.Stop(); RebuildGrid(); };
        }

        private void TxtJsonEditor_TextChanged(object sender, TextChangedEventArgs e)
        {
            _debounceTimer.Stop();
            _debounceTimer.Start();
        }

        // ==================== 核心：构建嵌套网格 ====================

        private void RebuildGrid()
        {
            GridContainer.Children.Clear();
            string json = TxtJsonEditor.Text?.Trim() ?? "";
            if (string.IsNullOrEmpty(json)) { SetStatus("就绪"); return; }

            try
            {
                var nodes = JsonGridParser.Parse(json);
                if (nodes.Count > 0)
                    RenderNode(nodes[0], GridContainer, 0);
                SetStatus("✅ 已同步");
            }
            catch { SetStatus("⚠️ JSON 格式错误"); }
        }

        /// <summary>递归渲染节点</summary>
        private void RenderNode(JsonGridNode node, Panel container, int depth)
        {
            if (node.Key == "root" && node.IsContainer)
            {
                foreach (var child in node.Children)
                    RenderNode(child, container, depth);
                if (node.HasTable)
                    RenderTable(node, container, depth);
                return;
            }

            if (node.HasTable)
            {
                // 有表格数据（对象数组 或 简单数组）→ 渲染表头 + 表格
                AddLabel(node.ExpandLabel, container, depth, true);
                RenderTable(node, container, depth);
            }
            else if (node.IsContainer)
            {
                // 普通对象
                AddLabel(node.ExpandLabel, container, depth, true);
                foreach (var child in node.Children)
                    RenderNode(child, container, depth + 1);
            }
            else
            {
                // 叶子节点
                RenderLeafNode(node, container, depth);
            }
        }

        /// <summary>渲染叶子节点 key: value</summary>
        private void RenderLeafNode(JsonGridNode node, Panel container, int depth)
        {
            var row = new Grid { Margin = new Thickness(depth * 20, 0, 0, 0) };
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(160) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            var keyBorder = new Border
            {
                BorderBrush = BorderColor,
                BorderThickness = new Thickness(0, 0, 0, 1),
                Padding = new Thickness(6, 4, 6, 4),
                Child = new TextBlock { Text = node.Key, FontWeight = FontWeights.SemiBold, Foreground = KeyColor, FontSize = 13, VerticalAlignment = VerticalAlignment.Center }
            };
            Grid.SetColumn(keyBorder, 0);
            row.Children.Add(keyBorder);

            var valBorder = new Border
            {
                BorderBrush = BorderColor,
                BorderThickness = new Thickness(0, 0, 0, 1),
                Padding = new Thickness(10, 4, 6, 4),
                Child = new TextBlock { Text = node.Value, Foreground = MakeBrush(node.ValueColor), FontSize = 13, VerticalAlignment = VerticalAlignment.Center }
            };
            Grid.SetColumn(valBorder, 1);
            row.Children.Add(valBorder);

            container.Children.Add(row);
        }

        /// <summary>添加标签文本</summary>
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

        // ==================== 通用表格渲染（对象数组 + 简单数组） ====================

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

            // 列：行号 + 数据列
            tableGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            foreach (var _ in node.TableColumns)
                tableGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto, MinWidth = 100 });

            // 行：表头 + 数据行
            tableGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            foreach (var _ in node.TableRows)
                tableGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            // ---- 表头 ----
            // ··· 按钮（点击导出 CSV）
            var menuBtn = CreateHeaderCell("···");
            menuBtn.Cursor = Cursors.Hand;
            menuBtn.MouseLeftButtonUp += (s, e) => ExportNodeToCsv(node);
            menuBtn.ToolTip = "导出此表格为 CSV";
            Grid.SetRow(menuBtn, 0);
            Grid.SetColumn(menuBtn, 0);
            tableGrid.Children.Add(menuBtn);

            for (int c = 0; c < node.TableColumns.Count; c++)
            {
                var header = CreateHeaderCell(node.TableColumns[c]);
                Grid.SetRow(header, 0);
                Grid.SetColumn(header, c + 1);
                tableGrid.Children.Add(header);
            }

            // ---- 数据行 ----
            for (int r = 0; r < node.TableRows.Count; r++)
            {
                var row = node.TableRows[r];
                int gridRow = r + 1;

                // 行号
                var rowNum = new Border
                {
                    Background = RowNumBg,
                    BorderBrush = BorderColor,
                    BorderThickness = new Thickness(0, 0, 1, 1),
                    Padding = new Thickness(8, 4, 8, 4),
                    Child = new TextBlock { Text = row.Index.ToString(), Foreground = Brushes.Gray, FontSize = 12, HorizontalAlignment = HorizontalAlignment.Center }
                };
                Grid.SetRow(rowNum, gridRow);
                Grid.SetColumn(rowNum, 0);
                tableGrid.Children.Add(rowNum);

                // 单元格
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

                    if (r % 2 == 0)
                        cellBorder.Background = EvenRowBg;

                    if (cell.IsNested)
                    {
                        // 嵌套结构
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
                                {
                                    foreach (var nested in cell.NestedChildren)
                                        RenderNestedContent(nested, nestedContainer);
                                }
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
                        cellBorder.Child = new TextBlock
                        {
                            Text = cell.Value,
                            Foreground = MakeBrush(cell.ValueColor),
                            FontSize = 12,
                            VerticalAlignment = VerticalAlignment.Center,
                            TextTrimming = TextTrimming.CharacterEllipsis,
                            MaxWidth = 280
                        };
                    }

                    Grid.SetRow(cellBorder, gridRow);
                    Grid.SetColumn(cellBorder, c + 1);
                    tableGrid.Children.Add(cellBorder);
                }
            }

            tableBorder.Child = tableGrid;
            container.Children.Add(tableBorder);
        }

        /// <summary>
        /// 渲染嵌套内容（展开后的子节点）
        /// </summary>
        private void RenderNestedContent(JsonGridNode nested, Panel container)
        {
            if (nested.HasTable)
            {
                // 有表格（对象数组 或 简单数组）→ 直接渲染表格
                RenderTable(nested, container, 0);
            }
            else if (nested.NodeType == "Object")
            {
                // 普通对象 → key-value
                foreach (var child in nested.Children)
                    RenderNode(child, container, 0);
            }
            else
            {
                RenderLeafNode(nested, container, 0);
            }
        }

        // ==================== ··· 按钮：导出节点为 CSV ====================

        private void ExportNodeToCsv(JsonGridNode node)
        {
            if (node.TableColumns.Count == 0 || node.TableRows.Count == 0)
            {
                MessageBox.Show("没有可导出的表格数据", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var sb = new StringBuilder();

            // 表头
            sb.AppendLine(string.Join(",", node.TableColumns.Select(EscapeCsvField)));

            // 数据行
            foreach (var row in node.TableRows)
            {
                var fields = new List<string>();
                foreach (var cell in row.Cells)
                {
                    if (cell.IsNested)
                        fields.Add(EscapeCsvField(cell.NestedSummary));
                    else
                        fields.Add(EscapeCsvField(cell.Value));
                }
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
                // 写入 UTF-8 BOM 以便 Excel 正确识别中文
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
            try { TxtJsonEditor.Text = JsonToolService.Beautify(TxtJsonEditor.Text); SetStatus("✅ 格式化完成"); }
            catch (JsonException ex) { SetStatus($"❌ {ex.Message}"); }
        }

        private void BtnMinify_Click(object sender, RoutedEventArgs e)
        {
            try { TxtJsonEditor.Text = JsonToolService.Minify(TxtJsonEditor.Text); SetStatus("✅ 压缩完成"); }
            catch (JsonException ex) { SetStatus($"❌ {ex.Message}"); }
        }

        private void BtnValidate_Click(object sender, RoutedEventArgs e)
        {
            var (isValid, message, _) = JsonToolService.Validate(TxtJsonEditor.Text);
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
                TxtJsonEditor.Text = content;
                SetStatus($"✅ 已导入: {Path.GetFileName(dlg.FileName)}");
            }
        }

        private void BtnExportJson_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TxtJsonEditor.Text)) return;
            var dlg = new SaveFileDialog { Filter = "JSON 文件|*.json", FileName = $"export_{DateTime.Now:yyyyMMdd_HHmmss}.json" };
            if (dlg.ShowDialog() == true)
            {
                File.WriteAllText(dlg.FileName, TxtJsonEditor.Text, Encoding.UTF8);
                SetStatus("✅ 已导出");
            }
        }

        private void BtnExportCsv_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TxtJsonEditor.Text)) return;
            try
            {
                string csv = JsonToolService.JsonToCsv(TxtJsonEditor.Text);
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
