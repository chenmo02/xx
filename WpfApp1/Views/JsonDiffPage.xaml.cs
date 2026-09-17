using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using WpfApp1.Controls;
using WpfApp1.Services;

namespace WpfApp1.Views
{
    public partial class JsonDiffPage : Page
    {
        private List<DiffEntry> _allDiffs = [];
        private readonly List<DiffEntry> _unchanged = [];
        private bool _hasCompared;
        private bool _syncingSelection;
        private TextBox TxtJsonA => EditorA.Input;
        private TextBox TxtJsonB => EditorB.Input;

        public JsonDiffPage()
        {
            InitializeComponent();
            TxtJsonA.TextChanged += Input_Changed;
            TxtJsonB.TextChanged += Input_Changed;
        }

        private void Page_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (ContentLayout != null)
            {
                ContentLayout.Height = Math.Max(580, ActualHeight - 40);
                ContentLayout.Width = Math.Max(580, ActualWidth - 38 -
                    (ActualHeight < 620 ? SystemParameters.VerticalScrollBarWidth : 0));
            }
        }

        private void EditorHeader_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            var header = (Grid)sender;
            var actions = header.Children[1];
            bool stacked = header.ActualWidth < 360;
            Grid.SetRow(actions, stacked ? 1 : 0);
            Grid.SetColumn(actions, stacked ? 0 : 1);
            Grid.SetColumnSpan(actions, stacked ? 2 : 1);
        }

        private void Input_Changed(object sender, TextChangedEventArgs e)
        {
            bool hadResults = _hasCompared;
            _hasCompared = false;
            _allDiffs.Clear();
            _unchanged.Clear();
            RenderDiffResults();
            TxtDiffSummary.Text = hadResults ? "●  内容已修改，请重新对比" : "●  等待对比";
            TxtDiffSummary.Foreground = Br("#8094B1");
        }

        private void BtnImportA_Click(object sender, RoutedEventArgs e) => ImportJsonTo(TxtJsonA);
        private void BtnImportB_Click(object sender, RoutedEventArgs e) => ImportJsonTo(TxtJsonB);

        private void ImportJsonTo(TextBox target)
        {
            var dlg = new OpenFileDialog { Filter = "JSON 文件|*.json|所有文件|*.*" };
            if (dlg.ShowDialog() != true) return;
            try
            {
                var content = File.ReadAllText(dlg.FileName);
                try { content = JsonToolService.Beautify(content); } catch (JsonException) { }
                target.Text = content;
                ToastService.Show(this, "JSON 文件已导入");
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                MessageBox.Show($"无法读取文件：{ex.Message}", "导入失败", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void BtnBeautifyA_Click(object sender, RoutedEventArgs e) => BeautifyBox(TxtJsonA);
        private void BtnBeautifyB_Click(object sender, RoutedEventArgs e) => BeautifyBox(TxtJsonB);

        private void BeautifyBox(TextBox box)
        {
            try
            {
                box.Text = JsonToolService.Beautify(box.Text);
                ToastService.Show(this, "JSON 格式化完成");
            }
            catch (JsonException) { MessageBox.Show("JSON 格式错误，无法格式化", "提示", MessageBoxButton.OK, MessageBoxImage.Warning); }
        }

        private void BtnClearAll_Click(object sender, RoutedEventArgs e)
        {
            TxtJsonA.Clear();
            TxtJsonB.Clear();
            TxtDiffSummary.Text = "●  等待对比";
        }

        private void BtnSwap_Click(object sender, RoutedEventArgs e)
        {
            bool compareAgain = _hasCompared;
            (TxtJsonA.Text, TxtJsonB.Text) = (TxtJsonB.Text, TxtJsonA.Text);
            if (compareAgain) BtnCompare_Click(sender, e);
        }

        private void BtnCompare_Click(object sender, RoutedEventArgs e)
        {
            string jsonA = TxtJsonA.Text, jsonB = TxtJsonB.Text;
            if (string.IsNullOrWhiteSpace(jsonA) && string.IsNullOrWhiteSpace(jsonB))
            {
                TxtDiffSummary.Text = "●  请先在两侧粘贴或导入 JSON";
                TxtDiffSummary.Foreground = Br("#D78A26");
                TxtJsonA.Focus();
                return;
            }

            JsonElement? elA = null, elB = null;
            if (!TryParseInput(jsonA, "A", out elA) || !TryParseInput(jsonB, "B", out elB)) return;
            // Validate both sides before changing either input; locations refer to the formatted text.
            if (elA != null) jsonA = JsonToolService.Beautify(jsonA);
            if (elB != null) jsonB = JsonToolService.Beautify(jsonB);
            TxtJsonA.Text = jsonA;
            TxtJsonB.Text = jsonB;
            _allDiffs = [];
            _unchanged.Clear();
            CompareElements(elA, elB, "$");
            var linesA = FindPathLines(jsonA);
            var linesB = FindPathLines(jsonB);
            int index = 0;
            foreach (var entry in _allDiffs.Concat(_unchanged))
            {
                entry.Index = ++index;
                entry.LineA = linesA.TryGetValue(entry.Path, out var lineA) ? lineA : null;
                entry.LineB = linesB.TryGetValue(entry.Path, out var lineB) ? lineB : null;
            }
            _hasCompared = true;
            RenderDiffResults();
            if (DiffNavigation.Items.Count > 0) DiffNavigation.SelectedIndex = 0;
            ToastService.Show(this, $"对比完成，发现 {_allDiffs.Count} 项差异");
        }

        private bool TryParseInput(string json, string side, out JsonElement? value)
        {
            value = null;
            if (string.IsNullOrWhiteSpace(json)) return true;
            try
            {
                using var doc = JsonDocument.Parse(json);
                value = doc.RootElement.Clone();
                return true;
            }
            catch (JsonException ex)
            {
                TxtDiffSummary.Text = $"●  JSON {side} 格式错误：行 {ex.LineNumber + 1}，字节位置 {ex.BytePositionInLine + 1}";
                TxtDiffSummary.ToolTip = ex.Message;
                TxtDiffSummary.Foreground = Br("#E05A6D");
                var editor = side == "A" ? EditorA : EditorB;
                editor.GoToLine((int)(ex.LineNumber ?? 0) + 1);
                editor.Input.Focus();
                return false;
            }
        }

        // Token locations avoid ambiguous text searches for repeated or escaped keys.
        private static Dictionary<string, int> FindPathLines(string json)
        {
            var result = new Dictionary<string, int>();
            if (string.IsNullOrWhiteSpace(json)) return result;
            byte[] bytes = Encoding.UTF8.GetBytes(json);
            var lineStarts = new List<int> { 0 };
            for (int i = 0; i < bytes.Length; i++)
            {
                if (bytes[i] == '\n' || (bytes[i] == '\r' && (i + 1 == bytes.Length || bytes[i + 1] != '\n')))
                    lineStarts.Add(i + 1);
            }
            var reader = new Utf8JsonReader(bytes);
            reader.Read();
            ReadNode(ref reader, "$", result, lineStarts);
            return result;
        }

        private static void ReadNode(ref Utf8JsonReader reader, string path,
            Dictionary<string, int> result, List<int> lineStarts)
        {
            int found = lineStarts.BinarySearch((int)reader.TokenStartIndex);
            result[path] = found >= 0 ? found + 1 : ~found;
            if (reader.TokenType == JsonTokenType.StartObject)
            {
                while (reader.Read() && reader.TokenType != JsonTokenType.EndObject)
                {
                    string key = reader.GetString()!;
                    reader.Read();
                    ReadNode(ref reader, PropertyPath(path, key), result, lineStarts);
                }
            }
            else if (reader.TokenType == JsonTokenType.StartArray)
            {
                int index = 0;
                while (reader.Read() && reader.TokenType != JsonTokenType.EndArray)
                    ReadNode(ref reader, $"{path}[{index++}]", result, lineStarts);
            }
        }

        private static string PropertyPath(string parent, string key) =>
            key.Length > 0 && (char.IsLetter(key[0]) || key[0] == '_') &&
            key.All(c => char.IsLetterOrDigit(c) || c == '_')
                ? $"{parent}.{key}"
                : $"{parent}[{JsonSerializer.Serialize(key)}]";

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
            if (elA.ValueKind != elB.ValueKind && !(IsBoolean(elA) && IsBoolean(elB)))
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
                    else _unchanged.Add(new DiffEntry { Path = path, Type = DiffType.Unchanged, OldValue = valA, NewValue = valB });
                    break;
            }
        }

        private static bool IsBoolean(JsonElement value) => value.ValueKind is JsonValueKind.True or JsonValueKind.False;

        private void CompareObjects(JsonElement a, JsonElement b, string path)
        {
            var keysA = a.EnumerateObject().Select(p => p.Name).ToHashSet();
            var keysB = b.EnumerateObject().Select(p => p.Name).ToHashSet();
            if (keysA.Count == 0 && keysB.Count == 0)
                _unchanged.Add(new DiffEntry { Path = path, Type = DiffType.Unchanged, OldValue = "{}", NewValue = "{}" });

            // A 有 B 没有 → 删除
            foreach (var key in keysA.Except(keysB))
            {
                var childPath = PropertyPath(path, key);
                CollectAll(a.GetProperty(key), childPath, DiffType.Removed);
            }

            // B 有 A 没有 → 新增
            foreach (var key in keysB.Except(keysA))
            {
                var childPath = PropertyPath(path, key);
                CollectAll(b.GetProperty(key), childPath, DiffType.Added);
            }

            // 两者都有 → 递归对比
            foreach (var key in keysA.Intersect(keysB))
            {
                var childPath = PropertyPath(path, key);
                CompareElements(a.GetProperty(key), b.GetProperty(key), childPath);
            }
        }

        private void CompareArrays(JsonElement a, JsonElement b, string path)
        {
            int lenA = a.GetArrayLength();
            int lenB = b.GetArrayLength();
            int maxLen = Math.Max(lenA, lenB);
            if (maxLen == 0)
                _unchanged.Add(new DiffEntry { Path = path, Type = DiffType.Unchanged, OldValue = "[]", NewValue = "[]" });

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
                case JsonValueKind.Object when el.EnumerateObject().Any():
                    foreach (var prop in el.EnumerateObject())
                        CollectAll(prop.Value, PropertyPath(path, prop.Name), type);
                    break;

                case JsonValueKind.Array when el.GetArrayLength() > 0:
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



        private void RenderDiffResults()
        {
            if (DiffDetails == null) return;
            var filtered = _allDiffs.Where(d => d.Type switch
            {
                DiffType.Added => ChkShowAdded.IsChecked == true,
                DiffType.Removed => ChkShowRemoved.IsChecked == true,
                DiffType.Modified => ChkShowModified.IsChecked == true,
                DiffType.TypeChanged => ChkShowTypeChanged.IsChecked == true,
                _ => false
            }).ToList();
            var selected = DiffNavigation.SelectedItem as DiffEntry;
            _syncingSelection = true;
            DiffNavigation.ItemsSource = filtered;
            var details = ChkOnlyDiff.IsChecked == true ? filtered : filtered.Concat(_unchanged).ToList();
            DiffDetails.ItemsSource = details;
            if (selected != null && filtered.Contains(selected))
            {
                DiffNavigation.SelectedItem = selected;
                DiffDetails.SelectedItem = selected;
            }
            _syncingSelection = false;
            CountAdded.Text = _allDiffs.Count(d => d.Type == DiffType.Added).ToString();
            CountRemoved.Text = _allDiffs.Count(d => d.Type == DiffType.Removed).ToString();
            CountModified.Text = _allDiffs.Count(d => d.Type == DiffType.Modified).ToString();
            CountTypeChanged.Text = _allDiffs.Count(d => d.Type == DiffType.TypeChanged).ToString();
            NavCount.Text = $"{filtered.Count} 处差异";
            NavEmpty.Visibility = filtered.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
            NavEmpty.Text = !_hasCompared ? "对比后可点击差异定位" : _allDiffs.Count == 0 ? "没有发现差异" : "当前筛选无差异";
            EmptyMessage.Visibility = details.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
            EmptyMessage.Text = !_hasCompared ? "在两侧粘贴或导入 JSON，点击「开始对比」查看差异"
                : _allDiffs.Count == 0 ? "两个 JSON 完全相同，没有任何差异" : "当前筛选条件下没有记录，请调整统计卡片或显示选项";
            BtnCopyReport.IsEnabled = BtnExportReport.IsEnabled = _hasCompared;
            TxtDiffSummary.ToolTip = null;
            if (_hasCompared)
            {
                TxtDiffSummary.Text = $"●  已完成对比  ·  {_allDiffs.Count} 处差异  ·  " +
                    (string.IsNullOrWhiteSpace(TxtJsonA.Text) || string.IsNullOrWhiteSpace(TxtJsonB.Text)
                        ? "单侧为空，按整体新增 / 删除处理" : "数据格式有效");
                TxtDiffSummary.Foreground = Br("#16A578");
            }
            ApplyMarkers(EditorA, filtered, true);
            ApplyMarkers(EditorB, filtered, false);
            if (DiffNavigation.SelectedItem is DiffEntry entry)
            {
                EditorA.GoToLine(entry.LineA);
                EditorB.GoToLine(entry.LineB);
            }
        }

        private static void ApplyMarkers(JsonDiffEditor editor, List<DiffEntry> entries, bool original)
        {
            var markers = new Dictionary<int, (Brush, string)>();
            foreach (var entry in entries)
            {
                int? line = original ? entry.LineA : entry.LineB;
                if (line is int number) markers.TryAdd(number, (entry.Accent, entry.Symbol));
            }
            editor.SetMarkers(markers);
        }

        private void DiffNavigation_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_syncingSelection || DiffNavigation.SelectedItem is not DiffEntry entry) return;
            SelectEntry(entry);
        }

        private void DiffDetails_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_syncingSelection || DiffDetails.SelectedItem is not DiffEntry entry) return;
            SelectEntry(entry);
        }

        private void SelectEntry(DiffEntry entry)
        {
            _syncingSelection = true;
            DiffNavigation.SelectedItem = entry.Type == DiffType.Unchanged ? null : entry;
            DiffDetails.SelectedItem = entry;
            DiffDetails.ScrollIntoView(entry);
            if (DiffNavigation.SelectedItem != null) DiffNavigation.ScrollIntoView(entry);
            EditorA.GoToLine(entry.LineA);
            EditorB.GoToLine(entry.LineB);
            _syncingSelection = false;
        }

        private void Filter_Changed(object sender, RoutedEventArgs e)
        {
            if (_hasCompared) RenderDiffResults();
        }

        private void BtnCopyReport_Click(object sender, RoutedEventArgs e)
        {
            if (!_hasCompared) return;
            try
            {
                Clipboard.SetText(BuildReport());
                ToastService.Show(this, "差异报告已复制到剪贴板");
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                ToastService.Show(this, "剪贴板暂时不可用，请重试");
            }
        }

        private void BtnExportReport_Click(object sender, RoutedEventArgs e)
        {
            if (!_hasCompared) return;
            var dlg = new SaveFileDialog
            {
                Filter = "文本报告|*.txt",
                FileName = $"JSON对比报告_{DateTime.Now:yyyyMMdd_HHmmss}.txt"
            };
            if (dlg.ShowDialog() != true) return;
            try
            {
                File.WriteAllText(dlg.FileName, BuildReport(), Encoding.UTF8);
                ToastService.Show(this, "差异报告已导出");
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                MessageBox.Show($"无法保存报告：{ex.Message}", "导出失败", MessageBoxButton.OK, MessageBoxImage.Warning);
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

        private static SolidColorBrush Br(string hex) => new((Color)ColorConverter.ConvertFromString(hex)!);
        private enum DiffType { Added, Removed, Modified, TypeChanged, Unchanged }

        private class DiffEntry
        {
            public int Index { get; set; }
            public int? LineA { get; set; }
            public int? LineB { get; set; }
            public string LineLabel => Type == DiffType.Removed ? $"A 行 {LineA}" : $"B 行 {LineB}";
            public string TypeLabel => Type switch
            {
                DiffType.Added => "新增", DiffType.Removed => "删除", DiffType.Modified => "修改",
                DiffType.TypeChanged => "类型变化", _ => "未变化"
            };
            public string Symbol => Type switch
            {
                DiffType.Added => "+", DiffType.Removed => "−", DiffType.Modified => "~",
                DiffType.TypeChanged => "⇄", _ => "="
            };
            public string TypeDisplay => $"{Symbol}  {TypeLabel}";
            public Brush Accent => Br(Type switch
            {
                DiffType.Added => "#14AD7C", DiffType.Removed => "#F25C73", DiffType.Modified => "#ED9223",
                DiffType.TypeChanged => "#9956EF", _ => "#8B9DB5"
            });
            public Brush RowBackground => Br(Type switch
            {
                DiffType.Added => "#F1FCF8", DiffType.Removed => "#FFF5F7", DiffType.Modified => "#FFFAF2",
                DiffType.TypeChanged => "#F8F3FF", _ => "#FFFFFF"
            });
            public string OldDisplay => Type == DiffType.Added ? "—" : Type == DiffType.TypeChanged ? $"({OldType}) {OldValue}" : OldValue;
            public string NewDisplay => Type == DiffType.Removed ? "—" : Type == DiffType.TypeChanged ? $"({NewType}) {NewValue}" : NewValue;
            public string Path { get; set; } = "";
            public DiffType Type { get; set; }
            public string OldValue { get; set; } = "";
            public string NewValue { get; set; } = "";
            public string OldType { get; set; } = "";
            public string NewType { get; set; } = "";
        }
    }
}
