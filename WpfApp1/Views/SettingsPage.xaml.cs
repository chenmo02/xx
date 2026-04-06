using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;

namespace WpfApp1.Views
{
    public partial class SettingsPage : Page
    {
        // ── Win32 剪贴板 API ──
        [DllImport("user32.dll")] private static extern bool OpenClipboard(IntPtr hWndNewOwner);
        [DllImport("user32.dll")] private static extern bool CloseClipboard();
        [DllImport("user32.dll")] private static extern bool EmptyClipboard();
        [DllImport("user32.dll")] private static extern IntPtr SetClipboardData(uint uFormat, IntPtr hMem);
        private const uint CF_UNICODETEXT = 13;

        private readonly string _configPath;
        private DispatcherTimer? _timestampTimer;

        public SettingsPage()
        {
            InitializeComponent();

            // 配置文件路径：exe 同目录
            var exeDir = AppDomain.CurrentDomain.BaseDirectory;
            _configPath = Path.Combine(exeDir, "settings.json");

            LoadSettings();
            StartTimestampTimer();
            LoadAboutInfo();
        }

        // ═══════════════════════════════════════
        // 设置持久化
        // ═══════════════════════════════════════

        private void LoadSettings()
        {
            try
            {
                if (!File.Exists(_configPath)) return;

                var json = File.ReadAllText(_configPath, Encoding.UTF8);
                var cfg = JsonSerializer.Deserialize<Dictionary<string, string>>(json);
                if (cfg == null) return;

                if (cfg.TryGetValue("DefaultDbType", out var dbType))
                    CmbDefaultDbType.SelectedIndex = dbType == "SQLServer" ? 1 : 0;

                if (cfg.TryGetValue("TempTablePrefix", out var prefix))
                    TxtTempTablePrefix.Text = prefix;

                if (cfg.TryGetValue("BatchSize", out var batch))
                    TxtBatchSize.Text = batch;

                if (cfg.TryGetValue("DefaultExportPath", out var exportPath))
                    TxtDefaultExportPath.Text = exportPath;
            }
            catch
            {
                // 配置文件损坏时忽略
            }
        }

        private void BtnSaveSettings_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var cfg = new Dictionary<string, string>
                {
                    ["DefaultDbType"] = CmbDefaultDbType.SelectedIndex == 1 ? "SQLServer" : "PostgreSQL",
                    ["TempTablePrefix"] = TxtTempTablePrefix.Text.Trim(),
                    ["BatchSize"] = TxtBatchSize.Text.Trim(),
                    ["DefaultExportPath"] = TxtDefaultExportPath.Text.Trim()
                };

                var json = JsonSerializer.Serialize(cfg, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(_configPath, json, Encoding.UTF8);

                ShowToast("✅ 设置已保存");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnSelectExportPath_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFolderDialog { Title = "选择默认导出路径" };
            if (dialog.ShowDialog() == true)
                TxtDefaultExportPath.Text = dialog.FolderName;
        }

        // ═══════════════════════════════════════
        // 时间戳转换
        // ═══════════════════════════════════════

        private void StartTimestampTimer()
        {
            UpdateCurrentTimestamp();
            _timestampTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1) };
            _timestampTimer.Tick += (_, _) => UpdateCurrentTimestamp();
            _timestampTimer.Start();
        }

        private void UpdateCurrentTimestamp()
        {
            var now = DateTimeOffset.Now;
            TxtCurrentTimestamp.Text = $"{now.ToUnixTimeSeconds()}  （{now:yyyy-MM-dd HH:mm:ss}）";
        }

        private void BtnCopyCurrentTimestamp_Click(object sender, RoutedEventArgs e)
        {
            var ts = DateTimeOffset.Now.ToUnixTimeSeconds().ToString();
            SafeCopyToClipboard(ts);
            ShowToast("✅ 已复制当前时间戳");
        }

        private void BtnTimestampToDate_Click(object sender, RoutedEventArgs e)
        {
            var input = TxtTimestampInput.Text.Trim();
            if (string.IsNullOrEmpty(input))
            {
                TxtTimestampResult.Text = "请输入时间戳";
                TxtTimestampResult.Foreground = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#E53E3E"));
                return;
            }

            if (!long.TryParse(input, out var ts))
            {
                TxtTimestampResult.Text = "❌ 无效的时间戳";
                TxtTimestampResult.Foreground = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#E53E3E"));
                return;
            }

            try
            {
                DateTimeOffset dto;
                string unit;

                // 自动识别秒/毫秒：大于 1e12 认为是毫秒
                if (ts > 9_999_999_999L)
                {
                    dto = DateTimeOffset.FromUnixTimeMilliseconds(ts).ToLocalTime();
                    unit = "毫秒";
                }
                else
                {
                    dto = DateTimeOffset.FromUnixTimeSeconds(ts).ToLocalTime();
                    unit = "秒";
                }

                TxtTimestampResult.Text = $"{dto:yyyy-MM-dd HH:mm:ss.fff}  （{unit}级）";
                TxtTimestampResult.Foreground = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#333"));
            }
            catch
            {
                TxtTimestampResult.Text = "❌ 转换失败，时间戳超出范围";
                TxtTimestampResult.Foreground = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#E53E3E"));
            }
        }

        private void BtnDateToTimestamp_Click(object sender, RoutedEventArgs e)
        {
            var input = TxtDateInput.Text.Trim();
            if (string.IsNullOrEmpty(input))
            {
                TxtDateResult.Text = "请输入日期时间";
                TxtDateResult.Foreground = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#E53E3E"));
                return;
            }

            // 支持多种常见格式
            string[] formats = [
                "yyyy-MM-dd HH:mm:ss",
                "yyyy-MM-dd HH:mm",
                "yyyy-MM-dd",
                "yyyy/MM/dd HH:mm:ss",
                "yyyy/MM/dd HH:mm",
                "yyyy/MM/dd",
                "yyyyMMdd HHmmss",
                "yyyyMMdd",
                "MM/dd/yyyy HH:mm:ss",
                "MM/dd/yyyy"
            ];

            if (DateTimeOffset.TryParseExact(input, formats, CultureInfo.InvariantCulture,
                    DateTimeStyles.AssumeLocal, out var dto))
            {
                TxtDateResult.Text = $"秒: {dto.ToUnixTimeSeconds()}    毫秒: {dto.ToUnixTimeMilliseconds()}";
                TxtDateResult.Foreground = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#333"));
            }
            else
            {
                TxtDateResult.Text = "❌ 无法识别的日期格式";
                TxtDateResult.Foreground = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#E53E3E"));
            }
        }

        // ═══════════════════════════════════════
        // Base64 编解码
        // ═══════════════════════════════════════

        private void BtnBase64Encode_Click(object sender, RoutedEventArgs e)
        {
            var input = TxtBase64Input.Text;
            if (string.IsNullOrEmpty(input)) return;

            try
            {
                var bytes = Encoding.UTF8.GetBytes(input);
                TxtBase64Output.Text = Convert.ToBase64String(bytes);
            }
            catch (Exception ex)
            {
                TxtBase64Output.Text = $"❌ 编码失败：{ex.Message}";
            }
        }

        private void BtnBase64Decode_Click(object sender, RoutedEventArgs e)
        {
            var input = TxtBase64Output.Text.Trim();
            if (string.IsNullOrEmpty(input)) return;

            try
            {
                var bytes = Convert.FromBase64String(input);
                TxtBase64Input.Text = Encoding.UTF8.GetString(bytes);
            }
            catch (Exception ex)
            {
                TxtBase64Input.Text = $"❌ 解码失败：{ex.Message}";
            }
        }

        private void BtnCopyBase64Result_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(TxtBase64Output.Text))
            {
                SafeCopyToClipboard(TxtBase64Output.Text);
                ShowToast("✅ 已复制");
            }
        }

        // ═══════════════════════════════════════
        // UUID / GUID 生成
        // ═══════════════════════════════════════

        private void BtnGenerateUuid_Click(object sender, RoutedEventArgs e)
        {
            if (!int.TryParse(TxtUuidCount.Text.Trim(), out var count) || count < 1)
                count = 1;
            if (count > 100) count = 100;

            var uppercase = ChkUuidUppercase.IsChecked == true;
            var noDash = ChkUuidNoDash.IsChecked == true;

            var sb = new StringBuilder();
            for (int i = 0; i < count; i++)
            {
                var uuid = Guid.NewGuid().ToString();
                if (noDash) uuid = uuid.Replace("-", "");
                if (uppercase) uuid = uuid.ToUpperInvariant();
                if (i > 0) sb.AppendLine();
                sb.Append(uuid);
            }

            TxtUuidResult.Text = sb.ToString();
        }

        private void BtnCopyUuid_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(TxtUuidResult.Text))
            {
                SafeCopyToClipboard(TxtUuidResult.Text);
                ShowToast("✅ 已复制");
            }
        }

        // ═══════════════════════════════════════
        // 文本哈希计算
        // ═══════════════════════════════════════

        private void BtnCalcHash_Click(object sender, RoutedEventArgs e)
        {
            var input = TxtHashInput.Text;
            if (string.IsNullOrEmpty(input))
            {
                TxtHashMd5.Text = TxtHashSha1.Text = TxtHashSha256.Text = "";
                return;
            }

            var bytes = Encoding.UTF8.GetBytes(input);

            TxtHashMd5.Text = BitConverter.ToString(MD5.HashData(bytes)).Replace("-", "").ToLowerInvariant();
            TxtHashSha1.Text = BitConverter.ToString(SHA1.HashData(bytes)).Replace("-", "").ToLowerInvariant();
            TxtHashSha256.Text = BitConverter.ToString(SHA256.HashData(bytes)).Replace("-", "").ToLowerInvariant();
        }

        private void BtnCopyHashMd5_Click(object sender, RoutedEventArgs e) => CopyHashResult(TxtHashMd5.Text);
        private void BtnCopyHashSha1_Click(object sender, RoutedEventArgs e) => CopyHashResult(TxtHashSha1.Text);
        private void BtnCopyHashSha256_Click(object sender, RoutedEventArgs e) => CopyHashResult(TxtHashSha256.Text);

        private void CopyHashResult(string text)
        {
            if (!string.IsNullOrEmpty(text))
            {
                SafeCopyToClipboard(text);
                ShowToast("✅ 已复制");
            }
        }

        // ═══════════════════════════════════════
        // URL 编解码
        // ═══════════════════════════════════════

        private void BtnUrlEncode_Click(object sender, RoutedEventArgs e)
        {
            var input = TxtUrlInput.Text;
            if (string.IsNullOrEmpty(input)) return;
            TxtUrlOutput.Text = Uri.EscapeDataString(input);
        }

        private void BtnUrlDecode_Click(object sender, RoutedEventArgs e)
        {
            var input = TxtUrlOutput.Text.Trim();
            if (string.IsNullOrEmpty(input)) return;

            try
            {
                TxtUrlInput.Text = Uri.UnescapeDataString(input);
            }
            catch (Exception ex)
            {
                TxtUrlInput.Text = $"❌ 解码失败：{ex.Message}";
            }
        }

        private void BtnCopyUrlResult_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(TxtUrlOutput.Text))
            {
                SafeCopyToClipboard(TxtUrlOutput.Text);
                ShowToast("✅ 已复制");
            }
        }

        // ═══════════════════════════════════════
        // 正则表达式测试
        // ═══════════════════════════════════════

        private void BtnRegexTest_Click(object sender, RoutedEventArgs e)
        {
            var pattern = TxtRegexPattern.Text;
            var input = TxtRegexInput.Text;

            if (string.IsNullOrEmpty(pattern))
            {
                TxtRegexResult.Text = "请输入正则表达式";
                TxtRegexMatchCount.Text = "";
                return;
            }

            if (string.IsNullOrEmpty(input))
            {
                TxtRegexResult.Text = "请输入测试文本";
                TxtRegexMatchCount.Text = "";
                return;
            }

            try
            {
                var options = RegexOptions.None;
                if (ChkRegexIgnoreCase.IsChecked == true)
                    options |= RegexOptions.IgnoreCase;

                var matches = Regex.Matches(input, pattern, options, TimeSpan.FromSeconds(5));
                TxtRegexMatchCount.Text = $"找到 {matches.Count} 个匹配";

                if (matches.Count == 0)
                {
                    TxtRegexResult.Text = "（无匹配）";
                    return;
                }

                var sb = new StringBuilder();
                int idx = 0;
                foreach (Match m in matches)
                {
                    sb.AppendLine($"[{idx}] 位置 {m.Index}, 长度 {m.Length}: \"{m.Value}\"");

                    // 显示捕获组
                    if (m.Groups.Count > 1)
                    {
                        for (int g = 1; g < m.Groups.Count; g++)
                        {
                            sb.AppendLine($"     组{g}: \"{m.Groups[g].Value}\"");
                        }
                    }
                    idx++;
                }

                TxtRegexResult.Text = sb.ToString().TrimEnd();
            }
            catch (RegexParseException ex)
            {
                TxtRegexResult.Text = $"❌ 正则语法错误：{ex.Message}";
                TxtRegexMatchCount.Text = "";
            }
            catch (RegexMatchTimeoutException)
            {
                TxtRegexResult.Text = "❌ 匹配超时（可能存在灾难性回溯）";
                TxtRegexMatchCount.Text = "";
            }
        }

        // ═══════════════════════════════════════
        // 文本对比
        // ═══════════════════════════════════════

        private void BtnDiffCompare_Click(object sender, RoutedEventArgs e)
        {
            var textA = TxtDiffA.Text ?? "";
            var textB = TxtDiffB.Text ?? "";

            var linesA = textA.Split('\n').Select(l => l.TrimEnd('\r')).ToArray();
            var linesB = textB.Split('\n').Select(l => l.TrimEnd('\r')).ToArray();

            var results = new List<DiffLine>();
            int maxLen = Math.Max(linesA.Length, linesB.Length);
            int addCount = 0, removeCount = 0, sameCount = 0;

            for (int i = 0; i < maxLen; i++)
            {
                var a = i < linesA.Length ? linesA[i] : null;
                var b = i < linesB.Length ? linesB[i] : null;

                if (a == b)
                {
                    results.Add(new DiffLine($"  {i + 1:D3} | {a}", "#555"));
                    sameCount++;
                }
                else
                {
                    if (a != null)
                    {
                        results.Add(new DiffLine($"- {i + 1:D3} | {a}", "#E53E3E"));
                        removeCount++;
                    }
                    if (b != null)
                    {
                        results.Add(new DiffLine($"+ {i + 1:D3} | {b}", "#38A169"));
                        addCount++;
                    }
                }
            }

            TxtDiffSummary.Text = $"共 {maxLen} 行，相同 {sameCount} 行，删除 {removeCount} 行，新增 {addCount} 行";
            DiffResultPanel.Visibility = Visibility.Visible;

            // 用 StackPanel 手动构建带颜色的行
            DiffResultList.Items.Clear();
            foreach (var line in results)
            {
                var tb = new TextBlock
                {
                    Text = line.Text,
                    FontFamily = new System.Windows.Media.FontFamily("Consolas, Courier New"),
                    FontSize = 12.5,
                    Foreground = new System.Windows.Media.SolidColorBrush(
                        (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(line.Color)),
                    Margin = new Thickness(0, 1, 0, 1)
                };
                DiffResultList.Items.Add(tb);
            }
        }

        private record DiffLine(string Text, string Color);

        // ═══════════════════════════════════════
        // 关于信息
        // ═══════════════════════════════════════

        private void LoadAboutInfo()
        {
            TxtAboutRuntime.Text = $".NET {Environment.Version.Major}";
            TxtAboutOS.Text = $"Win {Environment.OSVersion.Version.Major}.{Environment.OSVersion.Version.Build}";
            TxtAboutDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
        }

        // ═══════════════════════════════════════
        // 剪贴板工具（Win32 原生 API + WPF 回退）
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
                    finally
                    {
                        CloseClipboard();
                    }
                }
                else
                {
                    // 回退到 WPF
                    Clipboard.SetText(text);
                }
            }
            catch
            {
                try { Clipboard.SetText(text); } catch { }
            }
        }

        // ═══════════════════════════════════════
        // Toast 提示
        // ═══════════════════════════════════════

        private void ShowToast(string message)
        {
            // 简单实现：在页面顶部短暂显示消息
            var popup = new Border
            {
                Background = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#333")),
                CornerRadius = new CornerRadius(6),
                Padding = new Thickness(16, 8, 16, 8),
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Top,
                Margin = new Thickness(0, 10, 0, 0),
                Opacity = 0.92,
                Child = new TextBlock
                {
                    Text = message,
                    Foreground = System.Windows.Media.Brushes.White,
                    FontSize = 13
                }
            };

            // 找到页面的根 Grid 或 ScrollViewer 的父级
            if (this.Content is ScrollViewer sv)
            {
                var grid = new Grid();
                this.Content = grid;
                grid.Children.Add(sv);
                grid.Children.Add(popup);

                var timer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1.5) };
                timer.Tick += (_, _) =>
                {
                    grid.Children.Remove(popup);
                    timer.Stop();
                };
                timer.Start();
            }
        }
    }
}
