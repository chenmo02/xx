using Microsoft.Win32;
using System.Data;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using CsvHelper;
using CsvHelper.Configuration;

namespace WpfApp1.Views
{
    public partial class CsvViewerPage : Page
    {
        private DataTable? _data;
        private readonly List<(int row, int col)> _matchPositions = [];
        private int _currentMatchIndex = -1;

        public CsvViewerPage()
        {
            InitializeComponent();
        }

        // ==================== 打开文件 ====================

        private async void BtnOpen_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Filter = "CSV 文件|*.csv|文本文件|*.txt;*.tsv|所有文件|*.*",
                Title = "打开 CSV 文件"
            };

            if (dlg.ShowDialog() != true) return;

            BtnOpen.IsEnabled = false;
            TxtStatus.Text = "⏳ 正在加载...";

            try
            {
                string filePath = dlg.FileName;
                var fileInfo = new FileInfo(filePath);
                string sizeStr = fileInfo.Length > 1024 * 1024
                    ? $"{fileInfo.Length / 1024.0 / 1024.0:F1} MB"
                    : $"{fileInfo.Length / 1024.0:F1} KB";

                // 检测编码
                var encoding = DetectEncoding(filePath);
                string encodingName = encoding.EncodingName;

                // 检测分隔符
                string delimiter = DetectDelimiter(filePath, encoding);
                string delimiterName = delimiter switch
                {
                    "," => "逗号",
                    "\t" => "Tab",
                    ";" => "分号",
                    "|" => "管道符",
                    _ => delimiter
                };

                // 异步解析
                _data = await Task.Run(() => ParseCsv(filePath, encoding, delimiter));

                DgCsv.ItemsSource = _data.DefaultView;

                TxtFileInfo.Text = $"📄 {Path.GetFileName(filePath)}  |  {sizeStr}  |  {_data.Rows.Count} 行 × {_data.Columns.Count} 列  |  分隔符: {delimiterName}";
                TxtEncoding.Text = $"编码: {encodingName}";
                TxtStatus.Text = $"✅ 已加载 {_data.Rows.Count} 行";
                TxtRowCol.Text = $"共 {_data.Rows.Count} 行, {_data.Columns.Count} 列";

                // 清除搜索
                ClearSearch();
            }
            catch (Exception ex)
            {
                TxtStatus.Text = $"❌ 加载失败";
                MessageBox.Show($"文件加载失败：\n{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                BtnOpen.IsEnabled = true;
            }
        }

        // ==================== 搜索功能 ====================

        private void TxtSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter) DoSearch();
        }

        private void BtnSearch_Click(object sender, RoutedEventArgs e) => DoSearch();

        private void BtnNextMatch_Click(object sender, RoutedEventArgs e)
        {
            if (_matchPositions.Count == 0) return;

            _currentMatchIndex = (_currentMatchIndex + 1) % _matchPositions.Count;
            NavigateToMatch(_currentMatchIndex);
        }

        private void BtnClearSearch_Click(object sender, RoutedEventArgs e) => ClearSearch();

        private void DoSearch()
        {
            if (_data == null || _data.Rows.Count == 0)
            {
                TxtMatchInfo.Text = "请先打开文件";
                return;
            }

            string keyword = TxtSearch.Text.Trim();
            if (string.IsNullOrEmpty(keyword))
            {
                TxtMatchInfo.Text = "请输入搜索关键词";
                return;
            }

            _matchPositions.Clear();
            _currentMatchIndex = -1;

            // 全表搜索
            for (int r = 0; r < _data.Rows.Count; r++)
            {
                for (int c = 0; c < _data.Columns.Count; c++)
                {
                    string val = _data.Rows[r][c]?.ToString() ?? "";
                    if (val.Contains(keyword, StringComparison.OrdinalIgnoreCase))
                    {
                        _matchPositions.Add((r, c));
                    }
                }
            }

            if (_matchPositions.Count == 0)
            {
                TxtMatchInfo.Text = $"未找到 \"{keyword}\"";
                TxtMatchInfo.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F73131")!);
                TxtStatus.Text = $"🔍 未找到匹配项";
            }
            else
            {
                _currentMatchIndex = 0;
                TxtMatchInfo.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#4E6EF2")!);
                NavigateToMatch(0);
            }
        }

        private void NavigateToMatch(int index)
        {
            if (index < 0 || index >= _matchPositions.Count || _data == null) return;

            var (row, col) = _matchPositions[index];

            TxtMatchInfo.Text = $"匹配 {index + 1}/{_matchPositions.Count}";
            TxtStatus.Text = $"🔍 第 {row + 1} 行, 第 {col + 1} 列 ({_data.Columns[col].ColumnName})";

            // 滚动到目标行并选中
            DgCsv.SelectedIndex = row;
            DgCsv.ScrollIntoView(DgCsv.Items[row]);

            // 尝试选中具体单元格
            DgCsv.Focus();
            DgCsv.UpdateLayout();

            try
            {
                if (DgCsv.Columns.Count > col)
                {
                    var cellInfo = new DataGridCellInfo(DgCsv.Items[row], DgCsv.Columns[col]);
                    DgCsv.CurrentCell = cellInfo;
                }
            }
            catch { }
        }

        private void ClearSearch()
        {
            TxtSearch.Text = "";
            _matchPositions.Clear();
            _currentMatchIndex = -1;
            TxtMatchInfo.Text = "";
            TxtMatchInfo.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#999")!);
        }

        // ==================== CSV 解析 ====================

        private static DataTable ParseCsv(string filePath, Encoding encoding, string delimiter)
        {
            var dt = new DataTable();

            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                Delimiter = delimiter,
                HasHeaderRecord = true,
                MissingFieldFound = null,
                BadDataFound = null,
                TrimOptions = TrimOptions.Trim
            };

            using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = new StreamReader(fs, encoding);
            using var csv = new CsvReader(reader, config);

            csv.Read();
            csv.ReadHeader();

            if (csv.HeaderRecord != null)
            {
                for (int i = 0; i < csv.HeaderRecord.Length; i++)
                {
                    string colName = csv.HeaderRecord[i]?.Trim() ?? $"列{i + 1}";
                    if (string.IsNullOrWhiteSpace(colName)) colName = $"列{i + 1}";
                    if (dt.Columns.Contains(colName)) colName = $"{colName}_{i + 1}";
                    dt.Columns.Add(colName, typeof(string));
                }
            }

            while (csv.Read())
            {
                var row = dt.NewRow();
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    try { row[i] = csv.GetField(i) ?? ""; }
                    catch { row[i] = ""; }
                }
                dt.Rows.Add(row);
            }

            return dt;
        }

        // ==================== 编码检测 ====================

        private static Encoding DetectEncoding(string filePath)
        {
            var bom = new byte[4];
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                _ = fs.ReadAtLeast(bom, 4, throwOnEndOfStream: false);
            }

            if (bom[0] == 0xEF && bom[1] == 0xBB && bom[2] == 0xBF) return Encoding.UTF8;
            if (bom[0] == 0xFF && bom[1] == 0xFE) return Encoding.Unicode;
            if (bom[0] == 0xFE && bom[1] == 0xFF) return Encoding.BigEndianUnicode;

            try
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                var utf8 = new UTF8Encoding(false, true);
                using var fs2 = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var sr = new StreamReader(fs2, utf8, true);
                sr.ReadToEnd();
                return Encoding.UTF8;
            }
            catch
            {
                return Encoding.GetEncoding("GBK");
            }
        }

        // ==================== 分隔符检测 ====================

        private static string DetectDelimiter(string filePath, Encoding encoding)
        {
            using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = new StreamReader(fs, encoding);
            string? firstLine = reader.ReadLine();
            if (string.IsNullOrEmpty(firstLine)) return ",";

            int comma = firstLine.Count(c => c == ',');
            int tab = firstLine.Count(c => c == '\t');
            int semi = firstLine.Count(c => c == ';');
            int pipe = firstLine.Count(c => c == '|');

            int max = Math.Max(Math.Max(comma, tab), Math.Max(semi, pipe));
            if (max == 0) return ",";
            if (max == tab) return "\t";
            if (max == semi) return ";";
            if (max == pipe) return "|";
            return ",";
        }
    }
}
