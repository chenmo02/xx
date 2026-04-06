using Microsoft.Win32;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using WpfApp1.Services;

namespace WpfApp1.Views
{
    public partial class DataImportPage : Page
    {
        private string? _currentFilePath;
        private string? _currentSheetName;
        private DataTable? _currentData;

        public DataImportPage()
        {
            InitializeComponent();
        }

        // ==================== Step 1: 选择文件 ====================

        private async void BtnSelectFile_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "数据文件|*.xlsx;*.xls;*.csv|Excel 文件|*.xlsx;*.xls|CSV 文件|*.csv",
                Title = "选择数据文件"
            };

            if (dialog.ShowDialog() != true) return;

            _currentFilePath = dialog.FileName;
            TxtFilePath.Text = dialog.FileName;
            TxtFilePath.Foreground = Brushes.Black;

            // 清空旧数据
            _currentData = null;
            _currentSheetName = null;
            DgPreview.ItemsSource = null;
            TxtSqlOutput.Text = "";
            TxtSqlStats.Text = "";

            string ext = Path.GetExtension(dialog.FileName).ToLower();

            if (ext is ".xlsx" or ".xls")
            {
                // Excel：加载 Sheet 列表
                BtnSelectFile.IsEnabled = false;
                try
                {
                    var sheets = await Task.Run(() => FileParserService.GetSheetNames(dialog.FileName));

                    CbSheetList.ItemsSource = sheets;
                    CbSheetList.SelectedIndex = 0;
                    PanelSheetSelect.Visibility = Visibility.Visible;

                    TxtFileInfo.Text = $"📄 Excel 文件，共 {sheets.Count} 个工作表";
                }
                catch (Exception ex)
                {
                    TxtFileInfo.Text = $"❌ 读取失败: {ex.Message}";
                    MessageBox.Show($"读取 Excel 失败：\n{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                finally
                {
                    BtnSelectFile.IsEnabled = true;
                }
            }
            else
            {
                // CSV：直接加载
                PanelSheetSelect.Visibility = Visibility.Collapsed;
                await LoadDataAsync(dialog.FileName, null);
            }
        }

        // ==================== Sheet 切换 ====================

        private async void CbSheetList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CbSheetList.SelectedItem == null || string.IsNullOrEmpty(_currentFilePath)) return;

            _currentSheetName = CbSheetList.SelectedItem.ToString();
            await LoadDataAsync(_currentFilePath, _currentSheetName);
        }

        // ==================== 加载数据并预览 ====================

        private async Task LoadDataAsync(string filePath, string? sheetName)
        {
            BtnSelectFile.IsEnabled = false;
            BtnGenerateSql.IsEnabled = false;
            TxtPreviewInfo.Text = "⏳ 正在解析数据...";

            try
            {
                // 获取文件信息
                var info = await Task.Run(() => FileParserService.GetFileInfo(filePath, sheetName));
                string sizeStr = info.fileSize > 1024 * 1024
                    ? $"{info.fileSize / 1024.0 / 1024.0:F1} MB"
                    : $"{info.fileSize / 1024.0:F1} KB";

                string sheetLabel = string.IsNullOrEmpty(sheetName) ? "" : $"  |  Sheet: {sheetName}";
                TxtFileInfo.Text = $"📄 {Path.GetFileName(filePath)}  |  {sizeStr}  |  {info.totalRows} 行 × {info.totalCols} 列{sheetLabel}";

                // 加载预览（前50行）
                var previewData = await Task.Run(() => FileParserService.ParseFile(filePath, sheetName, 50));
                DgPreview.ItemsSource = previewData.DefaultView;
                TxtPreviewInfo.Text = $"显示前 {previewData.Rows.Count} 行（共 {info.totalRows} 行），选择 Sheet 后可切换";

                // 全量加载
                _currentData = await Task.Run(() => FileParserService.ParseFile(filePath, sheetName));
                TxtPreviewInfo.Text = $"✅ 已加载 {_currentData.Rows.Count} 行 × {_currentData.Columns.Count} 列，可以生成 SQL";
                TxtPreviewInfo.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#10B981"));
            }
            catch (Exception ex)
            {
                TxtPreviewInfo.Text = $"❌ 解析失败: {ex.Message}";
                TxtPreviewInfo.Foreground = Brushes.Red;
                _currentData = null;
            }
            finally
            {
                BtnSelectFile.IsEnabled = true;
                BtnGenerateSql.IsEnabled = _currentData != null;
            }
        }

        // ==================== 数据库类型切换 ====================

        private void DbType_Changed(object sender, RoutedEventArgs e)
        {
            if (TxtTableName == null || TxtTableHint == null) return;

            if (RbPostgres.IsChecked == true)
            {
                if (TxtTableName.Text.StartsWith("#"))
                    TxtTableName.Text = "temp_import";
                TxtTableHint.Text = "提示：PostgreSQL 使用 CREATE TEMP TABLE 创建会话级临时表";
            }
            else
            {
                if (TxtTableName.Text.StartsWith("temp_"))
                    TxtTableName.Text = "#temp_import";
                TxtTableHint.Text = "提示：SQL Server 使用 # 前缀创建临时表，## 为全局临时表";
            }

            // 如果已有数据，清空旧 SQL
            TxtSqlOutput.Text = "";
            TxtSqlStats.Text = "";
        }

        // ==================== 生成 SQL ====================

        private async void BtnGenerateSql_Click(object sender, RoutedEventArgs e)
        {
            if (_currentData == null || _currentData.Rows.Count == 0)
            {
                MessageBox.Show("请先选择数据文件并确保有数据！", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(TxtTableName.Text))
            {
                MessageBox.Show("请填写临时表名称！", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!int.TryParse(TxtBatchSize.Text, out int batchSize) || batchSize <= 0)
            {
                MessageBox.Show("每批行数必须为正整数！", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var dbType = RbPostgres.IsChecked == true
                ? SqlGeneratorService.DbType.PostgreSQL
                : SqlGeneratorService.DbType.SqlServer;

            string tableName = TxtTableName.Text.Trim();
            bool dropIfExists = ChkDropIfExists.IsChecked == true;
            bool batchInsert = ChkBatchInsert.IsChecked == true;

            BtnGenerateSql.IsEnabled = false;
            BtnGenerateSql.Content = "⏳ 生成中...";
            TxtSqlOutput.Text = "正在生成 SQL 语句，请稍候...";

            try
            {
                var data = _currentData;
                string sql = await Task.Run(() =>
                    SqlGeneratorService.GenerateFullSql(dbType, tableName, data, dropIfExists, batchInsert, batchSize));

                TxtSqlOutput.Text = sql;

                // 统计信息
                int lineCount = sql.Split('\n').Length;
                double sizeKb = System.Text.Encoding.UTF8.GetByteCount(sql) / 1024.0;
                TxtSqlStats.Text = $"✅ 生成完成  |  {dbType}  |  表: {tableName}  |  {_currentData.Rows.Count} 行数据  |  SQL {lineCount} 行  |  {sizeKb:F1} KB";
            }
            catch (Exception ex)
            {
                TxtSqlOutput.Text = $"生成失败: {ex.Message}";
                MessageBox.Show($"SQL 生成失败：\n{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                BtnGenerateSql.IsEnabled = true;
                BtnGenerateSql.Content = "⚡ 生成 SQL";
            }
        }

        // ==================== Win32 剪贴板 API ====================

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

        private const uint CF_UNICODETEXT = 13;
        private const uint GMEM_MOVEABLE = 0x0002;

        /// <summary>
        /// 使用 Win32 API 写入剪贴板，绕过 WPF/OLE 剪贴板锁定问题
        /// </summary>
        private static bool CopyToClipboardNative(string text)
        {
            if (!OpenClipboard(IntPtr.Zero))
                return false;

            try
            {
                EmptyClipboard();

                var bytes = (text.Length + 1) * 2; // UTF-16
                var hGlobal = GlobalAlloc(GMEM_MOVEABLE, (UIntPtr)bytes);
                if (hGlobal == IntPtr.Zero) return false;

                var target = GlobalLock(hGlobal);
                if (target == IntPtr.Zero) return false;

                Marshal.Copy(text.ToCharArray(), 0, target, text.Length);
                // 写入末尾 null 终止符
                Marshal.WriteInt16(target, text.Length * 2, 0);
                GlobalUnlock(hGlobal);

                SetClipboardData(CF_UNICODETEXT, hGlobal);
                return true;
            }
            finally
            {
                CloseClipboard();
            }
        }

        // ==================== 复制到剪贴板 ====================

        private void BtnCopySql_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TxtSqlOutput.Text))
            {
                MessageBox.Show("没有可复制的 SQL 语句，请先生成！", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            bool success = false;

            // 方式1：Win32 原生 API（最可靠）
            try
            {
                success = CopyToClipboardNative(TxtSqlOutput.Text);
            }
            catch
            {
                success = false;
            }

            // 方式2：回退到 WPF 方式（带重试）
            if (!success)
            {
                for (int i = 0; i < 5; i++)
                {
                    try
                    {
                        Clipboard.SetDataObject(TxtSqlOutput.Text, true);
                        success = true;
                        break;
                    }
                    catch
                    {
                        Thread.Sleep(150);
                    }
                }
            }

            if (success)
            {
                BtnCopySql.Content = "✅ 已复制！";
                var timer = new System.Windows.Threading.DispatcherTimer
                {
                    Interval = TimeSpan.FromSeconds(2)
                };
                timer.Tick += (s, _) =>
                {
                    BtnCopySql.Content = "📋 复制到剪贴板";
                    ((System.Windows.Threading.DispatcherTimer)s!).Stop();
                };
                timer.Start();
            }
            else
            {
                MessageBox.Show("剪贴板被其他程序占用，复制失败。\n\n替代方案：点击 SQL 文本框，Ctrl+A 全选后 Ctrl+C 复制。",
                    "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        // ==================== 保存为文件 ====================

        private void BtnSaveSql_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TxtSqlOutput.Text))
            {
                MessageBox.Show("没有可保存的 SQL 语句，请先生成！", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var dialog = new SaveFileDialog
            {
                Filter = "SQL 文件|*.sql|文本文件|*.txt",
                FileName = $"{TxtTableName.Text}_{DateTime.Now:yyyyMMdd_HHmmss}.sql",
                Title = "保存 SQL 文件"
            };

            if (dialog.ShowDialog() == true)
            {
                File.WriteAllText(dialog.FileName, TxtSqlOutput.Text, System.Text.Encoding.UTF8);
                MessageBox.Show($"SQL 文件已保存到：\n{dialog.FileName}", "保存成功", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
    }
}
