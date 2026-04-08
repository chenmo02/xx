using Microsoft.Win32;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using WpfApp1.Services;

namespace WpfApp1.Views
{
    public partial class DataImportPage : Page
    {
        private static readonly Brush SuccessBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#10B981"));
        private static readonly Brush ErrorBrush = Brushes.Red;
        private static readonly Brush MutedBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#9CA3AF"));

        private string? _currentFilePath;
        private string? _currentSheetName;
        private string? _currentDbfEncoding;
        private DataTable? _currentData;
        private ImportSettings _importSettings = new();
        private string? _lastSuggestedTableName;
        private bool _isBusy;
        private bool _settingsSubscribed;
        private bool _suppressDbfEncodingReload;

        public DataImportPage()
        {
            InitializeComponent();
            Unloaded += Page_Unloaded;
        }

        private void Page_Loaded(object sender, RoutedEventArgs e)
        {
            if (!_settingsSubscribed)
            {
                ImportSettingsService.SettingsSaved += ImportSettingsService_SettingsSaved;
                _settingsSubscribed = true;
            }

            LoadImportSettings(forceTableName: true);

            if (CbDbfEncoding.Items.Count == 0)
            {
                _suppressDbfEncodingReload = true;
                CbDbfEncoding.ItemsSource = FileParserService.GetDbfEncodings();
                CbDbfEncoding.SelectedItem = "UTF-8";
                _suppressDbfEncodingReload = false;
            }

            ApplyDatabaseHint(forceTableName: true);
            UpdateStatus("就绪");
            Keyboard.Focus(this);
        }

        private void Page_Unloaded(object sender, RoutedEventArgs e)
        {
            if (_settingsSubscribed)
            {
                ImportSettingsService.SettingsSaved -= ImportSettingsService_SettingsSaved;
                _settingsSubscribed = false;
            }
        }

        private void ImportSettingsService_SettingsSaved(object? sender, ImportSettings settings)
        {
            Dispatcher.Invoke(() =>
            {
                _importSettings = ImportSettingsService.Normalize(settings);
                ApplyImportSettingsToUi(forceTableName: false);
            });
        }

        private async void BtnSelectFile_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "数据文件|*.xlsx;*.xls;*.csv;*.dbf|Excel 文件|*.xlsx;*.xls|CSV 文件|*.csv|DBF 文件|*.dbf",
                Title = "选择数据文件"
            };

            ApplyDefaultExportPath(dialog);

            if (dialog.ShowDialog() == true)
            {
                await LoadSelectedFileAsync(dialog.FileName);
            }
        }

        private async Task LoadSelectedFileAsync(string filePath)
        {
            if (_isBusy)
            {
                return;
            }

            ResetCurrentData();

            _currentFilePath = filePath;
            _currentSheetName = null;
            TxtFilePath.Text = filePath;
            TxtFilePath.Foreground = Brushes.Black;

            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            PanelDbfEncoding.Visibility = ext == ".dbf" ? Visibility.Visible : Visibility.Collapsed;
            PanelSheetSelect.Visibility = ext is ".xlsx" or ".xls" ? Visibility.Visible : Visibility.Collapsed;

            if (ext == ".dbf")
            {
                _currentDbfEncoding = CbDbfEncoding.SelectedItem?.ToString() ?? "UTF-8";
            }
            else
            {
                _currentDbfEncoding = null;
            }

            if (ext is ".xlsx" or ".xls")
            {
                List<string> sheets = [];
                try
                {
                    SetBusyState(true, "正在读取工作表列表...");
                    sheets = await Task.Run(() => FileParserService.GetSheetNames(filePath));
                    CbSheetList.ItemsSource = sheets;
                    if (sheets.Count > 0)
                    {
                        _currentSheetName = sheets[0];
                        CbSheetList.SelectedIndex = 0;
                    }
                    else
                    {
                        _currentSheetName = null;
                        CbSheetList.SelectedIndex = -1;
                    }

                    TxtFileInfo.Text = $"文件: {Path.GetFileName(filePath)}  |  工作表: {sheets.Count}";
                }
                catch (Exception ex)
                {
                    HandleLoadError(ex, "读取 Excel 工作表失败");
                }
                finally
                {
                    SetBusyState(false, "就绪");
                }

                if (sheets.Count > 0 && !string.IsNullOrWhiteSpace(_currentSheetName))
                {
                    await LoadDataAsync(filePath, _currentSheetName);
                }
            }
            else
            {
                PanelSheetSelect.Visibility = Visibility.Collapsed;
                await LoadDataAsync(filePath, sheetName: null);
            }
        }

        private async void CbSheetList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isBusy || CbSheetList.SelectedItem == null || string.IsNullOrWhiteSpace(_currentFilePath))
            {
                return;
            }

            _currentSheetName = CbSheetList.SelectedItem.ToString();
            await LoadDataAsync(_currentFilePath, _currentSheetName);
        }

        private async void CbDbfEncoding_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_suppressDbfEncodingReload || _isBusy || string.IsNullOrWhiteSpace(_currentFilePath))
            {
                return;
            }

            if (!string.Equals(Path.GetExtension(_currentFilePath), ".dbf", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            _currentDbfEncoding = CbDbfEncoding.SelectedItem?.ToString() ?? "UTF-8";
            await LoadDataAsync(_currentFilePath, sheetName: null);
        }

        private async Task LoadDataAsync(string filePath, string? sheetName)
        {
            try
            {
                SetBusyState(true, "正在读取文件...");
                ShowProgress(0);

                var progress = new Progress<ImportProgressInfo>(info =>
                {
                    UpdateStatus(info.Stage);
                    ShowProgress(info.Percentage);
                });

                _currentData = await Task.Run(() => FileParserService.ParseFile(filePath, sheetName, 0, _currentDbfEncoding, progress));
                DgPreview.ItemsSource = _currentData.DefaultView;

                long fileSize = new FileInfo(filePath).Length;
                UpdateFileInfo(filePath, sheetName, fileSize, _currentData.Rows.Count, _currentData.Columns.Count);
                ShowLoadHints(_currentData.Rows.Count);

                TxtPreviewInfo.Text = $"已加载 {_currentData.Rows.Count:N0} 行 × {_currentData.Columns.Count:N0} 列，可直接编辑并重新生成 SQL。";
                TxtPreviewInfo.Foreground = SuccessBrush;
                UpdateStatus("文件加载完成");
                UpdateExportButtons();
            }
            catch (Exception ex)
            {
                HandleLoadError(ex, "加载数据失败");
            }
            finally
            {
                HideProgress();
                SetBusyState(false, "就绪");
            }
        }

        private void DbType_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (TxtTableHint is null || TxtTableName is null || TxtSqlOutput is null || TxtSqlStats is null)
            {
                return;
            }

            ApplyDatabaseHint(forceTableName: false);
            TxtSqlOutput.Clear();
            TxtSqlStats.Text = "";
        }

        private async void BtnGenerateSql_Click(object sender, RoutedEventArgs e)
        {
            if (_isBusy)
            {
                return;
            }

            if (_currentData == null || _currentData.Rows.Count == 0)
            {
                MessageBox.Show("请先导入数据文件，并确保表格中存在数据。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(TxtTableName.Text))
            {
                MessageBox.Show("表名不能为空。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!int.TryParse(TxtBatchSize.Text, out int batchSize) || batchSize <= 0)
            {
                MessageBox.Show("每批行数必须是正整数。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            CommitPendingGridEdits();
            DataTable dataSnapshot = _currentData.Copy();

            if (dataSnapshot.Rows.Count > 100000)
            {
                var result = MessageBox.Show(
                    $"当前数据量为 {dataSnapshot.Rows.Count:N0} 行，生成 SQL 可能较慢且文件较大。\n\n是否继续生成？",
                    "大数据量警告",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);

                if (result != MessageBoxResult.Yes)
                {
                    UpdateStatus("已取消生成");
                    return;
                }
            }

            SqlGeneratorService.DbType dbType = GetSelectedDbType();
            string tableName = TxtTableName.Text.Trim();
            bool dropIfExists = ChkDropIfExists.IsChecked == true;
            bool batchInsert = ChkBatchInsert.IsChecked == true;
            bool limitFieldLength = ChkLimitFieldLength.IsChecked == true;

            try
            {
                SetBusyState(true, "正在生成 SQL...");
                ShowProgress(15);
                TxtSqlOutput.Text = "正在生成 SQL，请稍候...";

                string sql = await Task.Run(() => SqlGeneratorService.GenerateFullSql(
                    dbType,
                    tableName,
                    dataSnapshot,
                    dropIfExists: dropIfExists,
                    batchInsert: batchInsert,
                    batchSize: batchSize,
                    limitStringLength: limitFieldLength));

                TxtSqlOutput.Text = sql;
                TxtSqlOutput.ScrollToHome();

                int lineCount = sql.Split(Environment.NewLine).Length;
                double sizeKb = System.Text.Encoding.UTF8.GetByteCount(sql) / 1024d;
                TxtSqlStats.Text = $"数据库: {dbType}  |  表名: {tableName}  |  数据: {dataSnapshot.Rows.Count:N0} 行  |  SQL: {lineCount:N0} 行  |  {sizeKb:F1} KB";

                UpdateStatus("SQL 生成完成");
                ShowProgress(100);
            }
            catch (Exception ex)
            {
                TxtSqlOutput.Text = $"生成失败: {ex.Message}";
                MessageBox.Show($"SQL 生成失败：\n{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                UpdateStatus("SQL 生成失败");
            }
            finally
            {
                HideProgress();
                SetBusyState(false, "就绪");
            }
        }

        private void BtnCopySql_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TxtSqlOutput.Text))
            {
                MessageBox.Show("当前没有可复制的 SQL。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            bool copied = TryCopyToClipboard(TxtSqlOutput.Text);
            if (!copied)
            {
                MessageBox.Show("剪贴板当前被占用，请稍后再试。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            BtnCopySql.Content = "已复制";
            UpdateStatus("SQL 已复制到剪贴板");

            var timer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(1.5)
            };
            timer.Tick += (_, _) =>
            {
                BtnCopySql.Content = "复制 SQL";
                timer.Stop();
            };
            timer.Start();
        }

        private void BtnSaveSql_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TxtSqlOutput.Text))
            {
                MessageBox.Show("当前没有可保存的 SQL。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var dialog = new SaveFileDialog
            {
                Filter = "SQL 文件|*.sql|文本文件|*.txt",
                FileName = $"{TxtTableName.Text}_{DateTime.Now:yyyyMMdd_HHmmss}.sql",
                Title = "保存 SQL 文件"
            };

            ApplyDefaultExportPath(dialog);

            if (dialog.ShowDialog() == true)
            {
                ExportService.ExportSql(dialog.FileName, TxtSqlOutput.Text);
                UpdateStatus($"SQL 已保存: {dialog.FileName}");
            }
        }

        private void BtnExportCsv_Click(object sender, RoutedEventArgs e)
        {
            if (_currentData == null)
            {
                MessageBox.Show("当前没有可导出的表格数据。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            CommitPendingGridEdits();

            var dialog = new SaveFileDialog
            {
                Filter = "CSV 文件|*.csv",
                FileName = $"{TxtTableName.Text}_{DateTime.Now:yyyyMMdd_HHmmss}.csv",
                Title = "导出 CSV"
            };

            ApplyDefaultExportPath(dialog);

            if (dialog.ShowDialog() == true)
            {
                ExportService.ExportCsv(dialog.FileName, _currentData);
                UpdateStatus($"CSV 已导出: {dialog.FileName}");
            }
        }

        private void BtnExportJson_Click(object sender, RoutedEventArgs e)
        {
            if (_currentData == null)
            {
                MessageBox.Show("当前没有可导出的表格数据。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            CommitPendingGridEdits();

            var dialog = new SaveFileDialog
            {
                Filter = "JSON 文件|*.json",
                FileName = $"{TxtTableName.Text}_{DateTime.Now:yyyyMMdd_HHmmss}.json",
                Title = "导出 JSON"
            };

            if (dialog.ShowDialog() == true)
            {
                ExportService.ExportJson(dialog.FileName, _currentData);
                UpdateStatus($"JSON 已导出: {dialog.FileName}");
            }
        }

        private void Page_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[]? files = e.Data.GetData(DataFormats.FileDrop) as string[];
                if (files != null && files.Length > 0 && IsSupportedFile(files[0]))
                {
                    e.Effects = DragDropEffects.Copy;
                    return;
                }
            }

            e.Effects = DragDropEffects.None;
        }

        private async void Page_Drop(object sender, DragEventArgs e)
        {
            if (_isBusy || !e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                return;
            }

            string[]? files = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (files == null || files.Length == 0 || !IsSupportedFile(files[0]))
            {
                return;
            }

            await LoadSelectedFileAsync(files[0]);
        }

        private async void Page_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (Keyboard.Modifiers == ModifierKeys.Control && e.Key == Key.O)
            {
                e.Handled = true;
                BtnSelectFile_Click(sender, new RoutedEventArgs());
                return;
            }

            if (Keyboard.Modifiers == ModifierKeys.Control && e.Key == Key.S)
            {
                e.Handled = true;
                BtnSaveSql_Click(sender, new RoutedEventArgs());
                return;
            }

            if (Keyboard.Modifiers == ModifierKeys.Control && e.Key == Key.C && TxtSqlOutput.IsFocused)
            {
                e.Handled = true;
                BtnCopySql_Click(sender, new RoutedEventArgs());
                return;
            }

            if (e.Key == Key.F5 && !_isBusy && _currentData != null)
            {
                e.Handled = true;
                BtnGenerateSql_Click(sender, new RoutedEventArgs());
                return;
            }

            if (Keyboard.Modifiers == ModifierKeys.Control && e.Key == Key.R && !string.IsNullOrWhiteSpace(_currentFilePath))
            {
                e.Handled = true;
                await LoadDataAsync(_currentFilePath, _currentSheetName);
            }
        }

        private void DgPreview_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = (e.Row.GetIndex() + 1).ToString();
        }

        private void ResetCurrentData()
        {
            _currentData = null;
            DgPreview.ItemsSource = null;
            TxtPreviewInfo.Text = "请先选择数据文件";
            TxtPreviewInfo.Foreground = MutedBrush;
            TxtSqlOutput.Clear();
            TxtSqlStats.Text = "";
            UpdateExportButtons();
        }

        private void UpdateFileInfo(string filePath, string? sheetName, long fileSize, int totalRows, int totalCols)
        {
            string fileName = Path.GetFileName(filePath);
            string sizeLabel = fileSize >= 1024 * 1024
                ? $"{fileSize / 1024d / 1024d:F2} MB"
                : $"{fileSize / 1024d:F1} KB";
            string sheetLabel = string.IsNullOrWhiteSpace(sheetName) ? "" : $"  |  Sheet: {sheetName}";

            TxtFileInfo.Text = $"文件: {fileName}  |  大小: {sizeLabel}  |  行数: {totalRows:N0}  |  列数: {totalCols:N0}{sheetLabel}";
        }

        private void ShowLoadHints(int totalRows)
        {
            if (totalRows > 100000)
            {
                UpdateStatus($"当前数据量 {totalRows:N0} 行，加载和生成 SQL 可能较慢。");
            }
            else if (totalRows > 50000)
            {
                UpdateStatus($"数据量 {totalRows:N0} 行，建议优先使用批量 INSERT。");
            }
            else
            {
                UpdateStatus("正在加载数据...");
            }
        }

        private void CommitPendingGridEdits()
        {
            DgPreview.CommitEdit(DataGridEditingUnit.Cell, true);
            DgPreview.CommitEdit(DataGridEditingUnit.Row, true);
        }

        private SqlGeneratorService.DbType GetSelectedDbType()
        {
            if (CbDatabaseType is null)
            {
                return SqlGeneratorService.DbType.PostgreSQL;
            }

            string selected = (CbDatabaseType.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "PostgreSQL";
            return selected switch
            {
                "SQL Server" => SqlGeneratorService.DbType.SqlServer,
                "MySQL" => SqlGeneratorService.DbType.MySQL,
                "Oracle" => SqlGeneratorService.DbType.Oracle,
                _ => SqlGeneratorService.DbType.PostgreSQL
            };
        }

        private void LoadImportSettings(bool forceTableName)
        {
            _importSettings = ImportSettingsService.Load();
            ApplyImportSettingsToUi(forceTableName);
        }

        private void ApplyImportSettingsToUi(bool forceTableName)
        {
            SelectDbType(_importSettings.DefaultDbType);
            TxtBatchSize.Text = _importSettings.BatchSize.ToString();
            ChkDropIfExists.IsChecked = _importSettings.DropIfExists;
            ChkBatchInsert.IsChecked = _importSettings.BatchInsert;
            ChkLimitFieldLength.IsChecked = _importSettings.LimitFieldLength;
            ApplyDatabaseHint(forceTableName);
        }

        private void SelectDbType(string dbType)
        {
            foreach (ComboBoxItem item in CbDatabaseType.Items)
            {
                string content = item.Content?.ToString() ?? string.Empty;
                if (string.Equals(content, dbType, StringComparison.OrdinalIgnoreCase))
                {
                    CbDatabaseType.SelectedItem = item;
                    return;
                }
            }

            CbDatabaseType.SelectedIndex = 0;
        }

        private void ApplyDatabaseHint(bool forceTableName)
        {
            if (TxtTableHint is null || TxtTableName is null || CbDatabaseType is null)
            {
                return;
            }

            SqlGeneratorService.DbType dbType = GetSelectedDbType();
            TxtTableHint.Text = dbType switch
            {
                SqlGeneratorService.DbType.SqlServer => "提示：SQL Server 临时表建议使用 # 前缀，批量 INSERT 自动限制为 1000 行。",
                SqlGeneratorService.DbType.MySQL => "提示：MySQL 使用 CREATE TEMPORARY TABLE，布尔值会输出为 1 / 0。",
                SqlGeneratorService.DbType.Oracle => "提示：Oracle 使用 CREATE GLOBAL TEMPORARY TABLE，日期会输出为 TO_DATE(...)。",
                _ => "提示：PostgreSQL 使用 CREATE TEMPORARY TABLE，布尔值会输出为 TRUE / FALSE。"
            };

            string suggestedName = string.IsNullOrWhiteSpace(_importSettings.DefaultTableName)
                ? SqlGeneratorService.GetDefaultTableName(dbType)
                : _importSettings.DefaultTableName.Trim();
            string current = TxtTableName.Text.Trim();
            var defaultNames = Enum.GetValues<SqlGeneratorService.DbType>()
                .SelectMany(type => new[]
                {
                    SqlGeneratorService.GetDefaultTableName(type),
                    _importSettings.DefaultTableName
                })
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            if (forceTableName || string.IsNullOrWhiteSpace(current) || defaultNames.Contains(current) || string.Equals(current, _lastSuggestedTableName, StringComparison.OrdinalIgnoreCase))
            {
                TxtTableName.Text = suggestedName;
            }

            _lastSuggestedTableName = suggestedName;
        }

        private void ApplyDefaultExportPath(FileDialog dialog)
        {
            if (!string.IsNullOrWhiteSpace(_importSettings.DefaultExportPath) && Directory.Exists(_importSettings.DefaultExportPath))
            {
                dialog.InitialDirectory = _importSettings.DefaultExportPath;
            }
        }

        private void UpdateExportButtons()
        {
            if (BtnGenerateSql is null || BtnExportCsv is null || BtnExportJson is null || BtnSaveSql is null || BtnCopySql is null || TxtSqlOutput is null)
            {
                return;
            }

            bool hasData = _currentData != null && _currentData.Columns.Count > 0;
            BtnGenerateSql.IsEnabled = !_isBusy && hasData;
            BtnExportCsv.IsEnabled = !_isBusy && hasData;
            BtnExportJson.IsEnabled = !_isBusy && hasData;
            BtnSaveSql.IsEnabled = !_isBusy && !string.IsNullOrWhiteSpace(TxtSqlOutput.Text);
            BtnCopySql.IsEnabled = !_isBusy && !string.IsNullOrWhiteSpace(TxtSqlOutput.Text);
        }

        private void SetBusyState(bool isBusy, string status)
        {
            _isBusy = isBusy;
            if (BtnSelectFile != null) BtnSelectFile.IsEnabled = !isBusy;
            if (CbSheetList != null) CbSheetList.IsEnabled = !isBusy;
            if (CbDbfEncoding != null) CbDbfEncoding.IsEnabled = !isBusy;
            if (CbDatabaseType != null) CbDatabaseType.IsEnabled = !isBusy;
            if (DgPreview != null) DgPreview.IsEnabled = !isBusy;
            UpdateExportButtons();
            UpdateStatus(status);
            if (BtnGenerateSql != null) BtnGenerateSql.Content = isBusy ? "处理中..." : "生成 SQL";
        }

        private void UpdateStatus(string message)
        {
            if (TxtStatus != null)
            {
                TxtStatus.Text = message;
            }
        }

        private void ShowProgress(int value)
        {
            if (ProgressLoad != null)
            {
                ProgressLoad.Visibility = Visibility.Visible;
                ProgressLoad.Value = Math.Max(0, Math.Min(100, value));
            }
        }

        private void HideProgress()
        {
            if (ProgressLoad != null)
            {
                ProgressLoad.Visibility = Visibility.Collapsed;
                ProgressLoad.Value = 0;
            }
        }

        private void HandleLoadError(Exception ex, string title)
        {
            _currentData = null;
            DgPreview.ItemsSource = null;
            TxtPreviewInfo.Text = $"解析失败: {ex.Message}";
            TxtPreviewInfo.Foreground = ErrorBrush;
            UpdateStatus("加载失败");
            UpdateExportButtons();
            MessageBox.Show($"{title}：\n{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        private static bool IsSupportedFile(string filePath)
        {
            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            return ext is ".xlsx" or ".xls" or ".csv" or ".dbf";
        }

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

        private static bool TryCopyToClipboard(string text)
        {
            try
            {
                if (OpenClipboard(IntPtr.Zero))
                {
                    try
                    {
                        EmptyClipboard();

                        int bytes = (text.Length + 1) * 2;
                        IntPtr hGlobal = GlobalAlloc(GMEM_MOVEABLE, (UIntPtr)bytes);
                        if (hGlobal == IntPtr.Zero)
                        {
                            return false;
                        }

                        IntPtr target = GlobalLock(hGlobal);
                        if (target == IntPtr.Zero)
                        {
                            return false;
                        }

                        Marshal.Copy(text.ToCharArray(), 0, target, text.Length);
                        Marshal.WriteInt16(target, text.Length * 2, 0);
                        GlobalUnlock(hGlobal);
                        _ = SetClipboardData(CF_UNICODETEXT, hGlobal);
                        return true;
                    }
                    finally
                    {
                        CloseClipboard();
                    }
                }
            }
            catch
            {
            }

            for (int i = 0; i < 5; i++)
            {
                try
                {
                    Clipboard.SetDataObject(text, true);
                    return true;
                }
                catch
                {
                    Thread.Sleep(120);
                }
            }

            return false;
        }
    }
}
