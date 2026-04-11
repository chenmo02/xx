using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using ExcelDataReader;
using Microsoft.Win32;
using WpfApp1.Services;

namespace WpfApp1.Views
{
    public partial class DataValidationPage : Page
    {
        // ── 步骤状态 ────────────────────────────────────────────
        private int _step = 1;

        // ── 数据 ────────────────────────────────────────────────
        private List<DvTargetColumn> _targetColumns = [];
        private DvSourceData? _sourceData;
        private ObservableCollection<DvMappingRow> _mappings = [];
        private DvValidationResult? _lastResult;

        // ── 异步校验 ─────────────────────────────────────────────
        private CancellationTokenSource? _cts;

        // ── DataGrid 绑定数据 ─────────────────────────────────────
        /// <summary>源字段列表，供字段映射 ComboBox 使用</summary>
        public List<string> SourceHeaders { get; private set; } = [];

        /// <summary>映射方式选项列表</summary>
        public List<string> MappingTypeOptions { get; } = ["源字段映射", "固定值", "忽略"];

        public DataValidationPage()
        {
            InitializeComponent();
            DataContext = this;
            UpdateStepUI();
            UpdateStructQuery();

            // 拦截大文本粘贴，显示遮罩
            DataObject.AddPastingHandler(TxtDdl, OnDdlPasting);
            DataObject.AddPastingHandler(TxtInsert, OnInsertPasting);
        }

        private void DataGrid_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (sender is not DataGrid grid) return;
            if (Keyboard.Modifiers != ModifierKeys.Control || e.Key != Key.C) return;

            e.Handled = CopyCurrentCell(grid) || CopyCurrentRow(grid);
        }

        private void DataGrid_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is not DataGrid grid) return;
            if (e.OriginalSource is not DependencyObject source) return;

            var cell = FindVisualParent<DataGridCell>(source);
            if (cell != null)
            {
                var row = FindVisualParent<DataGridRow>(cell);
                if (row != null)
                {
                    grid.SelectedItem = row.Item;
                    grid.CurrentCell = new DataGridCellInfo(row.Item, cell.Column);
                    cell.Focus();
                }
                return;
            }

            var dataGridRow = FindVisualParent<DataGridRow>(source);
            if (dataGridRow != null)
            {
                grid.SelectedItem = dataGridRow.Item;
                dataGridRow.Focus();
            }
        }

        private void CopyCurrentCellMenu_Click(object sender, RoutedEventArgs e)
        {
            if (TryGetContextMenuGrid(sender, out var grid))
                CopyCurrentCell(grid);
        }

        private void CopyCurrentRowMenu_Click(object sender, RoutedEventArgs e)
        {
            if (TryGetContextMenuGrid(sender, out var grid))
                CopyCurrentRow(grid);
        }

        private bool TryGetContextMenuGrid(object sender, out DataGrid grid)
        {
            grid = null!;
            if (sender is not FrameworkElement element) return false;
            if (element.Parent is not ContextMenu menu) return false;
            if (menu.PlacementTarget is not DataGrid dataGrid) return false;
            grid = dataGrid;
            return true;
        }

        private bool CopyCurrentCell(DataGrid grid)
        {
            if (grid.CurrentCell.Column == null || grid.CurrentCell.Item == null)
            {
                SetStatus("请先选中要复制的单元格", true);
                return false;
            }

            string text = GetCellText(grid, grid.CurrentCell.Item, grid.CurrentCell.Column);
            Clipboard.SetText(text);
            SetStatus($"已复制单元格内容：{grid.CurrentCell.Column.Header}");
            return true;
        }

        private bool CopyCurrentRow(DataGrid grid)
        {
            var item = grid.SelectedItem ?? grid.CurrentCell.Item;
            if (item == null)
            {
                SetStatus("请先选中要复制的行", true);
                return false;
            }

            var orderedColumns = grid.Columns.OrderBy(c => c.DisplayIndex).ToList();
            string rowText = string.Join("\t", orderedColumns.Select(c => GetCellText(grid, item, c)));
            Clipboard.SetText(rowText);
            SetStatus("已复制当前行");
            return true;
        }

        private string GetCellText(DataGrid grid, object item, DataGridColumn column)
        {
            if (grid == DgMapping && item is DvMappingRow mapping)
                return GetMappingCellText(mapping, column.DisplayIndex);

            if (grid == DgIssues && item is DvIssue issue)
                return GetIssueCellText(issue, column.DisplayIndex);

            return string.Empty;
        }

        private static string GetMappingCellText(DvMappingRow mapping, int columnIndex)
        {
            return columnIndex switch
            {
                0 => mapping.RowIndex.ToString(),
                1 => mapping.TargetColumnName,
                2 => mapping.TargetDisplayType,
                3 => mapping.IsRequired ? "是" : "否",
                4 => mapping.MappingTypeStr,
                5 => mapping.MappingType switch
                {
                    DvMappingType.Source => mapping.SourceColumnName ?? string.Empty,
                    DvMappingType.Constant => mapping.ConstantValue ?? string.Empty,
                    _ => string.Empty
                },
                6 => mapping.MatchMethodText,
                7 => mapping.ConfidenceText,
                8 => mapping.IsConfirmed ? "是" : "否",
                _ => string.Empty
            };
        }

        private static string GetIssueCellText(DvIssue issue, int columnIndex)
        {
            return columnIndex switch
            {
                0 => issue.RowNumber.ToString(),
                1 => issue.PrimaryKeyDisplay ?? string.Empty,
                2 => issue.SourceColumnName ?? string.Empty,
                3 => issue.TargetColumnName,
                4 => issue.TargetDataType,
                5 => issue.LevelText,
                6 => issue.ErrorType,
                7 => issue.ActualValue ?? string.Empty,
                8 => issue.Message,
                _ => string.Empty
            };
        }

        private static T? FindVisualParent<T>(DependencyObject? child) where T : DependencyObject
        {
            while (child != null)
            {
                if (child is T target) return target;
                child = VisualTreeHelper.GetParent(child);
            }

            return null;
        }

        // ══════════════════════════════════════════════════════════
        //  步骤导航
        // ══════════════════════════════════════════════════════════

        private void BtnNext_Click(object sender, RoutedEventArgs e)
        {
            if (!CanGoNext(out string? err))
            {
                SetStatus(err ?? "请先完成当前步骤", true);
                return;
            }
            _step++;
            if (_step == 3) BuildMappings();
            UpdateStepUI();
            SetStatus("");
        }

        private void BtnBack_Click(object sender, RoutedEventArgs e)
        {
            if (_step > 1) _step--;
            UpdateStepUI();
            SetStatus("");
        }

        private bool CanGoNext(out string? err)
        {
            err = null;
            switch (_step)
            {
                case 1:
                    if (_targetColumns.Count == 0)
                    { err = "请先解析 DDL 或导入结构 Excel"; return false; }
                    if (string.IsNullOrWhiteSpace(TxtTableName.Text))
                    { err = "请填写目标表名"; return false; }
                    return true;

                case 2:
                    if (_sourceData == null || _sourceData.RowCount == 0)
                    { err = "请先解析 INSERT 语句或导入数据 Excel"; return false; }
                    return true;

                case 3:
                    // 检查必填字段是否全部映射并确认（自动生成字段如 uuid 除外）
                    var unconfirmed = _mappings.Where(m => !m.IsConfirmed).ToList();
                    var requiredIgnored = _mappings
                        .Where(m => m.IsRequired && m.MappingType == DvMappingType.Ignore && !m.IsAutoGenCandidate).ToList();
                    if (requiredIgnored.Any())
                    {
                        TxtMappingHint.Visibility = Visibility.Visible;
                        err = $"有 {requiredIgnored.Count} 个必填字段未映射";
                        return false;
                    }
                    if (unconfirmed.Any())
                    { err = $"还有 {unconfirmed.Count} 个字段未确认，请点击[全部确认]或逐行确认"; return false; }
                    TxtMappingHint.Visibility = Visibility.Collapsed;
                    return true;

                default:
                    return false; // Step 4 没有"下一步"
            }
        }

        private void UpdateStepUI()
        {
            // 内容面板
            Panel1.Visibility = _step == 1 ? Visibility.Visible : Visibility.Collapsed;
            Panel2.Visibility = _step == 2 ? Visibility.Visible : Visibility.Collapsed;
            Panel3.Visibility = _step == 3 ? Visibility.Visible : Visibility.Collapsed;
            Panel4.Visibility = _step == 4 ? Visibility.Visible : Visibility.Collapsed;

            // 导航按钮
            BtnBack.Visibility = _step > 1 ? Visibility.Visible : Visibility.Collapsed;
            BtnNext.Visibility = _step < 4 ? Visibility.Visible : Visibility.Collapsed;

            // 步骤指示器
            RefreshStepCircle(Step1Circle, Step1Label, "1", 1);
            RefreshStepCircle(Step2Circle, Step2Label, "2", 2);
            RefreshStepCircle(Step3Circle, Step3Label, "3", 3);
            RefreshStepCircle(Step4Circle, Step4Label, "4", 4);

            // 连接线颜色
            Line12.Background = _step > 1
                ? new SolidColorBrush(Color.FromRgb(16, 185, 129))
                : new SolidColorBrush(Color.FromRgb(209, 213, 219));
            Line23.Background = _step > 2
                ? new SolidColorBrush(Color.FromRgb(16, 185, 129))
                : new SolidColorBrush(Color.FromRgb(209, 213, 219));
            Line34.Background = _step > 3
                ? new SolidColorBrush(Color.FromRgb(16, 185, 129))
                : new SolidColorBrush(Color.FromRgb(209, 213, 219));
        }

        private void RefreshStepCircle(Border circle, TextBlock label, string num, int stepNum)
        {
            if (_step == stepNum)
            {
                circle.Background = new SolidColorBrush(Color.FromRgb(78, 110, 242));
                label.Foreground = new SolidColorBrush(Color.FromRgb(78, 110, 242));
                label.FontWeight = FontWeights.SemiBold;
            }
            else if (_step > stepNum)
            {
                circle.Background = new SolidColorBrush(Color.FromRgb(16, 185, 129));
                label.Foreground = new SolidColorBrush(Color.FromRgb(16, 185, 129));
                label.FontWeight = FontWeights.Normal;
                // 显示勾号
                if (circle.Child is TextBlock tb) tb.Text = "✓";
            }
            else
            {
                circle.Background = new SolidColorBrush(Color.FromRgb(209, 213, 219));
                label.Foreground = new SolidColorBrush(Color.FromRgb(156, 163, 175));
                label.FontWeight = FontWeights.Normal;
                if (circle.Child is TextBlock tb) tb.Text = num;
            }
        }

        // ══════════════════════════════════════════════════════════
        //  Step 1：结构输入
        // ══════════════════════════════════════════════════════════

        // ── 大文本粘贴拦截 ─────────────────────────────────────
        private const int PasteThreshold = 5000; // 超过此字符数显示遮罩

        private async void OnDdlPasting(object sender, DataObjectPastingEventArgs e)
        {
            if (e.DataObject.GetDataPresent(typeof(string)))
            {
                var text = (string)e.DataObject.GetData(typeof(string));
                if (text != null && text.Length > PasteThreshold)
                {
                    e.CancelCommand(); // 取消默认粘贴
                    DdlPasteOverlay.Visibility = Visibility.Visible;
                    await Task.Delay(30); // 让遮罩渲染
                    TxtDdl.Text = text;
                    TxtDdl.CaretIndex = text.Length;
                    DdlPasteOverlay.Visibility = Visibility.Collapsed;
                }
            }
        }

        private async void OnInsertPasting(object sender, DataObjectPastingEventArgs e)
        {
            if (e.DataObject.GetDataPresent(typeof(string)))
            {
                var text = (string)e.DataObject.GetData(typeof(string));
                if (text != null && text.Length > PasteThreshold)
                {
                    e.CancelCommand();
                    InsertPasteOverlay.Visibility = Visibility.Visible;
                    await Task.Delay(30);
                    TxtInsert.Text = text;
                    TxtInsert.CaretIndex = text.Length;
                    InsertPasteOverlay.Visibility = Visibility.Collapsed;
                }
            }
        }

        private void StructMode_Changed(object sender, RoutedEventArgs e)
        {
            if (DdlCard == null || StructExcelCard == null) return;
            bool isDdl = RbDdlMode.IsChecked == true;

            // 左侧 DDL 卡片
            DdlCard.Padding = new Thickness(isDdl ? 2 : 0);
            DdlCard.Background = isDdl
                ? new SolidColorBrush(Color.FromRgb(78, 110, 242))
                : Brushes.Transparent;

            // 右侧 SQL 查询卡片 — 与左侧对称
            StructExcelCard.Padding = new Thickness(isDdl ? 0 : 2);
            StructExcelCard.Background = isDdl
                ? Brushes.Transparent
                : new SolidColorBrush(Color.FromRgb(78, 110, 242));
            // 选中时用外层蓝色边框，隐藏内层灰色边框
            StructExcelInner.BorderThickness = new Thickness(isDdl ? 1 : 0);
        }

        private void DbType_Changed(object sender, RoutedEventArgs e)
        {
            UpdateStructQuery();
        }

        private void UpdateStructQuery()
        {
            if (TxtStructQuery == null) return;
            string tableName = string.IsNullOrWhiteSpace(TxtTableName?.Text)
                ? "你的表名"
                : TxtTableName!.Text.Trim();

            if (RbSqlServer?.IsChecked == true)
            {
                TxtStructQuery.Text =
                    "SELECT\r\n" +
                    "    c.ORDINAL_POSITION AS ordinal_position,\r\n" +
                    "    c.COLUMN_NAME AS column_name,\r\n" +
                    "    c.DATA_TYPE AS data_type,\r\n" +
                    "    c.CHARACTER_MAXIMUM_LENGTH AS character_maximum_length,\r\n" +
                    "    c.NUMERIC_PRECISION AS numeric_precision,\r\n" +
                    "    c.NUMERIC_SCALE AS numeric_scale,\r\n" +
                    "    c.IS_NULLABLE AS is_nullable\r\n" +
                    "FROM INFORMATION_SCHEMA.COLUMNS c\r\n" +
                    $"WHERE c.TABLE_NAME = '{tableName}'\r\n" +
                    "ORDER BY c.ORDINAL_POSITION;";
            }
            else
            {
                TxtStructQuery.Text =
                    "SELECT\r\n" +
                    "    ordinal_position,\r\n" +
                    "    column_name,\r\n" +
                    "    data_type,\r\n" +
                    "    character_maximum_length,\r\n" +
                    "    numeric_precision,\r\n" +
                    "    numeric_scale,\r\n" +
                    "    is_nullable\r\n" +
                    "FROM information_schema.columns\r\n" +
                    $"WHERE table_name = LOWER('{tableName}')\r\n" +
                    "ORDER BY ordinal_position;";
            }
        }

        private void BtnCopyStructQuery_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(TxtStructQuery.Text))
            {
                Clipboard.SetText(TxtStructQuery.Text);
                SetStatus("查询语句已复制到剪贴板");
            }
        }

        private void BtnClearDdl_Click(object sender, RoutedEventArgs e)
        {
            TxtDdl.Clear();
            TxtDdlStatus.Text = "";
            _targetColumns.Clear();
            SetStatus("");
        }

        private void BtnParseDdl_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dbType = RbSqlServer.IsChecked == true ? DvDbType.SqlServer : DvDbType.PostgreSql;

                // 自动检测 DDL 中的数据库类型特征
                var detected = DdlParser.DetectDbType(TxtDdl.Text);
                if (detected.HasValue && detected.Value != dbType)
                {
                    string selectedName = dbType == DvDbType.SqlServer ? "SQL Server" : "PostgreSQL";
                    string detectedName = detected.Value == DvDbType.SqlServer ? "SQL Server" : "PostgreSQL";
                    var result = MessageBox.Show(
                        $"您选择的数据库类型是 {selectedName}，但建表语句特征更像 {detectedName}。\n\n" +
                        $"数据库类型不正确会导致字段类型识别错误，影响校验结果。\n\n" +
                        $"是否自动切换为 {detectedName} 并继续解析？",
                        "数据库类型不匹配",
                        MessageBoxButton.YesNoCancel,
                        MessageBoxImage.Warning);

                    if (result == MessageBoxResult.Yes)
                    {
                        // 自动切换
                        if (detected.Value == DvDbType.SqlServer)
                            RbSqlServer.IsChecked = true;
                        else
                            RbPostgreSql.IsChecked = true;
                        dbType = detected.Value;
                    }
                    else if (result == MessageBoxResult.Cancel)
                    {
                        return; // 取消解析
                    }
                    // No = 继续用用户选的类型
                }

                _targetColumns = DdlParser.Parse(TxtDdl.Text, dbType);

                // 自动提取表名
                var extractedName = DdlParser.ExtractTableName(TxtDdl.Text);
                if (!string.IsNullOrEmpty(extractedName))
                {
                    TxtTableName.Text = extractedName;
                    UpdateStructQuery();
                }

                TxtDdlStatus.Text = $"⌛️ 已解析 {_targetColumns.Count} 个字段";
                TxtDdlStatus.Foreground = new SolidColorBrush(Color.FromRgb(16, 185, 129));
                SetStatus($"DDL 解析成功，共 {_targetColumns.Count} 个字段" +
                    (!string.IsNullOrEmpty(extractedName) ? $"，表名：{extractedName}" : ""));
            }
            catch (Exception ex)
            {
                TxtDdlStatus.Text = $"⌛️ 解析失败: {ex.Message}";
                TxtDdlStatus.Foreground = new SolidColorBrush(Color.FromRgb(220, 38, 38));
                _targetColumns.Clear();
            }
        }

        private void BtnImportStructExcel_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Filter = "Excel 文件|*.xlsx;*.xls",
                Title = "选择结构 Excel"
            };
            if (dlg.ShowDialog() != true) return;
            try
            {
                var dbType = RbSqlServer.IsChecked == true ? DvDbType.SqlServer : DvDbType.PostgreSql;
                _targetColumns = ReadStructExcel(dlg.FileName, dbType);
                TxtStructExcelInfo.Text = $"⌛️ {Path.GetFileName(dlg.FileName)} — 已读取 {_targetColumns.Count} 个字段";
                SetStatus($"结构 Excel 导入成功，共 {_targetColumns.Count} 个字段");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"读取失败:\n{ex.Message}", "导入错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private List<DvTargetColumn> ReadStructExcel(string path, DvDbType dbType)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            var result = new List<DvTargetColumn>();
            using var stream = File.OpenRead(path);
            using var reader = ExcelReaderFactory.CreateReader(stream);
            var ds = reader.AsDataSet(new ExcelDataSetConfiguration
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration { UseHeaderRow = true }
            });
            var dt = ds.Tables[0];

            // 别名映射
            int Col(params string[] aliases)
            {
                foreach (var a in aliases)
                    for (int i = 0; i < dt.Columns.Count; i++)
                        if (string.Equals(dt.Columns[i].ColumnName, a, StringComparison.OrdinalIgnoreCase))
                            return i;
                return -1;
            }

            int iName = Col("column_name", "字段名", "列名");
            int iType = Col("data_type", "字段类型", "类型");
            int iLen = Col("character_maximum_length", "长度", "字符长度");
            int iPrec = Col("numeric_precision", "精度");
            int iScale = Col("numeric_scale", "小数位");
            int iNull = Col("is_nullable", "是否可空");

            if (iName < 0 || iType < 0) throw new InvalidOperationException("未找到 column_name / data_type 列");

            for (int r = 0; r < dt.Rows.Count; r++)
            {
                var row = dt.Rows[r];
                var colName = row[iName]?.ToString()?.Trim() ?? "";
                var dataType = row[iType]?.ToString()?.Trim().ToLower() ?? "";
                if (string.IsNullOrEmpty(colName)) continue;

                int? len = iLen >= 0 && int.TryParse(row[iLen]?.ToString(), out int l) ? l : null;
                int? prec = iPrec >= 0 && int.TryParse(row[iPrec]?.ToString(), out int p) ? p : null;
                int? scale = iScale >= 0 && int.TryParse(row[iScale]?.ToString(), out int s) ? s : null;
                bool nullable = iNull < 0 || row[iNull]?.ToString()?.Trim().Equals("YES", StringComparison.OrdinalIgnoreCase) == true;

                result.Add(new DvTargetColumn
                {
                    OrdinalPosition = r + 1,
                    ColumnName = colName,
                    OriginalDataType = dataType,
                    NormalizedType = SchemaNormalizer.Normalize(dataType, dbType),
                    MaxLength = len,
                    NumericPrecision = prec,
                    NumericScale = scale,
                    IsNullable = nullable,
                    DatabaseType = dbType
                });
            }
            return result;
        }

        // ══════════════════════════════════════════════════════════
        //  Step 2：数据输入
        // ══════════════════════════════════════════════════════════

        private void DataMode_Changed(object sender, RoutedEventArgs e)
        {
            if (InsertCard == null || DataExcelCard == null) return;
            bool isInsert = RbInsertMode.IsChecked == true;

            // 左侧 INSERT 卡片
            InsertCard.Padding = new Thickness(isInsert ? 2 : 0);
            InsertCard.Background = isInsert
                ? new SolidColorBrush(Color.FromRgb(78, 110, 242))
                : Brushes.Transparent;

            // 右侧 Excel 卡片 — 与左侧对称
            DataExcelCard.Padding = new Thickness(isInsert ? 0 : 2);
            DataExcelCard.Background = isInsert
                ? Brushes.Transparent
                : new SolidColorBrush(Color.FromRgb(78, 110, 242));
            DataExcelInner.BorderThickness = new Thickness(isInsert ? 1 : 0);
        }

        private void BtnClearInsert_Click(object sender, RoutedEventArgs e)
        {
            TxtInsert.Clear();
            TxtInsertStatus.Text = "";
            _sourceData = null;
            SourceHeaders = [];
            SetStatus("");
        }

        private async void BtnParseInsert_Click(object sender, RoutedEventArgs e)
        {
            var sql = TxtInsert.Text;
            if (string.IsNullOrWhiteSpace(sql))
            {
                TxtInsertStatus.Text = "⌛️ SQL 语句为空";
                TxtInsertStatus.Foreground = new SolidColorBrush(Color.FromRgb(220, 38, 38));
                return;
            }

            BtnNext.IsEnabled = false;
            TxtInsertStatus.Text = "⌛️ 解析中...";
            TxtInsertStatus.Foreground = new SolidColorBrush(Color.FromRgb(100, 116, 139));

            try
            {
                var (headers, rows, warning) = await Task.Run(() => InsertStatementParser.Parse(sql));
                _sourceData = new DvSourceData { Headers = headers, Rows = rows };
                SourceHeaders = [.. headers];
                TxtInsertStatus.Text = $"⌛️ 已解析 {rows.Count} 行 × {headers.Count} 列";
                TxtInsertStatus.Foreground = new SolidColorBrush(Color.FromRgb(16, 185, 129));
                if (warning != null)
                    SetStatus($"注意：{warning}");
                else
                    SetStatus($"INSERT 解析成功，共 {rows.Count} 行");
            }
            catch (Exception ex)
            {
                TxtInsertStatus.Text = $"? 解析失败: {ex.Message}";
                TxtInsertStatus.Foreground = new SolidColorBrush(Color.FromRgb(220, 38, 38));
                _sourceData = null;
            }
            finally
            {
                BtnNext.IsEnabled = true;
            }
        }

        private void BtnImportDataExcel_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Filter = "Excel 文件|*.xlsx;*.xls",
                Title = "选择数据 Excel"
            };
            if (dlg.ShowDialog() != true) return;
            try
            {
                _sourceData = ReadDataExcel(dlg.FileName);
                SourceHeaders = [.. _sourceData.Headers];
                TxtDataExcelInfo.Text = $"⌛️ {Path.GetFileName(dlg.FileName)} — {_sourceData.RowCount} 行 × {_sourceData.Headers.Count} 列";
                SetStatus($"数据 Excel 导入成功，共 {_sourceData.RowCount} 行");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"读取失败:\n{ex.Message}", "导入错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private DvSourceData ReadDataExcel(string path)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using var stream = File.OpenRead(path);
            using var reader = ExcelReaderFactory.CreateReader(stream);
            var ds = reader.AsDataSet(new ExcelDataSetConfiguration
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration { UseHeaderRow = true }
            });
            var dt = ds.Tables[0];
            var headers = dt.Columns.Cast<System.Data.DataColumn>().Select(c => c.ColumnName).ToList();
            var rows = new List<IReadOnlyList<string?>>();
            foreach (System.Data.DataRow row in dt.Rows)
            {
                var vals = new List<string?>();
                for (int i = 0; i < headers.Count; i++)
                {
                    var v = row[i];
                    vals.Add(v == null || v == DBNull.Value ? null : v.ToString());
                }
                rows.Add(vals);
            }
            return new DvSourceData { Headers = headers, Rows = rows };
        }

        // ══════════════════════════════════════════════════════════
        //  Step 3：字段映射
        // ══════════════════════════════════════════════════════════

        private void BuildMappings()
        {
            _mappings = new ObservableCollection<DvMappingRow>(
                FieldMatcherService.AutoMap(_targetColumns, _sourceData!.Headers));
            DgMapping.ItemsSource = _mappings;

            // 填充主键下拉选项
            var pkOptions = new List<string> { "(无)" };
            pkOptions.AddRange(SourceHeaders);
            CbPk1.ItemsSource = pkOptions;
            CbPk2.ItemsSource = pkOptions;
            CbPk1.SelectedIndex = 0;
            CbPk2.SelectedIndex = 0;

            UpdateMappingInfo();
        }

        private void BtnAutoMap_Click(object sender, RoutedEventArgs e)
        {
            BuildMappings();
            SetStatus("已重新自动映射");
        }

        private void BtnNextUnconfirmed_Click(object sender, RoutedEventArgs e)
        {
            if (_mappings.Count == 0) return;
            int start = DgMapping.SelectedIndex < 0 ? 0 : DgMapping.SelectedIndex + 1;
            for (int i = 0; i < _mappings.Count; i++)
            {
                int idx = (start + i) % _mappings.Count;
                if (!_mappings[idx].IsConfirmed)
                {
                    DgMapping.SelectedIndex = idx;
                    DgMapping.ScrollIntoView(_mappings[idx]);
                    return;
                }
            }
            SetStatus("所有字段均已确认");
        }

        private void BtnIgnoreAllUuid_Click(object sender, RoutedEventArgs e)
        {
            int count = 0;
            foreach (var m in _mappings.Where(m => m.IsAutoGenCandidate))
            {
                m.MappingType = DvMappingType.Ignore;
                m.IsConfirmed = true;
                count++;
            }
            UpdateMappingInfo();
            SetStatus(count > 0 ? $"已忽略 {count} 个 UUID 字段" : "没有 UUID 类型字段");
        }

        private void BtnConfirmAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var m in _mappings)
                m.IsConfirmed = true;
            TxtMappingHint.Visibility = Visibility.Collapsed;
            UpdateMappingInfo();
            SetStatus("全部映射已确认");
        }

        private void MappingType_Changed(object sender, SelectionChangedEventArgs e)
        {
            UpdateMappingInfo();
        }

        /// <summary>获取用户选择的主键列名列表（排除"(无)"）</summary>
        private List<string> GetSelectedPkColumns()
        {
            var pks = new List<string>();
            if (CbPk1.SelectedItem is string pk1 && pk1 != "(无)")
                pks.Add(pk1);
            if (CbPk2.SelectedItem is string pk2 && pk2 != "(无)" && !pks.Contains(pk2))
                pks.Add(pk2);
            return pks;
        }

        private void UpdateMappingInfo()
        {
            if (_mappings.Count == 0) return;
            int mapped = _mappings.Count(m => m.MappingType != DvMappingType.Ignore);
            int confirmed = _mappings.Count(m => m.IsConfirmed);
            int required = _mappings.Count(m => m.IsRequired && m.MappingType == DvMappingType.Ignore && !m.IsAutoGenCandidate);
            int autoGen = _mappings.Count(m => m.IsAutoGenCandidate && m.MappingType == DvMappingType.Ignore);
            RunMappingInfo.Text =
                $"  共 {_mappings.Count} 个字段，已映射 {mapped}，已确认 {confirmed}" +
                (autoGen > 0 ? $"，?? {autoGen} 个自动生成字段已忽略" : "") +
                (required > 0 ? $"，? {required} 个必填未映射" : "");
        }

        // ══════════════════════════════════════════════════════════
        //  Step 4：校验
        // ══════════════════════════════════════════════════════════

        private async void BtnRunValidation_Click(object sender, RoutedEventArgs e)
        {
            if (_sourceData == null || _targetColumns.Count == 0) return;

            // 重置上一次结果
            _lastResult = null;

            // UI 状态
            BtnRunValidation.IsEnabled = false;
            BtnExportReport.Visibility = Visibility.Collapsed;
            SummaryCard.Visibility = Visibility.Collapsed;
            ProgressPanel.Visibility = Visibility.Visible;
            PbProgress.Value = 0;
            TxtProgress.Text = "正在准备校验...";
            DgIssues.ItemsSource = null;

            _cts = new CancellationTokenSource();
            var sw = System.Diagnostics.Stopwatch.StartNew();
            var progress = new Progress<(int current, int total, int errors)>(p =>
            {
                if (p.total > 0)
                    PbProgress.Value = (double)p.current / p.total * 100;
                TxtProgress.Text = $"校验中 {p.current:N0}/{p.total:N0} 行 ({sw.Elapsed.TotalSeconds:F0}s)，{p.errors:N0} 个错误";
            });

            try
            {
                var pkColumns = GetSelectedPkColumns();
                bool skipIntFormat = ChkSkipIntFormat.IsChecked == true;
                bool skipUuidFormat = ChkSkipUuidFormat.IsChecked == true;
                bool skipDateTimeFormat = ChkSkipDateTimeFormat.IsChecked == true;
                _lastResult = await ValidationEngine.RunAsync(
                    _targetColumns, _sourceData, _mappings,
                    pkColumns.Count > 0 ? pkColumns : null,
                    skipIntFormat,
                    skipUuidFormat,
                    skipDateTimeFormat,
                    progress, _cts.Token);
            }
            catch (OperationCanceledException)
            {
                SetStatus("校验已取消");
            }
            finally
            {
                sw.Stop();
                BtnRunValidation.IsEnabled = true;
                ProgressPanel.Visibility = Visibility.Collapsed;
            }

            var lastResult = _lastResult;
            if (lastResult == null) return;

            // 显示摘要
            TxtTotalRows.Text = lastResult.TotalRows.ToString("N0");
            TxtErrorCount.Text = lastResult.ErrorCount.ToString("N0");
            TxtRawErrorCount.Text = lastResult.RawErrorCount.ToString("N0");
            TxtWarningCount.Text = lastResult.WarningCount.ToString("N0");
            TxtElapsed.Text = $"{lastResult.Elapsed.TotalSeconds:F1}s";
            SummaryCard.Visibility = Visibility.Visible;

            // 显示错误列表
            DgIssues.ItemsSource = lastResult.Issues;

            BtnExportReport.Visibility = Visibility.Visible;

            SetStatus(lastResult.ErrorCount == 0 && lastResult.WarningCount == 0
                ? "校验通过，无错误无警告"
                : $"校验完成：{lastResult.ErrorCount} 条异常记录，{lastResult.RawErrorCount} 项错误，{lastResult.WarningCount} 条警告记录");
        }

        private void BtnCancelValidation_Click(object sender, RoutedEventArgs e)
        {
            _cts?.Cancel();
            TxtProgress.Text = "正在取消...";
            BtnCancelValidation.IsEnabled = false;
        }

        private void BtnExportReport_Click(object sender, RoutedEventArgs e)
        {
            var lastResult = _lastResult;
            if (lastResult == null) return;
            var dlg = new SaveFileDialog
            {
                Filter = "Excel 文件|*.xlsx",
                FileName = $"验证报告_{TxtTableName.Text}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx",
                Title = "保存验证报告"
            };
            if (dlg.ShowDialog() != true) return;
            try
            {
                var dbType = RbSqlServer.IsChecked == true ? DvDbType.SqlServer : DvDbType.PostgreSql;
                ValidationReportService.Generate(
                    lastResult, _targetColumns, _mappings,
                    dbType, TxtTableName.Text, dlg.FileName);
                SetStatus($"报告已导出：{dlg.FileName}");
                // 询问是否打开
                if (MessageBox.Show("报告已保存，是否立即打开？", "导出成功",
                    MessageBoxButton.YesNo, MessageBoxImage.Information) == MessageBoxResult.Yes)
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(dlg.FileName) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导出失败:\n{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ── 工具 ────────────────────────────────────────────────
        private void SetStatus(string msg, bool isError = false)
        {
            TxtStatus.Text = msg;
            TxtStatus.Foreground = isError
                ? new SolidColorBrush(Color.FromRgb(220, 38, 38))
                : new SolidColorBrush(Color.FromRgb(100, 116, 139));
        }
    }
}

