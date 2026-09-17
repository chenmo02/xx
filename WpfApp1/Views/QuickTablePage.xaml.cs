using System.IO;
using System.Windows;
using System.Windows.Controls;
using WpfApp1.Services;

namespace WpfApp1.Views;

public partial class QuickTablePage : Page
{
    private List<QuickTableColumn> _columns = [];

    public QuickTablePage()
    {
        InitializeComponent();
        CmbAllType.SelectedIndex = 0;
        var settings = ImportSettingsService.Load();
        CmbDbType.SelectedIndex = settings.DefaultDbType switch
        {
            "SQL Server" => 1,
            "MySQL" => 2,
            "Oracle" => 3,
            _ => 0
        };
        TxtTableName.Text = string.Empty;
    }

    private void BtnParse_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var parsed = InsertStatementParser.Parse(TxtInsert.Text);
            string tableName = InsertStatementParser.ExtractTableName(TxtInsert.Text);
            TxtTableName.Text = StripIdentifier(tableName);
            _columns = QuickTableService.InferColumns(parsed.Headers, parsed.Rows);
            ApplyCaseToColumns();
            GridColumns.ItemsSource = _columns;
            TxtStatus.Text = $"已解析 {_columns.Count} 个字段、{parsed.Rows.Count} 行" +
                             (string.IsNullOrWhiteSpace(parsed.Warning) ? string.Empty : $"；警告：{parsed.Warning}");
            TxtOutput.Clear();
            ToastService.Show(this, TxtStatus.Text);
        }
        catch (Exception ex)
        {
            TxtStatus.Text = "解析失败：" + ex.Message;
            MessageBox.Show(ex.Message, "解析 INSERT 失败", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private void BtnApplyType_Click(object sender, RoutedEventArgs e)
    {
        ApplySelectedType(showEmptyMessage: true);
    }

    private void BtnApplyLength_Click(object sender, RoutedEventArgs e)
    {
        if (_columns.Count == 0)
        {
            MessageBox.Show("请先解析 INSERT SQL。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        if (!int.TryParse(TxtAllLength.Text.Trim(), out int length) || length <= 0)
        {
            MessageBox.Show("字段长度必须是大于 0 的整数。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            TxtAllLength.Focus();
            return;
        }

        foreach (var column in _columns) column.Length = length;
        GridColumns.Items.Refresh();
        TxtStatus.Text = $"已将 {_columns.Count} 个字段长度设置为 {length}";
        ToastService.Show(this, TxtStatus.Text);
    }

    private void CmbAllType_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (IsInitialized) ApplySelectedType(showEmptyMessage: false);
    }

    private void ApplySelectedType(bool showEmptyMessage)
    {
        if (_columns.Count == 0)
        {
            if (showEmptyMessage) MessageBox.Show("请先解析 INSERT SQL。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        string? type = (CmbAllType.SelectedItem as ComboBoxItem)?.Content?.ToString()
                       ?? CmbAllType.SelectedItem?.ToString();
        type = QuickTableService.NormalizeType(type);
        foreach (var column in _columns) column.Type = type;
        GridColumns.Items.Refresh();
        TxtStatus.Text = $"已将 {_columns.Count} 个字段设置为 {type}";
        if (showEmptyMessage) ToastService.Show(this, TxtStatus.Text);
    }

    private void BtnGenerate_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            if (_columns.Count == 0) throw new InvalidOperationException("请先解析 INSERT SQL。");
            ApplyCaseToColumns();
            foreach (var column in _columns)
            {
                string normalizedType = QuickTableService.NormalizeType(column.Type);
                if (!string.Equals(column.Type, normalizedType, StringComparison.Ordinal))
                    column.Type = normalizedType;
            }
            TxtOutput.Text = QuickTableService.GenerateSql(GetDbType(), TxtTableName.Text,
                _columns, ChkDrop.IsChecked == true, ChkQuote.IsChecked == true);
            TxtStatus.Text = "CREATE TABLE SQL 已生成";
            ToastService.Show(this, "建表 SQL 已生成");
        }
        catch (Exception ex)
        {
            TxtStatus.Text = "生成失败：" + ex.Message;
            MessageBox.Show(ex.Message, "生成 SQL 失败", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private void BtnCopy_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(TxtOutput.Text)) { BtnGenerate_Click(sender, e); }
        if (string.IsNullOrWhiteSpace(TxtOutput.Text)) return;
        try
        {
            Clipboard.SetText(TxtOutput.Text);
            ToastService.Show(this, "SQL 已复制到剪贴板");
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            ToastService.Show(this, "剪贴板暂时不可用，请重试");
        }
    }

    private void BtnExport_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(TxtOutput.Text)) { BtnGenerate_Click(sender, e); }
        if (string.IsNullOrWhiteSpace(TxtOutput.Text)) return;
        var dialog = new Microsoft.Win32.SaveFileDialog { Filter = "SQL 文件 (*.sql)|*.sql|所有文件 (*.*)|*.*", FileName = "create_table.sql" };
        if (dialog.ShowDialog() == true)
        {
            File.WriteAllText(dialog.FileName, TxtOutput.Text);
            TxtStatus.Text = "已导出：" + dialog.FileName;
            ToastService.Show(this, "建表 SQL 已导出");
        }
    }

    private void ApplyCaseToColumns()
    {
        string mode = CmbCase.SelectedIndex switch { 1 => "大写", 2 => "小写", _ => "保持原样" };
        foreach (var column in _columns) column.Name = QuickTableService.ApplyCase(column.OriginalName, mode);
    }

    private SqlGeneratorService.DbType GetDbType() => CmbDbType.SelectedIndex switch
    {
        1 => SqlGeneratorService.DbType.SqlServer,
        2 => SqlGeneratorService.DbType.MySQL,
        3 => SqlGeneratorService.DbType.Oracle,
        _ => SqlGeneratorService.DbType.PostgreSQL
    };

    private static string StripIdentifier(string name) => string.Join('.', name.Split('.', StringSplitOptions.RemoveEmptyEntries)
        .Select(part => part.Trim().Trim('[', ']', '`', '"')));

    private void TxtTableName_TextChanged(object sender, TextChangedEventArgs e)
    {

    }

    private void TxtAllLength_TextChanged(object sender, TextChangedEventArgs e)
    {

    }
}
