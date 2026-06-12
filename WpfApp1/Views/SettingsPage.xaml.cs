using System;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using WpfApp1.Services;

namespace WpfApp1.Views
{
    public partial class SettingsPage : Page
    {
        public SettingsPage()
        {
            InitializeComponent();

            LoadSettings();
        }

        // ═══════════════════════════════════════
        // 设置持久化
        // ═══════════════════════════════════════

        private void LoadSettings()
        {
            ImportSettings settings = ImportSettingsService.Load();
            SelectDbType(settings.DefaultDbType);
            TxtDefaultTableName.Text = settings.DefaultTableName;
            TxtDefaultBatchSize.Text = settings.BatchSize.ToString(CultureInfo.InvariantCulture);
            ChkDefaultDropIfExists.IsChecked = settings.DropIfExists;
            ChkDefaultBatchInsert.IsChecked = settings.BatchInsert;
            ChkDefaultLimitFieldLength.IsChecked = settings.LimitFieldLength;
            TxtDefaultExportPath.Text = settings.DefaultExportPath;
        }

        private void BtnSaveSettings_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!int.TryParse(TxtDefaultBatchSize.Text.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int batchSize) || batchSize <= 0)
                {
                    MessageBox.Show("每批 INSERT 行数必须是大于 0 的整数。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    TxtDefaultBatchSize.Focus();
                    return;
                }

                string exportPath = TxtDefaultExportPath.Text.Trim();
                if (!string.IsNullOrWhiteSpace(exportPath))
                {
                    Directory.CreateDirectory(exportPath);
                }

                var settings = new ImportSettings
                {
                    DefaultDbType = GetSelectedDbType(),
                    DefaultTableName = TxtDefaultTableName.Text.Trim(),
                    BatchSize = batchSize,
                    DropIfExists = ChkDefaultDropIfExists.IsChecked == true,
                    BatchInsert = ChkDefaultBatchInsert.IsChecked == true,
                    LimitFieldLength = ChkDefaultLimitFieldLength.IsChecked == true,
                    DefaultExportPath = exportPath
                };

                settings = ImportSettingsService.Normalize(settings);
                ImportSettingsService.Save(settings);
                TxtDefaultTableName.Text = settings.DefaultTableName;

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

        private void SelectDbType(string dbType)
        {
            foreach (ComboBoxItem item in CmbDefaultDbType.Items)
            {
                string content = item.Content?.ToString() ?? string.Empty;
                if (string.Equals(content, dbType, StringComparison.OrdinalIgnoreCase))
                {
                    CmbDefaultDbType.SelectedItem = item;
                    return;
                }
            }

            CmbDefaultDbType.SelectedIndex = 0;
        }

        private string GetSelectedDbType() => (CmbDefaultDbType.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "PostgreSQL";

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
