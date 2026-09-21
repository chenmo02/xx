using System.Data;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using WpfApp1.Services;
using WpfApp1.Views;

internal static class Program
{
    private const BindingFlags PrivateInstance = BindingFlags.Instance | BindingFlags.NonPublic;
    private static object? Call(DataImportPage page, string method, params object?[] args) =>
        typeof(DataImportPage).GetMethod(method, PrivateInstance)!.Invoke(page, args);
    private static T Field<T>(DataImportPage page, string name) =>
        (T)typeof(DataImportPage).GetField(name, PrivateInstance)!.GetValue(page)!;
    private static T Control<T>(DataImportPage page, string name) => (T)page.FindName(name);
    private static void Check(bool condition, string message)
    {
        if (!condition) throw new Exception(message);
    }

    [STAThread]
    private static int Main()
    {
        var app = new Application();
        foreach (string resource in new[] { "LightTheme", "ScrollBarStyles", "ToastStyles", "ButtonStyles", "DataGridStyles", "InputStyles", "CardStyles", "CommonStyles" })
            app.Resources.MergedDictionaries.Add(new ResourceDictionary
            {
                Source = new Uri($"pack://application:,,,/CCToolbox;component/Themes/{resource}.xaml")
            });
        var page = new DataImportPage();
        var window = new Window
        {
            Content = page, Width = 1400, Height = 950,
            Left = -30000, Top = -30000, ShowInTaskbar = false, WindowStyle = WindowStyle.None
        };
        int exitCode = 1;
        window.Show();
        app.Dispatcher.InvokeAsync(async () =>
        {
            string directory = Path.Combine(Path.GetTempPath(), "CCToolbox-import-test-" + Guid.NewGuid());
            Directory.CreateDirectory(directory);
            try
            {
                await Run(page, window, directory);
                exitCode = 0;
                Console.WriteLine("PASS");
            }
            catch (Exception ex) { Console.Error.WriteLine(ex); }
            finally
            {
                Directory.Delete(directory, recursive: true);
                window.Close();
                app.Shutdown();
            }
        });
        app.Run();
        return exitCode;
    }

    private static async Task Run(DataImportPage page, Window window, string directory)
    {
        string csv = Path.Combine(directory, "large.csv");
        await Task.Run(() =>
        {
            using var writer = new StreamWriter(csv, false, new UTF8Encoding(true));
            writer.WriteLine("id,name,code,amount,date,department,description,notes");
            for (int row = 0; row < 160000; row++)
                writer.WriteLine($"{row},测试项目{row},00123,1200.50,2026-09-20,部门{row % 20},\"带逗号,及单引号'的文本\",\"第一行\n第二行\"");
        });

        double longestGap = 0;
        int ticks = 0;
        bool sawLoadCompletion = false, sawLoadFade = false, sawLoadCollapse = false;
        var loadingPanel = Control<Grid>(page, "FileLoadingPanel");
        var loadingProgress = Control<ProgressBar>(page, "ProgressLoad");
        var clock = Stopwatch.StartNew();
        double lastTick = clock.Elapsed.TotalMilliseconds;
        var timer = new DispatcherTimer(DispatcherPriority.Normal) { Interval = TimeSpan.FromMilliseconds(25) };
        timer.Tick += (_, _) =>
        {
            double now = clock.Elapsed.TotalMilliseconds;
            longestGap = Math.Max(longestGap, now - lastTick);
            lastTick = now;
            ticks++;
            sawLoadCompletion |= loadingProgress.IsVisible && loadingProgress.Value == 100;
            sawLoadFade |= loadingPanel.IsVisible && loadingPanel.Opacity > 0 && loadingPanel.Opacity < 1;
            sawLoadCollapse |= loadingPanel.IsVisible && loadingPanel.Height > 0 && loadingPanel.Opacity == 0;
        };
        timer.Start();
        await (Task)Call(page, "LoadSelectedFileAsync", csv)!;
        await Dispatcher.Yield(DispatcherPriority.ApplicationIdle);
        DataTable data = Field<DataTable>(page, "_currentData");
        Check(data.Rows.Count == 160000, "CSV import lost rows");
        Check(sawLoadCompletion && sawLoadFade && sawLoadCollapse, "Loading must show completion, fade, and smoothly collapse");
        Check(!loadingPanel.IsVisible && loadingPanel.Opacity == 1 && double.IsNaN(loadingPanel.Height), "Loading animation state was not reset");
        Console.WriteLine($"Import 160,000 rows: {clock.Elapsed.TotalSeconds:F2}s");

        data.Rows[159999][1] = "LAST_ROW_EDITED";
        Control<TextBox>(page, "TxtTableName").Text = "large_import";
        var output = Control<TextBox>(page, "TxtSqlOutput");
        foreach (bool batch in new[] { true, false })
        {
            Control<CheckBox>(page, "ChkBatchInsert").IsChecked = batch;
            var elapsed = Stopwatch.StartNew();
            // Bypass only the existing >100,000-row confirmation dialog.
            Task generation = (Task)Call(page, "GenerateSqlAsync")!;
            Check(!Control<DataGrid>(page, "DgPreview").IsEnabled, "Editing enabled during generation");
            Check(!Control<Button>(page, "BtnCopySql").IsEnabled, "Copy enabled during generation");
            await (Task)Call(page, "LoadDataAsync", csv, null)!;
            Check(ReferenceEquals(data, Field<DataTable>(page, "_currentData")), "Reload replaced data during generation");
            await generation;
            output.BringIntoView();
            window.UpdateLayout();
            await Dispatcher.Yield(DispatcherPriority.ApplicationIdle);

            string full = Field<string>(page, "_generatedSql");
            Check(full.Contains("LAST_ROW_EDITED"), "Last edited row missing from full SQL");
            Check(full.Contains("带逗号,及单引号''的文本"), "Escaped field changed");
            Check(full.AsSpan().Count("LAST_ROW_EDITED") == 1, "Last row duplicated");
            Check(full.AsSpan().Count("'00123'") == 160000, "Generated SQL lost data rows");
            Check(output.Text.AsSpan().Count('\n') + 1 == 1000, "Large SQL must preview exactly 1,000 lines");
            Check(full.StartsWith(output.Text, StringComparison.Ordinal), "Preview differs from full SQL");
            Check(!output.Text.Contains("LAST_ROW_EDITED"), "Large document was put in preview");
            Check(Control<TextBlock>(page, "TxtSqlStats").Text.Contains("复制和保存包含完整 SQL"), "Missing preview notice");
            Check(Control<Button>(page, "BtnCopySql").IsEnabled && Control<Button>(page, "BtnSaveSql").IsEnabled, "Full SQL export unavailable");

            string export = Path.Combine(directory, "full.sql");
            await Task.Run(() => ExportService.ExportSql(export, full));
            Check(await File.ReadAllTextAsync(export) == full, "Saved SQL was truncated");
            for (int i = 0; i < 10; i++)
            {
                output.ScrollToVerticalOffset(i * 60);
                window.UpdateLayout();
                await Task.Delay(30);
            }
            Console.WriteLine($"Generate/render/save/scroll batch={batch}: {elapsed.Elapsed.TotalSeconds:F2}s, SQL={full.Length:N0} chars, preview={output.Text.Length:N0} chars");
        }
        timer.Stop();
        Console.WriteLine($"UI heartbeat: {ticks} ticks, longest gap={longestGap:F0}ms");
        Check(ticks > 10 && longestGap < 1500, "UI blocked for more than 1.5 seconds");

        string smallCsv = Path.Combine(directory, "small.csv");
        await File.WriteAllTextAsync(smallCsv, "id,name\n1,small\n2,完整显示\n", new UTF8Encoding(true));
        await (Task)Call(page, "LoadSelectedFileAsync", smallCsv)!;
        Check(Field<string>(page, "_generatedSql") == "", "Loading another file retained old SQL");
        await (Task)Call(page, "GenerateSqlAsync")!;
        Check(output.Text == Field<string>(page, "_generatedSql") && output.Text.Contains("完整显示"), "Small file not displayed in full");
        Control<ComboBox>(page, "CbDatabaseType").SelectedIndex = 1;
        Check(Field<string>(page, "_generatedSql") == "" && !Control<Button>(page, "BtnSaveSql").IsEnabled, "Database change retained old SQL");

        // Short documents are still displayed in full; clear/reset must disable exports.
        var preview = typeof(DataImportPage).GetMethod("BuildSqlPreview", BindingFlags.Static | BindingFlags.NonPublic)!;
        string shortSql = "SELECT '测试', 1;\r\n";
        Check((string)preview.Invoke(null, [shortSql])! == shortSql, "Small SQL was truncated");
        string longLine = "SELECT '" + new string('中', 40000) + "';";
        Check((string)preview.Invoke(null, [longLine])! == longLine, "Character count incorrectly triggered large-file preview");
        foreach (string newline in new[] { "\n", "\r\n" })
        {
            string exactLimit = string.Join(newline, Enumerable.Repeat("SELECT '完整显示';", 5000));
            Check((string)preview.Invoke(null, [exactLimit])! == exactLimit, "Exactly 5,000 lines must display in full");
            string overLimit = exactLimit + newline + "SELECT '第5001行';";
            string limited = (string)preview.Invoke(null, [overLimit])!;
            Check(limited.Length < overLimit.Length && overLimit.StartsWith(limited, StringComparison.Ordinal), "5,001 lines must use a bounded preview");
            Check(limited == string.Join(newline, Enumerable.Repeat("SELECT '完整显示';", 1000)), "Preview must contain the first 1,000 complete lines");
        }
        string surrogateBoundary = new string('a', 31999) + "😀" + new string('\n', 5000);
        string longLinePreview = (string)preview.Invoke(null, [surrogateBoundary])!;
        Check(longLinePreview.StartsWith(new string('a', 31999) + "😀", StringComparison.Ordinal), "Long first line was truncated");
        Check(longLinePreview.AsSpan().Count('\n') + 1 == 1000, "Long lines must still allow a 1,000-line preview");
        Call(page, "ResetCurrentData");
        Check(Field<string>(page, "_generatedSql") == "" && output.Text == "", "Reset retained old SQL");
        Check(!Control<Button>(page, "BtnCopySql").IsEnabled && !Control<Button>(page, "BtnSaveSql").IsEnabled, "Empty result can be exported");
    }
}
