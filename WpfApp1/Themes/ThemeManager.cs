using System.Windows;

namespace WpfApp1.Themes;

/// <summary>
/// 全局主题切换管理器。
/// 通过替换 Application.Resources 中的主题颜色字典实现运行时切换。
/// </summary>
public static class ThemeManager
{
    private const string ThemeSuffix = "Theme.xaml";

    public static string CurrentTheme { get; private set; } = "Light";

    /// <summary>
    /// 切换到指定主题（"Light" 或 "Dark"）。
    /// </summary>
    public static void SwitchTheme(string themeName)
    {
        var app = Application.Current;
        if (app == null) return;

        var dictionaries = app.Resources.MergedDictionaries;

        // 移除旧的主题字典
        for (int i = dictionaries.Count - 1; i >= 0; i--)
        {
            var source = dictionaries[i].Source?.OriginalString;
            if (source != null && source.Contains(ThemeSuffix))
            {
                dictionaries.RemoveAt(i);
                break;
            }
        }

        // 加入新的主题字典
        var newTheme = new ResourceDictionary
        {
            Source = new Uri($"Themes/{themeName}{ThemeSuffix}", UriKind.Relative)
        };
        dictionaries.Insert(0, newTheme);

        CurrentTheme = themeName;
    }
}
