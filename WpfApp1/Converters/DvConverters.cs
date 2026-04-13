using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;
using WpfApp1.Services;

namespace WpfApp1.Converters
{
    /// <summary>bool → Visibility（true=Visible）</summary>
    public sealed class BoolVisConverter : IValueConverter
    {
        public static readonly BoolVisConverter Instance = new();
        public object Convert(object v, Type t, object p, CultureInfo c)
            => v is true ? Visibility.Visible : Visibility.Collapsed;
        public object ConvertBack(object v, Type t, object p, CultureInfo c)
            => v is Visibility.Visible;
    }

    /// <summary>bool → Visibility（true=Collapsed）</summary>
    public sealed class InverseBoolVisConverter : IValueConverter
    {
        public static readonly InverseBoolVisConverter Instance = new();
        public object Convert(object v, Type t, object p, CultureInfo c)
            => v is true ? Visibility.Collapsed : Visibility.Visible;
        public object ConvertBack(object v, Type t, object p, CultureInfo c)
            => v is Visibility.Collapsed;
    }

    /// <summary>bool → "是"/"否"</summary>
    public sealed class BoolYesNoConverter : IValueConverter
    {
        public static readonly BoolYesNoConverter Instance = new();
        public object Convert(object v, Type t, object p, CultureInfo c)
            => v is true ? "是" : "";
        public object ConvertBack(object v, Type t, object p, CultureInfo c)
            => throw new NotImplementedException();
    }

    /// <summary>IsRequired → 前景色（必填红色）</summary>
    public sealed class RequiredColorConverter : IValueConverter
    {
        public static readonly RequiredColorConverter Instance = new();
        public object Convert(object v, Type t, object p, CultureInfo c)
            => v is true
                ? new SolidColorBrush(Color.FromRgb(220, 38, 38))
                : new SolidColorBrush(Color.FromRgb(107, 114, 128));
        public object ConvertBack(object v, Type t, object p, CultureInfo c)
            => throw new NotImplementedException();
    }

    /// <summary>DvValidationLevel → 前景色</summary>
    public sealed class IssueColorConverter : IValueConverter
    {
        public static readonly IssueColorConverter Instance = new();
        public object Convert(object v, Type t, object p, CultureInfo c)
            => v is DvValidationLevel lv
                ? lv switch
                {
                    DvValidationLevel.Error => new SolidColorBrush(Color.FromRgb(220, 38, 38)),
                    DvValidationLevel.Warning => new SolidColorBrush(Color.FromRgb(217, 119, 6)),
                    _ => new SolidColorBrush(Color.FromRgb(107, 114, 128))
                }
                : DependencyProperty.UnsetValue;
        public object ConvertBack(object v, Type t, object p, CultureInfo c)
            => throw new NotImplementedException();
    }
}
