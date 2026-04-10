using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;

namespace WpfApp1.Behaviors
{
    public static class ScrollViewerAssist
    {
        public static readonly DependencyProperty EnableHiddenWheelScrollProperty =
            DependencyProperty.RegisterAttached(
                "EnableHiddenWheelScroll",
                typeof(bool),
                typeof(ScrollViewerAssist),
                new PropertyMetadata(false, OnEnableHiddenWheelScrollChanged));

        public static bool GetEnableHiddenWheelScroll(DependencyObject obj)
        {
            return (bool)obj.GetValue(EnableHiddenWheelScrollProperty);
        }

        public static void SetEnableHiddenWheelScroll(DependencyObject obj, bool value)
        {
            obj.SetValue(EnableHiddenWheelScrollProperty, value);
        }

        private static void OnEnableHiddenWheelScrollChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is not ScrollViewer scrollViewer)
            {
                return;
            }

            if ((bool)e.NewValue)
            {
                scrollViewer.PreviewMouseWheel += ScrollViewer_PreviewMouseWheel;
            }
            else
            {
                scrollViewer.PreviewMouseWheel -= ScrollViewer_PreviewMouseWheel;
            }
        }

        private static void ScrollViewer_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (sender is not ScrollViewer scrollViewer || scrollViewer.ScrollableHeight <= 0)
            {
                return;
            }

            if (IsInsideNestedScrollableControl(e.OriginalSource as DependencyObject, scrollViewer))
            {
                return;
            }

            double nextOffset = scrollViewer.VerticalOffset - (e.Delta / 3.0);
            nextOffset = Math.Max(0, Math.Min(nextOffset, scrollViewer.ScrollableHeight));
            scrollViewer.ScrollToVerticalOffset(nextOffset);
            e.Handled = true;
        }

        private static bool IsInsideNestedScrollableControl(DependencyObject? source, ScrollViewer owner)
        {
            DependencyObject? current = source;
            while (current != null && current != owner)
            {
                if (current is ScrollViewer)
                {
                    return true;
                }

                if (current is DataGrid or TextBoxBase or PasswordBox or RichTextBox or ListBox or ListView)
                {
                    return true;
                }

                current = GetParent(current);
            }

            return false;
        }

        private static DependencyObject? GetParent(DependencyObject current)
        {
            return current switch
            {
                Visual => VisualTreeHelper.GetParent(current),
                FrameworkContentElement frameworkContentElement => frameworkContentElement.Parent,
                _ => LogicalTreeHelper.GetParent(current)
            };
        }
    }
}
