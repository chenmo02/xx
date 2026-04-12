using System;
using System.Linq;
using System.Windows;
using System.Windows.Threading;
using WpfApp1.Views;

namespace WpfApp1.Services;

public static class AppMessageBox
{
    public static MessageBoxResult Show(string messageBoxText)
    {
        return Show(messageBoxText, "提示", MessageBoxButton.OK, MessageBoxImage.None);
    }

    public static MessageBoxResult Show(string messageBoxText, string caption)
    {
        return Show(messageBoxText, caption, MessageBoxButton.OK, MessageBoxImage.None);
    }

    public static MessageBoxResult Show(string messageBoxText, string caption, MessageBoxButton button)
    {
        return Show(messageBoxText, caption, button, MessageBoxImage.None);
    }

    public static MessageBoxResult Show(
        string messageBoxText,
        string caption,
        MessageBoxButton button,
        MessageBoxImage icon)
    {
        var app = Application.Current;
        if (app?.Dispatcher == null)
        {
            return System.Windows.MessageBox.Show(messageBoxText, caption, button, icon);
        }

        if (!app.Dispatcher.CheckAccess())
        {
            return app.Dispatcher.Invoke(
                () => ShowInternal(messageBoxText, caption, button, icon),
                DispatcherPriority.Send);
        }

        return ShowInternal(messageBoxText, caption, button, icon);
    }

    private static MessageBoxResult ShowInternal(
        string messageBoxText,
        string caption,
        MessageBoxButton button,
        MessageBoxImage icon)
    {
        var owner = GetDialogOwner();
        var dialog = new AppMessageDialog(messageBoxText, caption, button, icon)
        {
            Owner = owner
        };

        if (owner == null)
        {
            dialog.WindowStartupLocation = WindowStartupLocation.CenterScreen;
        }

        dialog.ShowDialog();
        return dialog.Result;
    }

    private static Window? GetDialogOwner()
    {
        var app = Application.Current;
        if (app == null)
        {
            return null;
        }

        return app.Windows
            .OfType<Window>()
            .FirstOrDefault(window => window.IsActive && window.IsVisible)
            ?? (app.MainWindow?.IsVisible == true ? app.MainWindow : null);
    }
}
