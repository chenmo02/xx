using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace WpfApp1.Views;

public partial class AppMessageDialog : Window
{
    private readonly MessageBoxButton _buttons;
    private MessageBoxResult _primaryResult = MessageBoxResult.OK;
    private MessageBoxResult _secondaryResult = MessageBoxResult.Cancel;
    private MessageBoxResult _tertiaryResult = MessageBoxResult.Cancel;

    public AppMessageDialog(
        string message,
        string caption,
        MessageBoxButton buttons,
        MessageBoxImage icon)
    {
        InitializeComponent();

        _buttons = buttons;
        Result = MessageBoxResult.None;

        Title = string.IsNullOrWhiteSpace(caption) ? "提示" : caption;
        CaptionText.Text = Title;
        MessageText.Text = string.IsNullOrWhiteSpace(message) ? "暂无提示内容。" : message.Trim();

        ConfigureVisual(icon);
        ConfigureButtons(buttons);

        Loaded += (_, _) => PrimaryButton.Focus();
    }

    public MessageBoxResult Result { get; private set; }

    private void ConfigureVisual(MessageBoxImage icon)
    {
        var palette = icon switch
        {
            MessageBoxImage.Error => new DialogPalette("#D9485F", "#FCEDEE", "#FFF6F7", "!", "系统检测到需要立即关注的问题。"),
            MessageBoxImage.Warning => new DialogPalette("#D97706", "#FFF5E8", "#FFF9F0", "!", "请留意本次操作带来的影响。"),
            MessageBoxImage.Question => new DialogPalette("#0F766E", "#E9FBF7", "#F3FCFA", "?", "请先确认再继续执行。"),
            MessageBoxImage.Information => new DialogPalette("#2563EB", "#EEF5FF", "#F8FBFF", "i", "操作结果和后续动作会在这里说明。"),
            _ => new DialogPalette("#475569", "#F1F5F9", "#F8FAFC", "i", "请查看本次提示的详细内容。")
        };

        var accentBrush = CreateBrush(palette.AccentColor);
        BadgeBorder.Background = accentBrush;
        BadgeText.Text = palette.BadgeText;
        SubCaptionText.Text = palette.SubTitle;
        HeaderGradientStart.Color = (Color)ColorConverter.ConvertFromString(palette.HeaderStartColor);
        HeaderGradientEnd.Color = (Color)ColorConverter.ConvertFromString(palette.HeaderEndColor);
        PrimaryButton.Background = accentBrush;
        PrimaryButton.BorderBrush = accentBrush;
    }

    private void ConfigureButtons(MessageBoxButton buttons)
    {
        HideButton(TertiaryButton);
        HideButton(SecondaryButton);

        switch (buttons)
        {
            case MessageBoxButton.OK:
                _primaryResult = MessageBoxResult.OK;
                ConfigureButton(PrimaryButton, "确定", true, false);
                HintText.Text = "按 Enter 确认并关闭当前弹窗。";
                break;

            case MessageBoxButton.OKCancel:
                _primaryResult = MessageBoxResult.OK;
                _secondaryResult = MessageBoxResult.Cancel;
                ConfigureButton(SecondaryButton, "取消", false, true);
                ConfigureButton(PrimaryButton, "确定", true, false);
                HintText.Text = "按 Enter 确认，按 Esc 取消。";
                break;

            case MessageBoxButton.YesNo:
                _primaryResult = MessageBoxResult.Yes;
                _secondaryResult = MessageBoxResult.No;
                ConfigureButton(SecondaryButton, "否", false, false);
                ConfigureButton(PrimaryButton, "是", true, false);
                HintText.Text = "按 Enter 选择“是”，按 Esc 选择“否”。";
                break;

            case MessageBoxButton.YesNoCancel:
                _primaryResult = MessageBoxResult.Yes;
                _secondaryResult = MessageBoxResult.No;
                _tertiaryResult = MessageBoxResult.Cancel;
                ConfigureButton(TertiaryButton, "取消", false, true);
                ConfigureButton(SecondaryButton, "否", false, false);
                ConfigureButton(PrimaryButton, "是", true, false);
                HintText.Text = "按 Enter 选择“是”，按 Esc 取消。";
                break;

            default:
                _primaryResult = MessageBoxResult.OK;
                ConfigureButton(PrimaryButton, "确定", true, false);
                HintText.Text = "按 Enter 确认并关闭当前弹窗。";
                break;
        }
    }

    private static void ConfigureButton(Button button, string text, bool isDefault, bool isCancel)
    {
        button.Content = text;
        button.Visibility = Visibility.Visible;
        button.IsDefault = isDefault;
        button.IsCancel = isCancel;
    }

    private static void HideButton(Button button)
    {
        button.Visibility = Visibility.Collapsed;
        button.IsDefault = false;
        button.IsCancel = false;
    }

    private void PrimaryButton_Click(object sender, RoutedEventArgs e)
    {
        CloseWithResult(_primaryResult);
    }

    private void SecondaryButton_Click(object sender, RoutedEventArgs e)
    {
        CloseWithResult(_secondaryResult);
    }

    private void TertiaryButton_Click(object sender, RoutedEventArgs e)
    {
        CloseWithResult(_tertiaryResult);
    }

    private void Window_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.Escape)
        {
            return;
        }

        e.Handled = true;
        CloseWithResult(GetDefaultDismissResult());
    }

    protected override void OnClosed(EventArgs e)
    {
        if (Result == MessageBoxResult.None)
        {
            Result = GetDefaultDismissResult();
        }

        base.OnClosed(e);
    }

    private void CloseWithResult(MessageBoxResult result)
    {
        Result = result;
        Close();
    }

    private MessageBoxResult GetDefaultDismissResult()
    {
        return _buttons switch
        {
            MessageBoxButton.OK => MessageBoxResult.OK,
            MessageBoxButton.OKCancel => MessageBoxResult.Cancel,
            MessageBoxButton.YesNo => MessageBoxResult.No,
            MessageBoxButton.YesNoCancel => MessageBoxResult.Cancel,
            _ => MessageBoxResult.Cancel
        };
    }

    private static SolidColorBrush CreateBrush(string color)
    {
        return new SolidColorBrush((Color)ColorConverter.ConvertFromString(color));
    }

    private sealed record DialogPalette(
        string AccentColor,
        string HeaderStartColor,
        string HeaderEndColor,
        string BadgeText,
        string SubTitle);
}
