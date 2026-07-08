using System.Windows;
using System.Windows.Input;

namespace WpfApp1.Views;

public partial class AppInputDialog : Window
{
    public AppInputDialog(string title, string prompt, string defaultValue = "")
    {
        InitializeComponent();

        Title = string.IsNullOrWhiteSpace(title) ? "输入" : title;
        CaptionText.Text = Title;
        PromptText.Text = string.IsNullOrWhiteSpace(prompt) ? "请输入内容：" : prompt;
        InputText.Text = defaultValue;

        Loaded += (_, _) =>
        {
            InputText.Focus();
            InputText.SelectAll();
        };
    }

    public string? InputValue { get; private set; }

    private void OkButton_Click(object sender, RoutedEventArgs e)
    {
        InputValue = InputText.Text;
        DialogResult = true;
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
    }

    private void Window_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.Escape)
        {
            return;
        }

        e.Handled = true;
        DialogResult = false;
    }
}
