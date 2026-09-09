using System.ComponentModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;
using WpfApp1.Services;

namespace WpfApp1.Views;

public partial class QuickNotesWindow : Window
{
    public QuickNotesWindow()
    {
        InitializeComponent();
        NoteEditor.Text = QuickNoteService.Load();
        Loaded += (_, _) =>
        {
            NoteEditor.Focus();
            NoteEditor.CaretIndex = NoteEditor.Text.Length;
        };
    }

    private void SaveButton_Click(object sender, RoutedEventArgs e)
    {
        var directory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        var dialog = new SaveFileDialog
        {
            Title = "保存随手记",
            Filter = "TXT 文件 (*.txt)|*.txt|所有文件 (*.*)|*.*",
            DefaultExt = ".txt",
            AddExtension = true,
            InitialDirectory = directory,
            FileName = GetNextFileName(directory),
            OverwritePrompt = true
        };

        if (dialog.ShowDialog(this) != true) return;

        try
        {
            File.WriteAllText(dialog.FileName, NoteEditor.Text);
            SaveNotes();
            SaveStatus.Text = $"已保存：{Path.GetFileName(dialog.FileName)}";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"保存失败：\n{ex.Message}", "保存失败", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e) => Close();

    private void Window_Closing(object? sender, CancelEventArgs e) => SaveNotes();

    private void SaveNotes() => QuickNoteService.Save(NoteEditor.Text);

    private static string GetNextFileName(string directory)
    {
        for (var index = 0; ; index++)
        {
            var fileName = index == 0 ? "cc.txt" : $"cc{index}.txt";
            if (!File.Exists(Path.Combine(directory, fileName))) return fileName;
        }
    }

    private void TitleBar_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.OriginalSource is DependencyObject source && IsInsideButton(source)) return;
        if (e.LeftButton == MouseButtonState.Pressed) DragMove();
    }

    private static bool IsInsideButton(DependencyObject source)
    {
        for (var current = source; current != null; current = VisualTreeHelper.GetParent(current))
        {
            if (current is ButtonBase) return true;
        }

        return false;
    }

    private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if ((Keyboard.Modifiers & ModifierKeys.Control) == 0) return;

        switch (e.Key)
        {
            case Key.D:
                DuplicateCurrentLine();
                e.Handled = true;
                break;
            case Key.F:
                OpenFindBar(false);
                e.Handled = true;
                break;
            case Key.H:
                OpenFindBar(true);
                e.Handled = true;
                break;
        }
    }

    private void DuplicateCurrentLine()
    {
        var caret = NoteEditor.CaretIndex;
        var line = NoteEditor.GetLineIndexFromCharacterIndex(caret);
        var lineStart = NoteEditor.GetCharacterIndexFromLineIndex(line);
        var lineLength = NoteEditor.GetLineLength(line);
        var lineText = NoteEditor.Text.Substring(lineStart, lineLength);
        var insertion = Environment.NewLine + lineText;
        var insertAt = lineStart + lineLength;

        NoteEditor.Text = NoteEditor.Text.Insert(insertAt, insertion);
        NoteEditor.CaretIndex = Math.Min(caret + insertion.Length, NoteEditor.Text.Length);
        NoteEditor.Focus();
    }

    private void OpenFindBar(bool replaceMode)
    {
        FindBar.Visibility = Visibility.Visible;
        ReplaceRow.Visibility = replaceMode ? Visibility.Visible : Visibility.Collapsed;
        FindInput.Focus();
        FindInput.SelectAll();
    }

    private void CloseFindBar_Click(object sender, RoutedEventArgs e)
    {
        FindBar.Visibility = Visibility.Collapsed;
        NoteEditor.Focus();
    }

    private void FindNext_Click(object sender, RoutedEventArgs e) => FindMatch(true);

    private void FindPrevious_Click(object sender, RoutedEventArgs e) => FindMatch(false);

    private bool FindMatch(bool forward)
    {
        var term = FindInput.Text;
        if (string.IsNullOrEmpty(term))
        {
            SaveStatus.Text = "请输入查找内容";
            return false;
        }

        var text = NoteEditor.Text;
        var index = -1;
        if (forward)
        {
            var start = Math.Clamp(NoteEditor.SelectionStart + NoteEditor.SelectionLength, 0, text.Length);
            index = text.IndexOf(term, start, StringComparison.CurrentCultureIgnoreCase);
            if (index < 0) index = text.IndexOf(term, 0, StringComparison.CurrentCultureIgnoreCase);
        }
        else
        {
            var start = Math.Min(NoteEditor.SelectionStart - 1, text.Length - term.Length);
            if (start >= 0) index = text.LastIndexOf(term, start, StringComparison.CurrentCultureIgnoreCase);
            if (index < 0) index = text.LastIndexOf(term, StringComparison.CurrentCultureIgnoreCase);
        }

        if (index < 0)
        {
            SaveStatus.Text = "未找到匹配内容";
            return false;
        }

        NoteEditor.Select(index, term.Length);
        NoteEditor.Focus();
        SaveStatus.Text = "已找到匹配内容";
        return true;
    }

    private void ReplaceCurrent_Click(object sender, RoutedEventArgs e)
    {
        var term = FindInput.Text;
        if (string.IsNullOrEmpty(term))
        {
            SaveStatus.Text = "请输入查找内容";
            return;
        }

        if (NoteEditor.SelectedText.Equals(term, StringComparison.CurrentCultureIgnoreCase))
        {
            NoteEditor.SelectedText = ReplaceInput.Text;
            SaveStatus.Text = "已替换当前匹配内容";
        }
        else
        {
            FindMatch(true);
        }
    }

    private void ReplaceAll_Click(object sender, RoutedEventArgs e)
    {
        var term = FindInput.Text;
        if (string.IsNullOrEmpty(term))
        {
            SaveStatus.Text = "请输入查找内容";
            return;
        }

        var text = NoteEditor.Text;
        var count = 0;
        for (var index = 0; ; count++)
        {
            index = text.IndexOf(term, index, StringComparison.CurrentCultureIgnoreCase);
            if (index < 0) break;
            index += term.Length;
        }

        NoteEditor.Text = text.Replace(term, ReplaceInput.Text, StringComparison.CurrentCultureIgnoreCase);
        NoteEditor.CaretIndex = Math.Min(NoteEditor.CaretIndex, NoteEditor.Text.Length);
        NoteEditor.Focus();
        SaveStatus.Text = count == 0 ? "未找到匹配内容" : $"已替换 {count} 处内容";
    }
}
