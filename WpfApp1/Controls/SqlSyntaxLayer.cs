using System.Globalization;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace WpfApp1.Controls;

// Draw over the native read-only TextBox so selection, copying and scrolling stay native.
public sealed class SqlSyntaxLayer : FrameworkElement
{
    public static readonly DependencyProperty EditorProperty = DependencyProperty.Register(
        nameof(Editor), typeof(TextBox), typeof(SqlSyntaxLayer), new PropertyMetadata(null, EditorChanged));

    public TextBox? Editor
    {
        get => (TextBox?)GetValue(EditorProperty);
        set => SetValue(EditorProperty, value);
    }

    private static readonly Brush Plain = ColorBrush("#E2E8F0");
    private static readonly Brush Keyword = ColorBrush("#C4A0FF");
    private static readonly Brush DataType = ColorBrush("#61DBEF");
    private static readonly Brush String = ColorBrush("#9EE493");
    private static readonly Brush Number = ColorBrush("#FFCC80");
    private static readonly Brush Comment = ColorBrush("#8DA2BD");
    private static readonly Brush Identifier = ColorBrush("#82CCFF");
    private static readonly Brush Punctuation = ColorBrush("#F6C177");
    private static readonly HashSet<string> Keywords = new(
        ("CREATE GLOBAL TEMPORARY TEMP TABLE DROP IF EXISTS INSERT INTO VALUES ALL SELECT FROM " +
         "WHERE AS ON COMMIT PRESERVE DELETE ROWS NOT NULL DEFAULT PRIMARY KEY UNIQUE CONSTRAINT " +
         "ALTER ADD UPDATE SET WITH AND OR IN IS LIKE BETWEEN CASE WHEN THEN ELSE END BEGIN " +
         "DECLARE EXEC EXECUTE TRUE FALSE ORDER BY GROUP HAVING LIMIT OFFSET DISTINCT UNION JOIN " +
         "LEFT RIGHT INNER OUTER TOP TRUNCATE CASCADE RESTRICT REPLACE RETURNING").Split(' '),
        StringComparer.OrdinalIgnoreCase);
    private static readonly HashSet<string> Types = new(
        ("VARCHAR NVARCHAR VARCHAR2 NVARCHAR2 CHAR NCHAR TEXT NTEXT CLOB NCLOB INT INTEGER BIGINT " +
         "SMALLINT TINYINT NUMBER NUMERIC DECIMAL FLOAT DOUBLE REAL DATE DATETIME DATETIME2 TIMESTAMP " +
         "TIME BOOLEAN BOOL BIT BINARY VARBINARY BLOB BYTEA SERIAL BIGSERIAL UUID JSON JSONB MAX " +
         "COUNT SUM AVG MIN COALESCE CAST CONVERT TO_DATE TO_TIMESTAMP OBJECT_ID").Split(' '),
        StringComparer.OrdinalIgnoreCase);
    private static readonly Regex Tokens = new(
        "--[^\\r\\n]*|/\\*[\\s\\S]*?(?:\\*/|\\z)|'(?:''|[^'])*(?:'|\\z)|" +
        "\"(?:\"\"|[^\"])*(?:\"|\\z)|\\[(?:\\]\\]|[^\\]])*(?:\\]|\\z)|`(?:``|[^`])*(?:`|\\z)|" +
        @"\b\d+(?:\.\d+)?(?:[eE][+-]?\d+)?\b|[\p{L}_#@][\p{L}\p{N}_$#@]*|[(),;.*=+/<>%-]",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private readonly List<Token> _tokens = [];
    private string _source = "";
    private readonly Dictionary<int, (int Start, DrawingGroup Drawing)> _lines = [];
    private readonly Queue<int> _cachedLineOrder = new();
    private (double Width, double Dpi, string FontFamily, FontStyle Style, FontWeight Weight, FontStretch Stretch, double FontSize) _layout;
    private readonly record struct Token(int Start, int Length, Brush Color)
    {
        public int End => Start + Length;
    }

    public SqlSyntaxLayer()
    {
        IsHitTestVisible = false;
        SizeChanged += (_, _) => InvalidateVisual();
    }

    private static void EditorChanged(DependencyObject source, DependencyPropertyChangedEventArgs args)
    {
        var layer = (SqlSyntaxLayer)source;
        if (args.OldValue is TextBox previous)
        {
            previous.TextChanged -= layer.TextChanged;
            previous.RemoveHandler(ScrollViewer.ScrollChangedEvent, new ScrollChangedEventHandler(layer.ScrollChanged));
        }
        if (args.NewValue is TextBox current)
        {
            current.TextChanged += layer.TextChanged;
            current.AddHandler(ScrollViewer.ScrollChangedEvent, new ScrollChangedEventHandler(layer.ScrollChanged));
        }
        layer.UpdateTokens();
    }

    private void TextChanged(object sender, TextChangedEventArgs args) => UpdateTokens();
    private void ScrollChanged(object sender, ScrollChangedEventArgs args) => InvalidateVisual();

    private void UpdateTokens()
    {
        _tokens.Clear();
        ClearLines();
        _source = Editor?.Text ?? "";
        // Scan once per document, including multiline literals; paint only visible lines below.
        foreach (Match match in Tokens.Matches(_source))
        {
            string value = match.Value;
            Brush? color = value.StartsWith("--", StringComparison.Ordinal) || value.StartsWith("/*", StringComparison.Ordinal)
                ? Comment
                : value[0] == '\'' ? String
                : value[0] is '"' or '[' or '`' ? Identifier
                : char.IsDigit(value[0]) ? Number
                : Keywords.Contains(value) ? Keyword
                : Types.Contains(value) ? DataType
                : "(),;.*=+/<>%-".Contains(value[0]) ? Punctuation
                : null;
            if (color != null) _tokens.Add(new Token(match.Index, match.Length, color));
        }
        InvalidateVisual();
    }

    protected override void OnRender(DrawingContext dc)
    {
        var box = Editor;
        if (box == null || !box.IsArrangeValid || box.LineCount == 0) return;
        int first = box.GetFirstVisibleLineIndex();
        int last = Math.Min(box.GetLastVisibleLineIndex(), box.LineCount - 1);
        if (first < 0 || first > last) return;

        var scroll = box.Template.FindName("PART_ContentHost", box) as ScrollViewer;
        var bar = scroll?.Template.FindName("PART_VerticalScrollBar", scroll) as FrameworkElement;
        double right = ActualWidth - (scroll?.ComputedVerticalScrollBarVisibility == Visibility.Visible ? bar?.ActualWidth ?? 0 : 0);
        dc.PushClip(new RectangleGeometry(new Rect(0, 0, Math.Max(0, right), ActualHeight)));
        var origin = box.TranslatePoint(new Point(), this);
        var typeface = new Typeface(box.FontFamily, box.FontStyle, box.FontWeight, box.FontStretch);
        double dpi = VisualTreeHelper.GetDpi(this).PixelsPerDip;
        var layout = (right, dpi, box.FontFamily.Source, box.FontStyle, box.FontWeight, box.FontStretch, box.FontSize);
        if (_layout != layout)
        {
            ClearLines();
            _layout = layout;
        }

        // WPF's per-line text/geometry queries reformat text. Query the visible origin once,
        // then reuse immutable drawings and the TextBox's uniform line spacing while scrolling.
        var firstLine = GetLine(first, box, typeface, dpi);
        var rect = box.GetRectFromCharacterIndex(firstLine.Start);
        if (rect.IsEmpty) { dc.Pop(); return; }
        double lineHeight = rect.Height;
        if (first < last)
        {
            var nextRect = box.GetRectFromCharacterIndex(GetLine(first + 1, box, typeface, dpi).Start);
            if (!nextRect.IsEmpty) lineHeight = nextRect.Y - rect.Y;
        }
        for (int line = first; line <= last; line++)
        {
            var cached = GetLine(line, box, typeface, dpi);
            dc.PushTransform(new TranslateTransform(origin.X + rect.X, origin.Y + rect.Y + (line - first) * lineHeight));
            dc.DrawDrawing(cached.Drawing);
            dc.Pop();
        }
        dc.Pop();
    }

    private (int Start, DrawingGroup Drawing) GetLine(int line, TextBox box, Typeface typeface, double dpi)
    {
        if (_lines.TryGetValue(line, out var cached)) return cached;
        int start = box.GetCharacterIndexFromLineIndex(line);
        string source = _source.Substring(start, box.GetLineLength(line)).TrimEnd('\r', '\n');
        var drawing = new DrawingGroup();
        using var dc = drawing.Open();
        int low = 0, high = _tokens.Count;
        while (low < high)
        {
            int middle = (low + high) / 2;
            if (_tokens[middle].End <= start) low = middle + 1;
            else high = middle;
        }
        if (source.Length > 0)
        {
            var text = new FormattedText(source, CultureInfo.CurrentUICulture, FlowDirection.LeftToRight,
                typeface, box.FontSize, Plain, dpi);
            for (int i = low; i < _tokens.Count && _tokens[i].Start < start + source.Length; i++)
            {
                var token = _tokens[i];
                int from = Math.Max(0, token.Start - start);
                int to = Math.Min(source.Length, token.End - start);
                text.SetForegroundBrush(token.Color, from, to - from);
            }
            dc.DrawText(text, new Point());
        }
        dc.Close();
        drawing.Freeze();
        // Bound memory independently of SQL size; retain nearby lines for reverse scrolling.
        if (_lines.Count >= 256) _lines.Remove(_cachedLineOrder.Dequeue());
        _cachedLineOrder.Enqueue(line);
        cached = (start, drawing);
        _lines.Add(line, cached);
        return cached;
    }

    private void ClearLines()
    {
        _lines.Clear();
        _cachedLineOrder.Clear();
    }

    private static SolidColorBrush ColorBrush(string value)
    {
        var brush = new SolidColorBrush((Color)ColorConverter.ConvertFromString(value));
        brush.Freeze();
        return brush;
    }
}
