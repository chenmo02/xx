using System.Globalization;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace WpfApp1.Controls;

// Keep native TextBox editing, selection, IME and undo; draw syntax and markers separately.
public sealed class JsonDiffEditor : Grid
{
    public TextBox Input { get; }
    private readonly SyntaxLayer _syntax;
    private IReadOnlyDictionary<int, (Brush Color, string Symbol)> _markers =
        new Dictionary<int, (Brush, string)>();
    private int? _selectedLine;
    private readonly List<int> _lineStarts = [0];

    public JsonDiffEditor()
    {
        Background = Brush("#1C293B");
        ClipToBounds = true;
        Input = new TextBox
        {
            Style = null,
            AcceptsReturn = true,
            AcceptsTab = true,
            TextWrapping = TextWrapping.Wrap,
            FontFamily = new FontFamily("Consolas, Microsoft YaHei UI"),
            FontSize = 13,
            Background = Brushes.Transparent,
            Foreground = Brushes.Transparent,
            CaretBrush = Brushes.White,
            SelectionBrush = Brush("#426A9D"),
            SelectionOpacity = 0.55,
            BorderThickness = new Thickness(0),
            Padding = new Thickness(10, 10, 10, 12),
            Margin = new Thickness(48, 0, 0, 0),
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled
        };
        Children.Add(Input);
        _syntax = new SyntaxLayer(this) { IsHitTestVisible = false };
        Children.Add(_syntax);
        Input.TextChanged += (_, _) =>
        {
            _lineStarts.Clear();
            _lineStarts.Add(0);
            string source = Input.Text;
            for (int i = 0; i < source.Length; i++)
            {
                if (source[i] == '\n' || (source[i] == '\r' && (i + 1 == source.Length || source[i + 1] != '\n')))
                    _lineStarts.Add(i + 1);
            }
            _syntax.InvalidateVisual();
        };
        Input.AddHandler(ScrollViewer.ScrollChangedEvent,
            new ScrollChangedEventHandler((_, _) => _syntax.InvalidateVisual()));
        SizeChanged += (_, _) => _syntax.InvalidateVisual();
    }

    public void SetMarkers(IReadOnlyDictionary<int, (Brush Color, string Symbol)> markers)
    {
        _markers = markers;
        _selectedLine = null;
        _syntax.InvalidateVisual();
    }

    public void GoToLine(int? line)
    {
        _selectedLine = line;
        if (line is int number && number > 0 && number <= _lineStarts.Count)
        {
            Input.UpdateLayout();
            Input.CaretIndex = _lineStarts[number - 1];
            int visualLine = Input.GetLineIndexFromCharacterIndex(Input.CaretIndex);
            if (visualLine >= 0) Input.ScrollToLine(Math.Max(0, visualLine - 2));
        }
        _syntax.InvalidateVisual();
    }

    private static SolidColorBrush Brush(string value) =>
        new((Color)ColorConverter.ConvertFromString(value));

    private sealed class SyntaxLayer(JsonDiffEditor owner) : FrameworkElement
    {
        private static readonly Regex Tokens = new(
            "\"(?:\\\\.|[^\"\\\\])*\"|\\b(?:true|false|null)\\b|-?\\b\\d+(?:\\.\\d+)?(?:[eE][+-]?\\d+)?",
            RegexOptions.Compiled);
        private static readonly Brush Plain = Brush("#DAE5F1");
        private static readonly Brush Key = Brush("#85D3FF");
        private static readonly Brush String = Brush("#9DDC8B");
        private static readonly Brush Number = Brush("#F4B968");
        private static readonly Brush Literal = Brush("#5CD5F5");
        private static readonly Brush Muted = Brush("#8393A9");

        protected override void OnRender(DrawingContext dc)
        {
            var box = owner.Input;
            if (!box.IsArrangeValid) return;
            int lineCount = box.LineCount;
            if (lineCount <= 0) return;
            var first = box.GetFirstVisibleLineIndex();
            // WPF derives visible indexes from scroll geometry; after reflow the last
            // visible index can reach LineCount. Only draw lines in the actual layout.
            var last = Math.Min(box.GetLastVisibleLineIndex(), lineCount - 1);
            if (first < 0 || first > last) return;
            var origin = box.TranslatePoint(new Point(), this);
            var scroll = box.Template.FindName("PART_ContentHost", box) as ScrollViewer;
            var verticalBar = scroll?.Template.FindName("PART_VerticalScrollBar", scroll) as FrameworkElement;
            var horizontalBar = scroll?.Template.FindName("PART_HorizontalScrollBar", scroll) as FrameworkElement;
            double viewportRight = Math.Max(48, ActualWidth -
                (scroll?.ComputedVerticalScrollBarVisibility == Visibility.Visible ? verticalBar?.ActualWidth ?? 0 : 0));
            double viewportBottom = Math.Max(0, ActualHeight -
                (scroll?.ComputedHorizontalScrollBarVisibility == Visibility.Visible ? horizontalBar?.ActualHeight ?? 0 : 0));
            dc.PushClip(new RectangleGeometry(new Rect(0, 0, viewportRight,
                viewportBottom)));
            dc.DrawLine(new Pen(Brush("#304055"), 1), new Point(47, 8), new Point(47, ActualHeight));
            string document = box.Text;
            var logicalLines = new Dictionary<int, (string Source, MatchCollection Tokens)>();

            for (int line = first; line <= last; line++)
            {
                int start = box.GetCharacterIndexFromLineIndex(line);
                // WPF line indexes include soft wraps; markers and JSON paths use source lines.
                int found = owner._lineStarts.BinarySearch(start);
                int logicalLine = found >= 0 ? found : ~found - 1;
                int logicalStart = owner._lineStarts[logicalLine];
                bool continuation = start != logicalStart;
                var rect = box.GetRectFromCharacterIndex(start);
                if (rect.IsEmpty) continue;
                double top = origin.Y + rect.Y;
                double height = rect.Height;
                if (owner._markers.TryGetValue(logicalLine + 1, out var marker))
                {
                    dc.PushOpacity(0.14);
                    dc.DrawRectangle(marker.Color, null, new Rect(0, top, viewportRight, height));
                    dc.Pop();
                    if (!continuation) dc.DrawText(Text(marker.Symbol, marker.Color), new Point(5, top));
                }
                if (owner._selectedLine == logicalLine + 1)
                {
                    var outline = new Pen(Brush("#6DA5FF"), 1);
                    dc.DrawLine(outline, new Point(1, top), new Point(1, top + height));
                    dc.DrawLine(outline, new Point(viewportRight - 1, top), new Point(viewportRight - 1, top + height));
                    if (!continuation)
                        dc.DrawLine(outline, new Point(1, top), new Point(viewportRight - 1, top));
                    int nextStart = line + 1 < lineCount ? box.GetCharacterIndexFromLineIndex(line + 1) : document.Length;
                    bool lastSegment = logicalLine + 1 < owner._lineStarts.Count
                        ? nextStart >= owner._lineStarts[logicalLine + 1] : line == lineCount - 1;
                    if (lastSegment)
                        dc.DrawLine(outline, new Point(1, top + height), new Point(viewportRight - 1, top + height));
                }

                if (!continuation)
                {
                    var number = Text((logicalLine + 1).ToString(CultureInfo.InvariantCulture), Muted);
                    dc.DrawText(number, new Point(39 - number.Width, top));
                }
                string source = box.GetLineText(line).TrimEnd('\r', '\n');
                if (source.Length == 0) continue;
                var text = Text(source, Plain);
                if (!logicalLines.TryGetValue(logicalLine, out var logical))
                {
                    int logicalEnd = logicalLine + 1 < owner._lineStarts.Count ? owner._lineStarts[logicalLine + 1] : document.Length;
                    string fullLine = document[logicalStart..logicalEnd].TrimEnd('\r', '\n');
                    logical = (fullLine, Tokens.Matches(fullLine));
                    logicalLines.Add(logicalLine, logical);
                }
                int offset = start - logicalStart;
                foreach (Match match in logical.Tokens)
                {
                    int from = Math.Max(0, match.Index - offset);
                    int to = Math.Min(source.Length, match.Index + match.Length - offset);
                    if (to <= from) continue;
                    Brush color;
                    if (match.Value[0] == '"')
                    {
                        int end = match.Index + match.Length;
                        while (end < logical.Source.Length && char.IsWhiteSpace(logical.Source[end])) end++;
                        color = end < logical.Source.Length && logical.Source[end] == ':' ? Key : String;
                    }
                    else color = char.IsLetter(match.Value[0]) ? Literal : Number;
                    text.SetForegroundBrush(color, from, to - from);
                }
                dc.PushClip(new RectangleGeometry(new Rect(48, 0,
                    Math.Max(0, viewportRight - 48), ActualHeight)));
                dc.DrawText(text, new Point(origin.X + rect.X, top));
                dc.Pop();
            }
            dc.Pop();
        }

        private FormattedText Text(string value, Brush color) => new(value,
            CultureInfo.CurrentUICulture, FlowDirection.LeftToRight,
            new Typeface(owner.Input.FontFamily, FontStyles.Normal, FontWeights.Normal, FontStretches.Normal),
            owner.Input.FontSize, color, VisualTreeHelper.GetDpi(this).PixelsPerDip);
    }
}
