using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Media3D;

namespace WpfApp1.Behaviors;

// One controller on the outer viewer routes wheel input to the innermost scrollable viewport.
public sealed class SmoothWheelScroll
{
    private readonly ScrollViewer _owner;
    private ScrollViewer? _target;
    private double _destination;
    private double _position;
    private long _lastFrame;

    public SmoothWheelScroll(ScrollViewer owner)
    {
        _owner = owner;
        owner.PreviewMouseWheel += OnWheel;
        owner.PreviewMouseDown += (_, _) => Stop();
        owner.PreviewKeyDown += (_, _) => Stop();
        owner.Unloaded += (_, _) => Stop();
    }

    private void OnWheel(object sender, MouseWheelEventArgs e)
    {
        if (e.Handled || e.Delta == 0 || SystemParameters.WheelScrollLines == 0) return;
        ScrollViewer? viewer = null;
        for (DependencyObject? node = e.OriginalSource as DependencyObject; node != null; node = Parent(node))
        {
            // Let an open dropdown keep its own wheel handling.
            if (node is ComboBox) return;
            var candidate = node as ScrollViewer;
            if (candidate == null && node is TextBoxBase or DataGrid)
                candidate = FindViewer(node);
            if (candidate != null && candidate.ScrollableHeight > 0 &&
                (e.Delta < 0 ? candidate.VerticalOffset < candidate.ScrollableHeight - 0.5 : candidate.VerticalOffset > 0.5))
            {
                viewer = candidate;
                break;
            }
            if (node == _owner) break;
        }
        if (viewer == null) { Stop(); return; }
        double distance = SystemParameters.WheelScrollLines < 0
            ? viewer.ViewportHeight
            : SystemParameters.WheelScrollLines * 16.0;
        double movement = -e.Delta / 120.0 * distance;

        if (!SystemParameters.ClientAreaAnimation)
        {
            Stop();
            viewer.ScrollToVerticalOffset(Math.Clamp(viewer.VerticalOffset + movement, 0, viewer.ScrollableHeight));
        }
        else
        {
            if (_target != viewer)
            {
                Stop();
                _target = viewer;
                _destination = viewer.VerticalOffset;
                _position = viewer.VerticalOffset;
                _lastFrame = Stopwatch.GetTimestamp();
                CompositionTarget.Rendering += OnFrame;
            }
            if (Math.Sign(movement) != Math.Sign(_destination - viewer.VerticalOffset))
                _destination = viewer.VerticalOffset;
            _destination = Math.Clamp(_destination + movement, 0, viewer.ScrollableHeight);
        }
        e.Handled = true;
    }

    private void OnFrame(object? sender, EventArgs e)
    {
        var viewer = _target;
        if (viewer == null) return;
        if (!viewer.IsLoaded) { Stop(); return; }
        long now = Stopwatch.GetTimestamp();
        double elapsed = (now - _lastFrame) / (double)Stopwatch.Frequency;
        _lastFrame = now;
        _destination = Math.Clamp(_destination, 0, viewer.ScrollableHeight);
        double remaining = _destination - _position;
        if (Math.Abs(remaining) < 0.5)
        {
            viewer.ScrollToVerticalOffset(_destination);
            Stop();
            return;
        }
        _position += remaining * (1 - Math.Exp(-elapsed / 0.045));
        viewer.ScrollToVerticalOffset(_position);
    }

    private void Stop()
    {
        CompositionTarget.Rendering -= OnFrame;
        _target = null;
    }

    private static DependencyObject? Parent(DependencyObject node) => node switch
    {
        Visual or Visual3D => VisualTreeHelper.GetParent(node),
        FrameworkContentElement content => content.Parent,
        _ => LogicalTreeHelper.GetParent(node)
    };

    private static ScrollViewer? FindViewer(DependencyObject node)
    {
        for (int i = 0; i < VisualTreeHelper.GetChildrenCount(node); i++)
        {
            var child = VisualTreeHelper.GetChild(node, i);
            if (child is ScrollViewer viewer) return viewer;
            if (FindViewer(child) is ScrollViewer nested) return nested;
        }
        return null;
    }
}
