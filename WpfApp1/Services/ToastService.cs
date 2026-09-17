using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Threading;

namespace WpfApp1.Services;

public static class ToastService
{
    private static readonly ConditionalWeakTable<FrameworkElement, ToastAdorner> Active = new();

    public static void Show(FrameworkElement owner, string message)
    {
        if (!owner.Dispatcher.CheckAccess())
        {
            owner.Dispatcher.Invoke(() => Show(owner, message));
            return;
        }

        var target = owner is Window { Content: FrameworkElement content } ? content : owner;
        if (!target.IsLoaded) return;
        if (Active.TryGetValue(target, out var current))
        {
            current.Show(message);
            return;
        }

        var layer = AdornerLayer.GetAdornerLayer(target);
        if (layer == null) return;
        var toast = new ToastAdorner(target, layer);
        Active.Add(target, toast);
        layer.Add(toast);
        toast.Show(message);
    }

    private sealed class ToastAdorner : Adorner
    {
        private readonly FrameworkElement _owner;
        private readonly AdornerLayer _layer;
        private readonly Grid _visual = new();
        private readonly TextBlock _message;
        private readonly DispatcherTimer _timer = new() { Interval = TimeSpan.FromSeconds(1.5) };

        public ToastAdorner(FrameworkElement owner, AdornerLayer layer) : base(owner)
        {
            _owner = owner;
            _layer = layer;
            IsHitTestVisible = false;
            _message = new TextBlock { Style = (Style)owner.FindResource("AppToastTextStyle") };
            _visual.Children.Add(new Border { Style = (Style)owner.FindResource("AppToastStyle"), Child = _message });
            AddVisualChild(_visual);
            AddLogicalChild(_visual);
            _timer.Tick += (_, _) => Dismiss();
            owner.Unloaded += Owner_Unloaded;
        }

        public void Show(string message)
        {
            _message.Text = message;
            _timer.Stop();
            _timer.Start();
        }

        private void Owner_Unloaded(object sender, RoutedEventArgs e) => Dismiss();

        private void Dismiss()
        {
            _timer.Stop();
            _owner.Unloaded -= Owner_Unloaded;
            _layer.Remove(this);
            Active.Remove(_owner);
        }

        protected override int VisualChildrenCount => 1;
        protected override Visual GetVisualChild(int index) => _visual;
        protected override Size MeasureOverride(Size constraint)
        {
            _visual.Measure(AdornedElement.RenderSize);
            return AdornedElement.RenderSize;
        }
        protected override Size ArrangeOverride(Size finalSize)
        {
            _visual.Arrange(new Rect(finalSize));
            return finalSize;
        }
    }
}
