using System.Windows;
using System.Windows.Media;

namespace ProcurementControl.Views;

/// <summary>
/// Прикреплённое свойство скруглённой обрезки содержимого. Позволяет
/// вписать таблицы и другие прямоугольные элементы в карточки со
/// скруглёнными углами без артефактов по углам.
/// Использование: local:RoundedClip.Radius="10" на Border/DataGrid.
/// </summary>
public static class RoundedClip
{
    public static readonly DependencyProperty RadiusProperty = DependencyProperty.RegisterAttached(
        "Radius",
        typeof(double),
        typeof(RoundedClip),
        new PropertyMetadata(0.0, OnRadiusChanged));

    public static double GetRadius(DependencyObject obj)
        => (double)obj.GetValue(RadiusProperty);

    public static void SetRadius(DependencyObject obj, double value)
        => obj.SetValue(RadiusProperty, value);

    private static void OnRadiusChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is not FrameworkElement element)
        {
            return;
        }

        element.SizeChanged -= OnSizeChanged;
        element.SizeChanged += OnSizeChanged;
        ApplyClip(element, (double)e.NewValue);
    }

    private static void OnSizeChanged(object sender, SizeChangedEventArgs e)
    {
        if (sender is FrameworkElement element)
        {
            ApplyClip(element, GetRadius(element));
        }
    }

    private static void ApplyClip(FrameworkElement element, double radius)
    {
        if (radius <= 0 || element.ActualWidth <= 0 || element.ActualHeight <= 0)
        {
            element.Clip = null;
            return;
        }

        element.Clip = new RectangleGeometry(
            new Rect(0, 0, element.ActualWidth, element.ActualHeight),
            radius,
            radius);
    }
}
