using System.Windows;
using System.Windows.Controls;

namespace ProcurementControl.Navigation;

/// <summary>
/// Выбирает шаблон содержимого пункта навигации: обычная страница или разделитель.
/// </summary>
public sealed class NavItemTemplateSelector : DataTemplateSelector
{
    public DataTemplate? PageTemplate { get; set; }
    public DataTemplate? DividerTemplate { get; set; }

    public override DataTemplate? SelectTemplate(object? item, DependencyObject container)
        => item is NavigationItem { IsDivider: true } ? DividerTemplate : PageTemplate;
}

/// <summary>
/// Выбирает стиль контейнера (ListBoxItem) пункта навигации: кликабельная страница
/// или неинтерактивный разделитель.
/// </summary>
public sealed class NavItemContainerStyleSelector : StyleSelector
{
    public Style? PageStyle { get; set; }
    public Style? DividerStyle { get; set; }

    public override Style? SelectStyle(object? item, DependencyObject container)
        => item is NavigationItem { IsDivider: true } ? DividerStyle : PageStyle;
}
