namespace ProcurementControl.Navigation;

/// <summary>
/// Пункт левого навигационного меню. Либо ссылка на раздел (Page),
/// либо визуальный разделитель (IsDivider == true).
/// </summary>
public sealed class NavigationItem
{
    public AppPage Page { get; init; }

    public string Title { get; init; } = string.Empty;

    /// <summary>Символ шрифта Segoe MDL2 Assets для иконки (косметика).</summary>
    public string IconGlyph { get; init; } = string.Empty;

    public bool IsDivider { get; init; }

    public static NavigationItem Divider() => new() { IsDivider = true };

    public static NavigationItem For(AppPage page, string title, string iconGlyph = "")
        => new() { Page = page, Title = title, IconGlyph = iconGlyph };
}
