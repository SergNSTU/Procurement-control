namespace ProcurementControl.Models;

/// <summary>
/// Одна строка результата подбора аналогов: оригинальный компонент или его аналог.
/// </summary>
public sealed class CrossReferenceRow
{
    /// <summary>Тип замены: «Оригинал», «Другая упаковка», «Похожий» и т.п.</summary>
    public string Type { get; init; } = string.Empty;

    public string Mpn { get; init; } = string.Empty;

    /// <summary>Код позиции LCSC вида C8734.</summary>
    public string LcscCode { get; init; } = string.Empty;

    public string Brand { get; init; } = string.Empty;

    public string Package { get; init; } = string.Empty;

    public string Category { get; init; } = string.Empty;

    public string Stock { get; init; } = string.Empty;

    public string Moq { get; init; } = string.Empty;

    /// <summary>Лесенка цен в USD: «1+ $2.17 · 10+ $1.83 · 100+ $1.40».</summary>
    public string PriceText { get; init; } = string.Empty;

    public string Description { get; init; } = string.Empty;

    public string Link { get; init; } = string.Empty;

    public string Datasheet { get; init; } = string.Empty;

    /// <summary>Ссылка пригодна для открытия по двойному клику.</summary>
    public bool HasLink => !string.IsNullOrWhiteSpace(Link)
        && (Link.StartsWith("http://", StringComparison.OrdinalIgnoreCase)
            || Link.StartsWith("https://", StringComparison.OrdinalIgnoreCase));
}

/// <summary>Итог поиска аналогов: выбранный исходный компонент и строки для таблицы.</summary>
public sealed class CrossReferenceResult
{
    public string Query { get; init; } = string.Empty;

    /// <summary>Найден ли в LCSC компонент с точно таким же MPN.</summary>
    public bool ExactMatch { get; init; }

    public int CandidateCount { get; init; }

    /// <summary>MPN и код позиции, по которой LCSC подобрал аналоги.</summary>
    public string ResolvedModel { get; init; } = string.Empty;

    public string ResolvedCode { get; init; } = string.Empty;

    public List<CrossReferenceRow> Rows { get; init; } = new();
}
