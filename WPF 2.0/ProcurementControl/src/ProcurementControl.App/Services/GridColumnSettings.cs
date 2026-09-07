using System.Globalization;
using System.Windows.Controls;

namespace ProcurementControl.Services;

/// <summary>Сохраняет пользовательские ширины DataGrid в общих настройках SQLite.</summary>
public static class GridColumnSettings
{
    public static void Restore(DataGrid grid, string settingKey)
    {
        try
        {
            using var snapshot = PurchaseSnapshot.Create();
            var text = new SettingsRepository(snapshot.Connection).GetSetting(settingKey);
            if (string.IsNullOrWhiteSpace(text)) return;
            var widths = text.Split(';', StringSplitOptions.RemoveEmptyEntries)
                .Select(part => part.Split('=', 2))
                .Where(parts => parts.Length == 2 && double.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out _))
                .ToDictionary(parts => parts[0], parts => double.Parse(parts[1], CultureInfo.InvariantCulture), StringComparer.Ordinal);
            foreach (var column in grid.Columns)
                if (widths.TryGetValue(Key(column), out var width) && width >= 30) column.Width = width;
        }
        catch { }
    }

    public static void Save(DataGrid grid, string settingKey)
    {
        try
        {
            var values = grid.Columns.Select(column => Key(column) + "=" + Math.Round(column.ActualWidth).ToString(CultureInfo.InvariantCulture));
            PurchaseWriteRepository.SetSetting(settingKey, string.Join(";", values));
        }
        catch { }
    }

    private static string Key(DataGridColumn column) => !string.IsNullOrWhiteSpace(column.SortMemberPath)
        ? column.SortMemberPath : Convert.ToString(column.Header, CultureInfo.InvariantCulture) ?? string.Empty;
}
