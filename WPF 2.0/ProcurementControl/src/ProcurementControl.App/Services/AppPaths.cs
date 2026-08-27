using System.IO;

namespace ProcurementControl.Services;

/// <summary>
/// Резолвинг путей к общей папке данных.
///
/// Оригинальное PowerShell-приложение хранит данные в каталоге "data" в корне
/// портативной сборки (см. app/modules/PurchaseStore.ps1, Get-NotesDataRoot и др.,
/// которые строят пути от $script:AppRoot). Чтобы WPF-версия работала с теми же
/// данными, поднимаемся от расположения исполняемого файла вверх до каталога,
/// содержащего папки "app" и "data".
/// </summary>
public static class AppPaths
{
    private static string? _dataRoot;

    /// <summary>Корень портативной сборки (аналог $script:AppRoot).</summary>
    public static string AppRoot => ResolveAppRoot();

    /// <summary>Каталог данных: {AppRoot}/data.</summary>
    public static string DataRoot => Path.Combine(AppRoot, "data");

    /// <summary>Каталог заметок: {AppRoot}/data/notes.</summary>
    public static string NotesDirectory => Path.Combine(DataRoot, "notes");

    /// <summary>Каталог базы закупок: {AppRoot}/data/purchase_control.</summary>
    public static string PurchaseDataDirectory => Path.Combine(DataRoot, "purchase_control");

    /// <summary>Путь к SQLite-базе закупок (общая с оригиналом, только чтение).</summary>
    public static string PurchaseDatabasePath => Path.Combine(PurchaseDataDirectory, "purchase_control.sqlite");

    private static string ResolveAppRoot()
    {
        if (_dataRoot is not null)
        {
            return _dataRoot;
        }

        var current = new DirectoryInfo(AppContext.BaseDirectory);
        while (current is not null)
        {
            bool hasData = Directory.Exists(Path.Combine(current.FullName, "data"));
            bool hasApp = Directory.Exists(Path.Combine(current.FullName, "app"));
            if (hasData && hasApp)
            {
                _dataRoot = current.FullName;
                return _dataRoot;
            }
            current = current.Parent;
        }

        // Fallback: если корень не найден, используем папку рядом с исполняемым файлом.
        _dataRoot = AppContext.BaseDirectory;
        return _dataRoot;
    }
}
