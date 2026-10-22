using System.IO;

namespace ProcurementControl.Services;

/// <summary>
/// Резолвинг путей к общей папке данных.
///
/// Приложение хранит данные в каталоге "data" в корне портативной сборки
/// (data/purchase_control/purchase_control.sqlite). Поднимаемся от расположения
/// исполняемого файла вверх до каталога, содержащего "data/purchase_control":
/// в переносимой сборке он лежит рядом с exe, в рабочем пространстве — в корне
/// воркспейса. Раньше корень определялся по паре папок "app" и "data" (маркер
/// старой PowerShell-версии); старая версия удалена 2026-10-22, остался поиск
/// по data/purchase_control.
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
            if (Directory.Exists(Path.Combine(current.FullName, "data", "purchase_control")))
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
