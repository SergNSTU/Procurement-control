using System.IO;
using Microsoft.Data.Sqlite;

namespace ProcurementControl.Services;

/// <summary>
/// Безопасный доступ к базе закупок WinForms-приложения БЕЗ её изменения.
///
/// Оригинальная БД (data/purchase_control/purchase_control.sqlite) работает в режиме
/// WAL, и даже read-only подключение может выполнить WAL-recovery и затронуть файлы
/// -wal/-shm. Чтобы гарантированно не трогать оригинал, мы КОПИРУЕМ файлы .sqlite и
/// -wal во временную папку и открываем только копию (все служебные записи SQLite
/// происходят внутри копии). Исходная база остаётся нетронутой.
/// </summary>
public sealed class PurchaseSnapshot : IDisposable
{
    private readonly string _snapshotDir;
    private bool _disposed;

    public SqliteConnection Connection { get; }

    /// <summary>
    /// Путь к временной копии базы, если снапшот создан копированием.
    /// Для ReadOnly-фолбэка — <c>null</c>. Используется резервными копиями,
    /// чтобы положить в архив уже прочекпоинченную копию.
    /// </summary>
    public string? SnapshotDatabasePath => string.IsNullOrEmpty(_snapshotDir)
        ? null
        : Path.Combine(_snapshotDir, "purchase_control.sqlite");

    private PurchaseSnapshot(string snapshotDir, SqliteConnection connection)
    {
        _snapshotDir = snapshotDir;
        Connection = connection;
    }

    /// <summary>Создаёт временную копию БД и открывает подключение к копии.</summary>
    public static PurchaseSnapshot Create()
    {
        var source = AppPaths.PurchaseDatabasePath;
        if (!File.Exists(source))
        {
            throw new FileNotFoundException("База закупок не найдена: " + source, source);
        }

        try
        {
            return CreateFromCopy(source);
        }
        catch (IOException)
        {
            // Оригинал жёстко заблокирован (например, WinForms-приложение держит файл).
            // Запасной вариант: открыть оригинал строго в режиме ReadOnly — SQLite в этом
            // режиме ничего не пишет ни в .sqlite, ни в -wal/-shm.
            return CreateReadOnlyFallback(source);
        }
    }

    private static PurchaseSnapshot CreateFromCopy(string source)
    {
        var snapshotDir = Path.Combine(Path.GetTempPath(), "ProcurementControl.Wpf", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(snapshotDir);

        var mainDest = Path.Combine(snapshotDir, "purchase_control.sqlite");
        try
        {
            CopyLockedFile(source, mainDest);

            var walSource = source + "-wal";
            if (File.Exists(walSource))
            {
                CopyLockedFile(walSource, mainDest + "-wal");
            }
        }
        catch
        {
            // Не удалось скопировать — убираем временную папку и пробрасываем ошибку наверх.
            try { Directory.Delete(snapshotDir, recursive: true); } catch { }
            throw;
        }

        // Открываем копию в ReadWrite, чтобы WAL-recovery и checkpoint произошли
        // только внутри временной копии, а не на оригинальной базе.
        var connection = new SqliteConnection($"Data Source={mainDest}");
        connection.Open();

        using (var checkpoint = connection.CreateCommand())
        {
            checkpoint.CommandText = "PRAGMA wal_checkpoint(TRUNCATE);";
            checkpoint.ExecuteNonQuery();
        }

        return new PurchaseSnapshot(snapshotDir, connection);
    }

    /// <summary>
    /// Копирует файл, который может быть открыт другим процессом на запись
    /// (File.Copy использует FileShare.Read и на живой БД падает).
    /// </summary>
    private static void CopyLockedFile(string sourcePath, string destPath)
    {
        using var input = new FileStream(sourcePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var output = new FileStream(destPath, FileMode.Create, FileAccess.Write, FileShare.None);
        input.CopyTo(output);
    }

    private static PurchaseSnapshot CreateReadOnlyFallback(string source)
    {
        var builder = new SqliteConnectionStringBuilder
        {
            DataSource = source,
            Mode = SqliteOpenMode.ReadOnly,
        };
        var connection = new SqliteConnection(builder.ConnectionString);
        connection.Open();
        // _snapshotDir пустой — удалять нечего.
        return new PurchaseSnapshot(string.Empty, connection);
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }
        _disposed = true;

        try
        {
            // Microsoft.Data.Sqlite пулирует подключения: без ClearPool файл копии
            // остаётся открытым и временная папка не удаляется.
            SqliteConnection.ClearPool(Connection);
            Connection.Dispose();
        }
        catch
        {
            // Игнорируем ошибки закрытия копии.
        }

        try
        {
            if (!string.IsNullOrEmpty(_snapshotDir))
            {
                Directory.Delete(_snapshotDir, recursive: true);
            }
        }
        catch
        {
            // Временная папка может удаляться с задержкой — не критично.
        }
    }
}
