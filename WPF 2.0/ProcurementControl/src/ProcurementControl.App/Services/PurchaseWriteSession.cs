using System.IO;
using Microsoft.Data.Sqlite;

namespace ProcurementControl.Services;

/// <summary>
/// Короткая write-сессия для общей SQLite-базы. В отличие от
/// <see cref="PurchaseSnapshot"/>, эта сессия намеренно открывает оригинальную
/// базу, но не меняет PowerShell- или WinForms-код. Каждая операция выполняется
/// в отдельной транзакции BEGIN IMMEDIATE, поэтому второй экземпляр приложения
/// либо ждёт освобождения блокировки, либо получает понятную ошибку.
/// </summary>
public static class PurchaseWriteSession
{
    private const int BusyTimeoutSeconds = 10;
    private const int BeginAttempts = 4;

    /// <summary>
    /// Выполняет одну атомарную, недеструктивную операцию записи.
    /// Вызывающий код не должен сохранять переданное подключение.
    /// </summary>
    public static T Execute<T>(Func<SqliteConnection, T> operation)
    {
        ArgumentNullException.ThrowIfNull(operation);

        var databasePath = AppPaths.PurchaseDatabasePath;
        if (!File.Exists(databasePath))
        {
            throw new FileNotFoundException("База закупок не найдена: " + databasePath, databasePath);
        }

        var builder = new SqliteConnectionStringBuilder
        {
            DataSource = databasePath,
            Mode = SqliteOpenMode.ReadWrite,
            Pooling = false,
            DefaultTimeout = BusyTimeoutSeconds,
        };

        using var connection = new SqliteConnection(builder.ConnectionString);
        connection.Open();
        ConfigureConnection(connection);

        BeginImmediate(connection);
        var transactionStarted = true;
        try
        {
            var result = operation(connection);
            ExecuteControlCommand(connection, "COMMIT;");
            transactionStarted = false;
            return result;
        }
        catch
        {
            if (transactionStarted)
            {
                TryRollback(connection);
            }
            throw;
        }
    }

    private static void ConfigureConnection(SqliteConnection connection)
    {
        ExecuteControlCommand(connection, "PRAGMA foreign_keys = ON;");
        ExecuteControlCommand(connection, "PRAGMA busy_timeout = " + (BusyTimeoutSeconds * 1000) + ";");
    }

    private static void BeginImmediate(SqliteConnection connection)
    {
        SqliteException? lastBusyException = null;
        for (var attempt = 0; attempt < BeginAttempts; attempt++)
        {
            try
            {
                ExecuteControlCommand(connection, "BEGIN IMMEDIATE;");
                return;
            }
            catch (SqliteException ex) when (IsBusy(ex))
            {
                lastBusyException = ex;
                Thread.Sleep(150 * (attempt + 1));
            }
        }

        throw new InvalidOperationException(
            "База закупок занята основным приложением. Подождите несколько секунд и повторите действие.",
            lastBusyException);
    }

    private static bool IsBusy(SqliteException exception)
        => exception.SqliteErrorCode is 5 or 6;

    private static void TryRollback(SqliteConnection connection)
    {
        try
        {
            ExecuteControlCommand(connection, "ROLLBACK;");
        }
        catch
        {
            // Первичную ошибку записи нельзя скрывать ошибкой отката.
        }
    }

    private static void ExecuteControlCommand(SqliteConnection connection, string sql)
    {
        using var command = connection.CreateCommand();
        command.CommandText = sql;
        command.ExecuteNonQuery();
    }
}
