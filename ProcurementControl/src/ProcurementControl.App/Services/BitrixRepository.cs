using Microsoft.Data.Sqlite;
using ProcurementControl.Models;

namespace ProcurementControl.Services;

/// <summary>
/// Доступ к таблице bitrix_task_links — связь импортированной сделки
/// «Контроля закупки» с задачей проекта Bitrix. Таблица существовала и в
/// оригинале (PurchaseStore.ps1); блокировки задач больше не нужны.
/// Чтение идёт через снимок базы, запись — через PurchaseWriteSession.
/// </summary>
public static class BitrixRepository
{
    /// <summary>ID задач, уже связанных со сделками (повторный импорт не нужен).</summary>
    public static HashSet<int> GetLinkedTaskIds()
    {
        using var snapshot = PurchaseSnapshot.Create();
        var result = new HashSet<int>();
        using var command = snapshot.Connection.CreateCommand();
        command.CommandText = "SELECT bitrix_task_id FROM bitrix_task_links ORDER BY bitrix_task_id";
        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            result.Add(Convert.ToInt32(reader.GetValue(0)));
        }
        return result;
    }

    /// <summary>Задача Bitrix, связанная со сделкой; 0 — связи нет.</summary>
    public static int GetTaskIdForDeal(long dealId)
    {
        if (dealId <= 0)
        {
            return 0;
        }

        using var snapshot = PurchaseSnapshot.Create();
        using var command = snapshot.Connection.CreateCommand();
        command.CommandText = @"
SELECT bitrix_task_id FROM bitrix_task_links
WHERE deal_id = @deal_id
ORDER BY bitrix_task_id DESC
LIMIT 1;";
        PurchaseWriteRepository.Add(command, "@deal_id", dealId);
        var value = command.ExecuteScalar();
        return value is null || value is DBNull ? 0 : Convert.ToInt32(value);
    }

    /// <summary>
    /// Связывает сделку с задачей Bitrix (повторная вставка игнорируется).
    /// source: 'import' — автоимпорт, 'rfq_pipeline' — конвейер RFQ, 'linked' —
    /// прочее; вкладка «Bitrix API Test» показывает только импортные связи.
    /// Живая база старше схемы приложения: у bitrix_task_links бывают
    /// дополнительные NOT NULL-колонки (например, updated_at), поэтому список
    /// колонок берётся из PRAGMA и все обязательные заполняются.
    /// </summary>
    public static void LinkTask(long dealId, int taskId, string source = "linked")
    {
        if (dealId <= 0 || taskId <= 0)
        {
            return;
        }

        PurchaseWriteSession.Execute(connection =>
        {
            var now = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            var columns = new List<string>();
            var values = new List<string>();
            using (var info = connection.CreateCommand())
            {
                // notnull — служебное слово SQLite (оператор NOTNULL), поэтому в кавычках.
                info.CommandText = "SELECT name, type, \"notnull\", dflt_value FROM pragma_table_info('bitrix_task_links');";
                using var reader = info.ExecuteReader();
                while (reader.Read())
                {
                    var name = reader.GetString(0);
                    var type = reader.GetString(1);
                    var notNull = reader.GetInt64(2) == 1;
                    var hasDefault = !reader.IsDBNull(3);
                    if (name == "id" || (notNull && hasDefault))
                    {
                        continue;
                    }

                    var value = name switch
                    {
                        "deal_id" => "@deal_id",
                        "bitrix_task_id" => "@task_id",
                        "created_at" or "updated_at" => "@now",
                        "source" => "@source",
                        _ => notNull ? DefaultValueLiteral(type) : null,
                    };
                    if (value is null)
                    {
                        continue;
                    }

                    columns.Add(name);
                    values.Add(value);
                }
            }

            if (columns.Count == 0)
            {
                throw new InvalidOperationException("Таблица bitrix_task_links не найдена или пуста.");
            }

            using var command = connection.CreateCommand();
            command.CommandText = $@"
INSERT INTO bitrix_task_links({string.Join(", ", columns)})
SELECT {string.Join(", ", values)}
WHERE NOT EXISTS (SELECT 1 FROM bitrix_task_links WHERE bitrix_task_id = @task_id);";
            PurchaseWriteRepository.Add(command, "@deal_id", dealId);
            PurchaseWriteRepository.Add(command, "@task_id", taskId);
            PurchaseWriteRepository.Add(command, "@now", now);
            PurchaseWriteRepository.Add(command, "@source", source);
            command.ExecuteNonQuery();
            return 0;
        });
    }

    /// <summary>Заглушка для незнакомых NOT NULL-колонок без значений по умолчанию.</summary>
    private static string DefaultValueLiteral(string type)
        => type.ToUpperInvariant().Contains("INT") ? "0" : "''";

    /// <summary>Аналог Get-BitrixConfig: чтение ключей bitrix_* из settings.</summary>
    public static BitrixConfig GetConfig()
    {
        using var snapshot = PurchaseSnapshot.Create();
        var settings = new SettingsRepository(snapshot.Connection);
        var baseUrl = settings.GetSetting("bitrix_base_url");
        if (string.IsNullOrWhiteSpace(baseUrl))
        {
            baseUrl = "https://unirec.bitrix24.ru";
        }

        var token = settings.GetSetting("bitrix_webhook_token");
        return new BitrixConfig(
            BitrixService.NormalizeBaseUrl(baseUrl),
            ParseInt(settings.GetSetting("bitrix_webhook_user_id"), 165),
            token.Trim(),
            ParseInt(settings.GetSetting("bitrix_project_id"), 27));
    }

    /// <summary>Аналог ConvertTo-BitrixInt.</summary>
    private static int ParseInt(string value, int defaultValue)
        => int.TryParse(value.Trim(), out var parsed) ? parsed : defaultValue;
}
