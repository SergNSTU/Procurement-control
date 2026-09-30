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

    /// <summary>Связывает сделку с задачей Bitrix (повторная вставка игнорируется).</summary>
    public static void LinkTask(long dealId, int taskId)
    {
        if (dealId <= 0 || taskId <= 0)
        {
            return;
        }

        PurchaseWriteSession.Execute(connection =>
        {
            using var command = connection.CreateCommand();
            command.CommandText = @"
INSERT INTO bitrix_task_links(deal_id, bitrix_task_id, created_at)
SELECT @deal_id, @task_id, @created_at
WHERE NOT EXISTS (
    SELECT 1 FROM bitrix_task_links WHERE bitrix_task_id = @task_id);";
            PurchaseWriteRepository.Add(command, "@deal_id", dealId);
            PurchaseWriteRepository.Add(command, "@task_id", taskId);
            PurchaseWriteRepository.Add(command, "@created_at", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            command.ExecuteNonQuery();
            return 0;
        });
    }

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
