using ProcurementControl.Models;

namespace ProcurementControl.Services;

/// <summary>
/// Доступ к таблице bitrix_blocked_tasks — порт Get-BitrixBlockedTaskIds и
/// Block-BitrixTask (app/modules/PurchaseStore.ps1, строки 664-677).
/// Чтение идёт через снимок базы, запись — через PurchaseWriteSession.
/// </summary>
public static class BitrixRepository
{
    /// <summary>Аналог Get-BitrixBlockedTaskIds.</summary>
    public static HashSet<int> GetBlockedTaskIds()
    {
        using var snapshot = PurchaseSnapshot.Create();
        var result = new HashSet<int>();
        using var command = snapshot.Connection.CreateCommand();
        command.CommandText = "SELECT bitrix_task_id FROM bitrix_blocked_tasks ORDER BY bitrix_task_id";
        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            result.Add(Convert.ToInt32(reader.GetValue(0)));
        }
        return result;
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

    /// <summary>Аналог Block-BitrixTask: INSERT OR IGNORE в таблицу блокировок.</summary>
    public static void BlockTask(int taskId, string title)
    {
        if (taskId <= 0)
        {
            return;
        }

        PurchaseWriteSession.Execute(connection =>
        {
            using var command = connection.CreateCommand();
            command.CommandText = @"
INSERT OR IGNORE INTO bitrix_blocked_tasks(bitrix_task_id, title, blocked_at)
VALUES(@task_id, @title, @blocked_at);";
            command.Parameters.AddWithValue("@task_id", taskId);
            command.Parameters.AddWithValue("@title", title ?? string.Empty);
            command.Parameters.AddWithValue("@blocked_at", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            command.ExecuteNonQuery();
            return 0;
        });
    }

    /// <summary>Аналог ConvertTo-BitrixInt.</summary>
    private static int ParseInt(string value, int defaultValue)
        => int.TryParse(value.Trim(), out var parsed) ? parsed : defaultValue;
}
