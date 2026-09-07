using ProcurementControl.Models;

namespace ProcurementControl.Services;

/// <summary>
/// Чтение кандидатов на уведомление и запись их состояния — перенос
/// Get-DueNotificationCandidates, Mark-NotificationHandled и
/// Snooze-Notification (PurchaseStore.ps1, строки 1065-1201).
/// </summary>
public static class NotificationRepository
{
    /// <summary>
    /// Аналог Get-DueNotificationCandidates: задачи с напоминанием/дедлайном,
    /// напоминания сделок и ручные напоминания, просроченные на сегодня и не
    /// закрытые состоянием (не обработаны, отсрочка истекла).
    /// </summary>
    public static List<DueNotification> GetDueCandidates()
    {
        using var snapshot = PurchaseSnapshot.Create();
        var connection = snapshot.Connection;

        var today = DateTime.Today;
        var now = DateTime.Now;
        var items = new List<DueNotification>();

        // Состояния уведомлений: по ключу «источник|ид|вид|дата».
        var stateByKey = new Dictionary<string, (bool Handled, string SnoozeUntil, string LastShownAt)>(StringComparer.Ordinal);
        using (var states = connection.CreateCommand())
        {
            states.CommandText = @"
SELECT source, source_id, due_kind, due_date, handled, IFNULL(snooze_until, '') AS snooze_until
FROM notification_state";
            using var reader = states.ExecuteReader();
            while (reader.Read())
            {
                var key = string.Join("|",
                    Text(reader, 0),
                    ToLong(reader, 1).ToString(),
                    Text(reader, 2),
                    Text(reader, 3));
                stateByKey[key] = (ToLong(reader, 4) != 0, Text(reader, 5), Text(reader, 6));
            }
        }

        // Задачи: статусы, кроме «Выполнено» и «Не актуально».
        using (var components = connection.CreateCommand())
        {
            components.CommandText = @"
SELECT id, IFNULL(deal_number, '') AS deal_number,
       IFNULL(description, '') AS description,
       IFNULL(next_action, '') AS next_action,
       IFNULL(status, '') AS status,
       IFNULL(stage, '') AS stage,
       IFNULL(reminder_date, '') AS reminder_date,
       IFNULL(deadline_date, '') AS deadline_date
FROM component_deals
WHERE IFNULL(status, '') NOT IN ('Выполнено', 'Не актуально')
  AND (IFNULL(reminder_date, '') <> '' OR IFNULL(deadline_date, '') <> '')
ORDER BY id DESC";
            using var reader = components.ExecuteReader();
            while (reader.Read())
            {
                var id = ToLong(reader, 0);
                var due = new List<DueEntry>();
                foreach (var (kind, label, text) in new[]
                {
                    ("reminder", "Напоминание", Text(reader, 6)),
                    ("deadline", "Дедлайн", Text(reader, 7)),
                })
                {
                    var entry = TryDueEntry("component", id, kind, label, text, stateByKey, today, now);
                    if (entry is not null)
                    {
                        due.Add(entry);
                    }
                }

                if (due.Count == 0)
                {
                    continue;
                }

                var dealNumber = Text(reader, 1).Trim();
                var description = Text(reader, 2).Trim();
                var nextAction = Text(reader, 3).Trim();
                var title = dealNumber;
                if (string.IsNullOrWhiteSpace(title))
                {
                    title = description;
                }
                if (string.IsNullOrWhiteSpace(title))
                {
                    title = "Задача #" + id;
                }

                items.Add(new DueNotification
                {
                    Source = "component",
                    SourceId = id,
                    Title = title,
                    Description = description,
                    NextAction = nextAction,
                    Detail = nextAction,
                    Status = Text(reader, 4),
                    Stage = Text(reader, 5),
                }.WithDue(due));
            }
        }

        // Напоминания сделок (не архив, активные статусы).
        using (var deals = connection.CreateCommand())
        {
            deals.CommandText = @"
SELECT id, IFNULL(deal_number, '') AS deal_number,
       IFNULL(comment, '') AS comment,
       IFNULL(reminder_date, '') AS reminder_date
FROM deals
WHERE IFNULL(archived, 0) = 0
  AND IFNULL(status, '') NOT IN ('Выполнено', 'Не актуально')
  AND IFNULL(reminder_date, '') <> ''
ORDER BY reminder_date, id";
            using var reader = deals.ExecuteReader();
            while (reader.Read())
            {
                var id = ToLong(reader, 0);
                var entry = TryDueEntry("deal", id, "reminder", "Напоминание", Text(reader, 3), stateByKey, today, now);
                if (entry is null)
                {
                    continue;
                }

                items.Add(new DueNotification
                {
                    Source = "deal",
                    SourceId = id,
                    Title = Text(reader, 1),
                    Detail = Text(reader, 2).Trim(),
                    DealId = id,
                    SupplierId = 0,
                }.WithDue(entry));
            }
        }

        // Ручные напоминания (статус не Done).
        using (var reminders = connection.CreateCommand())
        {
            reminders.CommandText = @"
SELECT id, IFNULL(title, '') AS title, IFNULL(due_date, '') AS due_date,
       IFNULL(deal_id, 0) AS deal_id, IFNULL(supplier_id, 0) AS supplier_id
FROM reminders
WHERE status <> 'Done' AND IFNULL(due_date, '') <> ''
ORDER BY due_date, id";
            using var reader = reminders.ExecuteReader();
            while (reader.Read())
            {
                var id = ToLong(reader, 0);
                var entry = TryDueEntry("reminder", id, "manual", "Напоминание", Text(reader, 2), stateByKey, today, now);
                if (entry is null)
                {
                    continue;
                }

                items.Add(new DueNotification
                {
                    Source = "reminder",
                    SourceId = id,
                    Title = Text(reader, 1),
                    Detail = string.Empty,
                    DealId = ToLong(reader, 3),
                    SupplierId = ToLong(reader, 4),
                }.WithDue(entry));
            }
        }

        return items;
    }

    /// <summary>
    /// Аналог Mark-NotificationHandled: пометить все даты уведомления обработанными.
    /// </summary>
    public static void MarkHandled(DueNotification notification)
    {
        foreach (var due in notification.Due)
        {
            PurchaseWriteRepository.SetNotificationState(
                notification.Source, notification.SourceId, due.Kind, due.DateText,
                handled: true, snoozeUntil: null, shown: true);
        }
    }

    /// <summary>
    /// Аналог Snooze-Notification: отложить все даты уведомления до момента.
    /// </summary>
    public static void Snooze(DueNotification notification, DateTime until)
    {
        var untilText = until.ToString("yyyy-MM-dd HH:mm:ss");
        foreach (var due in notification.Due)
        {
            PurchaseWriteRepository.SetNotificationState(
                notification.Source, notification.SourceId, due.Kind, due.DateText,
                handled: false, snoozeUntil: untilText, shown: true);
        }
    }

    /// <summary>
    /// Проверка даты по правилам оригинала: дата распознана, не позже
    /// сегодняшнего дня, состояние отсутствует или ещё не закрыто
    /// (не обработано и отсрочка истекла).
    /// </summary>
    private static DueEntry? TryDueEntry(
        string source,
        long sourceId,
        string kind,
        string label,
        string dateText,
        IReadOnlyDictionary<string, (bool Handled, string SnoozeUntil, string LastShownAt)> stateByKey,
        DateTime today,
        DateTime now)
    {
        var date = PurchaseFormatting.ParseDate(dateText);
        if (date is null || date.Value.Date > today)
        {
            return null;
        }

        var normalized = date.Value.ToString("yyyy-MM-dd");
        var stateKey = $"{source}|{sourceId}|{kind}|{normalized}";
        var eligible = true;
        if (stateByKey.TryGetValue(stateKey, out var state))
        {
            eligible = !state.Handled;

            // Ручное напоминание остаётся открытым, пока пользователь не нажмёт
            // «Выполнено». Крестик карточки не должен прятать его навсегда:
            // повторно показываем такое напоминание на следующий день. Это также
            // возвращает к жизни старые записи, закрытые до появления этого правила.
            if (!eligible && source == "reminder" &&
                (!DateTime.TryParse(state.LastShownAt, out var lastShown) || lastShown.Date < today))
            {
                eligible = true;
            }

            if (eligible && !string.IsNullOrWhiteSpace(state.SnoozeUntil)
                && DateTime.TryParse(state.SnoozeUntil, out var snooze))
            {
                eligible = snooze <= now;
            }
            else if (eligible && !string.IsNullOrWhiteSpace(state.SnoozeUntil))
            {
                // Отсрочка не распознана — оригинал считал бы её истёкшей.
                eligible = true;
            }
        }

        if (!eligible)
        {
            return null;
        }

        return new DueEntry
        {
            Kind = kind,
            Label = label,
            Date = date.Value,
            DateText = normalized,
        };
    }

    private static string Text(Microsoft.Data.Sqlite.SqliteDataReader reader, int ordinal)
        => reader.IsDBNull(ordinal) ? string.Empty : Convert.ToString(reader.GetValue(ordinal)) ?? string.Empty;

    private static long ToLong(Microsoft.Data.Sqlite.SqliteDataReader reader, int ordinal)
        => reader.IsDBNull(ordinal) ? 0 : Convert.ToInt64(reader.GetValue(ordinal));
}

internal static class DueNotificationExtensions
{
    public static DueNotification WithDue(this DueNotification notification, params DueEntry[] entries)
    {
        foreach (var entry in entries)
        {
            notification.Due.Add(entry);
        }
        return notification;
    }

    public static DueNotification WithDue(this DueNotification notification, IEnumerable<DueEntry> entries)
    {
        foreach (var entry in entries)
        {
            notification.Due.Add(entry);
        }
        return notification;
    }
}
