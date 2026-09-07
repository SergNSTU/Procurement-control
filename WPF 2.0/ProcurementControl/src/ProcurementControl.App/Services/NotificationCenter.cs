using System.Windows;
using Microsoft.Data.Sqlite;
using ProcurementControl.Models;
using ProcurementControl.Views;

namespace ProcurementControl.Services;

/// <summary>
/// Очередь карточек уведомлений — перенос связки $notificationTimer,
/// Check-Notifications, Show-NextNotification и
/// Invoke-ActiveNotificationAction (RRFQComparer.ps1, строки 775-978).
/// Проверка раз в 60 секунд из главного окна; карточка показывается одна.
/// </summary>
public static class NotificationCenter
{
    /// <summary>Аналог $script:NotificationsEnabled (настройка notifications.enabled).</summary>
    public static bool Enabled { get; set; } = true;

    private static bool _checking;
    private static readonly HashSet<string> Keys = new(StringComparer.Ordinal);
    private static readonly Queue<DueNotification> Queue = new();
    private static DueNotification? _activeNotification;
    private static NotificationCardWindow? _activeCard;

    /// <summary>Нужно открыть цель уведомления на соответствующей странице.</summary>
    public static event Action<DueNotification>? OpenTargetRequested;

    /// <summary>Данные изменены действием карточки — страницы стоит обновить.</summary>
    public static event Action? DataChanged;

    /// <summary>
    /// Аналог Check-Notifications: если карточка уже показана или очередь не
    /// пуста — только показать следующую; иначе набрать новых кандидатов.
    /// </summary>
    public static void Check()
    {
        if (!Enabled || _checking)
        {
            return;
        }

        if (_activeCard is not null || Queue.Count > 0)
        {
            ShowNext();
            return;
        }

        _checking = true;
        try
        {
            foreach (var item in NotificationRepository.GetDueCandidates())
            {
                if (Keys.Add(GetKey(item)))
                {
                    Queue.Enqueue(item);
                }
            }

            ShowNext();
        }
        catch
        {
            // Ошибка чтения не должна ронять таймер — как пустой catch оригинала.
        }
        finally
        {
            _checking = false;
        }
    }

    /// <summary>Закрыть активную карточку при выходе из приложения.</summary>
    public static void Shutdown()
        => CloseCard();

    /// <summary>
    /// Аналог Invoke-ActiveNotificationAction (строки 775-825): обработка
    /// кнопок карточки. 'open' карточку не закрывает; 'close' помечает
    /// обработанным; 'later' откладывает на час; 'done' закрывает срок.
    /// </summary>
    public static void InvokeAction(string action)
    {
        var notification = _activeNotification;
        if (notification is null)
        {
            return;
        }

        try
        {
            if (action == "open")
            {
                OpenTargetRequested?.Invoke(notification);
                return;
            }

            if (action == "close")
            {
                NotificationRepository.MarkHandled(notification);
                Keys.Remove(GetKey(notification));
                CloseCard();
                ShowNext();
                return;
            }

            if (action == "later")
            {
                NotificationRepository.Snooze(notification, DateTime.Now.AddHours(1));
            }
            else if (notification.Source == "reminder")
            {
                PurchaseWriteRepository.SetReminderDone(notification.SourceId, done: true);
                NotificationRepository.MarkHandled(notification);
                DataChanged?.Invoke();
            }
            else if (notification.Source == "deal")
            {
                PurchaseWriteRepository.ClearDealReminder(notification.SourceId);
                NotificationRepository.MarkHandled(notification);
                DataChanged?.Invoke();
            }
            else
            {
                if (!CompleteTask(notification))
                {
                    // Отмена в диалоге завершения — карточка остаётся.
                    return;
                }
            }

            Keys.Remove(GetKey(notification));
            CloseCard();
            ShowNext();
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Уведомления");
        }
    }

    /// <summary>
    /// Аналог Show-NextNotification: взять из очереди, пометить даты
    /// показанными и открыть карточку.
    /// </summary>
    private static void ShowNext()
    {
        if (!Enabled || _activeCard is not null || Queue.Count == 0)
        {
            return;
        }

        var notification = Queue.Dequeue();
        _activeNotification = notification;
        foreach (var due in notification.Due)
        {
            PurchaseWriteRepository.SetNotificationState(
                notification.Source, notification.SourceId, due.Kind, due.DateText,
                handled: false, snoozeUntil: null, shown: true);
        }

        // Без владельца: карточка не должна сворачиваться вместе с главным окном.
        var card = new NotificationCardWindow(notification);
        _activeCard = card;
        card.Closed += (_, _) =>
        {
            if (ReferenceEquals(_activeCard, card))
            {
                _activeCard = null;
                _activeNotification = null;
            }
        };
        card.Show();
        card.Activate();
    }

    private static void CloseCard()
    {
        var card = _activeCard;
        _activeCard = null;
        _activeNotification = null;
        card?.Close();
    }

    /// <summary>Аналог Get-NotificationKey: «источник:ид:вид1:дата1,вид2:дата2».</summary>
    private static string GetKey(DueNotification notification)
        => notification.Source + ":" + notification.SourceId + ":" +
           string.Join(",", notification.Due.Select(d => d.Kind + ":" + d.DateText));

    /// <summary>
    /// Ветка «Выполнено» для задачи (строки 803-819 оригинала): диалог
    /// завершения, очистка сработавших дат и сохранение статуса/этапа.
    /// Возвращает false при отмене — тогда карточка не закрывается.
    /// </summary>
    private static bool CompleteTask(DueNotification notification)
    {
        string? status = null, stage = null;
        using (var snapshot = PurchaseSnapshot.Create())
        using (var command = snapshot.Connection.CreateCommand())
        {
            command.CommandText = "SELECT IFNULL(status, ''), IFNULL(stage, '') FROM component_deals WHERE id = @id";
            command.Parameters.AddWithValue("@id", notification.SourceId);
            using var reader = command.ExecuteReader();
            if (!reader.Read())
            {
                // Задача уже удалена — просто помечаем обработанным.
                NotificationRepository.MarkHandled(notification);
                DataChanged?.Invoke();
                return true;
            }

            status = reader.GetString(0);
            stage = reader.GetString(1);
        }

        var dialog = new TaskCompletionWindow(status, stage)
        {
            Owner = Application.Current.MainWindow,
        };
        if (dialog.ShowDialog() != true)
        {
            return false;
        }

        var kinds = notification.Due.Select(d => d.Kind).ToList();
        PurchaseWriteRepository.UpdateComponentDealCompletion(
            notification.SourceId,
            dialog.Status,
            dialog.Stage,
            clearReminderDate: kinds.Contains("reminder"),
            clearDeadlineDate: kinds.Contains("deadline"));
        NotificationRepository.MarkHandled(notification);
        DataChanged?.Invoke();
        return true;
    }
}
