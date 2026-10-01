using System.IO;

namespace ProcurementControl.Services;

/// <summary>
/// Интеграция с установленным Outlook через позднее связывание COM
/// (ProgID Outlook.Application, без interop-сборки — важно для portable-сборки).
/// Создаваемые объекты открываются через Display(), чтобы пользователь
/// проверил и отправил/сохранил их сам.
/// </summary>
public static class OutlookService
{
    private static object? _application;

    /// <summary>True, если на компьютере зарегистрирован Outlook.</summary>
    public static bool IsAvailable =>
        Type.GetTypeFromProgID("Outlook.Application") is not null;

    /// <summary>Запущенный (или запускаемый) экземпляр Outlook.Application.</summary>
    private static dynamic Application =>
        _application ??= Activator.CreateInstance(Type.GetTypeFromProgID("Outlook.Application")
            ?? throw new InvalidOperationException(
                "Outlook не найден на этом компьютере. Установите Outlook и попробуйте снова."))
        ?? throw new InvalidOperationException("Не удалось запустить Outlook.");

    /// <summary>Создаёт письмо-черновик и показывает его.</summary>
    public static void CreateMailDraft(string to, string subject, string body)
        => CreateMailDraft(to, subject, body, attachments: null);

    /// <summary>
    /// Создаёт письмо-черновик с вложениями и показывает его. Вложения — полные
    /// пути к существующим файлам. По-прежнему только Display(): письмо уходит,
    /// только если пользователь нажмёт «Отправить» сам.
    /// </summary>
    public static void CreateMailDraft(string to, string subject, string body, IReadOnlyList<string>? attachments)
    {
        var mail = Application.CreateItem(0); // olMailItem
        mail.To = to;
        mail.Subject = subject;
        mail.Body = body;
        if (attachments is not null)
        {
            foreach (var path in attachments.Where(File.Exists))
            {
                mail.Attachments.Add(path);
            }
        }

        mail.Display();
    }

    /// <summary>Ящик, с которого уходят письма поставщикам (настройка «отправка должна быть с почты comp»).</summary>
    public const string SupplierMailbox = "comp@npkexpert.ru";

    /// <summary>Идентификатор встроенной подписи в HTML-письме (cid).</summary>
    private const string SignatureCid = "unirec-signature";

    /// <summary>Подпись лежит рядом с exe (Content в csproj) — Outlook встраивает её в письмо.</summary>
    private static string SignatureImagePath => Path.Combine(AppContext.BaseDirectory, "Assets", "SignatureHarizin.png");

    /// <summary>
    /// Письмо поставщику: HTML-тело с подписью UniREC (картинка через cid) и
    /// отправитель comp@npkexpert.ru. Текст шаблона остаётся обычным текстом;
    /// при недоступной подписи/ящику письмо всё равно создаётся.
    /// </summary>
    public static void CreateSupplierRfqDraft(string to, string subject, string bodyText, IReadOnlyList<string>? attachments)
    {
        var mail = Application.CreateItem(0); // olMailItem
        mail.To = to;
        mail.Subject = subject;
        try { mail.SentOnBehalfOfName = SupplierMailbox; }
        catch { /* нет прав или не Exchange — останется ящик по умолчанию */ }

        var html = "<div style=\"font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#000000\">" +
                   System.Net.WebUtility.HtmlEncode(bodyText).Replace("\r\n", "<br>").Replace("\n", "<br>") +
                   "</div>";
        var signature = SignatureImagePath;
        if (File.Exists(signature))
        {
            var inline = mail.Attachments.Add(signature);
            // PR_ATTACH_CONTENT_ID + PR_ATTACHMENT_HIDDEN: картинка внутри письма, не в списке вложений.
            inline.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", SignatureCid);
            inline.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x7FFE000B", true);
            html += "<img src=\"cid:" + SignatureCid + "\" width=\"677\" height=\"250\">";
        }

        mail.HTMLBody = html;

        if (attachments is not null)
        {
            foreach (var path in attachments.Where(File.Exists))
            {
                mail.Attachments.Add(path);
            }
        }

        // Получатель-«простой SMTP» (например, mail.ru) отбивает письмо, если
        // From не совпадает с ящиком, под которым выполнена отправка. Поэтому
        // готовый черновик кладётся в «Черновики» самого ящика comp — отправка
        // из них идёт через этот аккаунт, а не через ящик по умолчанию.
        var moved = TryMoveToMailboxDrafts(mail);
        (moved ?? mail).Display();
    }

    /// <summary>
    /// Сохраняет письмо и переносит его в «Черновики» ящика SupplierMailbox.
    /// SendUsingAccount из внешнего процесса Outlook игнорирует (проверено),
    /// поэтому единственный надёжный путь — черновик в хранилище самого ящика.
    /// Возвращает перемещённый элемент или null, если перенос не удался.
    /// </summary>
    private static dynamic? TryMoveToMailboxDrafts(dynamic mail)
    {
        try
        {
            mail.Save();
            dynamic ns = Application.GetNamespace("MAPI");
            foreach (dynamic store in ns.Stores)
            {
                string? name = null;
                try { name = (string)store.DisplayName; } catch { /* свойство недоступно */ }
                if (!string.Equals(name, SupplierMailbox, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                dynamic drafts = store.GetDefaultFolder(16); // olFolderDrafts
                return mail.Move(drafts);
            }
        }
        catch
        {
            // Перенос не удался — черновик останется в стандартных «Черновиках».
        }
        return null;
    }

    /// <summary>Создаёт задачу Outlook и показывает её. Приоритет: Срочно/1 — высокая, 4/5 — низкая.</summary>
    public static void CreateTask(string subject, string body, DateTime? dueDate, string? priority = null)
    {
        var task = Application.CreateItem(3); // olTaskItem
        task.Subject = subject;
        task.Body = body;
        if (dueDate is { } due)
        {
            task.DueDate = due;
        }

        switch (priority)
        {
            case "Срочно":
            case "1":
            case "2":
                task.Importance = 2; // olImportanceHigh
                break;
            case "4":
            case "5":
                task.Importance = 0; // olImportanceLow
                break;
        }

        task.Display();
    }

    /// <summary>Создаёт встречу-напоминание в календаре на 9:00 указанного дня и показывает её.</summary>
    public static void CreateCalendarReminder(string subject, string body, DateTime date)
    {
        var start = date.Date.AddHours(9);
        var appointment = Application.CreateItem(1); // olAppointmentItem
        appointment.Subject = subject;
        appointment.Body = body;
        appointment.Start = start;
        appointment.End = start.AddMinutes(30);
        appointment.ReminderSet = true;
        appointment.Display();
    }
}
