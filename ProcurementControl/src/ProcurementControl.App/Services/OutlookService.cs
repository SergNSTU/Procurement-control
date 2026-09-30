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
    {
        var mail = Application.CreateItem(0); // olMailItem
        mail.To = to;
        mail.Subject = subject;
        mail.Body = body;
        mail.Display();
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
