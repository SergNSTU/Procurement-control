using System.Windows;

namespace ProcurementControl.Views;

/// <summary>
/// Аналог Show-TaskCompletionDialog (RRFQComparer.ps1, строки 649-702):
/// выбор статуса и этапа при закрытии задачи из карточки уведомления.
/// Возвращает (Статус, Этап) через ShowDialog(); при отмене — false.
/// </summary>
public partial class TaskCompletionWindow : Window
{
    /// <summary>Пункты комбобокса статуса (строка 674 оригинала).</summary>
    public IReadOnlyList<string> Statuses { get; } = new[]
    {
        "В работе", "Ожидание ответа", "На отслеживании", "Заказано",
        "Подано в оплату", "Выполнено", "Не актуально",
    };

    /// <summary>Пункты комбобокса этапа (строка 685 оригинала).</summary>
    public IReadOnlyList<string> Stages { get; } = new[]
    {
        "", "Запросил поставщиков", "RRFQ отправлено", "PI отправлен",
        "Контроль оплаты", "Отправлено в РФ",
    };

    public string Status { get; set; } = string.Empty;
    public string Stage { get; set; } = string.Empty;

    public TaskCompletionWindow(string currentStatus, string currentStage)
    {
        InitializeComponent();
        DataContext = this;

        // Текущее значение, если есть в списке; иначе первый пункт.
        Status = Statuses.Contains(currentStatus) ? currentStatus : Statuses[0];
        Stage = Stages.Contains(currentStage) ? currentStage : Stages[0];
    }

    private void SaveClicked(object sender, RoutedEventArgs e)
    {
        DialogResult = true;
    }

    private void CancelClicked(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
    }
}
