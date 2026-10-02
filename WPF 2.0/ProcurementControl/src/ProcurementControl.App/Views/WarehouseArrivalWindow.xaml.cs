using System.Windows;
using System.Windows.Input;
using ProcurementControl.Services;

namespace ProcurementControl.Views;

/// <summary>
/// Окно отчёта по файлу прихода со склада: сделанное в этом прогоне
/// (записанные/исправленные «Факты поступл.») и текущие вопросы — частичные
/// приходы, заказы, которых нет в файле. Каждый новый отчёт заменяет
/// содержимое: окно всегда показывает актуальное состояние.
/// По выбранному замечанию создаётся черновик письма поставщику.
/// Создаётся через WarehouseArrivalNotifier.
/// </summary>
public partial class WarehouseArrivalWindow : Window
{
    private readonly List<WarehouseArrivalIssue> _issues = new();

    public WarehouseArrivalWindow()
    {
        InitializeComponent();
    }

    /// <summary>Очищает окно перед загрузкой свежего отчёта.</summary>
    public void ResetReport()
    {
        _issues.Clear();
        FileText.Text = string.Empty;
        SummaryText.Text = string.Empty;
        IssuesList.Items.Clear();
    }

    /// <summary>Заполняет окно отчётом (перед этим вызвать ResetReport).</summary>
    public void AppendReport(WarehouseArrivalReport report)
    {
        if (report.FileSubject is not null)
        {
            FileText.Text = report.FileSubject +
                            (report.SavedPath is null ? string.Empty : "   →   " + report.SavedPath);
        }

        foreach (var line in report.Lines)
        {
            SummaryText.Text += (SummaryText.Text.Length == 0 ? string.Empty : Environment.NewLine) + line;
        }

        _issues.AddRange(report.Issues);
        RebuildIssueList();
    }

    private void RebuildIssueList()
    {
        IssuesList.Items.Clear();
        foreach (var issue in _issues)
        {
            IssuesList.Items.Add(issue.DisplayText);
        }

        var visible = _issues.Count > 0;
        IssuesList.Visibility = visible ? Visibility.Visible : Visibility.Collapsed;
        PushButton.Visibility = visible ? Visibility.Visible : Visibility.Collapsed;
        if (visible)
        {
            IssuesList.SelectedIndex = 0;
        }
    }

    private void PushClicked(object sender, RoutedEventArgs e)
    {
        var index = IssuesList.SelectedIndex;
        if (index < 0 || index >= _issues.Count)
        {
            return;
        }

        try
        {
            WarehouseArrivalImportService.PushSupplier(_issues[index]);
            ToastService.Show("Черновик письма создан — проверьте и отправьте сами", ToastKind.Success);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Пушнуть поставщика");
        }
    }

    private void OkClicked(object sender, RoutedEventArgs e) => Close();

    private void CloseClicked(object sender, RoutedEventArgs e) => Close();

    private void BorderDragMove(object sender, MouseButtonEventArgs e)
    {
        if (e.ButtonState == MouseButtonState.Pressed)
        {
            try
            {
                DragMove();
            }
            catch (InvalidOperationException)
            {
                // DragMove бросает исключение, если кнопка уже отпущена — не критично.
            }
        }
    }
}
