using System.Windows;
using ProcurementControl.Models;
using ProcurementControl.Services;

namespace ProcurementControl.Views;

/// <summary>
/// Карточка сработавшего уведомления — перенос формы из Show-NextNotification
/// (RRFQComparer.ps1, строки 827-961): 440 пикселей, без рамки, поверх всех
/// окон, в левом нижнем углу рабочей области. Показывается без владельца,
/// чтобы не сворачиваться вместе с главным окном.
/// </summary>
public partial class NotificationCardWindow : Window
{
    public NotificationCardWindow(DueNotification notification)
    {
        InitializeComponent();

        BodyText.Text = notification.Title;
        DueValue.Text = string.Join(", ", notification.Due.Select(d => d.Date.ToString("dd.MM.yyyy")));

        // Для задач карточка выше и содержит описание и следующее действие.
        if (notification.Source == "component")
        {
            Height = 264;
            DescriptionText.Text = "Описание: " + (string.IsNullOrWhiteSpace(notification.Description) ? "—" : notification.Description);
            NextActionText.Text = "Следующее действие: " + (string.IsNullOrWhiteSpace(notification.NextAction) ? "—" : notification.NextAction);
        }
        else
        {
            Height = 190;
            TaskDetails.Visibility = Visibility.Collapsed;
        }

        // Как $card.Add_Shown: позиция от рабочей области экрана.
        Loaded += (_, _) =>
        {
            var area = SystemParameters.WorkArea;
            Left = area.Left + 24;
            Top = area.Bottom - ActualHeight - 56;
        };
    }

    private void CloseClicked(object sender, RoutedEventArgs e)
        => NotificationCenter.InvokeAction("close");

    private void OpenClicked(object sender, RoutedEventArgs e)
        => NotificationCenter.InvokeAction("open");

    private void LaterClicked(object sender, RoutedEventArgs e)
        => NotificationCenter.InvokeAction("later");

    private void DoneClicked(object sender, RoutedEventArgs e)
        => NotificationCenter.InvokeAction("done");
}
