using System.Windows;

namespace ProcurementControl.Views;

/// <summary>
/// Аналог Show-SimpleInputDialog из оригинала: модальное окно с одним полем
/// ввода (новая сделка, добавление поставщика).
/// </summary>
public partial class InputDialogWindow : Window
{
    public InputDialogWindow(string title, string prompt, string initialValue = "")
    {
        InitializeComponent();
        Title = title;
        PromptText.Text = prompt;
        InputBox.Text = initialValue;
        Loaded += (_, _) =>
        {
            InputBox.Focus();
            InputBox.SelectAll();
        };
    }

    /// <summary>Введённое значение (актуально после DialogResult == true).</summary>
    public string Value => InputBox.Text.Trim();

    private void AcceptClick(object sender, RoutedEventArgs e)
    {
        DialogResult = true;
    }

    private void CancelClick(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
    }
}
