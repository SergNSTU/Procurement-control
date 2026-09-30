using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ProcurementControl.Models;
using ProcurementControl.Services;

namespace ProcurementControl.ViewModels;

/// <summary>
/// Диалог «Ручной ввод квоты» — перенос Show-ManualQuoteDialog и
/// Add-ManualQuoteToDecision (RrfqUi.ps1, строки 142-167 и 239-392).
/// Сохранённая квота сразу становится победителем выбранной строки.
/// </summary>
public partial class ManualQuoteViewModel : ObservableObject
{
    private readonly RrfqAnalysis _analysis;
    private readonly Decision _decision;

    public ManualQuoteViewModel(RrfqAnalysis analysis, Decision decision)
    {
        _analysis = analysis;
        _decision = decision;

        // Префилл из выбранной строки (строки 317-329 оригинала).
        MfgRussia = decision.MfgRussia;
        MfgChina = decision.MfgChina;
        QtyPacking = decision.QtyPacking;
        QtyToBuy = decision.QtyToBuy;
        ChinaRemark = decision.ChinaRemark;
        InfoText = $"Строка: {decision.Row}\r\nValue: {decision.Value}\r\n"
            + "После сохранения эта квота станет победителем для выбранной строки.";
    }

    /// <summary>Список поставщиков для комбобокса (аналог Get-SupplierNames).</summary>
    public IReadOnlyList<string> SupplierNames { get; } = RrfqConfigService.GetSupplierNames();

    /// <summary>Результат диалога: квота добавлена (аналог $form.Tag).</summary>
    public bool Saved { get; private set; }

    /// <summary>Вывод ошибок: по умолчанию MessageBox; подменяется в проверках.</summary>
    public Action<string, string> ShowError { get; set; } = (message, title) => MessageBox.Show(message, title);

    [ObservableProperty]
    private string _supplierText = RrfqConfigService.Current.OtherSupplier;

    [ObservableProperty]
    private string _price = string.Empty;

    [ObservableProperty]
    private string _leadTime = string.Empty;

    [ObservableProperty]
    private string _PN = string.Empty;

    [ObservableProperty]
    private string _DC = string.Empty;

    // Префилл из выбранной строки — заполняется в конструкторе.
    public string MfgRussia { get; set; } = string.Empty;

    public string MfgChina { get; set; } = string.Empty;

    public string QtyPacking { get; set; } = string.Empty;

    public string QtyToBuy { get; set; } = string.Empty;

    public string ChinaRemark { get; set; } = string.Empty;

    /// <summary>Инфо-бокс диалога (аналог $info).</summary>
    public string InfoText { get; }

    /// <summary>Аналог $btnOk: валидация цены и добавление квоты-победителя.</summary>
    [RelayCommand]
    private void Save()
    {
        try
        {
            if (string.IsNullOrWhiteSpace(Price))
            {
                throw new InvalidOperationException("Введите цену.");
            }

            var quote = RrfqEngine.NewManualQuote(
                _decision,
                SupplierText,
                ChinaRemark,
                PN,
                DC,
                MfgRussia,
                MfgChina,
                QtyPacking,
                QtyToBuy,
                Price,
                LeadTime);

            _analysis.Quotes.Add(quote);
            _decision.RfqRow.ManualWinnerId = quote.Id;
            RrfqEngine.RebuildDecisionsAfterManualMatch(_analysis);

            Saved = true;
            CloseRequested?.Invoke(this, EventArgs.Empty);
        }
        catch (Exception ex)
        {
            ShowError(ex.Message, "Ошибка ручной квоты");
        }
    }

    /// <summary>Сигнал окну закрыться (аналог $form.Close()).</summary>
    public event EventHandler? CloseRequested;
}
