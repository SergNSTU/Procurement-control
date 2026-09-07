using System.Collections.ObjectModel;
using System.Globalization;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using ProcurementControl.Models;

namespace ProcurementControl.Views;

/// <summary>Редактирует составляющие итоговой даты «Поступление комплектации» сделки.</summary>
public partial class DealCompletionReceiptWindow : Window
{
    public ObservableCollection<ReceiptRow> Rows { get; }

    public DealCompletionReceiptWindow(IEnumerable<SupplierRow> suppliers)
    {
        Rows = new ObservableCollection<ReceiptRow>(suppliers.Select(supplier => new ReceiptRow(
            supplier.Id,
            supplier.Supplier,
            ParseDate(supplier.ActualReceiptDate))));
        InitializeComponent();
        DataContext = this;
    }

    public IReadOnlyList<SupplierReceiptDateUpdate> Updates => Rows
        .Select(row => new SupplierReceiptDateUpdate(row.SupplierId, row.ReceiptDate))
        .ToList();

    private static DateTime? ParseDate(string value)
        => DateTime.TryParse(value, CultureInfo.GetCultureInfo("ru-RU"), DateTimeStyles.AllowWhiteSpaces, out var date)
            ? date
            : null;

    private void SaveClick(object sender, RoutedEventArgs e) => DialogResult = true;

    private void CancelClick(object sender, RoutedEventArgs e) => DialogResult = false;
}

public sealed partial class ReceiptRow : ObservableObject
{
    public long SupplierId { get; }
    public string Supplier { get; }

    [ObservableProperty]
    private DateTime? _receiptDate;

    public ReceiptRow(long supplierId, string supplier, DateTime? receiptDate)
    {
        SupplierId = supplierId;
        Supplier = supplier;
        _receiptDate = receiptDate;
    }
}
