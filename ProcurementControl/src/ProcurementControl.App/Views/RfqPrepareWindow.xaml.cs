using System.ComponentModel;
using System.IO;
using System.Windows;
using ProcurementControl.Services;

namespace ProcurementControl.Views;

/// <summary>
/// Подготовка поставщицкого RFQ: галочки «В RFQ» (предзаполнены зелёной
/// заливкой исходника) и правка строк. По кнопке «Собрать» строятся два
/// файла — RFQ для поставщиков и список «Для Компэла»; пути возвращаются
/// через SupplierFilePath/CompelFilePath при DialogResult == true.
/// </summary>
public partial class RfqPrepareWindow : Window
{
    private readonly string _outputDirectory;
    private readonly string _dealNumber;
    private bool _building;

    public string SupplierFilePath { get; private set; } = string.Empty;

    public string CompelFilePath { get; private set; } = string.Empty;

    public RfqPrepareWindow(string dealNumber, string client, RfqSourceReadResult source, string outputDirectory)
    {
        InitializeComponent();
        _dealNumber = dealNumber;
        _outputDirectory = outputDirectory;
        HeaderText.Text = string.IsNullOrWhiteSpace(client)
            ? "Сделка " + dealNumber + " — позиций в исходнике: " + source.Rows.Count
            : "Сделка " + dealNumber + " (" + client + ") — позиций в исходнике: " + source.Rows.Count;

        RowsGrid.ItemsSource = source.Rows;
        foreach (var row in source.Rows)
        {
            row.PropertyChanged += RowPropertyChanged;
        }

        if (!source.HasFillInfo)
        {
            HintText.Text += "\nЗелёная заливка не распознана — отметьте позиции вручную.";
        }

        UpdateSummary();
        Loaded += (_, _) => RowsGrid.Focus();
    }

    private void RowPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(RfqSourceRow.Included))
        {
            UpdateSummary();
        }
    }

    private void UpdateSummary()
    {
        if (RowsGrid.ItemsSource is not IReadOnlyList<RfqSourceRow> rows)
        {
            return;
        }

        var included = rows.Count(r => r.Included);
        SummaryText.Text = "В RFQ: " + included + "  ·  Для Компэла: " + (rows.Count - included);
    }

    private void IncludeAllClick(object sender, RoutedEventArgs e) => SetAllIncluded(true);

    private void IncludeNoneClick(object sender, RoutedEventArgs e) => SetAllIncluded(false);

    private void IncludeGreenOnlyClick(object sender, RoutedEventArgs e)
    {
        if (RowsGrid.ItemsSource is not IReadOnlyList<RfqSourceRow> rows)
        {
            return;
        }

        foreach (var row in rows)
        {
            row.Included = row.DetectedGreen;
        }
    }

    private void SetAllIncluded(bool included)
    {
        if (RowsGrid.ItemsSource is not IReadOnlyList<RfqSourceRow> rows)
        {
            return;
        }

        foreach (var row in rows)
        {
            row.Included = included;
        }
    }

    private void BuildClick(object sender, RoutedEventArgs e)
    {
        if (_building)
        {
            return;
        }

        try
        {
            if (RowsGrid.ItemsSource is not IReadOnlyList<RfqSourceRow> rows || rows.Count == 0)
            {
                throw new InvalidOperationException("Список строк пуст.");
            }

            if (rows.Count(r => r.Included) == 0)
            {
                throw new InvalidOperationException("Не отмечено ни одной позиции для поставщиков.");
            }

            _building = true;
            BuildButton.IsEnabled = false;
            ProgressService.Begin("RFQ: сборка файлов для поставщиков…");

            var supplierPath = RfqPrepareService.BuildSupplierFilePath(_outputDirectory, _dealNumber);
            var compelPath = RfqPrepareService.BuildCompelFilePath(_outputDirectory, _dealNumber);
            var worker = new Thread(() =>
            {
                Exception? error = null;
                try
                {
                    RfqPrepareService.WriteSupplierFile(rows, supplierPath);
                    RfqPrepareService.WriteCompelFile(rows, compelPath);
                }
                catch (Exception ex)
                {
                    error = ex;
                }

                Dispatcher.Invoke(() =>
                {
                    ProgressService.End();
                    _building = false;
                    BuildButton.IsEnabled = true;
                    if (error is not null)
                    {
                        ErrorLog.Append("excel", error);
                        MessageBox.Show(this, error.Message, "Сборка RFQ");
                        return;
                    }

                    SupplierFilePath = supplierPath;
                    CompelFilePath = compelPath;
                    DialogResult = true;
                });
            });
            worker.SetApartmentState(ApartmentState.STA);
            worker.IsBackground = true;
            worker.Start();
        }
        catch (Exception ex)
        {
            _building = false;
            BuildButton.IsEnabled = true;
            MessageBox.Show(this, ex.Message, "Сборка RFQ");
        }
    }
}
