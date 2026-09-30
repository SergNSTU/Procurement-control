using System.IO;
using System.Windows;
using ProcurementControl.Services;

namespace ProcurementControl.Views;

/// <summary>
/// Выбор вложения задачи Bitrix, которое пойдёт в конвейер RFQ. RFQ-исходник —
/// один файл, поэтому выбор одиночный; по умолчанию выделяется первый Excel.
/// </summary>
public partial class RfqBitrixFilesWindow : Window
{
    public RfqBitrixFilesWindow(string dealNumber, IReadOnlyList<BitrixService.BitrixTaskFile> files)
    {
        InitializeComponent();
        HeaderText.Text = "Сделка " + dealNumber + " — вложения задачи Bitrix (" + files.Count + "). Выберите файл заявки:";
        FilesList.ItemsSource = files
            .Select(f => new BitrixFileItem(f))
            .ToList();

        var firstExcel = ((List<BitrixFileItem>)FilesList.ItemsSource)
            .Find(item => item.IsExcel);
        if (firstExcel is not null)
        {
            FilesList.SelectedItem = firstExcel;
        }

        Loaded += (_, _) => FilesList.Focus();
    }

    public BitrixService.BitrixTaskFile? SelectedFile
        => FilesList.SelectedItem is BitrixFileItem item ? item.File : null;

    private void DownloadClick(object sender, RoutedEventArgs e)
    {
        if (SelectedFile is null)
        {
            MessageBox.Show(this, "Выберите файл заявки.", "RFQ из Bitrix");
            return;
        }

        DialogResult = true;
    }

    private sealed class BitrixFileItem
    {
        public BitrixFileItem(BitrixService.BitrixTaskFile file)
        {
            File = file;
            IsExcel = Path.GetExtension(file.Name) is ".xlsx" or ".xlsm" or ".xls";
        }

        public BitrixService.BitrixTaskFile File { get; }

        public bool IsExcel { get; }

        public string DisplayText
        {
            get
            {
                var name = string.IsNullOrWhiteSpace(File.Name) ? "Без имени (id " + File.AttachedObjectId + ")" : File.Name;
                var size = File.Size is { } bytes ? " · " + Math.Max(1, bytes / 1024) + " КБ" : string.Empty;
                return name + size + "  ·  " + File.Source;
            }
        }
    }
}
