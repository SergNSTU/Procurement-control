using System.ComponentModel;
using System.IO;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using ProcurementControl.Models;
using ProcurementControl.Services;

namespace ProcurementControl.Views;

/// <summary>Строка рассылки: поставщик сделки с адресом (редактируемым) и галкой черновика.</summary>
public partial class RfqSendRow : ObservableObject
{
    [ObservableProperty]
    private bool _draft = true;

    public long SupplierId { get; init; }

    public string Supplier { get; init; } = string.Empty;

    [ObservableProperty]
    private string _email = string.Empty;
}

/// <summary>
/// Рассылка RFQ поставщикам сделки: по галочке создаётся письмо-черновик
/// (OutlookService, только Display() — отправка остаётся за пользователем)
/// с файлом RFQ во вложении. DialogResult == true, если создан хотя бы один.
/// </summary>
public partial class RfqSendWindow : Window
{
    private const string DefaultBodyTemplate =
        "Hello dear,\r\n" +
        "Please send RRFQ";

    private const string TemplateSettingKey = "rfq.mail_template";

    private readonly string _dealNumber;
    private readonly string _client;
    private readonly IReadOnlyList<string> _attachments;
    private bool _sending;

    public int CreatedDrafts { get; private set; }

    public RfqSendWindow(string dealNumber, string client, IReadOnlyList<SupplierRow> suppliers, IReadOnlyList<string> attachments)
    {
        InitializeComponent();
        _dealNumber = dealNumber;
        _client = client;
        _attachments = attachments;

        HeaderText.Text = "Сделка " + dealNumber + (string.IsNullOrWhiteSpace(client) ? "" : " (" + client + ")");
        AttachmentText.Text = "Вложение: " + string.Join("; ", attachments.Select(Path.GetFileName));

        var rows = suppliers.Select(s => new RfqSendRow
        {
            SupplierId = s.Id,
            Supplier = s.Supplier,
            Email = SupplierDirectory.FindEmailBySupplier(s.Supplier) ?? string.Empty,
        }).ToList();
        SuppliersGrid.ItemsSource = rows;
        foreach (var row in rows)
        {
            row.PropertyChanged += (_, args) =>
            {
                if (args.PropertyName is nameof(RfqSendRow.Draft) or nameof(RfqSendRow.Email))
                {
                    UpdateSendButton();
                }
            };
        }

        BodyBox.Text = ReadBodyTemplate();
        UpdateSendButton();
    }

    private static string ReadBodyTemplate()
    {
        try
        {
            using var snapshot = PurchaseSnapshot.Create();
            var saved = new SettingsRepository(snapshot.Connection).GetSetting(TemplateSettingKey);
            if (!string.IsNullOrWhiteSpace(saved))
            {
                return saved.Replace("\\r\\n", "\r\n");
            }
        }
        catch (Exception ex)
        {
            ErrorLog.Append("settings", ex);
        }

        return DefaultBodyTemplate;
    }

    private void UpdateSendButton()
    {
        if (SuppliersGrid.ItemsSource is not IReadOnlyList<RfqSendRow> rows)
        {
            return;
        }

        SendButton.IsEnabled = !_sending && rows.Any(r => r.Draft && !string.IsNullOrWhiteSpace(r.Email));
    }

    private void SendClick(object sender, RoutedEventArgs e)
    {
        if (_sending)
        {
            return;
        }

        try
        {
            if (SuppliersGrid.ItemsSource is not IReadOnlyList<RfqSendRow> rows)
            {
                return;
            }

            var targets = rows.Where(r => r.Draft).ToList();
            if (targets.Count == 0)
            {
                throw new InvalidOperationException("Не выбран ни один поставщик.");
            }

            var missing = targets.Where(r => string.IsNullOrWhiteSpace(r.Email)).Select(r => r.Supplier).ToList();
            if (missing.Count > 0)
            {
                throw new InvalidOperationException("Не указан email у: " + string.Join(", ", missing) +
                    ".\nСнимите галочку или впишите адрес.");
            }

            if (MessageBox.Show(this, "Создать черновики писем: " + targets.Count + "?\r\n" +
                    "Письма только откроются в Outlook — отправка за вами.",
                    "Разослать RFQ", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
            {
                return;
            }

            _sending = true;
            SendButton.IsEnabled = false;
            ProgressService.Begin("RFQ: создание черновиков писем…");
            try
            {
                var created = 0;
                foreach (var target in targets)
                {
                    var body = BodyBox.Text
                        .Replace("{Deal}", _dealNumber)
                        .Replace("{Client}", string.IsNullOrWhiteSpace(_client) ? "—" : _client)
                        .Replace("{Supplier}", target.Supplier);
                    OutlookService.CreateSupplierRfqDraft(target.Email.Trim(), _dealNumber, body, _attachments);
                    created++;
                }

                CreatedDrafts = created;
                DialogResult = true;
            }
            finally
            {
                ProgressService.End();
                _sending = false;
                UpdateSendButton();
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Разослать RFQ");
        }
    }
}
