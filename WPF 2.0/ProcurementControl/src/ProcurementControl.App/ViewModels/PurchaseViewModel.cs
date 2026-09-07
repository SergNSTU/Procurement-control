using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;
using ProcurementControl.Models;
using ProcurementControl.Services;
using ProcurementControl.Views;

namespace ProcurementControl.ViewModels;

/// <summary>Пункт комбобокса фильтров.</summary>
public sealed record FilterOption(DealFilter Value, string Label)
{
    public override string ToString() => Label;
}

/// <summary>
/// Страница «Контроль закупки». Чтение идёт через снапшот БД (живёт всё время
/// страницы), запись — через PurchaseWriteRepository/PurchaseWriteSession
/// (BEGIN IMMEDIATE на реальной базе). Перенесённые действия: «Новая сделка»,
/// панель редактирования выбранной сделки, «В архив / вернуть» и «Добавить
/// поставщика» (обработчики $btnNewDeal, $btnSaveDealInfo, $btnArchiveDeal,
/// $btnAddPurchaseSupplier из RRFQComparer.ps1).
/// </summary>
public partial class PurchaseViewModel : ObservableObject
{
    /// <summary>Этапы сделки — пункты $cmbDealStatusEdit (строка 2061 оригинала).</summary>
    public IReadOnlyList<string> DealStatuses { get; } =
        new[] { "RFQ", "RRFQ", "PO", "PI", "Закупка", "Заказано", "В работе", "Ожидание" };

    /// <summary>Периоды — пункты $cmbDealPeriodEdit (строка 2072 оригинала).</summary>
    public IReadOnlyList<string> DealPeriods { get; } =
        new[] { "", "I квартал", "II квартал", "III квартал", "IV квартал" };

    public IReadOnlyList<string> DealPriorities { get; } =
        new[] { "", "Срочно", "1", "2", "3", "4", "5" };

    /// <summary>Значения выпадающего списка «Маски», как в исходном приложении.</summary>
    public IReadOnlyList<string> DealMasks { get; } = new[] { "", "Нет", "Да" };

    public IReadOnlyList<string> DealExecutors { get; } =
        new[] { "", "Евгений" };

    public IReadOnlyList<string> AssemblyLocations { get; } =
        new[] { "", "Китай", "РФ" };

    private PurchaseSnapshot? _snapshot;
    private PurchaseRepository? _repository;
    private List<PurchaseDealRow> _allDeals = new();

    [ObservableProperty]
    private string _searchText = string.Empty;

    [ObservableProperty]
    private FilterOption _selectedFilter;

    [ObservableProperty]
    private PurchaseDealRow? _selectedDeal;

    [ObservableProperty]
    private string _errorMessage = string.Empty;

    [ObservableProperty]
    private string _statusText = string.Empty;

    [ObservableProperty]
    private string _countText = string.Empty;

    // Поля панели редактирования выбранной сделки (аналог $dealEditPanel).
    [ObservableProperty]
    private string _editClient = string.Empty;

    [ObservableProperty]
    private string _editStatus = "RFQ";

    [ObservableProperty]
    private string _editPeriod = string.Empty;

    [ObservableProperty]
    private string _editComment = string.Empty;

    [ObservableProperty] private string _editDealNumber = string.Empty;
    [ObservableProperty] private string _editBoardCount = string.Empty;
    [ObservableProperty] private string _editPriority = string.Empty;
    [ObservableProperty] private string _editTrackingStatus = string.Empty;
    [ObservableProperty] private string _editExecutor = string.Empty;
    [ObservableProperty] private string _editReminderDate = string.Empty;
    [ObservableProperty] private string _editAssemblyLocation = string.Empty;

    [ObservableProperty]
    private string _archiveButtonText = "В архив";

    // Панель документов выбранной сделки (аналог $docsPanel/$docsGrid).
    [ObservableProperty]
    private SupplierRow? _selectedSupplier;

    [ObservableProperty]
    private PurchaseDocumentRow? _selectedDocument;

    // Карточки метрик.
    [ObservableProperty] private int _activeCount;
    [ObservableProperty] private int _overdueCount;
    [ObservableProperty] private int _workAndPaymentTaskCount;
    [ObservableProperty] private int _attentionCount;
    [ObservableProperty] private int _reminderCount;

    public IReadOnlyList<FilterOption> Filters { get; } =
        Enum.GetValues<DealFilter>().Select(f => new FilterOption(f, f.Label())).ToList();

    public ObservableCollection<PurchaseDealRow> Deals { get; } = new();

    public ObservableCollection<SupplierRow> Suppliers { get; } = new();

    public ObservableCollection<PurchaseDocumentRow> Documents { get; } = new();

    public PurchaseViewModel()
    {
        _selectedFilter = Filters[0];
        Load();
    }

    partial void OnSearchTextChanged(string value) => ApplyView();

    partial void OnSelectedFilterChanged(FilterOption value) => ApplyView();

    partial void OnSelectedDealChanged(PurchaseDealRow? value)
    {
        FillEditPanel(value);
        SelectedSupplier = null;
        LoadSuppliers();
        LoadDocuments();
    }

    [RelayCommand]
    private void Reload() => Load();

    /// <summary>
    /// Аналог ветки deal из Open-NotificationTarget (RRFQComparer.ps1,
    /// строки 747-763): очистить поиск, фильтр «Все включая архив», выделить сделку.
    /// </summary>
    public void OpenDeal(long id)
    {
        SearchText = string.Empty;
        SelectedFilter = Filters.First(f => f.Value == DealFilter.AllRecords);
        Reload();
        SelectedDeal = Deals.FirstOrDefault(r => r.Id == id);
    }

    /// <summary>Аналог $btnNewDeal: запрос номера и New-PurchaseDeal.</summary>
    [RelayCommand]
    private void NewDeal()
    {
        try
        {
            var dialog = new InputDialogWindow("Новая сделка", "Введите номер сделки / заказа")
            {
                Owner = Application.Current.MainWindow,
            };
            if (dialog.ShowDialog() != true || string.IsNullOrWhiteSpace(dialog.Value))
            {
                return;
            }

            var dealId = PurchaseWriteRepository.CreateOrUpdateDeal(new PurchaseDealDraft
            {
                DealNumber = dialog.Value,
            });
            Load();
            SelectedDeal = Deals.FirstOrDefault(r => r.Id == dealId);
            StatusText = "Сделка создана: " + dialog.Value;
            ToastService.Show(StatusText, ToastKind.Success);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Новая сделка");
        }
    }

    /// <summary>
    /// Аналог $btnSaveDealInfo: Update-PurchaseDeal с полями панели; поля,
    /// которые панель не редактирует (номер, платы, приоритет, отслеживание,
    /// исполнитель, напоминание, место сборки), передаются как null — «не менять».
    /// </summary>
    [RelayCommand]
    private void SaveDeal()
    {
        try
        {
            var deal = SelectedDeal;
            if (deal is null)
            {
                throw new InvalidOperationException("Выберите сделку.");
            }

            PurchaseWriteRepository.UpdateDeal(deal.Id, new PurchaseDealUpdate
            {
                Client = EditClient,
                Status = EditStatus,
                Comment = EditComment,
                Period = EditPeriod,
                DealNumber = EditDealNumber,
                BoardCount = EditBoardCount,
                Priority = EditPriority,
                TrackingStatus = EditTrackingStatus,
                Executor = EditExecutor,
                ReminderDate = EditReminderDate,
                AssemblyLocation = EditAssemblyLocation,
            });
            ApplyEditorValues(deal);
            RefreshEditedDealInPlace(deal);
            StatusText = "Сделка сохранена";
            ToastService.Show(StatusText, ToastKind.Success);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Сохранить сделку");
        }
    }

    /// <summary>Аналог Toggle-SelectedPurchaseDealArchive ($btnArchiveDeal).</summary>
    [RelayCommand]
    private void ToggleArchive()
    {
        try
        {
            var deal = SelectedDeal;
            if (deal is null)
            {
                throw new InvalidOperationException("Выберите сделку.");
            }

            var archived = deal.Archived != 1;
            PurchaseWriteRepository.SetDealArchived(deal.Id, archived);
            Load();
            SelectedDeal = Deals.FirstOrDefault(r => r.Id == deal.Id);
            StatusText = archived ? "Сделка отправлена в архив" : "Сделка возвращена из архива";
            ToastService.Show(StatusText, ToastKind.Info);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Архив");
        }
    }

    /// <summary>Аналог $btnAddPurchaseSupplier: запрос имени и Add-PurchaseSupplier.</summary>
    [RelayCommand]
    private void AddSupplier()
    {
        try
        {
            var deal = SelectedDeal;
            if (deal is null)
            {
                throw new InvalidOperationException("Выберите сделку.");
            }

            var dialog = new InputDialogWindow("Поставщик", "Введите наименование поставщика")
            {
                Owner = Application.Current.MainWindow,
            };
            if (dialog.ShowDialog() != true || string.IsNullOrWhiteSpace(dialog.Value))
            {
                return;
            }

            PurchaseWriteRepository.AddOrTouchSupplier(deal.Id, dialog.Value);
            Load();
            SelectedDeal = Deals.FirstOrDefault(r => r.Id == deal.Id);
            StatusText = "Поставщик добавлен: " + dialog.Value;
            ToastService.Show(StatusText, ToastKind.Success);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Добавить поставщика");
        }
    }

    /// <summary>Перенос удаления сделки в корзину с возможностью восстановления.</summary>
    [RelayCommand]
    private void DeleteDeal()
    {
        try
        {
            var deal = SelectedDeal ?? throw new InvalidOperationException("Выберите сделку.");
            var answer = MessageBox.Show(
                "Переместить сделку в корзину?\n" + deal.DealNumber + "\nДокументы и поставщики можно будет восстановить в Настройках.",
                "Корзина", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (answer != MessageBoxResult.Yes) return;

            PurchaseWriteRepository.DeleteDeal(deal.Id);
            Load();
            StatusText = "Сделка перемещена в корзину";
            ToastService.Show(StatusText, ToastKind.Info);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Удалить сделку");
        }
    }

    [RelayCommand]
    private void CopyDealNumber()
    {
        try
        {
            var deal = SelectedDeal ?? throw new InvalidOperationException("Выберите сделку.");
            Clipboard.SetText(deal.DealNumber);
            StatusText = "Номер сделки скопирован";
            ToastService.Show(StatusText, ToastKind.Info);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Скопировать номер сделки");
        }
    }

    /// <summary>Обновляет только количество плат — обработчик двойного щелчка в таблице.</summary>
    public void UpdateBoardCount(PurchaseDealRow deal, string boardCount)
    {
        try
        {
            PurchaseWriteRepository.UpdateDealBoardCount(deal.Id, boardCount);
            deal.BoardCount = boardCount;
            RefreshEditedDealInPlace(deal);
            StatusText = "Количество плат сохранено";
            ToastService.Show(StatusText, ToastKind.Success);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Кол-во плат");
        }
    }

    /// <summary>Сохраняет дату напоминания из календаря по двойному щелчку в таблице.</summary>
    public void UpdateDealReminderDate(PurchaseDealRow deal, string value)
    {
        try
        {
            PurchaseWriteRepository.UpdateDealReminderDate(deal.Id, value);
            deal.ReminderDate = value;
            RefreshEditedDealInPlace(deal);
            StatusText = "Напоминание сохранено";
            ToastService.Show(StatusText, ToastKind.Success);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Выбор даты");
        }
    }

    /// <summary>Сохраняет значение, изменённое прямо в строке таблицы сделок.</summary>
    public void UpdateDealGridField(PurchaseDealRow deal, string field)
    {
        try
        {
            var value = field switch
            {
                nameof(PurchaseDealRow.DealNumber) => deal.DealNumber,
                nameof(PurchaseDealRow.BoardCount) => deal.BoardCount,
                nameof(PurchaseDealRow.Client) => deal.Client,
                nameof(PurchaseDealRow.Priority) => deal.Priority,
                nameof(PurchaseDealRow.Status) => deal.Status,
                nameof(PurchaseDealRow.Comment) => deal.Comment,
                nameof(PurchaseDealRow.Period) => deal.Period,
                nameof(PurchaseDealRow.Executor) => deal.Executor,
                nameof(PurchaseDealRow.AssemblyLocation) => deal.AssemblyLocation,
                nameof(PurchaseDealRow.TrackingStatus) => deal.TrackingStatus,
                nameof(PurchaseDealRow.MasksText) => deal.MasksText,
                _ => throw new ArgumentException("Поле нельзя изменить из таблицы.", nameof(field)),
            };
            PurchaseWriteRepository.UpdateDealGridField(deal.Id, field, value);
            RefreshEditedDealInPlace(deal);
            StatusText = "Изменения сохранены";
            ToastService.Show(StatusText, ToastKind.Success);
        }
        catch (Exception ex)
        {
            ReloadSelection(deal.Id);
            MessageBox.Show(ex.Message, "Сохранить сделку");
        }
    }

    /// <summary>Сохраняет даты поступления по каждому поставщику выбранной сделки.</summary>
    public void UpdateCompletionReceiptDates(PurchaseDealRow deal, IEnumerable<SupplierReceiptDateUpdate> updates)
    {
        try
        {
            foreach (var update in updates)
                PurchaseWriteRepository.SetSupplierActualReceiptDate(update.SupplierId, update.ActualReceiptDate);
            ReloadSelection(deal.Id);
            StatusText = "Даты поступления сохранены";
            ToastService.Show(StatusText, ToastKind.Success);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Поступление комплектации");
        }
    }

    /// <summary>Открывает полную карточку поставщика и сохраняет её одной операцией.</summary>
    [RelayCommand]
    private void EditSupplier()
    {
        try
        {
            var supplier = SelectedSupplier ?? throw new InvalidOperationException("Выберите поставщика.");
            var dialog = new SupplierEditorWindow(supplier) { Owner = Application.Current.MainWindow };
            if (dialog.ShowDialog() != true) return;
            PurchaseWriteRepository.UpdateSupplier(supplier.Id, dialog.Update);
            ReloadSelection(supplier.DealId);
            StatusText = "Карточка поставщика сохранена";
            ToastService.Show(StatusText, ToastKind.Success);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Сохранить поставщика");
        }
    }

    /// <summary>Сохраняет один из рабочих флажков поставщика из таблицы без открытия карточки.</summary>
    public void UpdateSupplierFlags(SupplierRow supplier)
    {
        try
        {
            PurchaseWriteRepository.UpdateSupplier(supplier.Id, CreateSupplierUpdate(supplier));
            ReloadSelection(supplier.DealId);
            StatusText = "Статус поставщика сохранен";
            ToastService.Show(StatusText, ToastKind.Success);
        }
        catch (Exception ex)
        {
            ReloadSelection(supplier.DealId);
            MessageBox.Show(ex.Message, "Сохранить статус поставщика");
        }
    }

    /// <summary>Сохраняет дату поставщика из календаря таблицы.</summary>
    public void UpdateSupplierDate(SupplierRow supplier, string field, string value)
    {
        try
        {
            if (field == nameof(SupplierRow.InvoiceConfirmedDate)) supplier.InvoiceConfirmedDate = value;
            else if (field == nameof(SupplierRow.ComponentsReceiptDate)) supplier.ComponentsReceiptDate = value;
            else throw new ArgumentException("Поле даты не поддерживается.", nameof(field));

            PurchaseWriteRepository.UpdateSupplier(supplier.Id, CreateSupplierUpdate(supplier));
            ReloadSelection(supplier.DealId);
            StatusText = "Дата поставщика сохранена";
            ToastService.Show(StatusText, ToastKind.Success);
        }
        catch (Exception ex)
        {
            ReloadSelection(supplier.DealId);
            MessageBox.Show(ex.Message, "Сохранить дату поставщика");
        }
    }

    private static PurchaseSupplierUpdate CreateSupplierUpdate(SupplierRow supplier) => new()
    {
        InvoiceReceived = supplier.InvoiceReceived,
        InvoiceConfirmed = supplier.InvoiceConfirmed,
        SupplierOrderCreated = supplier.SupplierOrderCreated,
        ErpSupplierSent = supplier.ErpSupplierSent,
        ErpRogerSent = supplier.ErpRogerSent,
        PiAmountUsd = supplier.PiAmountUsd,
        PiAmountCny = supplier.PiAmountCny,
        PiAmountRub = supplier.PiAmountRub,
        PaidAmount = supplier.PaidAmount,
        DeliveryWeeks = supplier.DeliveryWeeks,
        PaymentSubmitted = supplier.PaymentSubmitted,
        Paid = supplier.Paid,
        InvoiceConfirmedDate = supplier.InvoiceConfirmedDate,
        ComponentsReceiptDate = supplier.ComponentsReceiptDate,
        ActualReceiptDate = supplier.ActualReceiptDate,
        Comment = supplier.Comment ?? string.Empty,
    };

    /// <summary>Удаляет поставщика и связанные с ним документы через корзину.</summary>
    [RelayCommand]
    private void DeleteSupplier()
    {
        try
        {
            var supplier = SelectedSupplier ?? throw new InvalidOperationException("Выберите поставщика.");
            if (MessageBox.Show("Переместить поставщика и его документы в корзину?\n" + supplier.Supplier,
                    "Удаление поставщика", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
                return;
            PurchaseWriteRepository.DeleteSupplier(supplier.Id);
            ReloadSelection(supplier.DealId);
            StatusText = "Поставщик перемещен в корзину";
            ToastService.Show(StatusText, ToastKind.Info);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Удаление поставщика");
        }
    }

    /// <summary>
    /// Аналог Add-PurchaseDocumentsFromPaths (RRFQComparer.ps1, строки 5601-5645):
    /// запрос типа документа, для поставщиковых типов — выбранный поставщик,
    /// затем Add-PurchaseDocument для каждого файла. Для RRFQ сразу предлагается
    /// добавить найденные квоты в «Базу квот».
    /// </summary>
    public void AddDocumentsFromPaths(IReadOnlyList<string> paths)
    {
        try
        {
            var deal = SelectedDeal;
            if (deal is null)
            {
                throw new InvalidOperationException("Выберите сделку.");
            }

            var files = paths.Where(p => !string.IsNullOrWhiteSpace(p)).Distinct().ToList();
            if (files.Count == 0)
            {
                return;
            }

            var typeDialog = new DocumentTypeWindow
            {
                Owner = Application.Current.MainWindow,
            };
            if (typeDialog.ShowDialog() != true || string.IsNullOrWhiteSpace(typeDialog.SelectedType))
            {
                return;
            }

            var documentType = typeDialog.SelectedType;
            var isDealLevel = PurchaseDocumentsStore.IsDealLevelDocumentType(documentType);
            var supplierId = isDealLevel ? 0 : SelectedSupplier?.Id ?? 0;
            if (!isDealLevel && supplierId <= 0)
            {
                throw new InvalidOperationException("Выберите поставщика для привязки файлов.");
            }
            var folderName = isDealLevel ? documentType : string.Empty;

            var rrfqSources = new List<(string Path, string Supplier)>();
            foreach (var file in files)
            {
                PurchaseDocumentsStore.AddDocument(deal.Id, supplierId, documentType, file, folderName);
                if (documentType == "RRFQ") rrfqSources.Add((file, SelectedSupplier?.Supplier ?? string.Empty));
            }

            Load();
            SelectedDeal = Deals.FirstOrDefault(r => r.Id == deal.Id);
            var addedQuotes = rrfqSources.Count == 0 ? 0 : ImportQuotes(rrfqSources, "Импорт из документов сделки");
            StatusText = "Загружено файлов: " + files.Count + (rrfqSources.Count == 0 ? string.Empty : "; квот добавлено: " + addedQuotes);
            ToastService.Show(StatusText, ToastKind.Success);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Загрузка документов");
        }
    }

    /// <summary>Аналог $btnPickPurchaseDocuments: OpenFileDialog с Multiselect.</summary>
    [RelayCommand]
    private void PickDocuments()
    {
        var dialog = new OpenFileDialog
        {
            Title = "Выберите документы",
            Filter = "All files (*.*)|*.*",
            Multiselect = true,
        };
        if (dialog.ShowDialog(Application.Current.MainWindow) == true)
        {
            AddDocumentsFromPaths(dialog.FileNames);
        }
    }

    /// <summary>Аналог пункта «Открыть в проводнике» контекстного меню $docsGrid.</summary>
    [RelayCommand]
    private void OpenDocumentInExplorer()
    {
        try
        {
            var document = SelectedDocument;
            if (document is null)
            {
                throw new InvalidOperationException("Выберите документ.");
            }

            if (!PurchaseDocumentsStore.OpenInExplorer(document.Id, document.StoredPath, document.OriginalName, document.FileHash))
            {
                throw new InvalidOperationException("Файл документа не найден: " + document.OriginalName);
            }
            LoadDocuments();
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Документы");
        }
    }

    /// <summary>Аналог пункта «Удалить»: Delete-PurchaseDocument через корзину.</summary>
    [RelayCommand]
    private void DeleteDocument()
    {
        try
        {
            var deal = SelectedDeal;
            var document = SelectedDocument;
            if (deal is null || document is null)
            {
                throw new InvalidOperationException("Выберите документ.");
            }

            var confirm = MessageBox.Show(
                "Отправить документ в корзину?\n" + document.OriginalName,
                "Удаление документа",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);
            if (confirm != MessageBoxResult.Yes)
            {
                return;
            }

            PurchaseDocumentsStore.DeleteDocument(document.Id);
            Load();
            SelectedDeal = Deals.FirstOrDefault(r => r.Id == deal.Id);
            StatusText = "Документ удален в корзину";
            ToastService.Show(StatusText, ToastKind.Info);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Удаление документа");
        }
    }

    /// <summary>Открывает файл документа по двойному щелчку, как $docsGrid.Add_CellDoubleClick.</summary>
    public void OpenDocument(PurchaseDocumentRow document)
    {
        try
        {
            if (!PurchaseDocumentsStore.OpenDocument(document.Id, document.StoredPath, document.OriginalName, document.FileHash))
            {
                throw new InvalidOperationException("Файл документа не найден: " + document.OriginalName);
            }
            LoadDocuments();
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Документы");
        }
    }

    /// <summary>Повторный импорт квот из выбранного документа типа RRFQ.</summary>
    [RelayCommand]
    private void LoadQuotesFromDocument()
    {
        try
        {
            var document = SelectedDocument ?? throw new InvalidOperationException("Выберите документ.");
            if (!string.Equals(document.DocumentType, "RRFQ", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Загрузка квот доступна только для документа типа RRFQ.");
            var resolved = PurchaseWriteSession.Execute(connection => PurchaseDocumentsStore.ResolveDocumentFile(
                connection, document.Id, document.StoredPath, document.OriginalName, document.FileHash));
            if (!resolved.Found || !File.Exists(resolved.Path))
                throw new FileNotFoundException("Файл документа не найден: " + document.OriginalName, resolved.Path);
            var count = ImportQuotes(new[] { (resolved.Path, document.Supplier) }, "Повторная загрузка квот из документа сделки");
            StatusText = "Квот добавлено: " + count;
            ToastService.Show(StatusText, ToastKind.Success);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Загрузка квот");
        }
    }

    /// <summary>Заполнить панель редактирования полями выбранной сделки.</summary>
    private void FillEditPanel(PurchaseDealRow? deal)
    {
        EditClient = deal?.Client ?? string.Empty;
        EditDealNumber = deal?.DealNumber ?? string.Empty;
        EditBoardCount = deal?.BoardCount ?? string.Empty;
        EditStatus = deal is not null && DealStatuses.Contains(deal.Status) ? deal.Status : "RFQ";
        EditPeriod = deal?.Period ?? string.Empty;
        EditComment = deal?.Comment ?? string.Empty;
        EditPriority = deal?.Priority ?? string.Empty;
        EditTrackingStatus = deal?.TrackingStatus ?? string.Empty;
        EditExecutor = deal?.Executor ?? string.Empty;
        EditReminderDate = deal?.ReminderDate ?? string.Empty;
        EditAssemblyLocation = deal?.AssemblyLocation ?? string.Empty;
        ArchiveButtonText = deal?.Archived == 1 ? "Вернуть" : "В архив";
    }

    private int ImportQuotes(IReadOnlyList<(string Path, string Supplier)> sources, string comment)
    {
        var quotes = RrfqQuoteImportService.ReadQuotes(sources);
        if (quotes.Count == 0) throw new InvalidOperationException("В RRFQ не найдено квот.");
        var dialog = new QuoteImportConfirmationWindow(quotes, sources.Select(source => source.Path))
        {
            Owner = Application.Current.MainWindow,
        };
        if (dialog.ShowDialog() != true) return 0;
        return RrfqQuoteImportService.SaveQuotes(quotes, sources.Select(source => source.Path), comment);
    }

    private void ReloadSelection(long dealId)
    {
        Load();
        SelectedDeal = Deals.FirstOrDefault(row => row.Id == dealId);
    }

    /// <summary>
    /// Сразу применяет фильтр и визуальные правила к изменённой строке.
    /// ApplyView использует уже загруженный список _allDeals, поэтому не сортирует
    /// сделки заново и не меняет их исходный порядок.
    /// </summary>
    private void RefreshEditedDealInPlace(PurchaseDealRow deal)
    {
        ApplyView();
        if (SelectedDeal?.Id == deal.Id) FillEditPanel(deal);
    }

    private void ApplyEditorValues(PurchaseDealRow deal)
    {
        deal.DealNumber = EditDealNumber;
        deal.BoardCount = EditBoardCount;
        deal.Client = EditClient;
        deal.Status = EditStatus;
        deal.Period = EditPeriod;
        deal.Priority = EditPriority;
        deal.TrackingStatus = EditTrackingStatus;
        deal.Executor = EditExecutor;
        deal.ReminderDate = EditReminderDate;
        deal.AssemblyLocation = EditAssemblyLocation;
        deal.Comment = EditComment;
    }

    private void Load()
    {
        try
        {
            _snapshot?.Dispose();
            _snapshot = PurchaseSnapshot.Create();
            _repository = new PurchaseRepository(_snapshot.Connection);

            _allDeals = _repository.GetDeals();
            var taskKeys = _repository.GetCockpitTaskKeys();
            var reminderCount = _repository.GetDealReminderRows().Count + _repository.GetManualReminderRows().Count;

            var metrics = PurchaseLogic.ComputeMetrics(_allDeals, taskKeys);
            ActiveCount = metrics.Active;
            OverdueCount = metrics.Overdue;
            WorkAndPaymentTaskCount = metrics.WorkAndPaymentTasks;
            AttentionCount = metrics.Attention;
            ReminderCount = reminderCount;

            ErrorMessage = string.Empty;
            ApplyView();
            LoadSuppliers();
            LoadDocuments();
        }
        catch (Exception ex)
        {
            _repository = null;
            ErrorMessage = "Не удалось прочитать данные: " + ex.Message;
        }
    }

    private void ApplyView()
    {
        var selectedId = SelectedDeal?.Id;

        var filtered = PurchaseLogic.FilterDeals(_allDeals, SelectedFilter.Value);
        var visible = PurchaseLogic.SearchDeals(filtered, SearchText);

        Deals.Clear();
        foreach (var row in visible)
        {
            Deals.Add(row);
        }

        CountText = "Всего: " + Deals.Count;
        SelectedDeal = Deals.FirstOrDefault(r => r.Id == selectedId);
    }

    private void LoadSuppliers()
    {
        Suppliers.Clear();
        var deal = SelectedDeal;
        if (deal is null || _repository is null)
        {
            return;
        }

        try
        {
            foreach (var supplier in _repository.GetSuppliers(deal.Id))
            {
                Suppliers.Add(supplier);
            }
        }
        catch (Exception ex)
        {
            ErrorMessage = "Не удалось прочитать поставщиков: " + ex.Message;
        }
    }

    /// <summary>Аналог Refresh-PurchaseDocuments: документы выбранной сделки.</summary>
    private void LoadDocuments()
    {
        var selectedId = SelectedDocument?.Id;
        Documents.Clear();
        var deal = SelectedDeal;
        if (deal is null)
        {
            return;
        }

        try
        {
            foreach (var document in PurchaseDocumentsStore.GetDocuments(deal.Id))
            {
                Documents.Add(document);
            }
            SelectedDocument = Documents.FirstOrDefault(d => d.Id == selectedId);
        }
        catch (Exception ex)
        {
            ErrorMessage = "Не удалось прочитать документы: " + ex.Message;
        }
    }
}
