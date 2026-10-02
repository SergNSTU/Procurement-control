using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;
using ProcurementControl.Models;
using ProcurementControl.Navigation;
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
    private IReadOnlyList<CockpitTaskKey> _cockpitTasks = Array.Empty<CockpitTaskKey>();
    private IReadOnlyList<SupplierPlanRow> _supplierPlanRows = Array.Empty<SupplierPlanRow>();
    private IReadOnlyList<DealReminderRow> _dealReminders = Array.Empty<DealReminderRow>();
    private IReadOnlyList<ManualReminderRow> _manualReminders = Array.Empty<ManualReminderRow>();
    private readonly DispatcherTimer _searchTimer;
    private PurchaseDealRow? _editorDeal;
    private bool _isLoadingEditor;
    private bool _isRestoringSelection;

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
    private bool _hasUnsavedChanges;

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
    [ObservableProperty] private int _expectingReceiptCount;
    [ObservableProperty] private int _reminderCount;

    /// <summary>
    /// Тестовая фича: автоматический импорт Excel-вложений из писем поставщиков
    /// (comp@npkexpert.ru) в документы активных сделок как RRFQ.
    /// </summary>
    [ObservableProperty]
    private bool _mailImportEnabled = OutlookMailImportService.IsEnabled;

    partial void OnMailImportEnabledChanged(bool value) => OutlookMailImportService.IsEnabled = value;

    /// <summary>
    /// Авто-импорт файла прихода со склада (письма «CE CT WE daily update»,
    /// раз в 2 часа): сверка заказов из комментариев поставщиков, «Факт
    /// поступл.», предупреждения о просрочках.
    /// </summary>
    [ObservableProperty]
    private bool _warehouseImportEnabled = WarehouseArrivalImportService.IsEnabled;

    partial void OnWarehouseImportEnabledChanged(bool value) => WarehouseArrivalImportService.IsEnabled = value;

    public IReadOnlyList<FilterOption> Filters { get; } =
        Enum.GetValues<DealFilter>().Select(f => new FilterOption(f, f.Label())).ToList();

    public ObservableCollection<PurchaseDealRow> Deals { get; } = new();

    public ObservableCollection<SupplierRow> Suppliers { get; } = new();

    public ObservableCollection<PurchaseDocumentRow> Documents { get; } = new();

    public PurchaseViewModel()
    {
        _selectedFilter = Filters[0];
        _searchTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(250) };
        _searchTimer.Tick += (_, _) =>
        {
            _searchTimer.Stop();
            ApplyView();
        };
        Load();
    }

    partial void OnSearchTextChanged(string value)
    {
        _searchTimer.Stop();
        _searchTimer.Start();
    }

    partial void OnSelectedFilterChanged(FilterOption value) => ApplyView();

    partial void OnSelectedDealChanged(PurchaseDealRow? value)
    {
        if (!_isRestoringSelection && HasUnsavedChanges && _editorDeal?.Id != value?.Id)
        {
            var answer = MessageBox.Show(
                "В карточке сделки есть несохранённые изменения. Отменить их и перейти к другой сделке?",
                "Несохранённые изменения",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);
            if (answer != MessageBoxResult.Yes)
            {
                _isRestoringSelection = true;
                SelectedDeal = _editorDeal;
                _isRestoringSelection = false;
                return;
            }
        }

        FillEditPanel(value);
        SelectedSupplier = null;
        LoadSuppliers();
        LoadDocuments();
    }

    partial void OnEditClientChanged(string value) => MarkEditorDirty();
    partial void OnEditDealNumberChanged(string value) => MarkEditorDirty();
    partial void OnEditBoardCountChanged(string value) => MarkEditorDirty();
    partial void OnEditStatusChanged(string value) => MarkEditorDirty();
    partial void OnEditPeriodChanged(string value) => MarkEditorDirty();
    partial void OnEditCommentChanged(string value) => MarkEditorDirty();
    partial void OnEditPriorityChanged(string value) => MarkEditorDirty();
    partial void OnEditTrackingStatusChanged(string value) => MarkEditorDirty();
    partial void OnEditExecutorChanged(string value) => MarkEditorDirty();
    partial void OnEditReminderDateChanged(string value) => MarkEditorDirty();
    partial void OnEditAssemblyLocationChanged(string value) => MarkEditorDirty();

    [RelayCommand]
    private void Reload() => Load();

    /// <summary>
    /// Ручная проверка ящика поставщиков без ожидания минутного таймера.
    /// Выполняется в фоновом STA-потоке (Outlook Object Model): иначе холодный
    /// запуск Outlook замораживал окно, а полоса прогресса не двигалась.
    /// Прогресс показывает сам OutlookMailImportService.Check.
    /// </summary>
    [RelayCommand]
    private void CheckMail()
    {
        var dispatcher = Application.Current?.Dispatcher;
        var worker = new Thread(() =>
        {
            var report = OutlookMailImportService.Check(manual: true);
            if (dispatcher is null || dispatcher.CheckAccess())
            {
                HandleMailReport(report);
            }
            else
            {
                dispatcher.Invoke(() => HandleMailReport(report));
            }
        });
        worker.SetApartmentState(ApartmentState.STA);
        worker.IsBackground = true;
        worker.Start();
    }

    /// <summary>
    /// Ручная проверка файла прихода со склада без ожидания двухчасового
    /// таймера. STA-поток — как в CheckMail: Outlook и Excel Object Model
    /// требуют STA, иначе замораживают окно.
    /// </summary>
    [RelayCommand]
    private void CheckWarehouse()
    {
        var dispatcher = Application.Current?.Dispatcher;
        var worker = new Thread(() =>
        {
            var report = WarehouseArrivalImportService.Check(manual: true);
            if (dispatcher is null || dispatcher.CheckAccess())
            {
                HandleWarehouseReport(report);
            }
            else
            {
                dispatcher.Invoke(() => HandleWarehouseReport(report));
            }
        });
        worker.SetApartmentState(ApartmentState.STA);
        worker.IsBackground = true;
        worker.Start();
    }

    private void HandleWarehouseReport(WarehouseArrivalReport report)
    {
        try
        {
            if (report.Error is not null)
            {
                ToastService.Show("Приход со склада: " + report.Error, ToastKind.Warn);
                return;
            }

            if (report.Disabled)
            {
                ToastService.Show("Импорт приходов выключен — включите галочку «Импорт приходов со склада»",
                    ToastKind.Info);
                return;
            }

            if (report.FileSubject is null)
            {
                ToastService.Show("Новых файлов прихода «CE CT WE» в почте нет", ToastKind.Info);
                return;
            }

            // Отчёт всегда показываем окном: там и сводка, и замечания.
            WarehouseArrivalNotifier.Show(report);
            if (report.FilledFacts > 0)
            {
                NotificationCenter.NotifyDataChanged();
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Приход со склада");
        }
    }

    private void HandleMailReport(MailImportReport report)
    {        try
        {
            if (report.Error is not null)
            {
                ToastService.Show("Импорт из почты: " + report.Error, ToastKind.Warn);
                return;
            }

            if (report.Imported.Count > 0)
            {
                // Постоянное уведомление: висит, пока пользователь не нажмёт ОК.
                MailImportNotifier.Show(report.Imported);
                NotificationCenter.NotifyDataChanged();
            }
            else
            {
                ToastService.Show(report.Disabled
                    ? "Импорт из почты выключен — включите галочку «Импорт RRFQ из почты»"
                    : "Почта проверена: новых Excel-вложений по активным сделкам нет (просмотрено: " + report.Scanned + ")",
                    ToastKind.Info);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Импорт из почты");
        }
    }

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
            // Флаг гасим ДО обновления списка: ApplyView пересобирает Deals и
            // временно сбрасывает выбор — при живом флаге в этот момент
            // срабатывает диалог «несохранённые изменения».
            HasUnsavedChanges = false;
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
            else if (field == nameof(SupplierRow.ActualReceiptDate)) supplier.ActualReceiptDate = value;
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

            foreach (var file in files)
            {
                PurchaseDocumentsStore.AddDocument(deal.Id, supplierId, documentType, file, folderName);
            }

            QuoteImportSaveResult? imported = null;
            if (documentType == "RRFQ")
            {
                var supplier = SelectedSupplier?.Supplier ?? string.Empty;
                imported = ImportQuotes(files.Select(file => (Path: file, Supplier: supplier)).ToList(),
                    "Импорт из документов сделки", deal);
            }

            Load();
            SelectedDeal = Deals.FirstOrDefault(r => r.Id == deal.Id);
            StatusText = imported is null
                ? "Загружено файлов: " + files.Count
                : $"Загружено файлов: {files.Count}; квоты: +{imported.Added}, обновлено {imported.Updated}, одинаковых {imported.Skipped}";
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
            if (document.DocumentType != "RRFQ")
                throw new InvalidOperationException("В аналитику загружаются только отдельные RRFQ поставщиков. Сводный RRFQ — итог сравнения и не импортируется как новая квота.");
            var resolved = PurchaseWriteSession.Execute(connection => PurchaseDocumentsStore.ResolveDocumentFile(
                connection, document.Id, document.StoredPath, document.OriginalName, document.FileHash));
            if (!resolved.Found || !File.Exists(resolved.Path))
                throw new FileNotFoundException("Файл документа не найден: " + document.OriginalName, resolved.Path);
            var result = ImportQuotes(new[] { (resolved.Path, document.Supplier) }, "Повторная загрузка квот из документа сделки", SelectedDeal!);
            StatusText = $"Квоты: добавлено {result.Added}, обновлено {result.Updated}, одинаковых {result.Skipped}, проверка {result.NeedsReview}";
            ToastService.Show(StatusText, ToastKind.Success);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Загрузка квот");
        }
    }

    [RelayCommand]
    private void ChangeDocumentBinding()
    {
        try
        {
            var deal = SelectedDeal ?? throw new InvalidOperationException("Выберите сделку.");
            var document = SelectedDocument ?? throw new InvalidOperationException("Выберите документ.");
            var dialog = new DocumentBindingWindow(document, Suppliers)
            {
                Owner = Application.Current.MainWindow,
            };
            if (dialog.ShowDialog() != true) return;
            PurchaseDocumentsStore.UpdateDocumentBinding(document.Id, dialog.SelectedType, dialog.SelectedSupplierId);
            LoadDocuments();
            StatusText = "Тип / поставщик документа обновлен";
            ToastService.Show(StatusText, ToastKind.Success);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Документы");
        }
    }

    [RelayCommand]
    private void MergeRrfq()
    {
        try
        {
            var deal = SelectedDeal ?? throw new InvalidOperationException("Выберите сделку.");
            var rfqs = Documents.Where(document => string.Equals(document.DocumentType, "RFQ", StringComparison.OrdinalIgnoreCase)).ToList();
            var responses = Documents.Where(document => string.Equals(document.DocumentType, "RRFQ", StringComparison.OrdinalIgnoreCase)).ToList();
            if (rfqs.Count == 0) throw new InvalidOperationException("Для сведения добавьте хотя бы один документ RFQ.");
            if (responses.Count == 0) throw new InvalidOperationException("Для сведения добавьте хотя бы один RRFQ поставщика.");

            var dialog = new RrfqMergeWindow(rfqs, responses)
            {
                Owner = Application.Current.MainWindow,
            };
            if (dialog.ShowDialog() != true) return;
            var rfq = dialog.SelectedRfq ?? throw new InvalidOperationException("Выберите RFQ.");
            var selectedResponses = dialog.SelectedResponses;
            if (selectedResponses.Count == 0) throw new InvalidOperationException("Выберите хотя бы один ответ поставщика.");
            var selectedUnassigned = selectedResponses.Where(document => document.SupplierId <= 0).ToList();
            if (selectedUnassigned.Count > 0)
                throw new InvalidOperationException("У выбранных RRFQ не указан поставщик: " + string.Join(", ", selectedUnassigned.Select(document => document.OriginalName)) +
                    ". Исключите их из состава или сначала назначьте поставщика через контекстное меню.");

            var selected = new[] { rfq }.Concat(selectedResponses).ToList();
            var resolved = new Dictionary<long, ResolvedDocumentFile>();
            foreach (var document in selected)
            {
                var file = PurchaseWriteSession.Execute(connection => PurchaseDocumentsStore.ResolveDocumentFile(
                    connection, document.Id, document.StoredPath, document.OriginalName, document.FileHash));
                if (!file.Found || !File.Exists(file.Path))
                    throw new FileNotFoundException("Не найден файл документа: " + document.OriginalName, file.Path);
                resolved[document.Id] = file;
            }

            var context = new RrfqDealComparisonContext(deal.Id, deal.DealNumber, rfq.Id, resolved[rfq.Id].Path,
                selectedResponses.Select(document => new RrfqDealSourceDocument(document.Id, document.SupplierId,
                    document.Supplier, resolved[document.Id].Path)).ToList());
            if (Application.Current.MainWindow?.DataContext is MainViewModel main)
                main.Navigate(AppPage.Rrfq);
            if (!RrfqImportQueue.OpenDealComparison(context)) return;
            StatusText = "Документы сделки переданы в сравнение RRFQ";
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Свести RRFQ");
        }
    }

    /// <summary>
    /// Конвейер RFQ по сделке (ветка rfq-pipeline): исходник заявки → подготовка
    /// поставщицкого файла (зелёные строки) → рассылка черновиков поставщикам.
    /// Чтение исходника — в фоновом STA-потоке (Excel COM), собранные файлы
    /// регистрируются в документах сделки.
    /// </summary>
    [RelayCommand]
    private void PrepareRfq()
    {
        try
        {
            var deal = SelectedDeal ?? throw new InvalidOperationException("Выберите сделку.");

            // Поставщиков на этапе подготовки не требуем — они нужны только
            // для рассылки, и их можно завести, пока RFQ уже собран.
            if (TrySendExistingRfq(deal))
            {
                return;
            }

            var dialog = new OpenFileDialog
            {
                Title = "Исходный RFQ (заявка клиента)",
                Filter = "Excel files (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls|All files (*.*)|*.*",
            };
            if (dialog.ShowDialog(Application.Current.MainWindow) != true)
            {
                return;
            }

            StartRfqPrepare(dialog.FileName, deal.Id, deal.DealNumber, deal.Client);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "RFQ поставщикам");
        }
    }

    /// <summary>
    /// Шаг 1 конвейера: исходник берётся из задачи Bitrix, связанной со сделкой
    /// (файлы задачи, чек-листа и комментариев; REST + вебхук из настроек Bitrix).
    /// Найденный RFQ прикрепляется к документам сделки как есть, без пересборки.
    /// </summary>
    [RelayCommand]
    private async Task LoadRfqFromBitrixAsync()
    {
        try
        {
            var deal = SelectedDeal ?? throw new InvalidOperationException("Выберите сделку.");

            if (TrySendExistingRfq(deal))
            {
                return;
            }

            var config = BitrixRepository.GetConfig();
            if (!config.IsConfigured)
            {
                throw new InvalidOperationException("Bitrix не настроен: нет вебхука (bitrix_webhook_token).");
            }

            var taskId = BitrixRepository.GetTaskIdForDeal(deal.Id);
            if (taskId <= 0)
            {
                // Сделка создана вручную, но задача в Bitrix может существовать:
                // ищем по номеру заявки в названии, опечатки номера — через
                // «похожие» (тот же хвост после дефиса) с подтверждением.
                ProgressService.Begin("Bitrix: поиск задачи по номеру " + deal.DealNumber + "…");
                BitrixTaskSearchResult search;
                try
                {
                    search = await BitrixService.FindTaskByRequestNumberAsync(config, deal.DealNumber);
                }
                finally
                {
                    ProgressService.End();
                }

                if (search.Task is null)
                {
                    throw new InvalidOperationException(
                        "В доступных вебхуку задачах Bitrix (" + search.Scanned + " шт.) нет задачи с номером " + deal.DealNumber +
                        " и похожих номеров. Проверьте номер сделки или загрузите RFQ файлом («RFQ поставщикам»).");
                }

                if (search.IsSimilarMatch)
                {
                    var answer = MessageBox.Show(
                        "Задачи с точным номером " + deal.DealNumber + " нет, но найдена похожая:\r\n\r\n" +
                        search.Task.Title + "\r\n\r\nСвязать сделку с ней и загрузить RFQ?",
                        "RFQ из Bitrix", MessageBoxButton.YesNo, MessageBoxImage.Question);
                    if (answer != MessageBoxResult.Yes)
                    {
                        return;
                    }
                }

                BitrixRepository.LinkTask(deal.Id, search.Task.TaskId, "rfq_pipeline");
                taskId = search.Task.TaskId;
                StatusText = "Сделка связана с задачей Bitrix " + taskId;
                ToastService.Show("Задача " + taskId + ": " + search.Task.Title, ToastKind.Info);
            }

            var dealId = deal.Id;
            var dealNumber = deal.DealNumber;
            var client = deal.Client;

            ProgressService.Begin("Bitrix: поиск файлов задачи " + taskId + "…");
            List<BitrixService.BitrixTaskFile> files;
            try
            {
                files = await BitrixService.GetTaskFilesAsync(config, taskId);
            }
            finally
            {
                ProgressService.End();
            }

            if (files.Count == 0)
            {
                MessageBox.Show(
                    "В задаче " + taskId + " не найдено вложений (файлы задачи, чек-лист, комментарии).\n" +
                    "Исходный RFQ приходит в чат задачи или прикрепляется к чек-листу.",
                    "RFQ из Bitrix");
                return;
            }

            // RFQ — исходный файл без цен. RRFQ (ответы поставщиков) исключаем,
            // остальное приоритизируем: RFQ_номер / RFQ номер / просто номер сделки.
            var candidates = files
                .Where(file => !IsRrfqName(file.Name))
                .OrderByDescending(file => ScoreRfqCandidate(file, dealNumber))
                .ToList();

            if (candidates.Count == 0)
            {
                MessageBox.Show(
                    "В задаче " + taskId + " нет исходного RFQ: найдены только ответы поставщиков (RRFQ).\n\n" +
                    "Исходник присылает Евгений С. в чат задачи (RFQ_номер, RFQ номер или просто номер сделки)\n" +
                    "либо он прикреплён к пункту чек-листа «Заполнение сводного RFQ».",
                    "RFQ из Bitrix");
                return;
            }

            BitrixService.BitrixTaskFile? selected = candidates[0];
            if (candidates.Count > 1)
            {
                var picker = new RfqBitrixFilesWindow(dealNumber, candidates) { Owner = Application.Current.MainWindow };
                if (picker.ShowDialog() != true || picker.SelectedFile is null)
                {
                    return;
                }

                selected = picker.SelectedFile;
            }

            var tempDirectory = Path.Combine(Path.GetTempPath(), "rfq-bitrix-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDirectory);
            var localPath = Path.Combine(tempDirectory,
                string.IsNullOrWhiteSpace(selected.Name) ? "task_" + taskId + "_file.xlsx" : selected.Name);

            ProgressService.Begin("Bitrix: скачивание " + (selected.Name.Length > 0 ? selected.Name : "вложения") + "…");
            try
            {
                await BitrixService.DownloadTaskFileAsync(config, selected, localPath);
            }
            finally
            {
                ProgressService.End();
            }

            StoreBitrixRfqAsIs(localPath, dealId, dealNumber, client);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "RFQ из Bitrix");
        }
    }

    /// <summary>
    /// Скачанный из Bitrix RFQ идёт в документы сделки как есть — без чтения,
    /// пересборки и списка «Для Компэла»; дальше — предложение разослать файл.
    /// </summary>
    private void StoreBitrixRfqAsIs(string sourcePath, long dealId, string dealNumber, string client)
    {
        try
        {
            var storedRfq = PurchaseDocumentsStore.AddDocument(dealId, 0, "RFQ", sourcePath, folderName: "RFQ");
            LoadDocuments();

            if (Suppliers.Count == 0)
            {
                // Поставщиков можно завести позже: повторное нажатие предложит
                // разослать готовый файл без пересборки.
                StatusText = "RFQ из Bitrix прикреплён к сделке; для рассылки добавьте поставщиков";
                MessageBox.Show(
                    "RFQ из Bitrix прикреплён к сделке как есть.\r\n\r\n" +
                    "Чтобы разослать: добавьте поставщиков сделки и снова нажмите\r\n" +
                    "«RFQ поставщикам» — приложение предложит отправить готовый файл.",
                    "RFQ из Bitrix");
                return;
            }

            var sendWindow = new RfqSendWindow(dealNumber, client, Suppliers.ToList(), new[] { storedRfq })
            {
                Owner = Application.Current.MainWindow,
            };
            if (sendWindow.ShowDialog() == true)
            {
                StatusText = "RFQ прикреплён, черновиков: " + sendWindow.CreatedDrafts;
                ToastService.Show("Черновиков создано: " + sendWindow.CreatedDrafts, ToastKind.Success);
            }
            else
            {
                StatusText = "RFQ из Bitrix прикреплён к сделке";
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "RFQ из Bitrix");
        }
        finally
        {
            // Хранилище документов держит свою копию — временная папка не нужна.
            try { Directory.Delete(Path.GetDirectoryName(sourcePath)!, recursive: true); } catch { }
        }
    }

    /// <summary>Общая часть конвейера: чтение исходника и открытие окна подготовки.</summary>
    private void StartRfqPrepare(string sourcePath, long dealId, string dealNumber, string client)
    {
        var suppliers = Suppliers.ToList();

        var dispatcher = Application.Current?.Dispatcher;
        ProgressService.Begin("RFQ: чтение исходника…");
        var worker = new Thread(() =>
        {
            Exception? error = null;
            RfqSourceReadResult? source = null;
            try
            {
                source = RfqPrepareService.ReadSource(sourcePath);
            }
            catch (Exception ex)
            {
                error = ex;
            }

            if (dispatcher is null || dispatcher.CheckAccess())
            {
                ContinueRfqPipeline(error, source, dealId, dealNumber, client, suppliers);
            }
            else
            {
                dispatcher.Invoke(() => ContinueRfqPipeline(error, source, dealId, dealNumber, client, suppliers));
            }
        });
        worker.SetApartmentState(ApartmentState.STA);
        worker.IsBackground = true;
        worker.Start();
    }

    /// <summary>
    /// Если к сделке уже прикреплён RFQ — предлагает разослать готовый файл без
    /// пересборки. Возвращает true, когда пользователь выбрал рассылку (или
    /// закрыл предложение рассылкой), false — нужно собирать новый.
    /// </summary>
    private bool TrySendExistingRfq(PurchaseDealRow deal)
    {
        var rfqDocument = Documents.FirstOrDefault(document =>
            string.Equals(document.DocumentType, "RFQ", StringComparison.OrdinalIgnoreCase));
        if (rfqDocument is null)
        {
            return false;
        }

        var answer = MessageBox.Show(
            "У сделки уже есть RFQ: " + rfqDocument.OriginalName + ".\r\n\r\n" +
            "Да — разослать готовый файл поставщикам\r\n" +
            "Нет — собрать новый RFQ из исходника",
            "RFQ поставщикам", MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (answer != MessageBoxResult.Yes)
        {
            return false;
        }

        OpenRfqSend(deal, rfqDocument);
        return true;
    }

    private void OpenRfqSend(PurchaseDealRow deal, PurchaseDocumentRow rfqDocument)
    {
        try
        {
            if (Suppliers.Count == 0)
            {
                MessageBox.Show(
                    "У сделки нет поставщиков. Добавьте их в карточке сделки (кнопка «Добавить поставщика»)\r\n" +
                    "и снова нажмите «RFQ поставщикам» — приложение предложит разослать готовый файл.",
                    "RFQ поставщикам");
                return;
            }

            var file = PurchaseWriteSession.Execute(connection => PurchaseDocumentsStore.ResolveDocumentFile(
                connection, rfqDocument.Id, rfqDocument.StoredPath, rfqDocument.OriginalName, rfqDocument.FileHash));
            if (!file.Found)
            {
                throw new InvalidOperationException("Файл RFQ не найден на диске: " + file.Path);
            }

            var sendWindow = new RfqSendWindow(deal.DealNumber, deal.Client, Suppliers.ToList(), new[] { file.Path })
            {
                Owner = Application.Current.MainWindow,
            };
            if (sendWindow.ShowDialog() == true)
            {
                StatusText = "Черновиков создано: " + sendWindow.CreatedDrafts;
                ToastService.Show("Черновиков создано: " + sendWindow.CreatedDrafts, ToastKind.Success);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "RFQ поставщикам");
        }
    }

    /// <summary>RRFQ — готовый файл с ценами от поставщика, не исходник для запроса.</summary>
    private static bool IsRrfqName(string name)
        => name.Contains("rrfq", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Приоритет файла-исходника: имена «RFQ_номер», «RFQ номер» или просто
    /// номер сделки; Excel — обязательный признак исходника. См. RfqBitrixFilesWindow.
    /// </summary>
    private static int ScoreRfqCandidate(BitrixService.BitrixTaskFile file, string dealNumber)
    {
        var baseName = Path.GetFileNameWithoutExtension(file.Name);
        var isExcel = Path.GetExtension(file.Name) is ".xlsx" or ".xlsm" or ".xls";
        var score = isExcel ? 1 : -20;
        if (baseName.StartsWith("rfq", StringComparison.OrdinalIgnoreCase))
        {
            score += 4;
        }

        if (file.Name.Contains(dealNumber, StringComparison.OrdinalIgnoreCase))
        {
            score += 4;
        }

        if (string.Equals(baseName, dealNumber, StringComparison.OrdinalIgnoreCase))
        {
            score += 2;
        }

        return score;
    }

    private void ContinueRfqPipeline(
        Exception? error,
        RfqSourceReadResult? source,
        long dealId,
        string dealNumber,
        string client,
        List<SupplierRow> suppliers)
    {
        ProgressService.End();
        try
        {
            if (error is not null)
            {
                throw error;
            }

            if (source is null || source.Rows.Count == 0)
            {
                throw new InvalidOperationException("В исходнике не найдено ни одной позиции.");
            }

            var prepareWindow = new RfqPrepareWindow(dealNumber, client, source, Path.GetTempPath())
            {
                Owner = Application.Current.MainWindow,
            };
            if (prepareWindow.ShowDialog() != true)
            {
                return;
            }

            // Собранные файлы регистрируются в документах сделки; письма
            // прикрепляют копию из хранилища, временные файлы больше не нужны.
            var storedRfq = PurchaseDocumentsStore.AddDocument(
                dealId, 0, "RFQ", prepareWindow.SupplierFilePath, folderName: "RFQ");
            try
            {
                PurchaseDocumentsStore.AddDocument(
                    dealId, 0, "Other", prepareWindow.CompelFilePath,
                    folderName: "RFQ",
                    displayName: Path.GetFileName(prepareWindow.CompelFilePath));
            }
            catch (Exception ex)
            {
                ErrorLog.Append("documents", ex);
            }

            LoadDocuments();

            if (Suppliers.Count == 0)
            {
                // Поставщиков можно завести позже: повторное нажатие «RFQ
                // поставщикам» предложит разослать готовый файл без пересборки.
                StatusText = "RFQ готов и прикреплён к сделке; для рассылки добавьте поставщиков";
                MessageBox.Show(
                    "RFQ собран и прикреплён к сделке (вместе со списком «Для Компэла»).\r\n\r\n" +
                    "Чтобы разослать: добавьте поставщиков сделки и снова нажмите\r\n" +
                    "«RFQ поставщикам» — приложение предложит отправить готовый файл.",
                    "RFQ поставщикам");
                return;
            }

            var sendWindow = new RfqSendWindow(dealNumber, client, suppliers, new[] { storedRfq })
            {
                Owner = Application.Current.MainWindow,
            };
            if (sendWindow.ShowDialog() == true)
            {
                StatusText = "RFQ подготовлен, черновиков: " + sendWindow.CreatedDrafts;
                ToastService.Show("Черновиков создано: " + sendWindow.CreatedDrafts, ToastKind.Success);
            }
            else
            {
                StatusText = "RFQ подготовлен и прикреплён к сделке";
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "RFQ поставщикам");
        }
    }

    /// <summary>Заполнить панель редактирования полями выбранной сделки.</summary>
    private void FillEditPanel(PurchaseDealRow? deal)
    {
        _isLoadingEditor = true;
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
        _editorDeal = deal;
        HasUnsavedChanges = false;
        _isLoadingEditor = false;
    }

    private void MarkEditorDirty()
    {
        if (_isLoadingEditor || SelectedDeal is null)
        {
            return;
        }

        HasUnsavedChanges = true;
        StatusText = "Есть несохранённые изменения";
    }

    private QuoteImportSaveResult ImportQuotes(IReadOnlyList<(string Path, string Supplier)> sources, string comment, PurchaseDealRow deal)
    {
        var quotes = RrfqQuoteImportService.ReadQuotes(sources);
        if (quotes.Count == 0) throw new InvalidOperationException("В RRFQ не найдено квот.");
        var dialog = new QuoteImportConfirmationWindow(quotes, sources.Select(source => source.Path))
        {
            Owner = Application.Current.MainWindow,
        };
        if (dialog.ShowDialog() != true) return new QuoteImportSaveResult(0, 0, 0, 0, 0);
        return RrfqQuoteImportService.SaveQuotesForDeal(quotes, sources.Select(source => source.Path), comment, deal.Id, deal.DealNumber);
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

    /// <summary>Содержимое плитки метрик: список по клику на плитку.</summary>
    public (string Title, List<TileLineRow> Lines) GetTileDetails(string tileKey)
        => PurchaseLogic.BuildTileLines(
            tileKey, _allDeals, _cockpitTasks, _supplierPlanRows, _dealReminders, _manualReminders, DateTime.Today);

    /// <summary>Переход к сделке по номеру — для строк списков плиток.</summary>
    public void OpenDealByNumber(string dealNumber)
    {
        var normalized = dealNumber.Trim();
        if (normalized.Length == 0)
        {
            return;
        }

        var row = _allDeals.FirstOrDefault(candidate =>
            string.Equals(candidate.DealNumber.Trim(), normalized, StringComparison.OrdinalIgnoreCase));
        if (row is not null)
        {
            OpenDeal(row.Id);
        }
    }

    private void Load()
    {
        try
        {
            _snapshot?.Dispose();
            _snapshot = PurchaseSnapshot.Create();
            _repository = new PurchaseRepository(_snapshot.Connection);

            _allDeals = _repository.GetDeals();
            _cockpitTasks = _repository.GetCockpitTaskKeys();
            _supplierPlanRows = _repository.GetSupplierPlanRows();
            _dealReminders = _repository.GetDealReminderRows();
            _manualReminders = _repository.GetManualReminderRows();
            var reminderCount = _dealReminders.Count + _manualReminders.Count;

            var metrics = PurchaseLogic.ComputeMetrics(_allDeals, _cockpitTasks, _supplierPlanRows);
            ActiveCount = metrics.Active;
            OverdueCount = metrics.Overdue;
            WorkAndPaymentTaskCount = metrics.WorkAndPaymentTasks;
            ExpectingReceiptCount = metrics.ExpectingReceipt;
            ReminderCount = reminderCount;

            ErrorMessage = string.Empty;
            ApplyView();
            LoadSuppliers();
            LoadDocuments();
        }
        catch (Exception ex)
        {
            _repository = null;
            ErrorLog.Append("purchase", ex);
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
