using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using ProcurementControl.Models;
using ProcurementControl.Navigation;
using ProcurementControl.Services;

namespace ProcurementControl.ViewModels;

/// <summary>
/// Центральный навигационный контроллер главного окна. Аналог оригинальной
/// функции Show-AppPage из RRFQComparer.ps1: хранит пункты левого меню и
/// переключает активную страницу (CurrentViewModel).
/// </summary>
public partial class MainViewModel : ObservableObject
{
    private readonly NotesService _notesService = new();
    private readonly Dictionary<AppPage, ObservableObject> _pageCache = new();
    private DispatcherTimer? _notificationTimer;
    private string? _lastMailImportError;
    private string? _lastBitrixImportError;
    private string? _lastWarehouseImportError;

    public ObservableCollection<NavigationItem> NavItems { get; } = new();

    [ObservableProperty]
    private NavigationItem? _selectedNavItem;

    [ObservableProperty]
    private ObservableObject? _currentViewModel;

    [ObservableProperty]
    private string _currentPageTitle = string.Empty;

    public MainViewModel()
    {
        BuildNavItems();
        // Стартуем на основном разделе «Контроль закупки».
        Navigate(AppPage.Purchase);
        InitializeNotifications();
        // Записи быстрого захвата обновляют кэшированные страницы,
        // как после действий карточек уведомлений.
        QuickCaptureService.DataChanged += RefreshPagesAfterNotification;
    }

    partial void OnSelectedNavItemChanged(NavigationItem? value)
    {
        if (value is null || value.IsDivider)
        {
            return;
        }
        ShowPage(value.Page);
    }

    /// <summary>Программный переход на раздел (вызывает выбор пункта меню).</summary>
    public void Navigate(AppPage page)
    {
        var item = NavItems.FirstOrDefault(n => !n.IsDivider && n.Page == page);
        if (item is not null)
        {
            SelectedNavItem = item;
        }
        else
        {
            ShowPage(page);
        }
    }

    private void ShowPage(AppPage page)
    {
        CurrentViewModel = GetOrCreateViewModel(page);
        CurrentPageTitle = GetTitle(page);
    }

    private ObservableObject GetOrCreateViewModel(AppPage page)
    {
        if (_pageCache.TryGetValue(page, out var cached))
        {
            return cached;
        }

        ObservableObject vm = page switch
        {
            // Перенесённые разделы.
            AppPage.Dashboard => new DashboardViewModel(),
            AppPage.Purchase => new PurchaseViewModel(),
            AppPage.Components => new TasksViewModel(),
            AppPage.Reminders => new RemindersViewModel(),
            AppPage.CompelParser => new CompelParserViewModel(),
            AppPage.ExcelCompare => new ExcelCompareViewModel(),
            AppPage.Bitrix => new BitrixViewModel(),
            AppPage.Rrfq => new RrfqViewModel(),
            AppPage.SupplierAnalytics => new SupplierAnalyticsViewModel(),
            AppPage.QuoteBase => new QuoteBaseViewModel(),
            AppPage.PriceSearch => new PriceSearchViewModel(),
            AppPage.CrossReference => new CrossReferenceViewModel(),
            AppPage.History => new HistoryViewModel(() => GetSelectedPurchaseDealId()),
            AppPage.Instructions => new InstructionsViewModel(),
            AppPage.Settings => new SettingsViewModel(),
            AppPage.Notes => new NotesViewModel(_notesService),
            _ => new PlaceholderViewModel(GetTitle(page))
        };

        _pageCache[page] = vm;
        return vm;
    }

    private string GetTitle(AppPage page)
        => NavItems.FirstOrDefault(n => !n.IsDivider && n.Page == page)?.Title ?? page.ToString();

    /// <summary>
    /// Аналог Get-SelectedDealId: сделка, выбранная в «Контроле закупки».
    /// Раздел может быть ещё не открыт — тогда возвращается 0.
    /// </summary>
    private long GetSelectedPurchaseDealId()
        => _pageCache.TryGetValue(AppPage.Purchase, out var page)
            && page is PurchaseViewModel purchase
            && purchase.SelectedDeal is not null
                ? purchase.SelectedDeal.Id
                : 0;

    /// <summary>
    /// Подключение системы карточек уведомлений — аналоги строки 1383 и таймера
    /// $notificationTimer (RRFQComparer.ps1, строки 5870-5872).
    /// </summary>
    private void InitializeNotifications()
    {
        // Настройка «Уведомления включены» читается при старте, как $script:NotificationsEnabled.
        try
        {
            using var snapshot = PurchaseSnapshot.Create();
            var settings = new SettingsRepository(snapshot.Connection);
            NotificationCenter.Enabled = settings.GetBooleanSetting("notifications.enabled", true);
        }
        catch (Exception exception)
        {
            ErrorLog.Append("settings", exception);
            NotificationCenter.Enabled = true;
        }

        NotificationCenter.OpenTargetRequested += OpenNotificationTarget;
        NotificationCenter.DataChanged += RefreshPagesAfterNotification;
        Application.Current.Exit += (_, _) =>
        {
            NotificationCenter.Shutdown();
            MailImportNotifier.Shutdown();
            WarehouseArrivalNotifier.Shutdown();
        };

        _notificationTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(60) };
        _notificationTimer.Tick += (_, _) =>
        {
            NotificationCenter.Check();
            CheckSupplierMail();
            CheckBitrixDeals();
            CheckWarehouseArrival();
        };
        _notificationTimer.Start();

        // Первичная проверка, чтобы просроченные пункты показали себя сразу.
        NotificationCenter.Check();
        CheckSupplierMail();
        CheckBitrixDeals();
        CheckWarehouseArrival();
    }

    /// <summary>
    /// Тестовая фича: раз в минуту смотреть ящик comp@npkexpert.ru и загружать
    /// Excel-вложения писем поставщиков по активным сделкам. Outlook Object Model
    /// требует STA: проверка выполняется в отдельном фоновом потоке, иначе холодный
    /// запуск Outlook и каждый обход ящика замораживают интерфейс — в том числе
    /// при старте приложения. Ошибка показывается один раз, пока не сменится —
    /// без спама тостами каждую минуту.
    /// </summary>
    private void CheckSupplierMail()
    {
        var worker = new Thread(() =>
        {
            try
            {
                var report = OutlookMailImportService.Check();
                if (report.Error is not null)
                {
                    if (report.Error != _lastMailImportError)
                    {
                        _lastMailImportError = report.Error;
                        ToastService.Show("Импорт из почты: " + report.Error, ToastKind.Warn);
                    }
                    return;
                }

                _lastMailImportError = null;
                if (report.Imported.Count > 0)
                {
                    // Постоянное уведомление: висит, пока пользователь не нажмёт ОК
                    // (сам переводит вызов в UI-поток).
                    MailImportNotifier.Show(report.Imported);
                    // Обновление кэшированных страниц — только в UI-потоке.
                    Application.Current?.Dispatcher.BeginInvoke(NotificationCenter.NotifyDataChanged);
                }
            }
            catch (Exception exception)
            {
                // Ошибка импорта не должна ронять общий таймер уведомлений.
                ErrorLog.Append("mail", exception);
            }
        });
        worker.SetApartmentState(ApartmentState.STA);
        worker.IsBackground = true;
        worker.Start();
    }

    /// <summary>
    /// Авто-проверка файла прихода со склада (папка «CE CT WE»): по своим часам
    /// раз в два часа (WarehouseArrivalImportService.DueForAutoCheck) тянет
    /// свежий Excel из почты, сверяет заказы из комментариев поставщиков.
    /// STA-поток — как для почтового импорта: Outlook и Excel Object Model
    /// требуют STA, иначе замораживают интерфейс.
    /// </summary>
    private void CheckWarehouseArrival()
    {
        if (!WarehouseArrivalImportService.DueForAutoCheck)
        {
            return;
        }

        var worker = new Thread(() =>
        {
            try
            {
                var report = WarehouseArrivalImportService.Check();
                if (report.Error is not null)
                {
                    if (report.Error != _lastWarehouseImportError)
                    {
                        _lastWarehouseImportError = report.Error;
                        ToastService.Show("Приход со склада: " + report.Error, ToastKind.Warn);
                    }
                    return;
                }

                _lastWarehouseImportError = null;
                // Проверка теперь каждый раз перечитывает свежий файл, и замечания
                // есть почти всегда — авто-окно показываем только когда записаны
                // новые «Факты поступл.», иначе оно висело бы каждые два часа.
                if (!report.Disabled && report.FilledFacts > 0)
                {
                    WarehouseArrivalNotifier.Show(report);
                    Application.Current?.Dispatcher.BeginInvoke(NotificationCenter.NotifyDataChanged);
                }
            }
            catch (Exception exception)
            {
                // Ошибка импорта не должна ронять общий таймер уведомлений.
                ErrorLog.Append("warehouse", exception);
            }
        });
        worker.SetApartmentState(ApartmentState.STA);
        worker.IsBackground = true;
        worker.Start();
    }

    /// <summary>
    /// Автоимпорт новых сделок из проекта Bitrix «Заявка на расчет КС в РФ»:
    /// по своим часам (раз в 10 минут, внутри BitrixImportService) тянет
    /// задачи, созданные после точки отсчёта, в «Контроль закупки». Ошибка
    /// показывается один раз, пока не сменится — без спама тостами.
    /// </summary>
    private async void CheckBitrixDeals()
    {
        try
        {
            if (!BitrixImportService.DueForAutoCheck)
            {
                return;
            }

            var report = await BitrixImportService.SyncAsync();
            if (report.Error is not null)
            {
                if (report.Error != _lastBitrixImportError)
                {
                    _lastBitrixImportError = report.Error;
                    ToastService.Show("Импорт из Bitrix: " + report.Error, ToastKind.Warn);
                }
                return;
            }

            _lastBitrixImportError = null;
            if (report.Imported.Count > 0)
            {
                ToastService.Show("Импорт из Bitrix (" + report.Imported.Count + "):\n" + string.Join("\n", report.Imported), ToastKind.Success);
                NotificationCenter.NotifyDataChanged();
            }
        }
        catch
        {
            // Ошибка импорта не должна ронять общий таймер уведомлений.
        }
    }

    /// <summary>
    /// Аналог Open-NotificationTarget (RRFQComparer.ps1, строки 704-764):
    /// показать главное окно и открыть страницу с целью уведомления.
    /// </summary>
    private void OpenNotificationTarget(DueNotification notification)
    {
        var window = Application.Current.MainWindow;
        if (window is not null)
        {
            window.Show();
            if (window.WindowState == WindowState.Minimized)
            {
                window.WindowState = WindowState.Normal;
            }

            window.Activate();
        }

        switch (notification.Source)
        {
            case "component":
                Navigate(AppPage.Components);
                (CurrentViewModel as TasksViewModel)?.OpenTask(notification.SourceId);
                break;
            case "reminder":
                Navigate(AppPage.Reminders);
                (CurrentViewModel as RemindersViewModel)?.OpenReminder(notification.SourceId);
                break;
            case "deal":
                Navigate(AppPage.Purchase);
                (CurrentViewModel as PurchaseViewModel)?.OpenDeal(notification.SourceId);
                break;
        }
    }

    /// <summary>
    /// Действия карточки меняют данные — кэшированные страницы обновляются,
    /// как Refresh-Reminders / Refresh-PurchaseDeals / Refresh-Components оригинала.
    /// </summary>
    private void RefreshPagesAfterNotification()
    {
        foreach (var vm in _pageCache.Values)
        {
            switch (vm)
            {
                case PurchaseViewModel purchase:
                    purchase.ReloadCommand.Execute(null);
                    break;
                case TasksViewModel tasks:
                    tasks.ReloadCommand.Execute(null);
                    break;
                case RemindersViewModel reminders:
                    reminders.ReloadCommand.Execute(null);
                    break;
                case BitrixViewModel bitrix:
                    bitrix.ReloadCommand.Execute(null);
                    break;
            }
        }
    }

    private void BuildNavItems()
    {
        // Порядок и подписи соответствуют оригиналу (RRFQComparer.ps1, строки 114-233).
        NavItems.Add(NavigationItem.For(AppPage.Dashboard, "Дашборд", "\uE80F"));
        NavItems.Add(NavigationItem.For(AppPage.Rrfq, "Сравнение RRFQ", "\uE8CB"));
        NavItems.Add(NavigationItem.For(AppPage.SupplierAnalytics, "Аналитика поставщиков", "\uE8A5"));
        NavItems.Add(NavigationItem.For(AppPage.Bitrix, "Bitrix API Test", "\uE753"));
        NavItems.Add(NavigationItem.For(AppPage.Purchase, "Контроль закупки", "\uE7BF"));
        NavItems.Add(NavigationItem.For(AppPage.Components, "Задачи", "\uE9D5"));
        NavItems.Add(NavigationItem.Divider());
        NavItems.Add(NavigationItem.For(AppPage.Reminders, "Напоминания", "\uE823"));
        NavItems.Add(NavigationItem.For(AppPage.QuoteBase, "База квот", "\uE8FD"));
        NavItems.Add(NavigationItem.For(AppPage.CompelParser, "Компэл Парсер", "\uE8C5"));
        NavItems.Add(NavigationItem.For(AppPage.PriceSearch, "Поиск цен", "\uE721"));
        NavItems.Add(NavigationItem.For(AppPage.CrossReference, "Поиск аналогов", "\uE71B"));
        NavItems.Add(NavigationItem.Divider());
        NavItems.Add(NavigationItem.For(AppPage.Instructions, "Инструкции", "\uE82D"));
        NavItems.Add(NavigationItem.For(AppPage.Settings, "Настройки", "\uE713"));
        NavItems.Add(NavigationItem.For(AppPage.Notes, "Заметки", "\uE70B"));
        NavItems.Add(NavigationItem.For(AppPage.History, "История", "\uE81C"));
        NavItems.Add(NavigationItem.For(AppPage.ExcelCompare, "Сравнение Excel", "\uE8CB"));
    }
}
