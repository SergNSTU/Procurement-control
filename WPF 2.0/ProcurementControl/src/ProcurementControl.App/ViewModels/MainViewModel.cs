using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
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
        // Оригинальное приложение стартует на разделе «Сравнение RRFQ».
        Navigate(AppPage.Rrfq);
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
            AppPage.Rrfq => new RrfqViewModel(),
            AppPage.QuoteBase => new QuoteBaseViewModel(),
            AppPage.Settings => new SettingsViewModel(),
            AppPage.Notes => new NotesViewModel(_notesService),
            _ => new PlaceholderViewModel(GetTitle(page))
        };

        _pageCache[page] = vm;
        return vm;
    }

    private string GetTitle(AppPage page)
        => NavItems.FirstOrDefault(n => !n.IsDivider && n.Page == page)?.Title ?? page.ToString();

    private void BuildNavItems()
    {
        // Порядок и подписи соответствуют оригиналу (RRFQComparer.ps1, строки 114-233).
        NavItems.Add(NavigationItem.For(AppPage.Dashboard, "Дашборд", "\uE80F"));
        NavItems.Add(NavigationItem.For(AppPage.Rrfq, "Сравнение RRFQ", "\uE8CB"));
        NavItems.Add(NavigationItem.For(AppPage.Bitrix, "Bitrix API Test", "\uE753"));
        NavItems.Add(NavigationItem.For(AppPage.Purchase, "Контроль закупки", "\uE7BF"));
        NavItems.Add(NavigationItem.For(AppPage.Components, "Задачи", "\uE9D5"));
        NavItems.Add(NavigationItem.Divider());
        NavItems.Add(NavigationItem.For(AppPage.Reminders, "Напоминания", "\uE823"));
        NavItems.Add(NavigationItem.For(AppPage.QuoteBase, "База квот", "\uE8FD"));
        NavItems.Add(NavigationItem.For(AppPage.CompelParser, "Компэл Парсер", "\uE8C5"));
        NavItems.Add(NavigationItem.For(AppPage.PriceSearch, "Поиск цен", "\uE721"));
        NavItems.Add(NavigationItem.Divider());
        NavItems.Add(NavigationItem.For(AppPage.Instructions, "Инструкции", "\uE82D"));
        NavItems.Add(NavigationItem.For(AppPage.Settings, "Настройки", "\uE713"));
        NavItems.Add(NavigationItem.For(AppPage.Notes, "Заметки", "\uE70B"));
        NavItems.Add(NavigationItem.For(AppPage.History, "История", "\uE81C"));
        NavItems.Add(NavigationItem.For(AppPage.ExcelCompare, "Сравнение Excel", "\uE8CB"));
    }
}
