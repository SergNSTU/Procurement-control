using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ProcurementControl.Models;
using ProcurementControl.Services;

namespace ProcurementControl.ViewModels;

/// <summary>
/// Read-only страница «Задачи» (аналог вкладки «Компоненты» оригинала).
/// Данные читаются одним снапшотом; поиск выполняется в памяти, чтобы не
/// копировать БД на каждый символ. Исходная база не изменяется.
/// </summary>
public partial class TasksViewModel : ObservableObject
{
    private List<TaskRow> _allRows = new();

    [ObservableProperty]
    private string _searchText = string.Empty;

    [ObservableProperty]
    private TaskRow? _selectedTask;

    [ObservableProperty]
    private string _errorMessage = string.Empty;

    [ObservableProperty]
    private string _countText = string.Empty;

    public string ReadOnlyNotice { get; } =
        "Режим чтения: изменение доступно в основной версии приложения";

    public ObservableCollection<TaskRow> Tasks { get; } = new();

    /// <summary>Записи выбранной задачи (поле notes) — только просмотр.</summary>
    public string SelectedNotes => SelectedTask?.Notes ?? string.Empty;

    public TasksViewModel()
    {
        Load();
    }

    partial void OnSearchTextChanged(string value)
    {
        ApplyFilter();
    }

    partial void OnSelectedTaskChanged(TaskRow? value)
    {
        OnPropertyChanged(nameof(SelectedNotes));
    }

    [RelayCommand]
    private void Reload()
    {
        Load();
    }

    private void Load()
    {
        try
        {
            using var snapshot = PurchaseSnapshot.Create();
            var repo = new PurchaseRepository(snapshot.Connection);
            _allRows = repo.GetTasks(string.Empty);

            ErrorMessage = string.Empty;
            ApplyFilter();
        }
        catch (Exception ex)
        {
            ErrorMessage = "Не удалось прочитать данные: " + ex.Message;
        }
    }

    /// <summary>Аналог LIKE '%...%' оригинала по сделке и описанию, без учёта регистра.</summary>
    private void ApplyFilter()
    {
        var needle = SearchText.Trim();

        var selectedId = SelectedTask?.Id;
        Tasks.Clear();
        foreach (var row in _allRows)
        {
            if (needle.Length == 0
                || row.DealNumber.Contains(needle, StringComparison.OrdinalIgnoreCase)
                || row.Description.Contains(needle, StringComparison.OrdinalIgnoreCase))
            {
                Tasks.Add(row);
            }
        }

        CountText = "Всего: " + Tasks.Count;
        SelectedTask = selectedId is null
            ? null
            : Tasks.FirstOrDefault(r => r.Id == selectedId);
        OnPropertyChanged(nameof(SelectedNotes));
    }
}
