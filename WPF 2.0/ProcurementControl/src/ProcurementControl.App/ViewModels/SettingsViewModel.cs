using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ProcurementControl.Models;
using ProcurementControl.Services;

namespace ProcurementControl.ViewModels;

/// <summary>
/// Read-only страница «Настройки» — перенос $settingsPage (RRFQComparer.ps1,
/// строки 3428-3552, 3562-3583, 5756-5842). Кнопки изменения данных
/// («Выбрать...», «Сохранить настройки», «Создать копию», «Восстановить
/// выбранное», «Удалить безвозвратно») в порт не входят. Снапшот БД живёт
/// всё время страницы; исходная база не изменяется.
/// </summary>
public partial class SettingsViewModel : ObservableObject
{
    private PurchaseSnapshot? _snapshot;
    private SettingsRepository? _repository;

    [ObservableProperty]
    private string _documentsRoot = string.Empty;

    [ObservableProperty]
    private bool _quickCaptureEnabled = true;

    [ObservableProperty]
    private string _errorMessage = string.Empty;

    [ObservableProperty]
    private string _recoveryHint = string.Empty;

    public ObservableCollection<RecoveryRow> RecoveryRows { get; } = new();

    public string QuickCaptureHint { get; } =
        "Хоткеи: Ctrl+Shift+T / Ctrl+Shift+R; агент: порт 8765";

    public string ReadOnlyNotice { get; } =
        "Режим чтения: изменение доступно в основной версии приложения";

    public SettingsViewModel()
    {
        Load();
    }

    /// <summary>Аналог $btnOpenRecoveryFolder: открыть папку копий в проводнике.</summary>
    [RelayCommand]
    private void OpenFolder()
    {
        try
        {
            RecoveryStore.OpenFolder();
            ErrorMessage = string.Empty;
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
        }
    }

    private void Load()
    {
        try
        {
            _snapshot?.Dispose();
            _snapshot = PurchaseSnapshot.Create();
            _repository = new SettingsRepository(_snapshot.Connection);

            DocumentsRoot = _repository.GetDocumentsRoot();
            QuickCaptureEnabled = _repository.GetBooleanSetting("quick_capture.enabled", defaultValue: true);

            RecoveryRows.Clear();
            var trash = _repository.GetTrashItems();
            foreach (var row in RecoveryStore.GetRecoveryRows(trash))
            {
                RecoveryRows.Add(row);
            }

            RecoveryHint = RecoveryStore.Hint(RecoveryRows.Count);
            ErrorMessage = string.Empty;
        }
        catch (Exception ex)
        {
            _repository = null;
            ErrorMessage = "Не удалось прочитать данные: " + ex.Message;
        }
    }
}
