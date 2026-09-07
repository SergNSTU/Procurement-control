using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ProcurementControl.Models;
using ProcurementControl.Services;

namespace ProcurementControl.ViewModels;

/// <summary>
/// Страница «Bitrix API Test» — перенос $bitrixPage (RRFQComparer.ps1,
/// строки 306-518). Интеграция по умолчанию отключена, как и в оригинале
/// (строка 1385): включается чекбоксом только на текущий сеанс.
/// </summary>
public partial class BitrixViewModel : ObservableObject
{
    private readonly DispatcherTimer _autoRefreshTimer;
    private BitrixConfig? _config;

    [ObservableProperty]
    private bool _integrationEnabled;

    [ObservableProperty]
    private string _statusText = string.Empty;

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(RefreshCommand))]
    private bool _isLoading;

    [ObservableProperty]
    private BitrixTaskRow? _selectedTask;

    public ObservableCollection<BitrixTaskRow> Tasks { get; } = new();

    public BitrixViewModel()
    {
        // Автообновление каждые 30 минут (аналог $bitrixTimer).
        _autoRefreshTimer = new DispatcherTimer { Interval = TimeSpan.FromMinutes(30) };
        _autoRefreshTimer.Tick += (_, _) =>
        {
            if (IntegrationEnabled && !IsLoading)
            {
                _ = RefreshAsync();
            }
        };
    }

    partial void OnIntegrationEnabledChanged(bool value)
    {
        RefreshCommand.NotifyCanExecuteChanged();
        if (value)
        {
            StatusText = "Подтягивание из Bitrix включено. Нажмите кнопку обновления.";
            _autoRefreshTimer.Start();
        }
        else
        {
            _autoRefreshTimer.Stop();
            Tasks.Clear();
            StatusText = "Подтягивание из Bitrix отключено.";
        }
    }

    /// <summary>Аналог $btnBitrixRefresh.Add_Click.</summary>
    [RelayCommand(CanExecute = nameof(CanRefresh))]
    private async Task RefreshAsync()
    {
        if (!IntegrationEnabled)
        {
            StatusText = "Подтягивание из Bitrix отключено.";
            return;
        }

        IsLoading = true;
        StatusText = "Загрузка...";
        Tasks.Clear();
        try
        {
            _config = BitrixRepository.GetConfig();
            var blocked = BitrixRepository.GetBlockedTaskIds();
            var rows = await BitrixService.GetTasksAsync(_config, blocked);
            foreach (var row in rows)
            {
                Tasks.Add(row);
            }

            StatusText = $"Задач: {Tasks.Count} в работе";
        }
        catch (Exception ex)
        {
            StatusText = "Ошибка: " + ex.Message;
        }
        finally
        {
            IsLoading = false;
        }
    }

    private bool CanRefresh() => IntegrationEnabled && !IsLoading;

    /// <summary>Аналог $bitrixCopyDealNumberMenuItem: номер до первого '/'.</summary>
    [RelayCommand]
    private void CopyNumber()
    {
        var task = SelectedTask;
        if (task is null)
        {
            return;
        }

        var match = Regex.Match(task.Title, "^\\s*([^/]+)");
        if (!match.Success)
        {
            return;
        }

        try
        {
            Clipboard.SetText(match.Groups[1].Value.Trim());
        }
        catch
        {
            // Буфер обмена может быть занят другим приложением.
        }
    }

    /// <summary>Аналог $bitrixBlockTaskMenuItem: подтверждение и блокировка.</summary>
    [RelayCommand]
    private void BlockTask()
    {
        var task = SelectedTask;
        if (task is null || task.TaskId <= 0)
        {
            return;
        }

        var answer = MessageBox.Show(
            $"Заблокировать задачу ID {task.TaskId}?", "Bitrix",
            MessageBoxButton.YesNo, MessageBoxImage.Warning);
        if (answer != MessageBoxResult.Yes)
        {
            return;
        }

        try
        {
            BitrixRepository.BlockTask(task.TaskId, task.Title);
            Tasks.Remove(task);
        }
        catch (Exception ex)
        {
            StatusText = "Ошибка: " + ex.Message;
        }
    }

    /// <summary>Аналог $bitrixGrid.Add_CellDoubleClick: открыть ссылку в браузере.</summary>
    [RelayCommand]
    private void OpenLink()
    {
        var task = SelectedTask;
        if (task is null || task.TaskId <= 0)
        {
            return;
        }

        var config = _config ?? BitrixRepository.GetConfig();
        var url = BitrixService.GetTaskViewUrl(config, task.TaskId);
        try
        {
            Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            StatusText = "Ошибка: " + ex.Message;
        }
    }
}
