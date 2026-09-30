using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ProcurementControl.Models;
using ProcurementControl.Services;

namespace ProcurementControl.ViewModels;

/// <summary>
/// Страница «Bitrix API Test» — контрольная панель автоимпорта сделок из
/// проекта Bitrix «Заявка на расчет КС в РФ» в «Контроль закупки». Сам импорт
/// выполняет BitrixImportService по общему таймеру приложения; здесь —
/// переключатель, ручная проверка и список уже импортированных сделок.
/// </summary>
public partial class BitrixViewModel : ObservableObject
{
    private BitrixConfig? _config;

    /// <summary>Автоимпорт включён (настройка bitrix.import.enabled, по умолчанию — да).</summary>
    [ObservableProperty]
    private bool _autoImportEnabled = BitrixImportService.IsEnabled;

    [ObservableProperty]
    private string _statusText = string.Empty;

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(CheckNowCommand))]
    private bool _isLoading;

    [ObservableProperty]
    private BitrixImportRow? _selectedRow;

    public ObservableCollection<BitrixImportRow> Rows { get; } = new();

    public BitrixViewModel()
    {
        Reload();
        StatusText = AutoImportEnabled
            ? "Автоимпорт включён: проверка раз в 10 минут."
            : "Автоимпорт выключен.";
    }

    partial void OnAutoImportEnabledChanged(bool value)
    {
        BitrixImportService.IsEnabled = value;
        StatusText = value
            ? "Автоимпорт включён: проверка раз в 10 минут."
            : "Автоимпорт выключен.";
    }

    /// <summary>Перечитывает импортированные сделки из базы (без похода в сеть).</summary>
    [RelayCommand]
    private void Reload()
    {
        Rows.Clear();
        foreach (var row in BitrixImportService.GetImportedRows())
        {
            Rows.Add(row);
        }
    }

    /// <summary>Ручная проверка проекта Bitrix без ожидания автопроверки.</summary>
    [RelayCommand(CanExecute = nameof(CanCheckNow))]
    private async Task CheckNowAsync()
    {
        IsLoading = true;
        StatusText = "Проверка проекта Bitrix...";
        try
        {
            _config = BitrixRepository.GetConfig();
            var report = await BitrixImportService.SyncAsync(manual: true);
            if (report.Error is not null)
            {
                StatusText = "Ошибка: " + report.Error;
                ToastService.Show("Импорт из Bitrix: " + report.Error, ToastKind.Warn);
                return;
            }

            if (report.Initialized)
            {
                StatusText = "Точка отсчёта установлена: задачи, созданные до этого момента, не импортируются.";
                ToastService.Show(StatusText, ToastKind.Info);
            }
            else
            {
                var skippedNote = report.PendingSkippedCount > 0
                    ? "; не распознано: " + report.PendingSkippedCount
                    : string.Empty;
                StatusText = "Импортировано сделок: " + report.Imported.Count
                    + "; уже были: " + report.AlreadyPresent + skippedNote;
                if (report.Skipped.Count > 0)
                {
                    StatusText += "\nНе распознаны (исправьте название в Bitrix): " + string.Join("; ", report.Skipped);
                }
            }

            if (report.Imported.Count > 0)
            {
                ToastService.Show(
                    "Импорт из Bitrix (" + report.Imported.Count + "):\n" + string.Join("\n", report.Imported),
                    ToastKind.Success);
                NotificationCenter.NotifyDataChanged();
            }

            Reload();
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

    private bool CanCheckNow() => !IsLoading;

    /// <summary>Номер сделки (до первого '/') в буфер обмена — как в оригинале вкладки.</summary>
    [RelayCommand]
    private void CopyNumber()
    {
        var row = SelectedRow;
        if (row is null)
        {
            return;
        }

        var match = Regex.Match(row.DealNumber, "^\\s*([^/]+)");
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

    /// <summary>Открыть задачу в Bitrix.</summary>
    [RelayCommand]
    private void OpenLink()
    {
        var row = SelectedRow;
        if (row is null || row.TaskId <= 0)
        {
            return;
        }

        var config = _config ?? BitrixRepository.GetConfig();
        try
        {
            Process.Start(new ProcessStartInfo(BitrixService.GetTaskViewUrl(config, row.TaskId))
            {
                UseShellExecute = true,
            });
        }
        catch (Exception ex)
        {
            StatusText = "Ошибка: " + ex.Message;
        }
    }
}
