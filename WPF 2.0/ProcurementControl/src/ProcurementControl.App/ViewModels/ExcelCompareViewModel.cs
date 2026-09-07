using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;
using ProcurementControl.Services;

namespace ProcurementControl.ViewModels;

/// <summary>
/// Страница «Сравнение Excel» — перенос $excelComparePage (RRFQComparer.ps1,
/// строки 243-305, 4704-4789). Сравнение идёт через Excel COM в UI-потоке,
/// как в оригинале; результат сохраняется рядом с базой данных.
/// </summary>
public partial class ExcelCompareViewModel : ObservableObject
{
    [ObservableProperty]
    private string _file1Text = string.Empty;

    [ObservableProperty]
    private string _file2Text = string.Empty;

    [ObservableProperty]
    private string _statusText = "Выберите два Excel-файла.";

    [ObservableProperty]
    private string _errorMessage = string.Empty;

    [ObservableProperty]
    private bool _isComparing;

    /// <summary>Аналог $btnExcelPick1.</summary>
    [RelayCommand]
    private void PickFile1()
    {
        var path = PickExcelFile();
        if (path is not null)
        {
            File1Text = path;
        }
    }

    /// <summary>Аналог $btnExcelPick2.</summary>
    [RelayCommand]
    private void PickFile2()
    {
        var path = PickExcelFile();
        if (path is not null)
        {
            File2Text = path;
        }
    }

    /// <summary>Аналог $btnExcelCompare: сравнить и сохранить результат.</summary>
    [RelayCommand(CanExecute = nameof(CanCompare))]
    private void Compare()
    {
        try
        {
            var output = ExcelCompareService.NewOutputPath();
            IsComparing = true;
            ErrorMessage = string.Empty;
            StatusText = "Сравнение...";

            var count = ExcelCompareService.CompareWorkbooks(File1Text.Trim(), File2Text.Trim(), output);
            StatusText = $"Готово. Несовпадений: {count}. Результат: {output}";
            MessageBox.Show($"Несовпадений: {count}\r\nРезультат сохранен:\r\n{output}", "Сравнение Excel");
            ToastService.Show("Сравнение Excel завершено. Несовпадений: " + count, ToastKind.Success);
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
            MessageBox.Show(ex.Message, "Сравнение Excel");
        }
        finally
        {
            IsComparing = false;
        }
    }

    private bool CanCompare()
        => !IsComparing
            && !string.IsNullOrWhiteSpace(File1Text)
            && !string.IsNullOrWhiteSpace(File2Text);

    partial void OnFile1TextChanged(string value) => CompareCommand.NotifyCanExecuteChanged();

    partial void OnFile2TextChanged(string value) => CompareCommand.NotifyCanExecuteChanged();

    /// <summary>Диалог выбора файла с фильтром оригинала.</summary>
    private static string? PickExcelFile()
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls|All files (*.*)|*.*",
        };
        return dialog.ShowDialog() == true ? dialog.FileName : null;
    }
}
