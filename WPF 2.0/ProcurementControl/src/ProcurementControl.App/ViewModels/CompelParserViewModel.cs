using System.Collections.ObjectModel;
using System.IO;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;
using ProcurementControl.Models;
using ProcurementControl.Services;

namespace ProcurementControl.ViewModels;

/// <summary>
/// Страница «Компэл Парсер» — перенос $compelParserPage (RRFQComparer.ps1,
/// строки 2769-3099). Читает выбранный пользователем HTML-файл расчёта Компэла,
/// показывает позиции и сводку; экспорт создаёт новый xlsx. База данных не используется.
/// </summary>
public partial class CompelParserViewModel : ObservableObject
{
    private CompelSummary? _summary;
    private string _lastExportPath = string.Empty;

    [ObservableProperty]
    private string _htmlPath = string.Empty;

    [ObservableProperty]
    private string _statusText = "Выберите сохраненный HTML-файл расчета SDS Compel.";

    [ObservableProperty]
    private bool _isStatusError;

    [ObservableProperty]
    private string _totalRowsText = "0";

    [ObservableProperty]
    private string _pricedRowsText = "0";

    [ObservableProperty]
    private string _missingRowsText = "0";

    [ObservableProperty]
    private string _selectedTotalUsdText = "0.00";

    [ObservableProperty]
    private string _cheapTotalUsdText = string.Empty;

    [ObservableProperty]
    private string _optimalTotalUsdText = string.Empty;

    [ObservableProperty]
    private bool _canOpenFolder;

    public ObservableCollection<CompelRow> Rows { get; } = new();

    /// <summary>Выбор файла сразу запускает разбор, как в оригинале (строки 3004-3019).</summary>
    [RelayCommand]
    private void ChooseHtml()
    {
        var dialog = new OpenFileDialog
        {
            Title = "Выберите HTML-файл SDS Compel",
            Filter = "HTML files (*.html;*.htm)|*.html;*.htm|All files (*.*)|*.*",
            Multiselect = false,
        };
        if (dialog.ShowDialog() != true)
        {
            return;
        }

        try
        {
            ParseFile(dialog.FileName);
        }
        catch (Exception ex)
        {
            ResetResult();
            SetStatus(ex.Message, isError: true);
        }
    }

    [RelayCommand]
    private void Parse()
    {
        try
        {
            ParseFile(HtmlPath);
        }
        catch (Exception ex)
        {
            SetStatus(ex.Message, isError: true);
        }
    }

    [RelayCommand(CanExecute = nameof(CanExport))]
    private void ExportExcel()
    {
        try
        {
            if (_summary is null || Rows.Count == 0)
            {
                throw new InvalidOperationException("Сначала разберите HTML-файл Компэла.");
            }

            var dialog = new SaveFileDialog
            {
                Title = "Сохранить Excel по Компэлу",
                Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                DefaultExt = "xlsx",
                AddExtension = true,
            };
            var sourceStem = string.IsNullOrWhiteSpace(HtmlPath)
                ? "compel"
                : Path.GetFileNameWithoutExtension(HtmlPath);
            dialog.FileName = sourceStem + "_compel_prices.xlsx";
            if (!string.IsNullOrWhiteSpace(HtmlPath))
            {
                var sourceDir = Path.GetDirectoryName(HtmlPath);
                if (!string.IsNullOrWhiteSpace(sourceDir) && Directory.Exists(sourceDir))
                {
                    dialog.InitialDirectory = sourceDir;
                }
            }
            if (dialog.ShowDialog() != true)
            {
                return;
            }

            SetStatus("Создаю Excel...", isError: false);
            var path = CompelExcelExporter.WriteWorkbook(dialog.FileName, Rows, _summary);
            _lastExportPath = path;
            CanOpenFolder = true;
            OpenExportFolderCommand.NotifyCanExecuteChanged();
            SetStatus("Excel готов: " + path, isError: false);
        }
        catch (Exception ex)
        {
            SetStatus(ex.Message, isError: true);
        }
    }

    [RelayCommand(CanExecute = nameof(CanOpenExportFolder))]
    private void OpenExportFolder()
    {
        try
        {
            if (string.IsNullOrWhiteSpace(_lastExportPath) || !File.Exists(_lastExportPath))
            {
                throw new InvalidOperationException("Сначала выгрузите Excel.");
            }
            CompelExcelExporter.RevealInExplorer(_lastExportPath);
        }
        catch (Exception ex)
        {
            SetStatus(ex.Message, isError: true);
        }
    }

    private bool CanExport() => _summary is not null && Rows.Count > 0;

    private bool CanOpenExportFolder() => CanOpenFolder;

    /// <summary>Перенос Invoke-CompelHtmlParse (строки 2977-3002).</summary>
    private void ParseFile(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new InvalidOperationException("Выберите HTML-файл расчета SDS Compel.");
        }

        SetStatus("Разбираю HTML...", isError: false);
        var result = CompelParserService.Parse(path);
        _summary = result.Summary;
        HtmlPath = path;

        Rows.Clear();
        foreach (var row in result.Rows)
        {
            Rows.Add(row);
        }

        RefreshSummaryMetrics(result.Summary);
        SetStatus(
            $"Готово: найдено {result.Summary.TotalRows} строк, {result.Summary.PricedRows} с ценой.",
            isError: false);
        ExportExcelCommand.NotifyCanExecuteChanged();
    }

    /// <summary>Перенос Refresh-CompelSummaryView (строки 2933-2952).</summary>
    private void RefreshSummaryMetrics(CompelSummary? summary)
    {
        if (summary is null)
        {
            TotalRowsText = "0";
            PricedRowsText = "0";
            MissingRowsText = "0";
            SelectedTotalUsdText = "0.00";
            CheapTotalUsdText = string.Empty;
            OptimalTotalUsdText = string.Empty;
            return;
        }

        TotalRowsText = summary.TotalRows.ToString();
        PricedRowsText = summary.PricedRows.ToString();
        MissingRowsText = summary.MissingRows.Count.ToString();
        SelectedTotalUsdText = FormatMetric(summary.SelectedTotalUsd);
        CheapTotalUsdText = summary.CheapTotalUsd is null ? string.Empty : FormatMetric(summary.CheapTotalUsd.Value);
        OptimalTotalUsdText = summary.OptimalTotalUsd is null ? string.Empty : FormatMetric(summary.OptimalTotalUsd.Value);
    }

    private static string FormatMetric(double value)
        => CompelRow.FormatNumber(value, 2);

    private void ResetResult()
    {
        _summary = null;
        Rows.Clear();
        RefreshSummaryMetrics(null);
        ExportExcelCommand.NotifyCanExecuteChanged();
    }

    /// <summary>Перенос Set-CompelStatus (строки 2927-2931): ошибка — красным.</summary>
    private void SetStatus(string text, bool isError)
    {
        StatusText = text;
        IsStatusError = isError;
    }
}
