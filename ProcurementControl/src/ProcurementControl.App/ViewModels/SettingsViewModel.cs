using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;
using ProcurementControl.Models;
using ProcurementControl.Services;

namespace ProcurementControl.ViewModels;

/// <summary>
/// Страница «Настройки» — перенос $settingsPage (RRFQComparer.ps1,
/// строки 3428-3552, 3562-3583, 5756-5842). Запись — порты обработчиков:
/// $btnChooseDocumentsRoot/$btnSaveSettings (строки 5756-5776), $btnCreateBackup,
/// $btnRestoreRecovery, $btnDeleteRecovery (строки 5777-5834). Полное восстановление
/// из архива (с заменой базы) в порт не входит; из корзины записи восстанавливаются.
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

    [ObservableProperty]
    private RecoveryRow? _selectedRecoveryRow;

    /// <summary>Пока копия создаётся, кнопка заблокирована — как $btnCreateBackup.Enabled в оригинале.</summary>
    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(CreateBackupCommand))]
    private bool _isCreatingBackup;

    public ObservableCollection<RecoveryRow> RecoveryRows { get; } = new();

    public string QuickCaptureHint {
        get => $"Хоткеи: {QuickCaptureService.HotkeyMode}+T (задача) / {QuickCaptureService.HotkeyMode}+R (напоминание)";
    }

    /// <summary>Версия сборки для страницы «Настройки»: номер из csproj + дата публикации.</summary>
    public string AppVersion { get; } = BuildAppVersion();

    private static string BuildAppVersion()
    {
        var assembly = typeof(SettingsViewModel).Assembly;
        var version = assembly.GetName().Version;
        var number = version is null ? "?" : version.ToString(3);
        try
        {
            var built = File.GetLastWriteTime(assembly.Location);
            if (built.Year > 2000)
            {
                return $"Версия {number} (сборка {built:dd.MM.yyyy HH:mm})";
            }
        }
        catch
        {
            // Путь к сборке может быть недоступен — покажем только номер.
        }

        return "Версия " + number;
    }

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

    /// <summary>Аналог $btnChooseDocumentsRoot (RRFQComparer.ps1, строки 5756-5763).</summary>
    [RelayCommand]
    private void ChooseDocumentsRoot()
    {
        var dialog = new OpenFolderDialog
        {
            Title = "Выберите папку для хранения документов",
        };
        if (Directory.Exists(DocumentsRoot))
        {
            dialog.InitialDirectory = DocumentsRoot;
        }

        if (dialog.ShowDialog() == true)
        {
            DocumentsRoot = dialog.FolderName;
        }
    }

    /// <summary>
    /// Аналог $btnSaveSettings (RRFQComparer.ps1, строки 5764-5776):
    /// портативный путь документов (аналог Set-PurchaseDocumentsRoot) и флаг быстрого захвата.
    /// </summary>
    [RelayCommand]
    private void SaveSettings()
    {
        try
        {
            if (string.IsNullOrWhiteSpace(DocumentsRoot))
            {
                throw new InvalidOperationException("Укажите папку документов.");
            }

            var resolved = ResolvePortablePath(DocumentsRoot);
            Directory.CreateDirectory(resolved);
            PurchaseWriteRepository.SetSetting("documents_root", ToPortablePath(resolved));
            PurchaseWriteRepository.SetSetting("quick_capture.enabled", QuickCaptureEnabled ? "1" : "0");
            DocumentsRoot = resolved;
            MessageBox.Show("Настройки сохранены.", "Настройки");
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Настройки");
        }
    }

    /// <summary>
    /// Аналог $btnCreateBackup (RRFQComparer.ps1, строки 5777-5788):
    /// ручная резервная копия через New-PurchaseRecoverySnapshot.
    /// </summary>
    [RelayCommand(CanExecute = nameof(CanCreateBackup))]
    private async Task CreateBackupAsync()
    {
        IsCreatingBackup = true;
        try
        {
            var path = await Task.Run(() => RecoveryStore.CreateSnapshot("Backup", "Создано вручную"));
            Load();
            MessageBox.Show("Резервная копия готова:\r\n" + path, "Резервная копия");
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Резервная копия");
        }
        finally
        {
            IsCreatingBackup = false;
        }
    }

    /// <summary>
    /// Аналог $btnRestoreRecovery (RRFQComparer.ps1, строки 5789-5821) для корзины.
    /// Для полных копий — восстановление через RecoveryStore.StageRestore: архив
    /// распаковывается сразу, замена данных происходит при следующем запуске.
    /// </summary>
    [RelayCommand]
    private void RestoreSelected()
    {
        try
        {
            if (SelectedRecoveryRow is null || string.IsNullOrWhiteSpace(SelectedRecoveryRow.RecoveryType))
            {
                throw new InvalidOperationException("Выберите резервную копию или запись в корзине.");
            }

            if (SelectedRecoveryRow.RecoveryType == "Backup")
            {
                var answer = MessageBox.Show(
                    "Заменить текущие данные (база закупок, файлы документов, заметки)\r\n" +
                    "содержимым копии от " + SelectedRecoveryRow.CreatedText + " (" + SelectedRecoveryRow.Reason + ")?\r\n\r\n" +
                    "Перед заменой будет создана страховочная копия текущих данных.\r\n" +
                    "После подтверждения приложение закроется, а восстановление\r\n" +
                    "выполнится при следующем запуске.",
                    "Восстановление данных", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                if (answer != MessageBoxResult.Yes)
                {
                    return;
                }

                RecoveryStore.StageRestore(SelectedRecoveryRow.SnapshotPath);
                MessageBox.Show(
                    "Копия подготовлена. Приложение будет закрыто.\r\nПри следующем запуске данные будут восстановлены.",
                    "Восстановление данных");
                Application.Current.Shutdown();
                return;
            }

            var answerTrash = MessageBox.Show(
                "Восстановить удалённую запись и её файлы из корзины?",
                "Восстановление данных", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (answerTrash != MessageBoxResult.Yes)
            {
                return;
            }

            PurchaseWriteRepository.RestoreTrashItem(SelectedRecoveryRow.TrashId);
            Load();
            MessageBox.Show("Удалённые данные восстановлены.", "Восстановление данных");
        }
        catch (Exception ex)
        {
            ErrorLog.Append("restore", ex);
            MessageBox.Show(ex.Message, "Восстановление данных");
        }
    }

    /// <summary>
    /// Аналог $btnDeleteRecovery (RRFQComparer.ps1, строки 5822-5834): безвозвратное удаление.
    /// </summary>
    [RelayCommand]
    private void DeleteSelected()
    {
        try
        {
            if (SelectedRecoveryRow is null || string.IsNullOrWhiteSpace(SelectedRecoveryRow.RecoveryType))
            {
                throw new InvalidOperationException("Выберите архив или запись корзины.");
            }

            var isTrash = SelectedRecoveryRow.RecoveryType == "Trash";
            var question = isTrash
                ? "Удалить выбранную запись из корзины без возможности восстановления?"
                : "Удалить выбранный архив без возможности восстановления?";
            var answer = MessageBox.Show(question, "Удалить безвозвратно", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (answer != MessageBoxResult.Yes)
            {
                return;
            }

            if (isTrash)
            {
                PurchaseWriteRepository.RemoveTrashItem(SelectedRecoveryRow.TrashId);
            }
            else
            {
                RecoveryStore.RemoveSnapshot(SelectedRecoveryRow.SnapshotPath);
            }

            SelectedRecoveryRow = null;
            Load();
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Удалить архив");
        }
    }

    private bool CanCreateBackup() => !IsCreatingBackup;

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
            RefreshRecovery();
            ErrorMessage = string.Empty;
        }
        catch (Exception ex)
        {
            _repository = null;
            ErrorLog.Append("settings", ex);
            ErrorMessage = "Не удалось прочитать данные: " + ex.Message;
        }
    }

    /// <summary>
    /// Перечитывает корзину и резервные копии (файлы + SELECT из уже открытого
    /// снапшота). Вызывается при каждом входе на страницу: автокопия создаётся
    /// в фоне при старте, и без этого список оставался бы тем, каким был при
    /// первом открытии Настроек.
    /// </summary>
    public void RefreshRecovery()
    {
        try
        {
            var trash = _repository?.GetTrashItems() ?? new List<TrashItem>();

            RecoveryRows.Clear();
            foreach (var row in RecoveryStore.GetRecoveryRows(trash))
            {
                RecoveryRows.Add(row);
            }

            RecoveryHint = RecoveryStore.Hint(RecoveryRows.Count);
        }
        catch (Exception ex)
        {
            ErrorLog.Append("settings", ex);
            ErrorMessage = "Не удалось обновить список копий: " + ex.Message;
        }
    }

    /// <summary>Аналог Resolve-PortablePath (PurchaseStore.ps1, строки 52-63).</summary>
    private static string ResolvePortablePath(string path)
    {
        var text = path.Trim();
        if (Path.IsPathRooted(text))
        {
            return text;
        }

        try
        {
            return Path.GetFullPath(Path.Combine(AppPaths.AppRoot, text));
        }
        catch
        {
            return Path.Combine(AppPaths.AppRoot, text);
        }
    }

    /// <summary>Аналог Convert-ToPortablePath (PurchaseStore.ps1, строки 65-79).</summary>
    private static string ToPortablePath(string path)
    {
        try
        {
            var full = Path.GetFullPath(path);
            var root = Path.GetFullPath(AppPaths.AppRoot);
            if (!root.EndsWith('\\'))
            {
                root += '\\';
            }

            if (full.StartsWith(root, StringComparison.OrdinalIgnoreCase))
            {
                return full[root.Length..];
            }
        }
        catch
        {
            // Оставляем как есть.
        }

        return path;
    }
}
