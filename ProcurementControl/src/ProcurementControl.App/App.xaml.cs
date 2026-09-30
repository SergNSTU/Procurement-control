using System.Globalization;
using System.Windows;
using ProcurementControl.Services;

namespace ProcurementControl;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        // Без этих обработчиков любое не пойманное исключение UI-потока молча
        // завершает процесс («иногда вылетает»). Ошибка пишется в лог, интерфейс
        // остаётся живым; фатальные — тоже логируются.
        DispatcherUnhandledException += (_, args) =>
        {
            ErrorLog.Append("ui", args.Exception);
            MessageBox.Show(
                "Ошибка интерфейса (записана в лог):\n" + args.Exception.Message,
                "Procurement Control",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            args.Handled = true;
        };
        AppDomain.CurrentDomain.UnhandledException += (_, args) =>
            ErrorLog.Append("fatal", args.ExceptionObject as Exception);
        TaskScheduler.UnobservedTaskException += (_, args) =>
        {
            ErrorLog.Append("task", args.Exception);
            args.SetObserved();
        };

        // Отложенное восстановление из резервной копии (Настройки → «Восстановить»):
        // выполняется до любых подключений к базе, пока файлы никем не заняты.
        try
        {
            RecoveryStore.ApplyPendingRestore();
        }
        catch (Exception exception)
        {
            ErrorLog.Append("restore", exception);
            MessageBox.Show(
                "Не удалось восстановить данные из резервной копии:\n" + exception.Message +
                "\nПодробности в data\\logs\\errors.log.",
                "Procurement Control",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }

        // Запуск «с нуля»: если общей базы закупок нет (например, программу
        // перенесли на другой компьютер без папки data), создаём каталоги и
        // пустую базу со схемой до открытия первого окна. MainWindow в
        // конструкторе сразу обращается к БД, поэтому подготовка должна
        // произойти раньше, чем StartupUri создаст окно.
        try
        {
            DatabaseInitializer.EnsureCreated();
        }
        catch (Exception exception)
        {
            MessageBox.Show(
                "Не удалось подготовить данные приложения:\n" + exception.Message,
                "Procurement Control",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            Shutdown();
            return;
        }

        RunStartupAutoBackup();

        base.OnStartup(e);
    }

    /// <summary>
    /// Автоматическая резервная копия раз в сутки: отметка о последней автокопии
    /// хранится в настройках базы. В фоновом потоке, чтобы не задерживать старт;
    /// сбой автокопии пишется в лог, но не мешает работе приложения.
    /// </summary>
    private static void RunStartupAutoBackup()
    {
        _ = Task.Run(() =>
        {
            try
            {
                string? lastDay;
                using (var snapshot = PurchaseSnapshot.Create())
                {
                    lastDay = new SettingsRepository(snapshot.Connection)
                        .GetSetting("recovery.last_auto_backup");
                }

                var today = DateTime.Now.ToString("yyyyMMdd", CultureInfo.InvariantCulture);
                if (lastDay == today)
                {
                    return;
                }

                RecoveryStore.CreateSnapshot("Backup", "Автоматическая копия при запуске");
                PurchaseWriteRepository.SetSetting("recovery.last_auto_backup", today);
                ErrorLog.Append("backup", "Автоматическая резервная копия создана.");
            }
            catch (Exception exception)
            {
                ErrorLog.Append("backup", exception);
            }
        });
    }
}
