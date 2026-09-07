using System.Windows;
using ProcurementControl.Services;

namespace ProcurementControl;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
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

        base.OnStartup(e);
    }
}
