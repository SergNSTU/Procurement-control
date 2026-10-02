using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using ProcurementControl;
using ProcurementControl.Navigation;
using ProcurementControl.Services;
using ProcurementControl.ViewModels;

namespace RealShots;

internal static class Program
{
    [STAThread]
    static void Main()
    {
        var workspace = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
        typeof(AppPaths).GetField("_dataRoot", BindingFlags.NonPublic | BindingFlags.Static)!.SetValue(null, workspace);
        var app = new ProcurementControl.App { ShutdownMode = ShutdownMode.OnExplicitShutdown };
        app.InitializeComponent();
        var window = new MainWindow { Width = 1480, Height = 900, WindowStartupLocation = WindowStartupLocation.Manual, Left = -3000, Top = -3000 };
        window.Show();
        window.UpdateLayout();
        var vm = (MainViewModel)window.DataContext;
        var output = Path.Combine(workspace, "tmp", "manual_realshots", "screens");
        Directory.CreateDirectory(output);
        Capture(window, Path.Combine(output, "01_rrfq_real.png"));
        foreach (var target in new[] { AppPage.CompelParser, AppPage.Purchase, AppPage.SupplierAnalytics, AppPage.Settings, AppPage.QuoteBase, AppPage.Components })
        {
            vm.SelectedNavItem = vm.NavItems.First(x => x.Page == target);
            window.UpdateLayout();
            Capture(window, Path.Combine(output, $"{(int)target:00}_{target}.png"));
        }
        window.Close();
        app.Shutdown();
    }

    static void Capture(FrameworkElement target, string path)
    {
        var width = (int)Math.Ceiling(target.ActualWidth);
        var height = (int)Math.Ceiling(target.ActualHeight);
        if (width < 100 || height < 100) throw new InvalidOperationException($"Invalid render size: {width}x{height}");
        var image = new RenderTargetBitmap(width, height, 96, 96, PixelFormats.Pbgra32);
        image.Render(target);
        var encoder = new PngBitmapEncoder();
        encoder.Frames.Add(BitmapFrame.Create(image));
        using var stream = File.Create(path);
        encoder.Save(stream);
    }
}
