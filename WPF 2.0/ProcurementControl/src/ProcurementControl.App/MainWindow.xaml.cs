using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using ProcurementControl.Services;
using ProcurementControl.ViewModels;

namespace ProcurementControl;

public partial class MainWindow : Window
{
    private const int WmGetMinMaxInfo = 0x0024;
    private const uint MonitorDefaultToNearest = 0x00000002;

    public MainWindow()
    {
        InitializeComponent();
        PurchaseWriteRepository.EnsureComponentOrderAmountColumn();
        DataContext = new MainViewModel();
        StateChanged += (_, _) => UpdateMaximizeIcon();
        // Глобальные хоткеи быстрого захвата требуют готового HWND.
        SourceInitialized += (_, _) =>
        {
            var source = (HwndSource)PresentationSource.FromVisual(this)!;
            source.AddHook(WindowProc);
            QuickCaptureService.Register(this);
        };
    }

    /// <summary>Аналог кнопки свернуть оригинальной формы.</summary>
    private void MinimizeWindow(object sender, RoutedEventArgs e)
        => WindowState = WindowState.Minimized;

    /// <summary>Развернуть/восстановить окно.</summary>
    private void ToggleMaximize(object sender, RoutedEventArgs e)
        => WindowState = WindowState == WindowState.Maximized
            ? WindowState.Normal
            : WindowState.Maximized;

    /// <summary>Крестик закрытия приложения.</summary>
    private void CloseWindow(object sender, RoutedEventArgs e)
        => Application.Current.Shutdown();

    private void UpdateMaximizeIcon()
    {
        // E923 = «Восстановить», E922 = «Развернуть» (Segoe MDL2 Assets).
        MaximizeButton.Content = WindowState == WindowState.Maximized ? "\uE923" : "\uE922";
        MaximizeButton.ToolTip = WindowState == WindowState.Maximized ? "Восстановить" : "Развернуть";
    }

    // Borderless WPF windows otherwise maximise to the whole monitor and can
    // cover the Windows taskbar. Restrict their maximised bounds to rcWork.
    private static IntPtr WindowProc(IntPtr hwnd, int message, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (message != WmGetMinMaxInfo)
        {
            return IntPtr.Zero;
        }

        var info = Marshal.PtrToStructure<MinMaxInfo>(lParam);
        var monitor = MonitorFromWindow(hwnd, MonitorDefaultToNearest);
        var monitorInfo = new MonitorInfo { Size = Marshal.SizeOf<MonitorInfo>() };
        if (monitor != IntPtr.Zero && GetMonitorInfo(monitor, ref monitorInfo))
        {
            var workArea = monitorInfo.WorkArea;
            var monitorArea = monitorInfo.MonitorArea;
            info.MaxPosition = new PointInt(workArea.Left - monitorArea.Left, workArea.Top - monitorArea.Top);
            info.MaxSize = new PointInt(workArea.Right - workArea.Left, workArea.Bottom - workArea.Top);
            Marshal.StructureToPtr(info, lParam, true);
        }

        handled = true;
        return IntPtr.Zero;
    }

    [DllImport("user32.dll")]
    private static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint flags);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern bool GetMonitorInfo(IntPtr monitor, ref MonitorInfo monitorInfo);

    [StructLayout(LayoutKind.Sequential)]
    private struct PointInt
    {
        public PointInt(int x, int y) => (X, Y) = (x, y);
        public int X;
        public int Y;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MinMaxInfo
    {
        public PointInt Reserved;
        public PointInt MaxSize;
        public PointInt MaxPosition;
        public PointInt MinTrackSize;
        public PointInt MaxTrackSize;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    private struct MonitorInfo
    {
        public int Size;
        public RectInt MonitorArea;
        public RectInt WorkArea;
        public int Flags;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct RectInt
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }
}
