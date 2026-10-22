using System.IO;
using System.Windows;
using System.Windows.Threading;

namespace QuickExcel.Desktop;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        DispatcherUnhandledException += OnUnhandled;
        var window = new MainWindow();
        MainWindow = window;
        window.Show();
        window.Loaded += async (_, _) => await window.StartAsync(e.Args);
    }

    private static void OnUnhandled(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        MessageBox.Show(e.Exception.Message, "QuickExcel", MessageBoxButton.OK, MessageBoxImage.Error);
        e.Handled = true;
    }
}
