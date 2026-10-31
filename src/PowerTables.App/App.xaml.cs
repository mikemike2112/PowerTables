using System.Windows;

namespace PowerTables.App;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        DispatcherUnhandledException += (_, ex) =>
        {
            MessageBox.Show(ex.Exception.Message, "PowerTables", MessageBoxButton.OK, MessageBoxImage.Error);
            ex.Handled = true;
        };
        var args = e.Args;
        if (args.Contains("--light")) Theme.Dark = false;
        Theme.Apply(Resources);
        var w = new MainWindow();
        w.Show();
        int bi = Array.IndexOf(args, "--bench");
        var file = bi >= 0 ? null : args.FirstOrDefault(a => !a.StartsWith("--"));
        if (file != null) w.OpenFile(file);
        if (bi >= 0 && bi + 2 < args.Length) w.RunBench(args[bi + 1], args[bi + 2]);
    }
}
