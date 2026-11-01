using System.Windows;

namespace PowerTables.App;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        DispatcherUnhandledException += (_, ex) =>
        {
            ThemedDialog.Notify(null, ex.Exception.Message);
            ex.Handled = true;
        };
        var args = e.Args;
        Theme.Load();
        if (args.Contains("--light")) Theme.Dark = false;
        Theme.Apply(Resources);
        var w = new MainWindow();
        w.Show();
        int di = Array.IndexOf(args, "--demo-dialog");
        if (di >= 0 && di + 1 < args.Length) Dispatcher.BeginInvoke(new Action(() => w.DemoDialog(args[di + 1])), System.Windows.Threading.DispatcherPriority.ApplicationIdle);
        int bi = Array.IndexOf(args, "--bench");
        var file = bi >= 0 ? null : args.FirstOrDefault(a => !a.StartsWith("--"));
        if (file != null) w.OpenFile(file);
        if (bi >= 0 && bi + 2 < args.Length) w.RunBench(args[bi + 1], args[bi + 2]);
    }
}
