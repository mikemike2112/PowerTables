using System.Diagnostics;
using System.Text.Json;

namespace PowerTables.Tests;

public static class TestUtil
{
    public static string Root
    {
        get
        {
            var d = new DirectoryInfo(AppContext.BaseDirectory);
            while (d != null && !File.Exists(Path.Combine(d.FullName, "PowerTables.sln"))) d = d.Parent;
            return d?.FullName ?? throw new InvalidOperationException("radice repo non trovata");
        }
    }

    public static string Corpus(params string[] p) => Path.Combine([Root, "tests", "corpus", .. p]);

    /// <summary>Esegue uno script PowerShell di tools/ (Excel COM) e restituisce (exitcode, output).</summary>
    public static (int Code, string Output) Ps(string script, params string[] args) => PsT(script, 300_000, args);

    public static (int Code, string Output) PsT(string script, int timeoutMs, params string[] args)
    {
        var pidFile = Path.Combine(Path.GetTempPath(), "pt-xl-" + Guid.NewGuid().ToString("N") + ".pid");
        var psi = new ProcessStartInfo("powershell.exe") { RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false };
        psi.Environment["PT_PIDFILE"] = pidFile;
        foreach (var a in new[] { "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", Path.Combine(Root, "tools", script) }.Concat(args)) psi.ArgumentList.Add(a);
        using var p = Process.Start(psi)!;
        var so = p.StandardOutput.ReadToEndAsync(); var se = p.StandardError.ReadToEndAsync();
        if (!p.WaitForExit(timeoutMs))
        {
            try { p.Kill(true); } catch { }
            // watchdog: termina SOLO l'istanza Excel avviata dallo script (PID registrato), mai per nome
            try { if (File.Exists(pidFile) && int.TryParse(File.ReadAllText(pidFile).Trim(), out var xp) && Process.GetProcessById(xp) is { ProcessName: "EXCEL" } x) x.Kill(); } catch { }
            throw new TimeoutException("script " + script);
        }
        try { File.Delete(pidFile); } catch { }
        return (p.ExitCode, so.Result + se.Result);
    }

    public static JsonDocument ReadJson(string path) => JsonDocument.Parse(File.ReadAllText(path).TrimStart('﻿'));
}
