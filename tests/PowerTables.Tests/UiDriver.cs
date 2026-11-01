using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Windows.Automation;

namespace PowerTables.Tests;

/// <summary>Pilota l'app reale con UI Automation per trovare i controlli e input REALE (SendInput) per mouse e tastiera.</summary>
public sealed class UiDriver : IDisposable
{
    [DllImport("user32.dll")] static extern int GetSystemMetrics(int i);
    static void SetCursorPos(int x, int y)
    {
        // spostamento come input reale (assoluto normalizzato), così l'app riceve WM_MOUSEMOVE veri
        int w = GetSystemMetrics(0), h = GetSystemMetrics(1);
        var i = new INPUT { type = 0, u = new InputUnion { mi = new MOUSEINPUT { dx = (int)((x + 0.5) * 65535 / w), dy = (int)((y + 0.5) * 65535 / h), flags = 0x8001 } } };
        SendInput(1, [i], Marshal.SizeOf<INPUT>());
    }
    [DllImport("user32.dll")] static extern uint SendInput(uint n, INPUT[] i, int size);
    [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
    void EnsureFocus() { if (GetForegroundWindow() != Proc.MainWindowHandle) { SetForegroundWindow(Proc.MainWindowHandle); Thread.Sleep(250); } }
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint f);
    [DllImport("user32.dll")] static extern bool SetProcessDpiAwarenessContext(IntPtr v);
    [StructLayout(LayoutKind.Sequential)] struct RECT { public int L, T, R, B; }
    [StructLayout(LayoutKind.Sequential)] struct INPUT { public uint type; public InputUnion u; }
    [StructLayout(LayoutKind.Explicit)] struct InputUnion { [FieldOffset(0)] public MOUSEINPUT mi; [FieldOffset(0)] public KEYBDINPUT ki; }
    [StructLayout(LayoutKind.Sequential)] struct MOUSEINPUT { public int dx, dy; public uint mouseData, flags, time; public IntPtr extra; }
    [StructLayout(LayoutKind.Sequential)] struct KEYBDINPUT { public ushort vk, scan; public uint flags, time; public IntPtr extra; }

    static UiDriver() { SetProcessDpiAwarenessContext(new IntPtr(-4)); } // per-monitor v2: coordinate fisiche uniformi con l'app

    public Process Proc { get; private set; } = null!;
    public AutomationElement Window { get; private set; } = null!;
    public string? DumpPath, SettingsPath;

    public static UiDriver Launch(string? file = null, string? settings = null, string[]? args = null)
    {
        var d = new UiDriver { DumpPath = Path.Combine(Path.GetTempPath(), "pt-dump-" + Guid.NewGuid().ToString("N") + ".json"), SettingsPath = settings ?? Path.Combine(Path.GetTempPath(), "pt-set-" + Guid.NewGuid().ToString("N") + ".json") };
        string exe = Path.Combine(TestUtil.Root, "src", "PowerTables.App", "bin", "Debug", "net10.0-windows", "PowerTables.App.exe");
        var psi = new ProcessStartInfo(exe) { UseShellExecute = false };
        psi.Environment["PT_TEST_HOOKS"] = "1"; psi.Environment["PT_TEST_DUMP"] = d.DumpPath; psi.Environment["PT_SETTINGS"] = d.SettingsPath;
        if (file != null) psi.ArgumentList.Add(file);
        foreach (var a in args ?? []) psi.ArgumentList.Add(a);
        d.Proc = Process.Start(psi)!;
        var sw = Stopwatch.StartNew();
        while (sw.Elapsed.TotalSeconds < 30)
        {
            Thread.Sleep(300);
            d.Proc.Refresh();
            if (d.Proc.HasExited) throw new InvalidOperationException("l'app è terminata all'avvio, codice " + d.Proc.ExitCode);
            if (d.Proc.MainWindowHandle != IntPtr.Zero)
            {
                d.Window = AutomationElement.FromHandle(d.Proc.MainWindowHandle);
                if (d.TryProbe(out _)) break;
            }
        }
        if (d.Window == null) throw new TimeoutException("finestra non trovata");
        SetForegroundWindow(d.Proc.MainWindowHandle);
        Thread.Sleep(400);
        return d;
    }

    public bool TryProbe(out Dictionary<string, string> p)
    {
        p = [];
        try
        {
            var e = Window.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.AutomationIdProperty, "Probe"));
            if (e == null) return false;
            foreach (var kv in e.Current.Name.Split(';', StringSplitOptions.RemoveEmptyEntries)) { int i = kv.IndexOf('='); if (i > 0) p[kv[..i]] = kv[(i + 1)..]; }
            return p.ContainsKey("sel");
        }
        catch (ElementNotAvailableException) { return false; }
    }

    public Dictionary<string, string> Probe() { Thread.Sleep(150); Assert.True(TryProbe(out var p), "sonda non disponibile"); return p; }

    /// <summary>Attende (fino a 6 s) che la sonda esponga key=value; altrimenti fallisce con il valore corrente.</summary>
    public void WaitProbe(string key, string value)
    {
        string cur = "", hist = "";
        for (int i = 0; i < 30; i++) { if (TryProbe(out var p) && (cur = p.GetValueOrDefault(key, "")) == value) return; hist += cur + (p.Count > 0 ? "[" + p.GetValueOrDefault("sel") + "|" + p.GetValueOrDefault("active") + "]" : "?") + ","; Thread.Sleep(200); }
        Assert.True(false, $"{key}: atteso {value}, storia {hist}");
    }

    /// <summary>Rilegge il dump finché la condizione è vera (fino a 6 s): l'app scrive il file a ogni modifica, in modo asincrono rispetto al test.</summary>
    public JsonElement DumpUntil(Func<JsonElement, bool> ok)
    {
        JsonElement d = default;
        for (int i = 0; i < 30; i++) { try { d = Dump(); if (ok(d)) return d; } catch (Exception e) when (e is KeyNotFoundException or InvalidOperationException) { } Thread.Sleep(200); }
        return d;
    }

    public JsonElement Dump()
    {
        Thread.Sleep(250);
        for (int i = 0; i < 20; i++) { try { return JsonDocument.Parse(File.ReadAllText(DumpPath!)).RootElement.Clone(); } catch (Exception e) when (e is IOException or JsonException) { Thread.Sleep(150); } }
        throw new InvalidOperationException("dump non disponibile");
    }

    public AutomationElement? Find(string name, int timeoutMs = 6000)
    {
        var sw = Stopwatch.StartNew();
        var cond = new OrCondition(new PropertyCondition(AutomationElement.NameProperty, name), new PropertyCondition(AutomationElement.AutomationIdProperty, name));
        var pid = new PropertyCondition(AutomationElement.ProcessIdProperty, Proc.Id);
        while (sw.ElapsedMilliseconds < timeoutMs)
        {
            var e = Window.FindFirst(TreeScope.Descendants, cond);
            if (e != null) return e;
            // popup e menu sono finestre di primo livello dello stesso processo
            foreach (AutomationElement w in AutomationElement.RootElement.FindAll(TreeScope.Children, pid))
            {
                e = w.FindFirst(TreeScope.Descendants, cond);
                if (e != null) return e;
            }
            Thread.Sleep(200);
        }
        return null;
    }

    /// <summary>Clic su un controllo; se il controllo "dopo" non compare (clic perso per il focus) ripete il clic fino a 3 volte.</summary>
    public void ClickUntil(string name, string expectedNext)
    {
        for (int i = 0; i < 3; i++) { Click(name); if (Find(expectedNext, 5000) != null) return; }
        throw new InvalidOperationException($"dopo il clic su {name} non compare {expectedNext}");
    }

    /// <summary>Esegue l'azione e riprova (fino a 3 volte) finché la condizione non è vera: l'input reale può perdersi se un'altra app prende il focus.</summary>
    public void Retry(Action act, Func<bool> done)
    {
        for (int i = 0; i < 3; i++) { act(); for (int w = 0; w < 10; w++) { if (done()) return; Thread.Sleep(200); } }
        Assert.True(done(), "azione non riuscita dopo 3 tentativi");
    }

    /// <summary>Seleziona la cella con un clic e digita il testo seguito da Invio; verifica il contenuto nel dump.</summary>
    public void EnterValue(int row, int col, string text, string expectedDisplay)
    {
        string a1 = ((char)('A' + col - 1)).ToString() + row;
        Retry(() => { ClickCell(row, col); Type(text); Key("{ENTER}"); }, () => { try { var d = Dump(); return d.GetProperty("cells").TryGetProperty(a1, out var c) && c.GetProperty("text").GetString() == expectedDisplay; } catch (Exception e) when (e is KeyNotFoundException or InvalidOperationException) { return false; } });
    }

    public void Click(string name)
    {
        var e = Find(name) ?? throw new InvalidOperationException("controllo non trovato: " + name);
        var r = e.Current.BoundingRectangle;
        ClickAt((int)(r.X + r.Width / 2), (int)(r.Y + r.Height / 2));
    }

    public void ClickAt(int x, int y, bool down = true, bool up = true, bool right = false)
    {
        EnsureFocus();
        SetCursorPos(x, y); Thread.Sleep(60);
        uint dn = right ? 0x0008u : 0x0002u, upf = right ? 0x0010u : 0x0004u;
        if (down) Mouse(dn); if (down && up) Thread.Sleep(40); if (up) Mouse(upf);
        Thread.Sleep(180);
    }

    public void DoubleClickAt(int x, int y) { ClickAt(x, y); ClickAt(x, y); }

    public void DragFrom(int x1, int y1, int x2, int y2)
    {
        EnsureFocus();
        SetCursorPos(x1, y1); Thread.Sleep(80); Mouse(0x0002);
        for (int i = 1; i <= 8; i++) { SetCursorPos(x1 + (x2 - x1) * i / 8, y1 + (y2 - y1) * i / 8); Thread.Sleep(40); }
        Thread.Sleep(80); Mouse(0x0004); Thread.Sleep(250);
    }

    static void Mouse(uint flags)
    {
        var i = new INPUT { type = 0, u = new InputUnion { mi = new MOUSEINPUT { flags = flags } } };
        SendInput(1, [i], Marshal.SizeOf<INPUT>());
    }

    public void Type(string text)
    {
        EnsureFocus();
        foreach (var ch in text)
        {
            var dn = new INPUT { type = 1, u = new InputUnion { ki = new KEYBDINPUT { scan = ch, flags = 0x0004 } } };
            var up = new INPUT { type = 1, u = new InputUnion { ki = new KEYBDINPUT { scan = ch, flags = 0x0004 | 0x0002 } } };
            SendInput(2, [dn, up], Marshal.SizeOf<INPUT>()); Thread.Sleep(25);
        }
        Thread.Sleep(150);
    }

    /// <summary>Combinazione di tasti: Key("^c"), Key("{ENTER}"), Key("+{TAB}"). ^ Ctrl, + Maiusc, % Alt.</summary>
    public void Key(string spec)
    {
        EnsureFocus();
        var mods = new List<ushort>(); ushort vk = 0;
        foreach (var c in spec.Where(c => c is '^' or '+' or '%')) mods.Add(c == '^' ? (ushort)0x11 : c == '+' ? (ushort)0x10 : (ushort)0x12);
        string rest = spec.TrimStart('^', '+', '%');
        vk = rest switch
        {
            "{ENTER}" => 0x0D, "{TAB}" => 0x09, "{ESC}" => 0x1B, "{DEL}" => 0x2E, "{F2}" => 0x71, "{DOWN}" => 0x28, "{UP}" => 0x26, "{LEFT}" => 0x25, "{RIGHT}" => 0x27, "{HOME}" => 0x24, "{END}" => 0x23,
            _ => (ushort)char.ToUpperInvariant(rest[0]),
        };
        var list = new List<INPUT>();
        foreach (var m in mods) list.Add(K(m, false));
        list.Add(K(vk, false)); list.Add(K(vk, true));
        for (int i = mods.Count - 1; i >= 0; i--) list.Add(K(mods[i], true));
        SendInput((uint)list.Count, [.. list], Marshal.SizeOf<INPUT>());
        Thread.Sleep(200);
    }

    static INPUT K(ushort vk, bool up) => new() { type = 1, u = new InputUnion { ki = new KEYBDINPUT { vk = vk, flags = up ? 2u : 0u } } };

    // ---- griglia: coordinate fisiche del centro di una cella con scorrimento 0 e metriche predefinite ----
    public (int X, int Y) CellCenter(int row, int col)
    {
        var p = Probe();
        var o = p["origin"].Split(','); double ox = double.Parse(o[0], System.Globalization.CultureInfo.InvariantCulture), oy = double.Parse(o[1], System.Globalization.CultureInfo.InvariantCulture);
        double s = double.Parse(p["scale"], System.Globalization.CultureInfo.InvariantCulture) * double.Parse(p["zoomscale"], System.Globalization.CultureInfo.InvariantCulture);
        return ((int)(ox + (50 + (col - 0.5) * 64) * s), (int)(oy + (22 + (row - 0.5) * 20) * s));
    }

    public void ClickCell(int row, int col) { var (x, y) = CellCenter(row, col); ClickAt(x, y); }

    public void Screenshot(string path)
    {
        GetWindowRect(Proc.MainWindowHandle, out var r);
        using var bmp = new System.Drawing.Bitmap(r.R - r.L, r.B - r.T);
        using (var g = System.Drawing.Graphics.FromImage(bmp)) { var hdc = g.GetHdc(); PrintWindow(Proc.MainWindowHandle, hdc, 2); g.ReleaseHdc(hdc); }
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        bmp.Save(path, System.Drawing.Imaging.ImageFormat.Png);
    }

    public void Dispose()
    {
        try { if (!Proc.HasExited) { Proc.CloseMainWindow(); if (!Proc.WaitForExit(3000)) { Proc.Kill(true); } } } catch { }
        try { Proc.Dispose(); } catch { }
    }

    /// <summary>Chiude senza dialogo "salva modifiche": segna il documento come non modificato non è possibile dall'esterno, quindi termina il solo processo avviato.</summary>
    public void KillOwn() { try { if (!Proc.HasExited) Proc.Kill(true); } catch { } }
}
