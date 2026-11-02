using System.Diagnostics;
using System.Text.Json;

namespace PowerTables.Tests;

/// <summary>Viste, Intestazioni e Nuova finestra dalla ribbon con mouse reale. Usa il blocco ui.lock.</summary>
[Collection("UiCollection"), Trait("Category", "Ui")]
public class ViewUiTests
{
    static int Mode(JsonElement d) => d.GetProperty("viewMode").GetInt32();

    [Fact]
    public void Views_Headings_NewWindow_ViaRibbon()
    {
        using var lk = UiLock.Acquire();
        int before = Process.GetProcessesByName("PowerTables.App").Length;
        using var ui = UiDriver.Launch();
        ui.EnterValue(1, 1, "x", "x");
        ui.ClickUntil("Scheda Visualizza", "Intestazioni");
        ui.Retry(() => ui.Click("StatusView_page"), () => { try { return Mode(ui.Dump()) == 1; } catch (Exception e) when (e is KeyNotFoundException or InvalidOperationException) { return false; } });
        ui.Retry(() => ui.Click("Anteprima interruzioni di pagina"), () => { try { return Mode(ui.Dump()) == 2; } catch (Exception e) when (e is KeyNotFoundException or InvalidOperationException) { return false; } });
        ui.Retry(() => ui.Click("Normale"), () => { try { return Mode(ui.Dump()) == 0; } catch (Exception e) when (e is KeyNotFoundException or InvalidOperationException) { return false; } });
        ui.Retry(() => ui.Click("Intestazioni"), () => { try { return !ui.Dump().GetProperty("headings").GetBoolean(); } catch (Exception e) when (e is KeyNotFoundException or InvalidOperationException) { return false; } });
        // il file non è salvato e ha dati: la nuova finestra parte vuota in un altro processo
        ui.Retry(() => ui.Click("Nuova finestra"), () => Process.GetProcessesByName("PowerTables.App").Length >= before + 2);
        ui.KillOwn();
        foreach (var p in Process.GetProcessesByName("PowerTables.App").Skip(before)) { try { p.Kill(true); } catch { } }
    }
}
