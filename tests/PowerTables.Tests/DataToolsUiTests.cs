using System.Text.Json;

namespace PowerTables.Tests;

/// <summary>Rimuovi duplicati, Definisci nome e Somma automatica dalla ribbon con mouse e tastiera reali. Usa il blocco ui.lock.</summary>
[Collection("UiCollection"), Trait("Category", "Ui")]
public class DataToolsUiTests
{
    static string Cell(JsonElement d, string a) => d.GetProperty("cells").TryGetProperty(a, out var c) ? c.GetProperty("text").GetString()! : "";

    [Fact]
    public void RemoveDuplicates_DefineName_AutoSum_ViaRibbon()
    {
        using var lk = UiLock.Acquire();
        using var ui = UiDriver.Launch();
        ui.EnterValue(1, 1, "Nome", "Nome"); ui.EnterValue(2, 1, "x", "x"); ui.EnterValue(3, 1, "y", "y"); ui.EnterValue(4, 1, "x", "x");

        // Dati > Rimuovi duplicati: l'area si ricava dalla cella attiva, intestazione presente
        ui.ClickCell(1, 1);
        ui.ClickUntil("Scheda Dati", "Rimuovi duplicati");
        ui.Retry(() => ui.Click("Rimuovi duplicati"), () => ui.Find("Dlg_Ok", 800) != null);
        ui.InvokeId("Dlg_Ok");                                   // conferma la finestra
        ui.Retry(() => { }, () => ui.Find("Dlg_Ok", 800) != null);
        ui.InvokeId("Dlg_Ok");                                   // chiude il riepilogo
        var d = ui.DumpUntil(x => Cell(x, "A4") == "");
        Assert.Equal("y", Cell(d, "A3"));
        ui.Key("^z");
        d = ui.DumpUntil(x => Cell(x, "A4") == "x");

        // Formule > Definisci nome, poi uso del nome in una formula
        ui.EnterValue(1, 3, "4", "4"); ui.EnterValue(2, 3, "6", "6");
        ui.ClickCell(1, 3); ui.Key("+{DOWN}");                  // C1:C2
        ui.ClickUntil("Scheda Formule", "Definisci nome");
        ui.Retry(() => ui.Click("Definisci nome"), () => ui.Find("NameName", 800) != null);
        ui.SetValueId("NameName", "Totale");
        ui.InvokeId("Dlg_Ok");
        ui.EnterValue(1, 5, "=SUM(Totale)", "10");

        // Somma automatica con una cella sola: usa i numeri sopra
        ui.EnterValue(1, 7, "1", "1"); ui.EnterValue(2, 7, "2", "2");
        ui.ClickCell(3, 7);
        ui.ClickUntil("Scheda Home", "Somma automatica");
        ui.Click("Somma automatica");
        d = ui.DumpUntil(x => Cell(x, "G3") == "3");
        ui.Key("^z");
        d = ui.DumpUntil(x => Cell(x, "G3") == "");
        ui.KillOwn();
    }
}
