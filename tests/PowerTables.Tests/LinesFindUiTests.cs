using System.Text.Json;

namespace PowerTables.Tests;

/// <summary>Inserisci/elimina righe e colonne e Trova/Sostituisci con mouse e tastiera reali. Usa il blocco ui.lock.</summary>
[Collection("UiCollection"), Trait("Category", "Ui")]
public class LinesFindUiTests
{
    static string Cell(JsonElement d, string a) => d.GetProperty("cells").TryGetProperty(a, out var c) ? c.GetProperty("text").GetString()! : "";
    static string F(JsonElement d, string a) => d.GetProperty("cells").TryGetProperty(a, out var c) && c.TryGetProperty("f", out var f) ? f.GetString() ?? "" : "";

    [Fact]
    public void InsertDeleteRowsCols_And_ReplaceAll_ViaRibbonAndShortcut()
    {
        using var lk = UiLock.Acquire();
        using var ui = UiDriver.Launch();
        ui.EnterValue(1, 1, "10", "10"); ui.EnterValue(2, 1, "20", "20"); ui.EnterValue(3, 1, "=A1+A2", "30");
        ui.EnterValue(1, 2, "gatto", "gatto"); ui.EnterValue(2, 2, "gatto nero", "gatto nero");

        // Home > Inserisci > Inserisci righe con A1 attiva: tutto scende di una riga, la formula segue
        ui.ClickCell(1, 1);
        ui.Retry(() => ui.Click("Cells_Insert"), () => ui.Find("Inserisci righe", 800) != null);
        ui.Click("Inserisci righe");
        var d = ui.DumpUntil(x => Cell(x, "A4") == "30");
        Assert.Equal("", Cell(d, "A1")); Assert.Equal("10", Cell(d, "A2"));
        Assert.Equal("=A2+A3", F(d, "A4"));
        // un solo Ctrl+Z annulla l'inserimento
        ui.Key("^z");
        d = ui.DumpUntil(x => Cell(x, "A3") == "30");
        Assert.Equal("=A1+A2", F(d, "A3")); Assert.Equal("10", Cell(d, "A1"));

        // Home > Elimina > Elimina colonne con B1 attiva: la colonna B sparisce, A resta
        ui.ClickCell(1, 2);
        ui.Retry(() => ui.Click("Cells_Delete"), () => ui.Find("Elimina colonne", 800) != null);
        ui.Click("Elimina colonne");
        d = ui.DumpUntil(x => Cell(x, "B1") == "");
        Assert.Equal("", Cell(d, "B2")); Assert.Equal("30", Cell(d, "A3"));
        ui.Key("^z");
        d = ui.DumpUntil(x => Cell(x, "B1") == "gatto");

        // Ctrl+H: Sostituisci tutto su tutto il foglio, un solo Ctrl+Z lo annulla
        ui.ClickCell(5, 5);
        ui.Retry(() => ui.Key("^h"), () => ui.Find("ReplaceText", 800) != null);
        ui.SetValueId("FindText", "gatto"); ui.SetValueId("ReplaceText", "cane");
        ui.InvokeId("Dlg_ReplaceAll");
        d = ui.DumpUntil(x => Cell(x, "B2") == "cane nero");
        Assert.Equal("cane", Cell(d, "B1"));
        Assert.Equal("2 sostituzioni.", ui.NameOf("FindStatus"));
        ui.ScreenCaptureWindowOf("FindText", Path.Combine(TestUtil.Root, "tests", "work", "docs", "sostituisci.png"));
        ui.InvokeId("Dlg_FindClose");
        ui.Key("^z");
        d = ui.DumpUntil(x => Cell(x, "B2") == "gatto nero");
        Assert.Equal("gatto", Cell(d, "B1"));
        ui.KillOwn();
    }
}
