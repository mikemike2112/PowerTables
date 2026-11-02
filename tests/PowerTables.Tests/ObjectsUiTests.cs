using System.Text.Json;

namespace PowerTables.Tests;

/// <summary>Inserisci grafico, forma, casella di testo e commento dalla ribbon con mouse e tastiera reali. Usa il blocco ui.lock.</summary>
[Collection("UiCollection"), Trait("Category", "Ui")]
public class ObjectsUiTests
{
    static bool Has(JsonElement d, string prop, string needle) => d.TryGetProperty(prop, out var a) && a.EnumerateArray().Any(x => x.GetString()!.Contains(needle));

    [Fact]
    public void Chart_Shape_TextBox_Comment_ViaRibbon()
    {
        using var lk = UiLock.Acquire();
        using var ui = UiDriver.Launch();
        ui.EnterValue(1, 1, "m", "m"); ui.EnterValue(1, 2, "v", "v");
        ui.EnterValue(2, 1, "a", "a"); ui.EnterValue(2, 2, "5", "5"); ui.EnterValue(3, 1, "b", "b"); ui.EnterValue(3, 2, "7", "7");

        ui.ClickCell(1, 1); ui.Key("+{DOWN}+{DOWN}+{RIGHT}");                    // A1:B3
        ui.ClickUntil("Scheda Inserisci", "Istogramma");
        ui.Retry(() => ui.Click("Istogramma"), () => { try { return Has(ui.Dump(), "objects", "Chart"); } catch (Exception e) when (e is KeyNotFoundException or InvalidOperationException) { return false; } });

        ui.ClickCell(10, 1);
        ui.ClickUntil("Scheda Inserisci", "Casella di testo");
        ui.Retry(() => ui.Click("Casella di testo"), () => ui.Find("ShapeText", 800) != null);
        ui.SetValueId("ShapeText", "Ciao");
        ui.InvokeId("Dlg_Ok");
        var d = ui.DumpUntil(x => Has(x, "objects", "TextBox"));

        ui.ClickCell(2, 2);
        ui.ClickUntil("Scheda Inserisci", "Commento");
        ui.Retry(() => ui.Click("Commento"), () => ui.Find("CommentText", 800) != null);
        ui.SetValueId("CommentText", "nota di prova");
        ui.InvokeId("Dlg_Ok");
        d = ui.DumpUntil(x => Has(x, "comments", "B2=nota di prova"));
        ui.Key("^z");
        d = ui.DumpUntil(x => !Has(x, "comments", "B2="));
        ui.KillOwn();
    }

    [Fact]
    public void AdvancedFilter_ViaRibbon()
    {
        using var lk = UiLock.Acquire();
        using var ui = UiDriver.Launch();
        ui.EnterValue(1, 1, "n", "n"); ui.EnterValue(2, 1, "a", "a"); ui.EnterValue(3, 1, "b", "b"); ui.EnterValue(1, 4, "n", "n"); ui.EnterValue(2, 4, "b", "b");
        ui.ClickCell(1, 1);
        ui.ClickUntil("Scheda Dati", "Avanzate");
        ui.Retry(() => ui.Click("Avanzate"), () => ui.Find("Filtro avanzato...", 800) != null);
        ui.Retry(() => ui.Click("Filtro avanzato..."), () => ui.Find("AdvCrit", 800) != null);
        ui.SetValueId("AdvList", "A1:A3"); ui.SetValueId("AdvCrit", "D1:D2");
        ui.InvokeId("Dlg_Ok");
        var d = ui.DumpUntil(x => x.GetProperty("hiddenRows").EnumerateArray().Any(h => h.GetInt32() == 2));
        Assert.DoesNotContain(d.GetProperty("hiddenRows").EnumerateArray(), h => h.GetInt32() == 3);
        ui.Key("^z");
        ui.DumpUntil(x => x.GetProperty("hiddenRows").GetArrayLength() == 0);
        ui.KillOwn();
    }
}
