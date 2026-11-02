using System.Text.Json;

namespace PowerTables.Tests;

/// <summary>Inserisci grafico, forma, casella di testo e commento dalla ribbon con mouse e tastiera reali. Usa il blocco ui.lock.</summary>
[Collection("UiCollection"), Trait("Category", "Ui")]
public class ObjectsUiTests
{
    static bool Has(JsonElement d, string prop, string needle) => d.TryGetProperty(prop, out var a) && a.EnumerateArray().Any(x => x.GetString()!.Contains(needle));

    static JsonElement Until(UiDriver ui, Func<JsonElement, bool> ok) { var d = ui.DumpUntil(ok); Assert.True(ok(d), "condizione non raggiunta; dump: " + d); return d; }

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
        var d = Until(ui, x => Has(x, "objects", "TextBox"));

        ui.ClickCell(2, 2);
        ui.ClickUntil("Scheda Inserisci", "Commento");
        ui.Retry(() => ui.Click("Commento"), () => ui.Find("CommentText", 800) != null);
        ui.SetValueId("CommentText", "nota di prova");
        ui.InvokeId("Dlg_Ok");
        d = Until(ui, x => Has(x, "comments", "B2=nota di prova"));
        ui.Key("^z");
        d = Until(ui, x => !Has(x, "comments", "B2="));
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
        var d = Until(ui, x => x.GetProperty("hiddenRows").EnumerateArray().Any(h => h.GetInt32() == 2));
        Assert.DoesNotContain(d.GetProperty("hiddenRows").EnumerateArray(), h => h.GetInt32() == 3);
        ui.Key("^z");
        Until(ui, x => x.GetProperty("hiddenRows").GetArrayLength() == 0);
        ui.KillOwn();
    }

    [Fact]
    public void GoalSeek_And_Forecast_ViaRibbon()
    {
        using var lk = UiLock.Acquire();
        using var ui = UiDriver.Launch();
        ui.EnterValue(1, 1, "2", "2"); ui.EnterValue(2, 1, "=A1*3+1", "7");
        ui.ClickCell(2, 1);
        ui.ClickUntil("Scheda Dati", "Analisi di simulazione");
        ui.Retry(() => ui.Click("Analisi di simulazione"), () => ui.Find("Ricerca obiettivo...", 800) != null);
        ui.Retry(() => ui.Click("Ricerca obiettivo..."), () => ui.Find("GsValue", 800) != null);
        ui.SetValueId("GsValue", "10"); ui.SetValueId("GsChange", "A1");
        ui.InvokeId("Dlg_Ok");
        var d = Until(ui, x => x.GetProperty("cells").GetProperty("A2").GetProperty("text").GetString() == "10");
        Assert.Equal("3", d.GetProperty("cells").GetProperty("A1").GetProperty("text").GetString());

        for (int i = 1; i <= 6; i++) { ui.EnterValue(i, 4, i.ToString(), i.ToString()); ui.EnterValue(i, 5, (10 + 2 * i).ToString(), (10 + 2 * i).ToString()); }
        ui.ClickCell(1, 4);
        ui.ClickUntil("Scheda Dati", "Foglio previsione");
        ui.Retry(() => ui.Click("Foglio previsione"), () => ui.Find("FcTimeline", 800) != null);
        ui.SetValueId("FcTimeline", "D1:D6"); ui.SetValueId("FcValues", "E1:E6");
        ui.InvokeId("Dlg_Ok");
        Until(ui, x => x.GetProperty("sheet").GetString() == "Previsione");
        ui.KillOwn();
    }

    [Fact]
    public void Pivot_And_Themes_ViaRibbon()
    {
        using var lk = UiLock.Acquire();
        using var ui = UiDriver.Launch();
        ui.EnterValue(1, 1, "r", "r"); ui.EnterValue(1, 2, "v", "v");
        ui.EnterValue(2, 1, "a", "a"); ui.EnterValue(2, 2, "5", "5"); ui.EnterValue(3, 1, "b", "b"); ui.EnterValue(3, 2, "7", "7"); ui.EnterValue(4, 1, "a", "a"); ui.EnterValue(4, 2, "1", "1");
        ui.ClickCell(1, 1);
        ui.ClickUntil("Scheda Inserisci", "Tabella pivot");
        ui.Retry(() => ui.Click("Tabella pivot"), () => ui.Find("PvRow", 800) != null);
        ui.SetValueId("PvRow", "r"); ui.SetValueId("PvValue", "v");
        ui.InvokeId("Dlg_Ok");
        string pe = "";
        var d = Until(ui, x => x.GetProperty("sheet").GetString() != "Foglio1");
        Assert.True(d.GetProperty("cells").GetProperty("A1").GetProperty("text").GetString() == "Etichette di riga", "errore dialogo: " + pe + " / " + d.ToString());
        Assert.Equal("6", d.GetProperty("cells").GetProperty("B2").GetProperty("text").GetString());
        Assert.Equal("13", d.GetProperty("cells").GetProperty("B4").GetProperty("text").GetString());
        ui.Key("^z");
        Until(ui, x => x.GetProperty("sheet").GetString() == "Foglio1");

        ui.ClickUntil("Scheda Layout di pagina", "Temi");
        ui.Retry(() => ui.Click("Temi"), () => ui.Find("Interfaccia chiara", 800) != null);
        ui.Retry(() => ui.Click("Interfaccia chiara"), () => { try { return !ui.Dump().GetProperty("dark").GetBoolean(); } catch (Exception e) when (e is KeyNotFoundException or InvalidOperationException) { return false; } });
        ui.KillOwn();
    }
}
