using System.Text.Json;

namespace PowerTables.Tests;

/// <summary>Test con l'app reale: UI Automation per trovare i controlli, mouse e tastiera reali (SendInput) per usarli.</summary>
[Collection("UiCollection"), Trait("Category", "Ui")]
public class UiTests
{
    // le catture vanno in tests/work; vengono copiate in docs/ (per il README) solo con PT_UPDATE_DOCS=1, così git status resta pulito
    static string Docs(string n) => Path.Combine(TestUtil.Root, Environment.GetEnvironmentVariable("PT_UPDATE_DOCS") == "1" ? "docs" : Path.Combine("tests", "work", "docs"), n);

    static string Cell(JsonElement d, string a) => d.GetProperty("cells").TryGetProperty(a, out var c) ? c.GetProperty("text").GetString()! : "";

    [Fact]
    public void Powers_Toggle_SwitchesTheme_Persists_AndSurvivesCorruptSettings()
    {
        using var lk = UiLock.Acquire();
        string settings = Path.Combine(Path.GetTempPath(), "pt-set-" + Guid.NewGuid().ToString("N") + ".json");
        using (var ui = UiDriver.Launch(settings: settings))
        {
            ui.WaitProbe("dark", "1");                          // predefinito scuro
            ui.ClickUntil("Tab_Powers", "DarkModeToggle");
            ui.Screenshot(Docs("screenshot-powers-dark.png"));
            ui.Click("DarkModeToggle");
            ui.WaitProbe("dark", "0");                          // il clic reale ha cambiato tutta l'interfaccia
            ui.Screenshot(Docs("screenshot-powers-light.png"));
            ui.ClickUntil("Tab_Home", "Incolla");
            ui.Screenshot(Docs("screenshot-light.png"));
            ui.ClickUntil("Tab_Powers", "DarkModeToggle"); ui.Click("DarkModeToggle");
            ui.WaitProbe("dark", "1");
            ui.Click("DarkModeToggle");                                     // lascia chiaro
            ui.KillOwn();
        }
        Assert.Contains("false", File.ReadAllText(settings));
        using (var ui = UiDriver.Launch(settings: settings))
        {
            ui.WaitProbe("dark", "0");                          // persistenza dopo riavvio
            ui.KillOwn();
        }
        File.WriteAllText(settings, "{{{ non è json \0\0");
        using (var ui = UiDriver.Launch(settings: settings))
        {
            ui.WaitProbe("dark", "1");                          // file corrotto: nessun crash, predefinito scuro
            ui.Screenshot(Docs("screenshot-dark.png"));
            ui.KillOwn();
        }
        File.WriteAllText(settings, "[1,2,3]");
        using (var ui = UiDriver.Launch(settings: settings)) { ui.WaitProbe("dark", "1"); ui.KillOwn(); }
    }

    [Fact]
    public void Grid_Typing_Formula_CopyPaste_Fill_Undo()
    {
        using var lk = UiLock.Acquire();
        using var ui = UiDriver.Launch();
        ui.EnterValue(1, 1, "10", "10");
        Assert.Equal("2,1", ui.Probe()["active"]);                          // Invio scende di una riga
        ui.EnterValue(2, 1, "20", "20");
        ui.EnterValue(3, 1, "=A1+A2", "30");
        var d = ui.DumpUntil(x => Cell(x, "A3") == "30");
        Assert.Equal("10", Cell(d, "A1")); Assert.Equal("30", Cell(d, "A3"));
        Assert.Equal("=A1+A2", d.GetProperty("cells").GetProperty("A3").GetProperty("f").GetString());
        // copia/incolla con adattamento della formula
        ui.Retry(() => { ui.ClickCell(3, 1); ui.Key("^c"); ui.ClickCell(3, 2); ui.Key("^v"); }, () => ui.Dump().GetProperty("cells").TryGetProperty("B3", out var b3) && b3.GetProperty("f").GetString() == "=B1+B2");
        d = ui.Dump();
        Assert.Equal("=B1+B2", d.GetProperty("cells").GetProperty("B3").GetProperty("f").GetString());
        // modifica della formula dalla barra della formula
        ui.ClickCell(3, 1);
        Assert.Equal("=A1+A2", ui.Probe()["bar"]);
        // riempimento: seleziono A1:A2 e trascino la maniglia fino alla riga 6
        ui.Retry(() =>
        {
            var (x1, y1) = ui.CellCenter(1, 1); var (x2, y2) = ui.CellCenter(2, 1);
            ui.DragFrom(x1, y1, x2, y2);
            ui.WaitProbe("sel", "1,1,2,1");
            var p = ui.Probe(); var o = p["origin"].Split(','); double s = double.Parse(p["scale"], System.Globalization.CultureInfo.InvariantCulture);
            int hx = (int)(double.Parse(o[0], System.Globalization.CultureInfo.InvariantCulture) + (50 + 64 - 2) * s), hy = (int)(double.Parse(o[1], System.Globalization.CultureInfo.InvariantCulture) + (22 + 40 - 2) * s);
            var (_, y6) = ui.CellCenter(6, 1);
            ui.DragFrom(hx, hy, hx, y6);
        }, () => Cell(ui.Dump(), "A6") == "60");
        d = ui.DumpUntil(x => Cell(x, "A6") == "60");
        Assert.Equal("60", Cell(d, "A6")); Assert.Equal("40", Cell(d, "A4"));
        ui.Key("^z");                                                       // annulla il riempimento
        d = ui.DumpUntil(x => Cell(x, "A6") == "");
        Assert.Equal("", Cell(d, "A6"));
        Assert.Equal("20", Cell(d, "A2"));
        ui.KillOwn();
    }

    [Fact]
    public void Ribbon_Sort_Filter_Merge_Freeze_Undo()
    {
        using var lk = UiLock.Acquire();
        using var ui = UiDriver.Launch();
        ui.EnterValue(1, 1, "c", "c"); ui.EnterValue(2, 1, "a", "a"); ui.EnterValue(3, 1, "b", "b");
        var (x1, y1) = ui.CellCenter(1, 1); var (x2, y2) = ui.CellCenter(3, 1);
        ui.DragFrom(x1, y1, x2, y2);
        ui.ClickUntil("Tab_Dati", "Ordina dalla A alla Z"); ui.Click("Ordina dalla A alla Z");
        var d = ui.DumpUntil(x => Cell(x, "A1") == "a");
        Assert.Equal("a", Cell(d, "A1")); Assert.Equal("c", Cell(d, "A3"));
        ui.Retry(() => { ui.ClickCell(5, 4); ui.Key("^z"); }, () => Cell(ui.Dump(), "A1") == "c");
        Assert.Equal("c", Cell(ui.Dump(), "A1"));
        // filtro
        ui.ClickCell(1, 1);
        ui.Retry(() => ui.Click("Filtro"), () => ui.TryProbe(out var pf) && pf["filter"] == "1");
        ui.Retry(() => ui.Click("Filtro"), () => ui.TryProbe(out var pf) && pf["filter"] == "0");
        // unione
        ui.ClickUntil("Tab_Home", "Unisci e centra");
        (x1, y1) = ui.CellCenter(5, 1); (x2, y2) = ui.CellCenter(5, 2);
        ui.DragFrom(x1, y1, x2, y2);
        ui.Click("Unisci e centra");
        d = ui.DumpUntil(x => x.GetProperty("merges").GetArrayLength() > 0);
        Assert.Contains("A5:B5", d.GetProperty("merges").EnumerateArray().Select(m => m.GetString()));
        // blocca riquadri
        ui.ClickUntil("Tab_Visualizza", "Blocca riquadri");
        ui.Retry(() => { ui.ClickUntil("Blocca riquadri", "Blocca riga superiore"); ui.Click("Blocca riga superiore"); }, () => ui.TryProbe(out var pr) && pr["frozen"] == "1,0");
        ui.Screenshot(Docs("screenshot-ui-session.png"));
        ui.KillOwn();
    }
}

[CollectionDefinition("UiCollection", DisableParallelization = true)]
public class UiCollection { }
