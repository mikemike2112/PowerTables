using System.Windows.Automation;

namespace PowerTables.Tests;

/// <summary>Pulsante ingrandisci e adattamento a varie larghezze: solo UI Automation, nessun input reale.</summary>
[Collection("UiCollection"), Trait("Category", "UiAutomation")]
public class ResponsiveUiTests
{
    static string Docs(string n) => Path.Combine(TestUtil.Root, Environment.GetEnvironmentVariable("PT_UPDATE_DOCS") == "1" ? "docs" : Path.Combine("tests", "work", "docs"), n);

    [Fact]
    public void MaximizeButton_ChangesIconAndTooltip_WithState()
    {
        using var ui = UiDriver.Launch();
        Assert.Equal("Ripristina giù", ui.NameOf("Ripristina giù"));                 // avvio a schermo intero: si può ripristinare
        ui.SetVisualState(WindowVisualState.Normal);
        Assert.NotNull(ui.Find("Ingrandisci"));                                      // finestra: si può ingrandire
        Assert.Null(ui.Find("Ripristina giù", 800));
        ui.SetVisualState(WindowVisualState.Maximized);
        Assert.NotNull(ui.Find("Ripristina giù"));
        ui.KillOwn();
    }

    [Fact]
    public void Ribbon_AdaptsFromWideToMinimumWidth_WithoutLosingCommands()
    {
        using var ui = UiDriver.Launch();
        ui.Screenshot(Docs("resp-massimizzata.png"));
        foreach (var w in new[] { 1100, 800, 640 })
        {
            ui.ResizeTo(w, 700);
            Assert.NotNull(ui.Find("Incolla", 3000));                                // il comando principale resta raggiungibile
            if (w <= 800) Assert.NotNull(ui.Find("RibbonOverflow", 3000));          // gruppi in più → menu a discesa
            if (w <= 800) Assert.NotNull(ui.Find("TabOverflow", 3000));             // schede in più → menu a discesa
            ui.Screenshot(Docs($"resp-{w}.png"));
        }
        ui.KillOwn();
    }

    static string LastInfo = "";
    static (bool Compact, bool Overflow) State(UiDriver ui)
    {
        Thread.Sleep(700);
        var p = ui.Probe();
        LastInfo = $"larghezza piena {p["rbfull"]} disponibile {p["rbavail"]}";
        return (p["rbcompact"] == "1", p["rbovf"] == "1");
    }

    /// <summary>Regressione: il calcolo adattivo usava misure in cache (WPF non invalida gli antenati quando cambia la visibilità di un discendente), quindi il ribbon restava compresso con mezza barra vuota.</summary>
    [Fact]
    public void Ribbon_ReExpands_AfterMaximize_AndCompactsAgain_WhenNarrowed()
    {
        using var lk = UiLock.Acquire();
        using var ui = UiDriver.Launch(TestUtil.Corpus("office-made", "mixed-budget.xlsx"));
        ui.ResizeTo(700, 600);
        var (c1, o1) = State(ui);
        Assert.True(c1 && o1, "stretta: gruppi compatti e menu » presenti");
        Assert.NotNull(ui.Find("RibbonOverflow", 2000));
        ui.SetVisualState(WindowVisualState.Maximized);
        var (c2, o2) = State(ui);
        Assert.False(c2, "massimizzata: i gruppi devono tornare estesi (" + LastInfo + ")");
        Assert.False(o2, "massimizzata: il menu » deve sparire");
        Assert.Null(ui.Find("RibbonOverflow", 800));
        Assert.Equal("0", ui.Probe()["modified"]);
        ui.Screenshot(Docs("resp-massimizzata-estesa.png"));
        // cambio scheda e ritorno, da massimizzata: resta esteso
        ui.ClickUntil("Tab_Dati", "Da testo/CSV");
        Assert.False(State(ui).Compact);
        ui.ClickUntil("Tab_Home", "Incolla");
        var (c3, o3) = State(ui);
        Assert.False(c3 || o3);
        // ripristina in finestra più piccola: si compatta di nuovo; poi di nuovo massimizzata: si riespande
        ui.ResizeTo(900, 600);
        var (c4, o4) = State(ui);
        Assert.True(c4 || o4, "finestra stretta: deve compattarsi");
        ui.SetVisualState(WindowVisualState.Maximized);
        var (c5, o5) = State(ui);
        Assert.False(c5 || o5);
        Assert.Equal("0", ui.Probe()["modified"]);          // ridimensionare e cambiare scheda non modifica il documento
        ui.KillOwn();
    }
}
