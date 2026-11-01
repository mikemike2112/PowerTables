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
}
