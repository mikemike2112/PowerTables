using System.Windows.Automation;

namespace PowerTables.Tests;

/// <summary>Scudo privacy nell'app reale, pilotato solo con UI Automation (nessun input reale).</summary>
[Collection("UiCollection"), Trait("Category", "UiAutomation")]
public class PrivacyUiTests
{
    static string Docs(string n) => Path.Combine(TestUtil.Root, Environment.GetEnvironmentVariable("PT_UPDATE_DOCS") == "1" ? "docs" : Path.Combine("tests", "work", "docs"), n);

    [Fact]
    public void Shield_FindsData_CreatesCleanCopy_OriginalUntouched()
    {
        string src = PrivacyShieldTests.MakeDirty(); string before = PrivacyShieldTests.Hash(src);
        string folder = Path.GetDirectoryName(src)!; string expected = Path.Combine(folder, "sporco (pulito).xlsx");
        if (File.Exists(expected)) File.Delete(expected);
        using var ui = UiDriver.Launch(src);
        Assert.True(ui.TryProbe(out var p0) && p0["modified"] == "0");
        // il comando si raggiunge dalla casella Cerca (UI Automation, nessun clic reale)
        ui.SetValueId("SearchBox", "Scudo");
        var res = ui.Find("Scudo privacy   ·   Powers") ?? throw new InvalidOperationException("risultato di ricerca non trovato");
        ((InvokePattern)res.GetCurrentPattern(InvokePattern.Pattern)).Invoke();
        Assert.NotNull(ui.Find("Priv_0"));                                           // le voci trovate hanno una casella ciascuna
        Assert.NotNull(ui.Find("Priv_Anonymize"));
        Thread.Sleep(400);
        ui.ScreenshotWindowOf("Priv_Make", Docs("scudo-privacy-analisi.png"));
        ui.InvokeId("Priv_Make");
        Assert.Equal("sporco (pulito)", ui.ValueOf("Dlg_FileName"));                 // nome proposto: nome (pulito)
        ui.InvokeId("Dlg_Save");
        var result = ui.Find("Priv_Result") ?? throw new InvalidOperationException("esito non mostrato");
        Thread.Sleep(400);
        ui.ScreenshotWindowOf("Dlg_Ok", Docs("scudo-privacy-esito.png"));
        ui.InvokeId("Dlg_Ok");
        ui.KillOwn();
        Assert.True(File.Exists(expected));
        Assert.Equal(before, PrivacyShieldTests.Hash(src));                          // originale identico byte per byte
        string text = PrivacyShieldTests.AllText(expected);
        foreach (var s in new[] { "Mario Rossi", "Anna Bianchi", "ACME Srl", "Luigi Verdi", "Mario.Rossi", "Cliente riservato" }) Assert.DoesNotContain(s, text);
    }
}
