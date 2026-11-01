using PowerTables.Core.IO;

namespace PowerTables.Tests;

public class SaveLogicTests
{
    [Fact]
    public void Map_ButtonsToOutcomes()
    {
        Assert.Equal(new SaveLogic.Outcome(true, true), SaveLogic.Map(SaveLogic.Choice.Save));
        Assert.Equal(new SaveLogic.Outcome(true, false), SaveLogic.Map(SaveLogic.Choice.DontSave));   // continua senza scrivere
        Assert.Equal(new SaveLogic.Outcome(false, false), SaveLogic.Map(SaveLogic.Choice.Cancel));    // non chiude e non salva
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("a/b")]
    [InlineData("a:b")]
    [InlineData("a?b")]
    [InlineData("nome.")]
    [InlineData("CON")]
    [InlineData("nul.xlsx")]
    [InlineData("LPT1")]
    public void InvalidNames_AreRejected(string n) => Assert.NotNull(SaveLogic.ValidateName(n));

    [Theory]
    [InlineData("bilancio")]
    [InlineData("Bilancio 2026.xlsx")]
    [InlineData("è.csv")]
    public void ValidNames_Accepted(string n) => Assert.Null(SaveLogic.ValidateName(n));

    [Fact]
    public void BuildPath_AddsExtensionOnce()
    {
        Assert.Equal(Path.Combine(@"C:\x", "a.xlsx"), SaveLogic.BuildPath(@"C:\x", "a", ".xlsx"));
        Assert.Equal(Path.Combine(@"C:\x", "a.XLSX"), SaveLogic.BuildPath(@"C:\x", " a.XLSX ", ".xlsx"));
    }

    [Fact]
    public void Folder_Validation_And_OverwriteDetection()
    {
        string dir = Path.Combine(Path.GetTempPath(), "pt-sl-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(dir);
        Assert.Null(SaveLogic.ValidateFolder(dir));
        Assert.NotNull(SaveLogic.ValidateFolder(Path.Combine(dir, "manca")));
        Assert.NotNull(SaveLogic.ValidateFolder(""));
        Assert.NotNull(SaveLogic.ValidateFolder("relativo"));
        string f = Path.Combine(dir, "a.xlsx");
        Assert.False(SaveLogic.NeedsOverwriteConfirmation(f));
        File.WriteAllText(f, "x");
        Assert.True(SaveLogic.NeedsOverwriteConfirmation(f));
        Directory.Delete(dir, true);
    }
}

/// <summary>Finestre di dialogo dell'app reale pilotate SOLO con UI Automation (Invoke/Value pattern): nessun input reale.</summary>
[Collection("UiCollection"), Trait("Category", "UiAutomation")]
public class SaveDialogUiTests
{
    static string Docs(string n) => Path.Combine(TestUtil.Root, Environment.GetEnvironmentVariable("PT_UPDATE_DOCS") == "1" ? "docs" : Path.Combine("tests", "work", "docs"), n);
    static string NewDir() { var d = Path.Combine(Path.GetTempPath(), "pt-dlg-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(d); return d; }

    static string Result(string p) { for (int i = 0; i < 40 && !File.Exists(p); i++) Thread.Sleep(150); Assert.True(File.Exists(p), "esito non scritto"); return File.ReadAllText(p); }

    [Fact]
    public void Screenshots_BothThemes_AndStructure()
    {
        foreach (var (dark, file) in new[] { (true, "dialog-salva-scuro.png"), (false, "dialog-salva-chiaro.png") })
        {
            // cartella finta solo per lo screenshot (niente dati personali nel repository); Annulla non la usa
            using var ui = UiDriver.LaunchDemo(@"save:C:\Documenti", dark, out var res);
            Assert.Equal("Nome file", ui.NameOf("Dlg_FileName"));
            Assert.Equal("Senza titolo", ui.ValueOf("Dlg_FileName"));
            Assert.NotNull(ui.Find("Dlg_Browse")); Assert.NotNull(ui.Find("Dlg_Cancel")); Assert.NotNull(ui.Find("Dlg_DontSave"));
            Thread.Sleep(500);
            ui.ScreenshotWindowOf("Dlg_Save", Docs(file));
            ui.InvokeId("Dlg_Cancel");
            Assert.StartsWith("Cancel|", Result(res));
            ui.KillOwn();
        }
    }

    [Fact]
    public void Cancel_And_DontSave_WriteNothing()
    {
        foreach (var (btn, expect) in new[] { ("Dlg_Cancel", "Cancel|"), ("Dlg_DontSave", "DontSave|") })
        {
            string dir = NewDir();
            using var ui = UiDriver.LaunchDemo("save:" + dir, true, out var res);
            ui.InvokeId(btn);
            Assert.StartsWith(expect, Result(res));
            ui.KillOwn();
            Assert.Empty(Directory.GetFileSystemEntries(dir));
            Directory.Delete(dir, true);
        }
    }

    [Fact]
    public void InvalidName_ShowsError_AndDoesNotClose_Then_Save_ReturnsPath()
    {
        string dir = NewDir();
        using var ui = UiDriver.LaunchDemo("save:" + dir, true, out var res);
        ui.SetValueId("Dlg_FileName", "a/b");
        ui.InvokeId("Dlg_Save");
        Assert.NotNull(ui.Find("Dlg_Error"));
        Assert.False(File.Exists(res));                                  // il dialogo non si è chiuso
        ui.SetValueId("Dlg_FileName", "bilancio");
        ui.InvokeId("Dlg_Save");
        Assert.Equal("Save|" + Path.Combine(dir, "bilancio.xlsx"), Result(res));
        ui.KillOwn();
        Assert.Empty(Directory.GetFileSystemEntries(dir));               // la finestra non scrive il file: lo fa l'app dopo
        Directory.Delete(dir, true);
    }

    [Fact]
    public void ExistingFile_RequiresOverwriteConfirmation_AndNoKeepsFile()
    {
        string dir = NewDir(); string f = Path.Combine(dir, "bilancio.xlsx"); File.WriteAllText(f, "originale");
        using var ui = UiDriver.LaunchDemo("save:" + dir, true, out var res);
        ui.SetValueId("Dlg_FileName", "bilancio");
        ui.InvokeId("Dlg_Save");
        Assert.NotNull(ui.Find("Dlg_Yes"));                              // conferma di sovrascrittura
        ui.InvokeId("Dlg_No");
        Assert.False(File.Exists(res));                                  // ancora nel dialogo di salvataggio
        Assert.Equal("originale", File.ReadAllText(f));
        ui.InvokeId("Dlg_Cancel");
        Assert.StartsWith("Cancel|", Result(res));
        Assert.Equal("originale", File.ReadAllText(f));
        ui.KillOwn(); Directory.Delete(dir, true);
    }

    [Fact]
    public void DocumentWithFile_ShowsSimplePrompt_AndErrorDialog()
    {
        string dir = NewDir();
        using (var ui = UiDriver.LaunchDemo("savefile:" + dir, false, out var res))
        {
            Assert.Null(ui.Find("Dlg_FileName", 1200));                  // niente campo nome: il file è già associato
            ui.InvokeId("Dlg_Save");
            Assert.Equal("Save|", Result(res));
            ui.KillOwn();
        }
        using (var ui = UiDriver.LaunchDemo("error", false, out var res2, "Dlg_Ok"))
        {
            Thread.Sleep(400);
            ui.ScreenshotWindowOf("Dlg_Ok", Docs("dialog-errore-chiaro.png"));
            ui.InvokeId("Dlg_Ok");
            Assert.Equal("ok", Result(res2));
            ui.KillOwn();
        }
        Directory.Delete(dir, true);
    }
}
