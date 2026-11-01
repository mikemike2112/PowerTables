using System.Text.Json;
using System.Windows.Automation;

namespace PowerTables.Tests;

/// <summary>Modifica di formule con mouse e tastiera reali: elenco funzioni con Tab e selezione dinamica dei riferimenti. Usa il blocco ui.lock.</summary>
[Collection("UiCollection"), Trait("Category", "Ui")]
public class FormulaEditUiTests
{
    static string Docs(string n) => Path.Combine(TestUtil.Root, Environment.GetEnvironmentVariable("PT_UPDATE_DOCS") == "1" ? "docs" : Path.Combine("tests", "work", "docs"), n);
    static string Cell(JsonElement d, string a) => d.GetProperty("cells").TryGetProperty(a, out var c) ? c.GetProperty("text").GetString()! : "";

    static void WaitBar(UiDriver ui, string expected)
    {
        string cur = "";
        for (int i = 0; i < 30; i++) { if (ui.TryProbe(out var p) && (cur = p.GetValueOrDefault("bar", "")) == expected) return; Thread.Sleep(200); }
        Assert.Equal(expected, cur);
    }

    static string Bar(UiDriver ui) => ui.Probe().GetValueOrDefault("bar", "");

    [Fact]
    public void Tab_CompletesFunction_AndItalianNamesWork()
    {
        using var lk = UiLock.Acquire();
        using var ui = UiDriver.Launch();
        ui.EnterValue(1, 1, "5", "5"); ui.EnterValue(2, 1, "7", "7");
        // "=SUM" mostra l'elenco, Tab completa con "(" e si prosegue a mano
        ui.Retry(() => { ui.ClickCell(3, 1); ui.Type("=SUM"); }, () => ui.Find("FormulaAutocomplete", 1500) != null);
        ui.ScreenCapture(Docs("formula-elenco-funzioni.png"));
        ui.Key("{TAB}");
        WaitBar(ui, "=SUM(");
        ui.Type("A1:A2)"); ui.Key("{ENTER}");
        var d = ui.DumpUntil(x => Cell(x, "A3") == "12");
        Assert.Equal("=SUM(A1:A2)", d.GetProperty("cells").GetProperty("A3").GetProperty("f").GetString());
        // nomi italiani e separatore ';' digitati a mano
        ui.EnterValue(4, 1, "=SOMMA(A1;A2;0,5)", "12,5");
        d = ui.Dump();
        Assert.Equal("=SUM(A1,A2,0.5)", d.GetProperty("cells").GetProperty("A4").GetProperty("f").GetString());
        ui.KillOwn();
    }

    [Fact]
    public void DynamicSelection_Click_Drag_Keyboard_AndNoAccidentalCommit()
    {
        using var lk = UiLock.Acquire();
        using var ui = UiDriver.Launch();
        for (int i = 1; i <= 4; i++) ui.EnterValue(i, 1, (i * 10).ToString(), (i * 10).ToString());
        // trascinamento: =SUM( poi A1→A3 con il mouse, il riferimento si aggiorna dal vivo
        ui.Retry(() => { ui.ClickCell(6, 3); ui.Type("=SUM("); }, () => Bar(ui) == "=SUM(");
        var (x1, y1) = ui.CellCenter(1, 1); var (x3, y3) = ui.CellCenter(3, 1);
        ui.DragFrom(x1, y1, x3, y3);
        WaitBar(ui, "=SUM(A1:A3");
        Assert.Equal("1", ui.Probe()["edit"]);                                   // ancora in modifica: nessun commit per sbaglio
        Assert.Equal("6,3", ui.Probe()["active"]);                               // la cella attiva è sempre quella in modifica
        ui.Screenshot(Docs("formula-selezione-dinamica.png"));
        // un clic in un punto dove non va un riferimento (dopo la selezione completata non si digita ')' ma si clicca): sostituisce il riferimento inserito
        ui.ClickCell(4, 1);
        WaitBar(ui, "=SUM(A4");
        ui.Type(")");
        WaitBar(ui, "=SUM(A4)");
        // clic dove non è ammesso un riferimento (dopo ')'): ignorato, nessun commit, testo invariato
        ui.ClickCell(2, 1);
        Thread.Sleep(300);
        Assert.Equal("=SUM(A4)", Bar(ui)); Assert.Equal("1", ui.Probe()["edit"]);
        // clic sulle intestazioni: ignorato
        var (hx, hy) = ui.CellCenter(1, 2); ui.ClickAt(hx, hy - 40);
        Assert.Equal("1", ui.Probe()["edit"]);
        ui.Key("{ENTER}");
        var d = ui.DumpUntil(x => Cell(x, "C6") == "40");
        Assert.Equal("=SUM(A4)", d.GetProperty("cells").GetProperty("C6").GetProperty("f").GetString());   // riferimenti relativi, nessun $ inserito da solo
        // tastiera: operatore poi frecce inseriscono il riferimento, Maiusc+freccia lo estende
        ui.Retry(() => { ui.ClickCell(8, 3); ui.Type("=SUM("); }, () => Bar(ui) == "=SUM(");
        ui.Key("{UP}");  WaitBar(ui, "=SUM(C7");
        ui.Key("{UP}");  WaitBar(ui, "=SUM(C6");
        ui.Key("+{UP}"); WaitBar(ui, "=SUM(C5:C6");
        ui.Key("{ESC}"); // chiude eventuale elenco; poi annulla la modifica
        ui.Key("{ESC}");
        ui.KillOwn();
    }

    [Fact]
    public void FormulasTab_Categories_And_InsertFunctionDialog_Work()
    {
        using var lk = UiLock.Acquire();
        using var ui = UiDriver.Launch();
        ui.ClickUntil("Tab_Formule", "Calcola adesso");
        Assert.Null(ui.Find("Tab_Disegno", 800)); Assert.Null(ui.Find("Tab_Revisione", 800));   // schede tolte
        // categoria Logiche: il menu elenca le funzioni, la scelta apre la modifica con =IF(
        ui.Retry(() => { ui.Click("Logiche"); }, () => ui.Find("IF  ·  SE", 1500) != null);
        ui.ScreenCapture(Docs("formule-categoria-logiche.png"));
        ui.Click("IF  ·  SE");
        WaitBar(ui, "=IF(");
        ui.Key("{ESC}"); ui.Key("{ESC}");
        // finestra "Inserisci funzione": ricerca in italiano e inserimento
        ui.Click("Inserisci funzione");
        ui.SetValueId("FuncSearch", "somma.se");
        Thread.Sleep(400);
        ui.ScreenCapture(Docs("formule-inserisci-funzione.png"));
        ui.InvokeId("Dlg_Insert");
        WaitBar(ui, "=SUMIF(");
        ui.Key("{ESC}"); ui.Key("{ESC}");
        ui.KillOwn();
    }
}
