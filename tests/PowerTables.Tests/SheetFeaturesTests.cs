using System.Globalization;
using System.IO.Compression;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using PowerTables.Core;
using PowerTables.Core.IO;

namespace PowerTables.Tests;

public class SheetFeaturesTests
{
    static readonly CultureInfo It = new("it-IT");
    static string Tmp(string n) { string d = Path.Combine(TestUtil.Root, "tests", "work", "features"); Directory.CreateDirectory(d); return Path.Combine(d, n); }
    static int Errors(string path) { using var d = SpreadsheetDocument.Open(path, false); return new OpenXmlValidator().Validate(d).Count(); }
    static string Part(string path, string entry) { using var z = ZipFile.OpenRead(path); return new StreamReader(z.GetEntry(entry)!.Open()).ReadToEnd(); }
    static (Workbook, Sheet) New() { var wb = XlsxWriter.CreateNew(); return (wb, wb.Sheets[0]); }

    // ---------- convalida ----------
    [Fact]
    public void ValidationNumberAndListAreEnforcedAndUndoable()
    {
        var (wb, s) = New();
        wb.SetValidation(s, 1, 1, 3, 1, new DvSpec("whole", "between", "1", "10", Err: "Da 1 a 10"));
        wb.SetValidation(s, 1, 2, 3, 2, new DvSpec("list", "between", "\"rosso,verde\""));
        Assert.Null(wb.CheckInput(s, 1, 1, "5", It));
        Assert.Equal("Da 1 a 10", wb.CheckInput(s, 2, 1, "11", It));
        Assert.NotNull(wb.CheckInput(s, 2, 1, "2,5", It));      // non intero
        Assert.NotNull(wb.CheckInput(s, 2, 1, "abc", It));
        Assert.Null(wb.CheckInput(s, 9, 1, "abc", It));          // fuori area
        Assert.Null(wb.CheckInput(s, 1, 2, "VERDE", It));
        Assert.NotNull(wb.CheckInput(s, 1, 2, "blu", It));
        Assert.Equal(["rosso", "verde"], wb.ValidationList(s, 1, 2));
        // l'area sovrapposta restringe la regola precedente
        wb.SetValidation(s, 2, 1, 2, 1, null);
        Assert.Null(wb.CheckInput(s, 2, 1, "99", It)); Assert.NotNull(wb.CheckInput(s, 3, 1, "99", It)); Assert.NotNull(wb.CheckInput(s, 1, 1, "99", It));
        wb.Undo();
        Assert.NotNull(wb.CheckInput(s, 2, 1, "99", It));
        wb.Undo(); wb.Undo();
        Assert.Null(wb.GetValidation(s, 1, 1));
        Assert.Throws<ArgumentException>(() => wb.SetValidation(s, 1, 1, 1, 1, new DvSpec("whole", "between", "1")));   // manca il massimo
    }

    [Fact]
    public void ValidationListFromRangeAndDecimalOperators()
    {
        var (wb, s) = New();
        wb.Do("x", () => { wb.SetValue(s, 1, 4, "a"); wb.SetValue(s, 2, 4, "b"); });
        wb.SetValidation(s, 1, 1, 1, 1, new DvSpec("list", "between", "$D$1:$D$2"));
        wb.SetValidation(s, 2, 1, 2, 1, new DvSpec("decimal", "greaterThan", "0"));
        Assert.Equal(["a", "b"], wb.ValidationList(s, 1, 1));
        Assert.Null(wb.CheckInput(s, 1, 1, "B", It)); Assert.NotNull(wb.CheckInput(s, 1, 1, "c", It));
        Assert.Null(wb.CheckInput(s, 2, 1, "0,5", It)); Assert.NotNull(wb.CheckInput(s, 2, 1, "-1", It));
    }

    [Fact]
    public void ValidationFollowsInsertedRowsAndSurvivesSave()
    {
        string p = Tmp("dv.xlsx");
        var (wb, s) = New();
        wb.SetValidation(s, 1, 1, 3, 1, new DvSpec("whole", "between", "1", "10"));
        wb.InsertLines(s, true, 1, 2);
        Assert.NotNull(wb.GetValidation(s, 3, 1)); Assert.Null(wb.GetValidation(s, 1, 1));
        wb.Undo();
        Assert.NotNull(wb.GetValidation(s, 1, 1));
        wb.Redo();
        XlsxWriter.Save(wb, p);
        Assert.Equal(0, Errors(p));
        Assert.Contains("sqref=\"A3:A5\"", Part(p, "xl/worksheets/sheet1.xml"));
        var wb2 = XlsxReader.Open(p);
        Assert.Equal("whole", wb2.GetValidation(wb2.Sheets[0], 4, 1)!.Type);
        Assert.NotNull(wb2.CheckInput(wb2.Sheets[0], 4, 1, "50", It));
    }

    // ---------- formattazione condizionale ----------
    [Fact]
    public void CondFormatRulesMatchAndUndo()
    {
        var (wb, s) = New();
        wb.Do("x", () => { wb.SetValue(s, 1, 1, 5.0); wb.SetValue(s, 2, 1, 15.0); wb.SetValue(s, 3, 1, "Mela rossa"); wb.SetValue(s, 4, 1, 10.0); });
        wb.AddCondFormat(s, 1, 1, 4, 1, new CfSpec("cellIs", "greaterThan", "10", null, null, "FF9C0006", "FFFFC7CE", false));
        wb.AddCondFormat(s, 1, 1, 4, 1, new CfSpec("containsText", "", "", null, "rossa", null, "FFFFEB9C", true));
        Assert.Null(wb.CondFormatAt(s, 1, 1));
        Assert.Equal(new CfLook("FF9C0006", "FFFFC7CE", false), wb.CondFormatAt(s, 2, 1));
        Assert.Equal(new CfLook("FF9C0006", "FFFFEB9C", true), wb.CondFormatAt(s, 3, 1));
        Assert.Null(wb.CondFormatAt(s, 4, 1));                   // uguale a 10, non maggiore
        wb.AddCondFormat(s, 4, 1, 4, 1, new CfSpec("cellIs", "between", "5", "10", null, "FF006100", null, false));
        Assert.Equal("FF006100", wb.CondFormatAt(s, 4, 1)!.Value.Font);
        wb.Undo(); Assert.Null(wb.CondFormatAt(s, 4, 1));
        wb.Undo(); Assert.False(wb.CondFormatAt(s, 3, 1)!.Value.Bold);
        wb.Redo(); Assert.True(wb.CondFormatAt(s, 3, 1)!.Value.Bold);
        wb.ClearCondFormat(s, 1, 1, 4, 1); Assert.Null(wb.CondFormatAt(s, 2, 1));
        wb.Undo(); Assert.NotNull(wb.CondFormatAt(s, 2, 1));
        Assert.Throws<ArgumentException>(() => wb.AddCondFormat(s, 1, 1, 1, 1, new CfSpec("cellIs", "greaterThan", "", null, null, "FF000000", null, false)));
        Assert.Throws<ArgumentException>(() => wb.AddCondFormat(s, 1, 1, 1, 1, new CfSpec("cellIs", "greaterThan", "1", null, null, null, null, false)));
    }

    [Fact]
    public void CondFormatSurvivesSaveAndFollowsDelete()
    {
        string p = Tmp("cf.xlsx");
        var (wb, s) = New();
        wb.Do("x", () => { for (int i = 1; i <= 5; i++) wb.SetValue(s, i, 1, (double)i); });
        wb.AddCondFormat(s, 2, 1, 5, 1, new CfSpec("cellIs", "greaterThan", "3", null, null, "FF9C0006", "FFFFC7CE", true));
        wb.DeleteLines(s, true, 1, 1);
        XlsxWriter.Save(wb, p);
        Assert.Equal(0, Errors(p));
        string sheet = Part(p, "xl/worksheets/sheet1.xml");
        Assert.Contains("sqref=\"A1:A4\"", sheet); Assert.Contains("<dxfs", Part(p, "xl/styles.xml"));
        var wb2 = XlsxReader.Open(p);
        Assert.NotNull(wb2.CondFormatAt(wb2.Sheets[0], 4, 1)); Assert.Null(wb2.CondFormatAt(wb2.Sheets[0], 2, 1));
        Assert.True(wb2.CondFormatAt(wb2.Sheets[0], 4, 1)!.Value.Bold);
    }

    // ---------- collegamenti ----------
    [Fact]
    public void HyperlinkExternalInternalAndRoundTrip()
    {
        string p = Tmp("hl.xlsx");
        var (wb, s) = New(); wb.AddSheet("Altro");
        wb.SetHyperlink(s, 1, 1, "example.org/a?b=1", null, "Sito");
        wb.SetHyperlink(s, 2, 1, null, "Altro!B2", null);
        Assert.Equal("Sito", s.GetValue(1, 1)); Assert.True(wb.Styles.Get(s.Get(1, 1).Style).Font.Underline);
        Assert.Equal("https://example.org/a?b=1", wb.GetHyperlink(s, 1, 1)!.Value.Url);
        Assert.Equal("Altro!B2", s.GetValue(2, 1));
        Assert.Throws<ArgumentException>(() => wb.SetHyperlink(s, 3, 1, "javascript:alert(1)", null, "x"));
        Assert.Throws<ArgumentException>(() => wb.SetHyperlink(s, 3, 1, "file:///c:/windows/system32/calc.exe", null, "x"));
        Assert.Throws<ArgumentException>(() => wb.SetHyperlink(s, 3, 1, null, "Nope!A1", null));
        XlsxWriter.Save(wb, p);
        Assert.Equal(0, Errors(p));
        var wb2 = XlsxReader.Open(p); var s2 = wb2.Sheets[0];
        Assert.Equal("https://example.org/a?b=1", wb2.GetHyperlink(s2, 1, 1)!.Value.Url);
        Assert.Equal("Altro!B2", wb2.GetHyperlink(s2, 2, 1)!.Value.Location);
        wb2.RemoveHyperlink(s2, 1, 1);
        Assert.Null(wb2.GetHyperlink(s2, 1, 1)); Assert.False(wb2.Styles.Get(s2.Get(1, 1).Style).Font.Underline);
        wb2.Undo(); Assert.NotNull(wb2.GetHyperlink(s2, 1, 1));
    }

    // ---------- imposta pagina ----------
    [Fact]
    public void PageSetupPersistsAndValidates()
    {
        string p = Tmp("ps.xlsx");
        var (wb, s) = New();
        Assert.Equal(Margins.Normal, wb.GetMargins(s)); Assert.False(wb.IsLandscape(s));
        wb.SetMargins(s, Margins.Narrow); wb.SetOrientation(s, true); wb.SetPaperSize(s, 8); wb.SetPrintGridLines(s, true);
        Assert.Throws<ArgumentException>(() => wb.SetPaperSize(s, 999));
        Assert.Throws<ArgumentException>(() => wb.SetMargins(s, new Margins(-1, 0, 0, 0, 0, 0)));
        XlsxWriter.Save(wb, p);
        Assert.Equal(0, Errors(p));
        var wb2 = XlsxReader.Open(p); var s2 = wb2.Sheets[0];
        Assert.Equal(Margins.Narrow, wb2.GetMargins(s2)); Assert.True(wb2.IsLandscape(s2)); Assert.Equal(8, wb2.PaperSize(s2)); Assert.True(wb2.PrintGridLines(s2));
        wb2.SetOrientation(s2, false); wb2.Undo(); Assert.True(wb2.IsLandscape(s2));
        wb2.SetPrintGridLines(s2, false); XlsxWriter.Save(wb2, p); Assert.Equal(0, Errors(p));
        Assert.DoesNotContain("printOptions", Part(p, "xl/worksheets/sheet1.xml"));
    }
}
