using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using PowerTables.Core;
using PowerTables.Core.IO;

namespace PowerTables.Tests;

public class SheetOpsTests
{
    static string Tmp(string n) { string d = Path.Combine(TestUtil.Root, "tests", "work", "sheetops"); Directory.CreateDirectory(d); return Path.Combine(d, n); }
    static int Errors(string path) { using var d = SpreadsheetDocument.Open(path, false); return new OpenXmlValidator().Validate(d).Count(); }

    [Fact]
    public void InsertMoveDeleteAreUndoableAndKeepFormulasAndLocalNames()
    {
        var wb = new Workbook(); var a = wb.AddSheet("A"); var b = wb.AddSheet("B");
        wb.Do("x", () => { wb.SetValue(b, 1, 1, 5.0); wb.SetFormula(a, 1, 1, "B!A1*2"); });
        wb.DefineName("LocB", "B!$A$1", b);
        Assert.Equal(10.0, a.GetValue(1, 1));

        var n = wb.InsertSheet("Nuovo", 0);
        Assert.Equal(["Nuovo", "A", "B"], wb.Sheets.Select(x => x.Name).ToArray());
        Assert.Equal(2, b.Index); Assert.Equal("B!$A$1", wb.LocalNames[(2, "LOCB")]);
        Assert.Equal(10.0, a.GetValue(1, 1));
        wb.Undo(); Assert.Equal(["A", "B"], wb.Sheets.Select(x => x.Name).ToArray()); Assert.Equal("B!$A$1", wb.LocalNames[(1, "LOCB")]);
        wb.Redo(); wb.Undo();

        wb.MoveSheet(a, 1);
        Assert.Equal(["B", "A"], wb.Sheets.Select(x => x.Name).ToArray()); Assert.Equal("B!$A$1", wb.LocalNames[(0, "LOCB")]);
        Assert.Equal(10.0, a.GetValue(1, 1));
        wb.Undo(); Assert.Equal(0, a.Index);

        wb.DeleteSheet(b);
        Assert.Equal(CellError.Ref, a.GetValue(1, 1)); Assert.Empty(wb.LocalNames);
        wb.Undo();
        Assert.Equal(10.0, a.GetValue(1, 1)); Assert.Equal("B!A1*2", a.Get(1, 1).F!.Text); Assert.Equal("B!$A$1", wb.LocalNames[(1, "LOCB")]);

        Assert.Throws<ArgumentException>(() => wb.InsertSheet("a"));          // nome duplicato
        Assert.Throws<ArgumentException>(() => wb.InsertSheet("x/y"));
        wb.DeleteSheet(b);
        Assert.Throws<InvalidOperationException>(() => wb.DeleteSheet(a));
    }

    [Fact]
    public void ReorderDeleteAndRestoreSurviveSaveAndValidate()
    {
        string p = Tmp("ops.xlsx");
        var wb = XlsxWriter.CreateNew(); var a = wb.Sheets[0];
        var b = wb.InsertSheet("Due"); var c = wb.InsertSheet("Tre");
        wb.Do("x", () => { wb.SetValue(a, 1, 1, "uno"); wb.SetValue(b, 1, 1, "due"); wb.SetValue(c, 1, 1, "tre"); });
        wb.DefineName("LocC", "Tre!$A$1", c);
        wb.SetPrintArea(c, (1, 1, 2, 2));
        XlsxWriter.Save(wb, p); Assert.Equal(0, Errors(p));

        wb.MoveSheet(c, -2);                 // Tre, Foglio1, Due
        XlsxWriter.Save(wb, p); Assert.Equal(0, Errors(p));
        var w2 = XlsxReader.Open(p);
        Assert.Equal(["Tre", "Foglio1", "Due"], w2.Sheets.Select(x => x.Name).ToArray());
        Assert.Equal("tre", w2.Sheets[0].GetValue(1, 1)); Assert.Equal("Tre!$A$1", w2.LocalNames[(0, "LOCC")]); Assert.NotNull(w2.Sheets[0].PrintArea);

        w2.DeleteSheet(w2.Sheets[1]);        // elimina Foglio1
        XlsxWriter.Save(w2, p); Assert.Equal(0, Errors(p));
        var w3 = XlsxReader.Open(p);
        Assert.Equal(["Tre", "Due"], w3.Sheets.Select(x => x.Name).ToArray()); Assert.Equal("due", w3.Sheets[1].GetValue(1, 1)); Assert.Equal("Tre!$A$1", w3.LocalNames[(0, "LOCC")]);

        w2.Undo();                           // riporta Foglio1 dopo il salvataggio
        XlsxWriter.Save(w2, p); Assert.Equal(0, Errors(p));
        var w4 = XlsxReader.Open(p);
        Assert.Equal(["Tre", "Foglio1", "Due"], w4.Sheets.Select(x => x.Name).ToArray()); Assert.Equal("uno", w4.Sheets[1].GetValue(1, 1));
    }

    [Fact]
    public void ClearFilterCriteriaKeepsArrows()
    {
        var wb = new Workbook(); var s = wb.AddSheet("S");
        wb.Do("x", () => { wb.SetValue(s, 1, 1, "h"); wb.SetValue(s, 2, 1, "a"); wb.SetValue(s, 3, 1, "b"); });
        wb.SetAutoFilter(s, 1, 1, 3, 1);
        wb.SetColumnFilter(s, 1, ["a"]);
        Assert.True(s.Rows[3].Hidden);
        wb.ClearFilterCriteria(s);
        Assert.NotNull(s.Filter); Assert.False(s.Rows[3].Hidden); Assert.Empty(s.Filter!.Allowed);
        wb.Undo(); Assert.True(s.Rows[3].Hidden);
    }

    [Fact]
    public void RenameSheetIsOneUndoStep()
    {
        var wb = new Workbook(); var s = wb.AddSheet("S"); wb.AddSheet("T");
        wb.Do("x", () => wb.SetFormula(wb.Sheets[1], 1, 1, "S!A1"));
        wb.RenameSheet(s, "Nuovo");
        Assert.Equal("Nuovo!A1", wb.Sheets[1].Get(1, 1).F!.Text);
        wb.Undo(); Assert.Equal("S", s.Name); Assert.Equal("S!A1", wb.Sheets[1].Get(1, 1).F!.Text);
    }
}
