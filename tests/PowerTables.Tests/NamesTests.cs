using System.IO.Compression;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using PowerTables.Core;
using PowerTables.Core.IO;

namespace PowerTables.Tests;

public class NamesTests
{
    static string Tmp(string n) { string d = Path.Combine(TestUtil.Root, "tests", "work", "names"); Directory.CreateDirectory(d); return Path.Combine(d, n); }
    static int Errors(string path) { using var d = SpreadsheetDocument.Open(path, false); return new OpenXmlValidator().Validate(d).Count(); }
    static string Part(string path, string entry) { using var z = ZipFile.OpenRead(path); return new StreamReader(z.GetEntry(entry)!.Open()).ReadToEnd(); }

    [Theory]
    [InlineData("Totale", true)] [InlineData("_x.y", true)] [InlineData("A1", false)] [InlineData("R1C1", false)] [InlineData("R", false)]
    [InlineData("due parole", false)] [InlineData("1abc", false)] [InlineData("", false)] [InlineData("TRUE", false)]
    public void ValidateName(string n, bool ok) => Assert.Equal(ok, Workbook.ValidateName(n) == null);

    [Fact]
    public void DefineNameFeedsFormulasAndUndoes()
    {
        var wb = new Workbook(); var s = wb.AddSheet("S");
        wb.Do("x", () => { wb.SetValue(s, 1, 1, 2.0); wb.SetValue(s, 2, 1, 3.0); wb.SetFormula(s, 1, 2, "SUM(Dati)"); });
        Assert.Equal(CellError.Name, s.GetValue(1, 2));
        wb.DefineName("Dati", "S!$A$1:$A$2");
        Assert.Equal(5.0, s.GetValue(1, 2));
        wb.DefineName("Dati", "S!$A$1:$A$1");
        Assert.Equal(2.0, s.GetValue(1, 2));
        wb.DeleteName("Dati");
        Assert.Equal(CellError.Name, s.GetValue(1, 2));
        wb.Undo(); Assert.Equal(2.0, s.GetValue(1, 2));
        wb.Undo(); Assert.Equal(5.0, s.GetValue(1, 2));
        wb.Undo(); Assert.Equal(CellError.Name, s.GetValue(1, 2)); Assert.Empty(wb.Names);
        Assert.Throws<ArgumentException>(() => wb.DefineName("A1", "S!$A$1"));
    }

    [Fact]
    public void LocalNameAndSheetRenameAreUndoable()
    {
        var wb = new Workbook(); var s = wb.AddSheet("S"); wb.AddSheet("T");
        wb.Do("x", () => wb.SetValue(s, 1, 1, 7.0));
        wb.DefineName("Loc", "S!$A$1", s);
        wb.RenameSheet(s, "Nuovo");
        Assert.Equal("Nuovo!$A$1", wb.LocalNames[(0, "LOC")]);
        wb.Undo();
        Assert.Equal("S", s.Name); Assert.Equal("S!$A$1", wb.LocalNames[(0, "LOC")]);
        wb.Redo(); Assert.Equal("Nuovo", s.Name);
    }

    [Fact]
    public void NamesAndPrintAreaSurviveSaveAndValidate()
    {
        string p = Tmp("names.xlsx");
        var wb = XlsxWriter.CreateNew(); var s = wb.Sheets[0];
        wb.Do("x", () => wb.SetValue(s, 1, 1, 1.0));
        wb.DefineName("Glob", "Foglio1!$A$1"); wb.DefineName("Loc", "Foglio1!$A$1:$A$3", s);
        wb.SetPrintArea(s, (1, 1, 5, 3));
        XlsxWriter.Save(wb, p);
        Assert.Equal(0, Errors(p));
        string book = Part(p, "xl/workbook.xml");
        Assert.Contains("name=\"Glob\"", book); Assert.Contains("localSheetId=\"0\"", book); Assert.Contains("Foglio1!$A$1:$C$5", book);
        var wb2 = XlsxReader.Open(p);
        Assert.Equal("Foglio1!$A$1", wb2.Names["Glob"]); Assert.Equal("Loc", wb2.LocalNameCase[(0, "LOC")]);
        Assert.Equal("Foglio1!$A$1:$C$5".Replace("$A$1:$C$5", "$A$1:$C$5"), wb2.Sheets[0].PrintArea!.Replace("$A$1:$C$5", "$A$1:$C$5"));
        wb2.DeleteName("Glob"); wb2.SetPrintArea(wb2.Sheets[0], null);
        XlsxWriter.Save(wb2, p);
        Assert.Equal(0, Errors(p));
        book = Part(p, "xl/workbook.xml");
        Assert.DoesNotContain("name=\"Glob\"", book); Assert.DoesNotContain("Print_Area", book); Assert.Contains("name=\"Loc\"", book);
    }

    [Fact]
    public void PrintAreaFollowsInsertedRows()
    {
        var wb = new Workbook(); var s = wb.AddSheet("S");
        wb.SetPrintArea(s, (2, 1, 4, 2));
        wb.InsertLines(s, true, 1, 2);
        Assert.Equal("S!$A$4:$B$6", s.PrintArea);
        wb.Undo(); Assert.Equal("S!$A$2:$B$4", s.PrintArea);
        wb.Undo(); Assert.Null(s.PrintArea);
    }
}
