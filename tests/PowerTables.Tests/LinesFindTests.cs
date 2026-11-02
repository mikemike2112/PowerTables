using System.IO.Compression;
using PowerTables.Core;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using PowerTables.Core.IO;

namespace PowerTables.Tests;

public class LinesFindTests
{
    static (Workbook, Sheet) New() { var wb = new Workbook(); return (wb, wb.AddSheet("S")); }

    [Fact]
    public void InsertRowsShiftsCellsAndFormulas()
    {
        var (wb, s) = New();
        wb.Do("x", () => { wb.SetValue(s, 1, 1, 1.0); wb.SetValue(s, 2, 1, 2.0); wb.SetValue(s, 3, 1, 3.0); wb.SetFormula(s, 4, 1, "SUM(A1:A3)+A2+$A$3"); });
        wb.InsertLines(s, true, 2, 2);
        Assert.Equal(1.0, s.GetValue(1, 1)); Assert.Null(s.GetValue(2, 1)); Assert.Equal(2.0, s.GetValue(4, 1));
        Assert.Equal("SUM(A1:A5)+A4+$A$5", s.Get(6, 1).F!.Text);
        Assert.Equal(11.0, s.GetValue(6, 1));
        wb.Undo();
        Assert.Equal("SUM(A1:A3)+A2+$A$3", s.Get(4, 1).F!.Text); Assert.Equal(2.0, s.GetValue(2, 1)); Assert.Equal(11.0, s.GetValue(4, 1));
        wb.Redo();
        Assert.Equal(11.0, s.GetValue(6, 1));
    }

    [Fact]
    public void DeleteRowsGivesRefAndShrinksRanges()
    {
        var (wb, s) = New();
        wb.Do("x", () => { for (int i = 1; i <= 5; i++) wb.SetValue(s, i, 1, (double)i); wb.SetFormula(s, 1, 2, "SUM(A1:A5)"); wb.SetFormula(s, 2, 2, "A3*2"); wb.SetFormula(s, 4, 2, "A4"); });
        wb.DeleteLines(s, true, 3, 1);
        Assert.Equal("SUM(A1:A4)", s.Get(1, 2).F!.Text);
        Assert.Equal(12.0, s.GetValue(1, 2));
        Assert.Equal("#REF!*2", s.Get(2, 2).F!.Text);
        Assert.Equal("A3", s.Get(3, 2).F!.Text);   // era A4 → riga 3, e la cella stessa è salita
        wb.Undo();
        Assert.Equal("A3*2", s.Get(2, 2).F!.Text); Assert.Equal(15.0, s.GetValue(1, 2));
    }

    [Fact]
    public void InsertColsAcrossSheetsMergesAndWidths()
    {
        var wb = new Workbook(); var a = wb.AddSheet("A"); var b = wb.AddSheet("B");
        wb.Do("x", () => { wb.SetValue(a, 1, 2, 5.0); wb.SetFormula(b, 1, 1, "A!B1+'A'!C1"); wb.Merge(a, 1, 2, 2, 3); wb.SetColWidth(a, 3, 20); });
        wb.InsertLines(a, false, 2, 1);
        Assert.Equal(5.0, a.GetValue(1, 3));
        Assert.Equal("A!C1+A!D1", b.Get(1, 1).F!.Text);
        Assert.Equal(new MergeRange(1, 3, 2, 4), a.Merges[0]);
        Assert.True(a.Cols.ContainsKey(4) && !a.Cols.ContainsKey(3));
        wb.DeleteLines(a, false, 3, 1);
        Assert.Equal("#REF!+A!C1", b.Get(1, 1).F!.Text.Replace("A!#REF!", "#REF!"));
    }

    [Fact]
    public void StructuralChangesSurviveSaveAndRichTextIsKept()
    {
        string dir = Path.Combine(TestUtil.Root, "tests", "work", "lines"); Directory.CreateDirectory(dir);
        string path = Path.Combine(dir, "rich.xlsx");
        var wb = XlsxReader.Open(TestUtil.Corpus("office-made", "rich-text.xlsx"));
        var s = wb.Sheets[0];
        string before = (string)s.GetValue(1, 1)!;
        wb.InsertLines(s, true, 1, 1);
        XlsxWriter.Save(wb, path);
        var w2 = XlsxReader.Open(path);
        Assert.Equal(before, w2.Sheets[0].GetValue(2, 1));
        Assert.Null(w2.Sheets[0].GetValue(1, 1));
        using var z = ZipFile.OpenRead(path);
        string xml = new StreamReader(z.Entries.First(e => e.FullName.StartsWith("xl/worksheets/sheet1")).Open()).ReadToEnd();
        Assert.Contains("<rPr><i />", xml);   // i runs del testo ricco ci sono ancora (inline)
    }

    [Fact]
    public void FindAndReplace()
    {
        var (wb, s) = New();
        wb.Do("x", () => { wb.SetValue(s, 1, 1, "Mela rossa"); wb.SetValue(s, 2, 2, "mela"); wb.SetValue(s, 3, 1, 12.0); wb.SetFormula(s, 3, 3, "A3*2"); });
        Assert.Equal((1, 1), wb.Find(s, "mela", 0, 0));
        Assert.Equal((2, 2), wb.Find(s, "mela", 1, 1));
        Assert.Equal((1, 1), wb.Find(s, "mela", 2, 2));          // giro completo
        Assert.Equal((2, 2), wb.Find(s, "mela", 1, 1, true, true)); // maiuscole: solo "mela"
        Assert.Null(wb.Find(s, "xyz", 1, 1));
        Assert.Equal((3, 3), wb.Find(s, "A3*", 1, 1));          // anche nelle formule
        Assert.Equal(2, wb.ReplaceAll(s, "mela", "pera"));
        Assert.Equal("pera rossa", s.GetValue(1, 1)); Assert.Equal("pera", s.GetValue(2, 2));
        Assert.Equal(1, wb.ReplaceAll(s, "A3", "A1", false, false));
        Assert.Equal("A1*2", s.Get(3, 3).F!.Text);
        Assert.Equal(1, wb.ReplaceAll(s, "12", "34", false, true));
        Assert.Equal(34.0, s.GetValue(3, 1));
        wb.Undo(); Assert.Equal(12.0, s.GetValue(3, 1));
    }

    static int Errors(string path) { using var d = SpreadsheetDocument.Open(path, false); return new OpenXmlValidator().Validate(d).Count(); }
    static string Part(string path, string entry) { using var z = ZipFile.OpenRead(path); return new StreamReader(z.GetEntry(entry)!.Open()).ReadToEnd(); }

    [Fact]
    public void ConditionalFormatValidationAndNamesFollowInsertAndDelete()
    {
        string dir = Path.Combine(TestUtil.Root, "tests", "work", "lines"); Directory.CreateDirectory(dir);
        foreach (var n in new[] { "conditional-format", "data-validation", "formula-named-range" })
        {
            string src = TestUtil.Corpus("office-made", n + ".xlsx"), outp = Path.Combine(dir, n + "-out.xlsx");
            var wb = XlsxReader.Open(src);
            wb.InsertLines(wb.Sheets[0], true, 1, 2);   // due righe in testa
            XlsxWriter.Save(wb, outp);
            Assert.True(Errors(outp) <= Errors(src), n);
            string sheet = Part(outp, "xl/worksheets/sheet1.xml"), book = Part(outp, "xl/workbook.xml");
            if (n == "conditional-format") { Assert.Contains("sqref=\"A3:A12\"", sheet); Assert.DoesNotContain("sqref=\"A1:A10\"", sheet); }
            if (n == "data-validation") { Assert.Contains("sqref=\"A3:A5\"", sheet); Assert.Contains("sqref=\"B3:B5\"", sheet); }
            if (n == "formula-named-range") Assert.Contains("Foglio1!$A$3:$A$5", book);
            // eliminazione: la regola sulle righe tolte sparisce, il resto si accorcia
            wb.DeleteLines(wb.Sheets[0], true, 3, 2);
            XlsxWriter.Save(wb, outp);
            Assert.True(Errors(outp) <= Errors(src), n + " del");
            if (n == "conditional-format") Assert.Contains("sqref=\"A3:A10\"", Part(outp, "xl/worksheets/sheet1.xml"));
            if (n == "data-validation") Assert.Contains("sqref=\"A3:A3\"", Part(outp, "xl/worksheets/sheet1.xml").Replace("sqref=\"A3\"", "sqref=\"A3:A3\""));
            if (n == "formula-named-range") Assert.Contains("Foglio1!$A$3:$A$3", Part(outp, "xl/workbook.xml").Replace("$A$3\"", "$A$3:$A$3\"").Replace("$A$3<", "$A$3:$A$3<"));
        }
    }
}
