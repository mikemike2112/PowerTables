using System.IO.Compression;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using PowerTables.Core;
using PowerTables.Core.IO;

namespace PowerTables.Tests;

public class TablesTests
{
    static string Tmp(string n) { string d = Path.Combine(TestUtil.Root, "tests", "work", "tables"); Directory.CreateDirectory(d); return Path.Combine(d, n); }
    static int Errors(string path) { using var d = SpreadsheetDocument.Open(path, false); return new OpenXmlValidator().Validate(d).Count(); }
    static string[] Entries(string path) { using var z = ZipFile.OpenRead(path); return z.Entries.Select(e => e.FullName).ToArray(); }
    static string Part(string path, string entry) { using var z = ZipFile.OpenRead(path); return new StreamReader(z.GetEntry(entry)!.Open()).ReadToEnd(); }

    static (Workbook, Sheet) Data()
    {
        var wb = XlsxWriter.CreateNew(); var s = wb.Sheets[0];
        wb.Do("x", () =>
        {
            wb.SetValue(s, 1, 1, "Nome"); wb.SetValue(s, 1, 2, "Nome"); wb.SetValue(s, 1, 3, 5.0);   // duplicato e numero → rinominati
            for (int r = 2; r <= 5; r++) { wb.SetValue(s, r, 1, "r" + r); wb.SetValue(s, r, 2, r * 1.0); wb.SetValue(s, r, 3, r * 2.0); }
        });
        return (wb, s);
    }

    [Fact]
    public void CreateTableFixesHeadersChecksOverlapAndUndoes()
    {
        var (wb, s) = Data();
        var t = wb.CreateTable(s, 1, 1, 5, 3, true);
        Assert.Equal("Tabella1", t.Name);
        Assert.Equal("Nome", s.GetValue(1, 1)); Assert.Equal("Colonna2", s.GetValue(1, 2)); Assert.Equal("Colonna3", s.GetValue(1, 3));
        Assert.Equal(wb.TableLookAt(s, 1, 1)!.Value.Fill, Workbook.TableStyles[0].Header);
        Assert.NotNull(wb.TableLookAt(s, 2, 1)); Assert.Null(wb.TableLookAt(s, 3, 1));
        Assert.Throws<ArgumentException>(() => wb.CreateTable(s, 3, 1, 4, 2, true));    // sovrapposta
        Assert.Throws<ArgumentException>(() => wb.CreateTable(s, 8, 1, 8, 2, true));    // intestazione senza dati
        Assert.Throws<ArgumentException>(() => wb.CreateTable(s, 8, 1, 9, 2, true, "Stile"));
        wb.Undo();
        Assert.Empty(s.Tables); Assert.Equal("Nome", s.GetValue(1, 2)); Assert.Equal(5.0, s.GetValue(1, 3));
        wb.Redo(); Assert.Single(s.Tables);
    }

    [Fact]
    public void TableWithoutHeaderInsertsHeaderRow()
    {
        var (wb, s) = Data();
        wb.CreateTable(s, 2, 1, 5, 2, false);
        var t = s.Tables[0];
        Assert.Equal((2, 1, 6, 2), (t.R1, t.C1, t.R2, t.C2));
        Assert.Equal("Colonna1", s.GetValue(2, 1)); Assert.Equal("r2", s.GetValue(3, 1));
        wb.Undo(); Assert.Equal("r2", s.GetValue(2, 1)); Assert.Empty(s.Tables);
    }

    [Fact]
    public void TableSurvivesSaveFollowsRowsAndHeaderEditsAndCanBeRemoved()
    {
        string p = Tmp("t.xlsx");
        var (wb, s) = Data();
        wb.CreateTable(s, 1, 1, 5, 3, true, "TableStyleMedium3", "Vendite");
        XlsxWriter.Save(wb, p);
        Assert.Equal(0, Errors(p));
        Assert.Contains("xl/tables/table1.xml", Entries(p));
        string tx = Part(p, "xl/tables/table1.xml");
        Assert.Contains("ref=\"A1:C5\"", tx); Assert.Contains("name=\"Vendite\"", tx); Assert.Contains("TableStyleMedium3", tx); Assert.Contains("tablePart", Part(p, "xl/worksheets/sheet1.xml"));

        var w2 = XlsxReader.Open(p); var s2 = w2.Sheets[0];
        Assert.Equal("Vendite", s2.Tables[0].Name); Assert.Equal((1, 1, 5, 3), (s2.Tables[0].R1, s2.Tables[0].C1, s2.Tables[0].R2, s2.Tables[0].C2));
        w2.InsertLines(s2, true, 1, 2);                     // due righe in testa: l'area scende
        w2.Do("h", () => w2.SetValue(s2, 3, 1, "Cliente")); // intestazione rinominata
        XlsxWriter.Save(w2, p);
        Assert.Equal(0, Errors(p));
        tx = Part(p, "xl/tables/table1.xml");
        Assert.Contains("ref=\"A3:C7\"", tx); Assert.Contains("name=\"Cliente\"", tx);

        w2.RemoveTable(s2, s2.Tables[0]);
        XlsxWriter.Save(w2, p);
        Assert.Equal(0, Errors(p));
        Assert.DoesNotContain("xl/tables/table1.xml", Entries(p)); Assert.DoesNotContain("tablePart", Part(p, "xl/worksheets/sheet1.xml"));
        w2.Undo();                                          // ripristina dopo il salvataggio
        XlsxWriter.Save(w2, p);
        Assert.Equal(0, Errors(p));
        Assert.Contains("ref=\"A3:C7\"", Part(p, Entries(p).First(e => e.StartsWith("xl/tables/table"))));
    }

    [Fact]
    public void SecondTableGetsDistinctIdAndName()
    {
        string p = Tmp("t2.xlsx");
        var (wb, s) = Data();
        wb.Do("x", () => { wb.SetValue(s, 1, 6, "A"); wb.SetValue(s, 2, 6, 1.0); });
        wb.CreateTable(s, 1, 1, 5, 3, true); wb.CreateTable(s, 1, 6, 2, 6, true);
        Assert.Equal(["Tabella1", "Tabella2"], s.Tables.Select(t => t.Name).ToArray());
        XlsxWriter.Save(wb, p);
        Assert.Equal(0, Errors(p));
        Assert.Equal(2, Entries(p).Count(e => e.StartsWith("xl/tables/table")));
    }
}
