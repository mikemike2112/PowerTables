using System.IO.Compression;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using PowerTables.Core;
using PowerTables.Core.IO;

namespace PowerTables.Tests;

public class PivotTests
{
    static string[] Entries(string path) { using var z = ZipFile.OpenRead(path); return z.Entries.Select(e => e.FullName).ToArray(); }
    static string[] Errors(string path) { using var d = SpreadsheetDocument.Open(path, false); return new OpenXmlValidator().Validate(d).Select(e => e.Path?.XPath + " " + e.Description).ToArray(); }

    static (Workbook, Sheet) Data()
    {
        var wb = XlsxWriter.CreateNew(); var s = wb.Sheets[0];
        wb.Do("x", () =>
        {
            wb.SetValue(s, 1, 1, "Reparto"); wb.SetValue(s, 1, 2, "Mese"); wb.SetValue(s, 1, 3, "Vendite");
            (string, string, double)[] d = [("Nord", "gen", 10), ("Sud", "gen", 5), ("Nord", "feb", 7), ("Est", "feb", 1), ("sud", "mar", 2)];
            for (int i = 0; i < d.Length; i++) { wb.SetValue(s, i + 2, 1, d[i].Item1); wb.SetValue(s, i + 2, 2, d[i].Item2); wb.SetValue(s, i + 2, 3, d[i].Item3); }
        });
        return (wb, s);
    }

    [Fact]
    public void SumAndCountGroupsValidationsAndUndo()
    {
        var (wb, s) = Data();
        var p = wb.CreatePivot(s, (1, 1, 6, 3), 1, 3, false, s, 1, 6);
        Assert.Equal("Etichette di riga", s.GetValue(1, 6)); Assert.Equal("Somma di Vendite", s.GetValue(1, 7));
        Assert.Equal(["Est", "Nord", "Sud"], p.RowItems.Select(x => (string)x).ToArray());
        Assert.Equal(1.0, s.GetValue(2, 7)); Assert.Equal(17.0, s.GetValue(3, 7)); Assert.Equal(7.0, s.GetValue(4, 7)); Assert.Equal(25.0, s.GetValue(5, 7)); Assert.Equal("Totale complessivo", s.GetValue(5, 6));
        wb.Undo(); Assert.Empty(s.Pivots); Assert.Null(s.GetValue(1, 6));
        var c = wb.CreatePivot(s, (1, 1, 6, 3), 1, 2, true, s, 10, 1);
        Assert.Equal("Conteggio di Mese", s.GetValue(10, 2)); Assert.Equal(5.0, s.GetValue(14, 2)); Assert.Equal("TabellaPivot1", c.Name);
        Assert.Throws<ArgumentException>(() => wb.CreatePivot(s, (1, 1, 6, 3), 1, 1, false, s, 20, 1));     // stessa colonna
        Assert.Throws<ArgumentException>(() => wb.CreatePivot(s, (1, 1, 6, 3), 1, 2, false, s, 20, 1));     // somma di testo
        Assert.Throws<ArgumentException>(() => wb.CreatePivot(s, (1, 1, 6, 3), 1, 3, false, s, 3, 1));      // sovrapposta ai dati
        Assert.Throws<ArgumentException>(() => wb.CreatePivot(s, (2, 1, 6, 3), 1, 3, false, s, 20, 1));     // niente intestazioni
    }

    [Fact]
    public void WrittenValidatesExcelOpensItAndUndoAfterSaveRemovesParts()
    {
        string dir = Path.Combine(TestUtil.Root, "tests", "work", "pivot"); if (Directory.Exists(dir)) Directory.Delete(dir, true); Directory.CreateDirectory(dir);
        string p = Path.Combine(dir, "p.xlsx");
        var (wb, s) = Data();
        wb.CreatePivot(s, (1, 1, 6, 3), 1, 3, false, s, 1, 6);
        XlsxWriter.Save(wb, p);
        Assert.Empty(Errors(p));
        Assert.Contains(Entries(p), e => e.StartsWith("xl/pivotTables/pivotTable")); Assert.Contains(Entries(p), e => e.Contains("pivotCacheDefinition"));
        var (code, outp) = TestUtil.PsT("excel-batch.ps1", 300_000, "-InDir", dir, "-OutDir", Path.Combine(dir, "json"));
        Assert.True(code == 0, outp);
        var j = TestUtil.ReadJson(Path.Combine(dir, "json", "p.xlsx.json")).RootElement;
        Assert.True(j.GetProperty("ok").GetBoolean(), j.GetProperty("error").ToString());
        Assert.Equal(1, j.GetProperty("sheets")[0].GetProperty("pivots").GetInt32());
        // Annulla dopo il salvataggio: le parti spariscono
        wb.Undo(); XlsxWriter.Save(wb, p);
        Assert.Empty(Errors(p)); Assert.DoesNotContain(Entries(p), e => e.StartsWith("xl/pivotTables/pivotTable")); Assert.DoesNotContain(Entries(p), e => e.Contains("pivotCacheDefinition"));
    }
}
