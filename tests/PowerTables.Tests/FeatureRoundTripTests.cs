using System.Globalization;
using System.Text.Json;
using PowerTables.Core;
using PowerTables.Core.IO;

namespace PowerTables.Tests;

/// <summary>Ogni funzionalità P0 di editing viene eseguita col clone, salvata e riaperta in Excel vero.</summary>
[Collection("OfficeCom"), Trait("Category", "OfficeCom")]
public class FeatureRoundTripTests
{
    static readonly CultureInfo It = CultureInfo.GetCultureInfo("it-IT");

    static JsonElement Run(string name, Action<Workbook> build, out Workbook wb)
    {
        var dir = Path.Combine(TestUtil.Root, "tests", "work", "feat", name);
        if (Directory.Exists(dir)) Directory.Delete(dir, true);
        Directory.CreateDirectory(dir);
        wb = XlsxWriter.CreateNew();
        build(wb);
        XlsxWriter.Save(wb, Path.Combine(dir, name + ".xlsx"));
        var outDir = Path.Combine(dir, "json");
        var (code, outp) = TestUtil.PsT("excel-batch.ps1", 300_000, "-InDir", dir, "-OutDir", outDir);
        Assert.True(code == 0, outp);
        return TestUtil.ReadJson(Path.Combine(outDir, name + ".xlsx.json")).RootElement.Clone();
    }

    static JsonElement Cell(JsonElement j, int sheet, int r, int c) =>
        j.GetProperty("sheets")[sheet].GetProperty("cells").EnumerateArray().First(x => x.GetProperty("r").GetInt32() == r && x.GetProperty("c").GetInt32() == c);

    [Fact]
    public void Formatting_Merge_Widths_Freeze_Filter_Sort_AreRealInExcel()
    {
        var j = Run("features", wb =>
        {
            var s = wb.Sheets[0]; s.Name = "Dati";
            wb.Do("dati", () =>
            {
                string[] h = ["Nome", "Gruppo", "Importo", "Data", "Quota"];
                for (int c = 0; c < h.Length; c++) wb.SetValue(s, 1, c + 1, h[c]);
                string[] n = ["Zeta", "Alfa", "Beta", "Gamma", "Delta", "Omega"]; string[] g = ["x", "y", "x", "y", "x", "y"];
                for (int i = 0; i < n.Length; i++)
                {
                    int r = i + 2;
                    wb.SetValue(s, r, 1, n[i]); wb.SetValue(s, r, 2, g[i]); wb.SetValue(s, r, 3, (i + 1) * 1234.5);
                    wb.SetInput(s, r, 4, $"{i + 1:00}/03/2024", It); wb.SetFormula(s, r, 5, $"C{r}/SUM($C$2:$C$7)");
                }
                wb.SetInput(s, 9, 1, "12,5%", It);
            });
            wb.FormatRange(s, 1, 1, 1, 5, x => x with { Font = x.Font with { Bold = true, Color = "FFFF0000" }, FillColor = "FFFFFF00", HAlign = "center" });
            wb.FormatRange(s, 2, 3, 7, 3, x => x with { NumFmtId = 4 });
            wb.FormatRange(s, 2, 5, 7, 5, x => x with { NumFmtId = 10 });
            wb.FormatRange(s, 2, 1, 2, 1, x => x with { Font = x.Font with { Italic = true, Underline = true, Name = "Arial", Size = 14 }, Wrap = true, Border = new BorderInfo(new("thin", null), new("thin", null), new("medium", null), new("thin", null)) });
            wb.Merge(s, 10, 1, 11, 3); wb.SetValue(s, 10, 1, "Unita");
            wb.SetColWidth(s, 1, 20); wb.SetColWidth(s, 3, 15.5); wb.SetRowHeight(s, 3, 33);
            wb.SetFreeze(s, 1, 1);
            wb.Sort(s, 1, 1, 7, 5, [(1, true)], true);
            wb.SetAutoFilter(s, 1, 1, 7, 5);
            wb.SetColumnFilter(s, 2, ["x"], It);
            var s2 = wb.AddSheet("Riepilogo");
            wb.Do("r", () => { wb.SetFormula(s2, 1, 1, "SUM(Dati!C2:C7)"); wb.SetFormula(s2, 2, 1, "XLOOKUP(\"Beta\",Dati!A2:A7,Dati!C2:C7)"); });
        }, out var w);
        var sh = j.GetProperty("sheets");
        Assert.Equal("Dati", sh[0].GetProperty("name").GetString()); Assert.Equal("Riepilogo", sh[1].GetProperty("name").GetString());
        Assert.Equal("Alfa", Cell(j, 0, 2, 1).GetProperty("v").GetString()); Assert.Equal("Zeta", Cell(j, 0, 7, 1).GetProperty("v").GetString());
        Assert.Equal("=C2/SUM($C$2:$C$7)", Cell(j, 0, 2, 5).GetProperty("f").GetString());
        var h = Cell(j, 0, 1, 1);
        Assert.True(h.GetProperty("b").GetBoolean()); Assert.Equal(255, h.GetProperty("fc").GetInt32()); Assert.Equal(65535, h.GetProperty("fill").GetInt32()); Assert.Equal(-4108, h.GetProperty("ha").GetInt32());
        Assert.Equal("#,##0.00", Cell(j, 0, 2, 3).GetProperty("nf").GetString());
        Assert.Equal("0.00%", Cell(j, 0, 2, 5).GetProperty("nf").GetString());
        Assert.True(Cell(j, 0, 7, 1).GetProperty("i").GetBoolean()); Assert.True(Cell(j, 0, 7, 1).GetProperty("wrap").GetBoolean());
        Assert.Equal("m/d/yyyy", Cell(j, 0, 2, 4).GetProperty("nf").GetString());
        Assert.Equal(0.125, Cell(j, 0, 9, 1).GetProperty("v").GetDouble());
        Assert.Equal("A10:C11", Cell(j, 0, 10, 1).GetProperty("merge").GetString());
        var cw = sh[0].GetProperty("colWidths").EnumerateArray().Select(x => x.GetDouble()).ToList();
        Assert.InRange(cw[0], 19.7, 20.1); Assert.InRange(cw[2], 15.3, 15.7);
        var rh = sh[0].GetProperty("rowHeights").EnumerateArray().Select(x => x.GetDouble()).ToList();
        Assert.Equal(33, rh[2], 1);
        Assert.True(j.GetProperty("freeze").GetProperty("on").GetBoolean()); Assert.Equal(1, j.GetProperty("freeze").GetProperty("row").GetInt32()); Assert.Equal(1, j.GetProperty("freeze").GetProperty("col").GetInt32());
        Assert.True(sh[0].GetProperty("filter").GetBoolean());
        Assert.Contains(rh, v => v == 0);
        Assert.Equal(7 * 1234.5 * 3, Cell(j, 1, 1, 1).GetProperty("v").GetDouble(), 6);
        Assert.Equal("=XLOOKUP(\"Beta\",Dati!A2:A7,Dati!C2:C7)", Cell(j, 1, 2, 1).GetProperty("f").GetString());
        Assert.Equal(3 * 1234.5, Cell(j, 1, 2, 1).GetProperty("v").GetDouble(), 6);
    }

    [Fact]
    public void Undo_Redo_RestoreFileState()
    {
        var wb = new Workbook(); var s = wb.AddSheet("S");
        wb.Do("a", () => wb.SetValue(s, 1, 1, 1.0));
        wb.Do("b", () => wb.SetFormula(s, 1, 2, "A1+1"));
        wb.Merge(s, 3, 1, 4, 2); wb.SetColWidth(s, 2, 30); wb.SetFreeze(s, 2, 0);
        wb.Undo(); wb.Undo(); wb.Undo();
        Assert.Empty(s.Merges); Assert.Equal(0, s.FrozenRows);
        wb.Redo(); wb.Redo(); wb.Redo();
        Assert.Single(s.Merges); Assert.Equal(2, s.FrozenRows); Assert.Equal(30 + 5.0 / 7, s.ColWidth(2), 1);
        wb.Undo(); wb.Undo(); wb.Undo(); wb.Undo(); wb.Undo();
        Assert.Equal(CellKind.Empty, s.Get(1, 1).Kind); Assert.Null(s.Get(1, 2).F);
    }

    [Fact]
    public void CsvRoundTripThroughExcel()
    {
        var j = Run("csvimport", wb =>
        {
            Csv.Import(wb, wb.Sheets[0], "Nome;Importo;Data\r\nAlfa;1.234,50;15/03/2024\r\nBeta;99,9;01/12/2023\r\n", ';', It);
        }, out _);
        Assert.Equal(1234.5, Cell(j, 0, 2, 2).GetProperty("v").GetDouble());
        Assert.Equal(45366, Cell(j, 0, 2, 3).GetProperty("v").GetDouble());
    }

    /// <summary>Modificare una cella nella stessa riga di una stringa a testo ricco non deve far perdere i runs formattati (verificato in Excel).</summary>
    [Fact]
    public void RichText_SurvivesEditOfAnotherCellInTheSameRow()
    {
        var dir = Path.Combine(TestUtil.Root, "tests", "work", "feat", "rich");
        if (Directory.Exists(dir)) Directory.Delete(dir, true);
        Directory.CreateDirectory(dir);
        string src = TestUtil.Corpus("office-made", "rich-text.xlsx");
        File.Copy(src, Path.Combine(dir, "orig.xlsx"));
        var wb = XlsxReader.Open(src);
        wb.Do("e", () => { wb.SetValue(wb.Sheets[0], 1, 2, 99.0); wb.SetValue(wb.Sheets[0], 2, 2, 100.0); });
        XlsxWriter.Save(wb, Path.Combine(dir, "edited.xlsx"));
        var outDir = Path.Combine(dir, "json");
        var (code, outp) = TestUtil.PsT("excel-batch.ps1", 300_000, "-InDir", dir, "-OutDir", outDir);
        Assert.True(code == 0, outp);
        var o = TestUtil.ReadJson(Path.Combine(outDir, "orig.xlsx.json")).RootElement; var e = TestUtil.ReadJson(Path.Combine(outDir, "edited.xlsx.json")).RootElement;
        foreach (var (r, c) in new[] { (1, 1), (2, 1) })
        {
            string a = Cell(o, 0, r, c).GetProperty("rt").GetString()!, b = Cell(e, 0, r, c).GetProperty("rt").GetString()!;
            Assert.NotEqual(a.Split(';')[0], a.Split(';')[1]);   // il testo ricco ha davvero formattazioni diverse nel file originale
            Assert.Equal(a, b);
        }
        Assert.Equal(99.0, Cell(e, 0, 1, 2).GetProperty("v").GetDouble());
        Assert.Equal("Grassetto e rosso corsivo", Cell(e, 0, 1, 1).GetProperty("v").GetString());
    }
}
