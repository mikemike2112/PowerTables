using System.Diagnostics;
using System.Text.Json;
using PowerTables.Core;
using PowerTables.Core.IO;

namespace PowerTables.Tests;

/// <summary>Test di carico reale: 1.000.000 di righe popolate (numeri, testo, date, formule).</summary>
[Trait("Category", "Load")]
public class LoadTests
{
    public static string LoadFile => Path.Combine(TestUtil.Root, "tests", "work", "load-1m.xlsx");
    const int Rows = 1_000_000;

    [Fact]
    public void GenerateLoadFile()
    {
        var sw = Stopwatch.StartNew();
        var wb = XlsxWriter.CreateNew(); var s = wb.Sheets[0]; s.Name = "Dati";
        uint money = wb.Styles.Derive(0, x => x with { NumFmtId = 4 });
        uint date = wb.Styles.Derive(0, x => x with { NumFmtId = 14 });
        uint bold = wb.Styles.Derive(0, x => x with { Font = x.Font with { Bold = true } });
        string[] hdr = ["ID", "Descrizione", "Importo", "Data", "Imponibile x1,22", "Etichetta"];
        for (int c = 0; c < hdr.Length; c++) s.Put(1, c + 1, Cell.FromValue(hdr[c], bold));
        var rnd = new Random(7);
        for (int r = 2; r <= Rows + 1; r++)
        {
            s.Put(r, 1, Cell.FromValue((double)(r - 1)));
            s.Put(r, 2, Cell.FromValue("Voce " + (r - 1) % 5000));
            s.Put(r, 3, Cell.FromValue(Math.Round(rnd.NextDouble() * 1000, 2), money));
            s.Put(r, 4, Cell.FromValue(43831.0 + rnd.Next(0, 2000), date));
            var c5 = new Cell { F = new FormulaInfo { Text = $"C{r}*1.22" } };
            s.Put(r, 5, c5);
            var c6 = new Cell { F = new FormulaInfo { Text = $"B{r}&\"-\"&A{r}" } };
            s.Put(r, 6, c6);
            s.DirtyRows.Add(r);
        }
        s.Put(1, 8, Cell.FromValue("Totale")); s.Put(1, 9, new Cell { F = new FormulaInfo { Text = $"SUM(E2:E{Rows + 1})" } });
        s.Put(2, 8, Cell.FromValue("Righe > 500")); s.Put(2, 9, new Cell { F = new FormulaInfo { Text = $"COUNTIF(C2:C{Rows + 1},\">500\")" } });
        s.Put(3, 8, Cell.FromValue("Cerca ID 777777")); s.Put(3, 9, new Cell { F = new FormulaInfo { Text = $"VLOOKUP(777777,A2:B{Rows + 1},2,FALSE)" } });
        s.DirtyRows.Add(1); s.DirtyRows.Add(2); s.DirtyRows.Add(3);
        long gen = sw.ElapsedMilliseconds; sw.Restart();
        wb.RecalcAll();
        long recalc = sw.ElapsedMilliseconds; sw.Restart();
        Directory.CreateDirectory(Path.GetDirectoryName(LoadFile)!);
        XlsxWriter.Save(wb, LoadFile);
        long save = sw.ElapsedMilliseconds;
        File.WriteAllText(Path.Combine(TestUtil.Root, "tests", "work", "load-gen.json"), JsonSerializer.Serialize(new { generateMs = gen, recalcAllFirstMs = recalc, saveMs = save, bytes = new FileInfo(LoadFile).Length, recalcFormulas = wb.LastRecalcCount }));
        
        Assert.Equal("Voce " + 777777 % 5000 + "-777777", s.GetValue(777778, 6));
    }
}
