using System.Globalization;
using PowerTables.Core;
using PowerTables.Core.IO;

namespace PowerTables.Tests;

/// <summary>Il motore ricalcola i file reali prodotti da Excel e deve ritrovare i valori in cache scritti da Excel stesso.</summary>
public class CorpusEngineTests
{
    public static IEnumerable<object[]> Files() => Directory.GetFiles(TestUtil.Corpus("office-made"), "*.xlsx").Select(f => new object[] { Path.GetFileName(f) }).OrderBy(x => (string)x[0]);

    // celle che per natura cambiano a ogni calcolo
    static bool Volatile(string f) => f.Contains("TODAY(") || f.Contains("NOW(") || f.Contains("RAND(");

    [Theory, MemberData(nameof(Files))]
    public void Recalc_ReproducesExcelCachedValues(string file)
    {
        var wb = XlsxReader.Open(Path.Combine(TestUtil.Corpus("office-made"), file));
        var before = new Dictionary<(int, long), Cell>();
        foreach (var s in wb.Sheets) foreach (var kv in s.CellMap) if (kv.Value.F is { Opaque: false } f && !Volatile(f.Text)) before[(s.Index, kv.Key)] = kv.Value;
        wb.RecalcAll();
        var bad = new List<string>();
        foreach (var ((si, key), old) in before)
        {
            var now = wb.Sheets[si].Get(Addr.Row(key), Addr.Col(key));
            bool same = old.Kind == now.Kind && (old.Kind == CellKind.Num ? Math.Abs(old.Num - now.Num) <= 1e-9 * Math.Max(1, Math.Abs(old.Num)) : old.Num.Equals(now.Num) && Equals(old.Obj, now.Obj));
            if (!same) bad.Add($"{wb.Sheets[si].Name}!{Addr.A1(Addr.Row(key), Addr.Col(key))} [{old.F!.Text}] Excel={Show(old)} mio={Show(now)}");
        }
        Assert.True(bad.Count == 0, file + ": " + string.Join("; ", bad.Take(8)));
    }

    static string Show(Cell c) => c.Kind switch { CellKind.Num => c.Num.ToString("R", CultureInfo.InvariantCulture), CellKind.Str => "\"" + c.Obj + "\"", CellKind.Bool => c.Num != 0 ? "TRUE" : "FALSE", CellKind.Err => Errors.Text((CellError)(int)c.Num), _ => "<vuoto>" };

    [Fact]
    public void ArrayAndDynamicFormulas_KeepFormulaAndCachedValue()
    {
        var wb = XlsxReader.Open(TestUtil.Corpus("office-made", "formula-dynamic-array.xlsx"));
        var s = wb.Sheets[0];
        var c1 = s.Get(1, 3);
        Assert.NotNull(c1.F); Assert.True(c1.F!.Opaque);            // SEQUENCE: formula originale intatta
        Assert.Equal(1.0, c1.Num);
        Assert.Equal(5.0, s.GetValue(5, 3));                          // valori "spill" restano valori
        wb.RecalcAll();
        Assert.Equal(1.0, s.Get(1, 3).Num);
        var p = Path.Combine(Path.GetTempPath(), "dyn-" + Guid.NewGuid().ToString("N") + ".xlsx");
        wb.Do("x", () => wb.SetValue(s, 20, 1, 1.0));
        XlsxWriter.Save(wb, p);
        var wb2 = XlsxReader.Open(p);
        Assert.True(wb2.Sheets[0].Get(1, 3).F!.Opaque);
        Assert.Equal(wb.Sheets[0].Get(1, 3).F!.RawF, wb2.Sheets[0].Get(1, 3).F!.RawF);
        File.Delete(p);
    }

    [Fact]
    public void UnknownFunction_KeepsCachedValue()
    {
        var wb = new Workbook(); var s = wb.AddSheet("S");
        s.Put(1, 1, new Cell { F = new FormulaInfo { Text = "FUTUREFN(1)" }, Kind = CellKind.Num, Num = 42 });
        wb.RecalcAll();
        Assert.Equal(42.0, s.GetValue(1, 1));
        wb.Do("x", () => wb.SetFormula(s, 2, 1, "FUTUREFN(2)"));
        Assert.Equal(CellError.Name, s.GetValue(2, 1)); // formula nuova digitata: #NAME? come Excel
    }
}
