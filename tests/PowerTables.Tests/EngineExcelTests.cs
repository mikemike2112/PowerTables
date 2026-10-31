using System.Globalization;
using System.Text.Json;
using PowerTables.Core;
using PowerTables.Core.IO;

namespace PowerTables.Tests;

/// <summary>Confronta il motore formule con i risultati calcolati da Excel vero (letti via COM con Formula/Value2).</summary>
public class EngineExcelTests
{
    static string CasesTxt => Path.Combine(TestUtil.Root, "tests", "PowerTables.Tests", "engine-cases.txt");
    static string Xlsx => TestUtil.Corpus("engine", "engine-cases.xlsx");
    static string Expected => TestUtil.Corpus("engine", "engine-expected.json");

    static string[] Cases() => File.ReadAllLines(CasesTxt).Where(l => l.Length > 0 && !l.StartsWith('#')).ToArray();

    static Workbook BuildWorkbook()
    {
        var wb = XlsxWriter.CreateNew();
        var d = wb.Sheets[0]; d.Name = "Data";
        var t = wb.AddSheet("T");
        wb.Do("fixture", () =>
        {
            double[] a = [5, 3, 8, 1, 9, 2, 7, 4, 6, 10];
            string[] b = ["apple", "Banana", "cherry", "date", "Elder", "fig", "grape", "apple", "Honey", "kiwi"];
            for (int i = 0; i < 10; i++)
            {
                wb.SetValue(d, i + 1, 1, a[i]); wb.SetValue(d, i + 1, 2, b[i]);
                wb.SetValue(d, i + 1, 3, i % 2 == 0 ? "x" : "y"); wb.SetValue(d, i + 1, 4, 44927.0 + i * 31);
                wb.SetValue(d, i + 1, 6, (i + 1) * 10.0);
            }
            wb.SetValue(d, 1, 5, 1.0); wb.SetValue(d, 2, 5, "a"); wb.SetValue(d, 3, 5, true); wb.SetValue(d, 5, 5, CellError.NA);
            wb.SetValue(d, 6, 5, 2.5); wb.SetValue(d, 7, 5, "3"); wb.SetValue(d, 8, 5, 0.0);
            for (int i = 0; i < 5; i++) { wb.SetValue(d, i + 1, 7, ((char)('a' + i)).ToString()); wb.SetValue(d, i + 1, 8, (i + 1) * 100.0); }
            var cases = Cases();
            for (int i = 0; i < cases.Length; i++) wb.SetFormula(t, i + 1, 1, cases[i]);
        });
        return wb;
    }

    [Fact, Trait("Category", "OfficeCom")]
    public void GenerateExpectedFromExcel()
    {
        Directory.CreateDirectory(Path.GetDirectoryName(Xlsx)!);
        XlsxWriter.Save(BuildWorkbook(), Xlsx);
        var (code, outp) = TestUtil.Ps("excel-dump.ps1", "-Path", Xlsx, "-Out", Expected, "-Recalc");
        Assert.True(code == 0, outp);
    }

    [Fact]
    public void EngineMatchesExcel()
    {
        var cases = Cases();
        using var exp = TestUtil.ReadJson(Expected);
        var tsheet = exp.RootElement.GetProperty("sheets").EnumerateArray().First(s => s.GetProperty("name").GetString() == "T");
        var byRow = tsheet.GetProperty("cells").EnumerateArray().ToDictionary(c => c.GetProperty("r").GetInt32());
        var wb = XlsxReader.Open(Xlsx);
        wb.RecalcAll();
        var t = wb.FindSheet("T")!;
        var bad = new List<string>(); int ok = 0, localeDiff = 0;
        for (int i = 0; i < cases.Length; i++)
        {
            var mine = t.GetValue(i + 1, 1);
            string m = Show(mine);
            if (!byRow.TryGetValue(i + 1, out var e)) { bad.Add($"{cases[i]} => excel vuoto, mio {m}"); continue; }
            string x = ShowExp(e);
            if (m == x) { ok++; continue; }
            if (mine is double dm && e.GetProperty("t").GetString() == "n")
            {
                double de = e.GetProperty("v").GetDouble();
                if (Math.Abs(dm - de) <= 1e-9 * Math.Max(1, Math.Abs(de))) { ok++; continue; }
            }
            if (mine is string sm && e.GetProperty("t").GetString() == "s" && sm == e.GetProperty("v").GetString()!.Replace(',', '.').Replace("VERO", "TRUE").Replace("FALSO", "FALSE")) { localeDiff++; continue; }
            bad.Add($"{cases[i]}  =>  mio {m}  |  Excel {x}");
        }
        File.WriteAllLines(Path.Combine(TestUtil.Root, "tests", "work", "engine-mismatch.txt"), bad);
        Assert.True(bad.Count == 0, $"{bad.Count}/{cases.Length} differenze (ok {ok}, locale {localeDiff}):\n" + string.Join("\n", bad.Take(60)));
    }

    static string Show(object? v) => v switch
    {
        null => "<vuoto>", double d => "n:" + d.ToString("R", CultureInfo.InvariantCulture), string s => "s:" + s, bool b => "b:" + b,
        CellError e => "e:" + Errors.Text(e), _ => v.ToString()!,
    };

    static string ShowExp(JsonElement e)
    {
        switch (e.GetProperty("t").GetString())
        {
            case "n": return "n:" + e.GetProperty("v").GetDouble().ToString("R", CultureInfo.InvariantCulture);
            case "s": return "s:" + e.GetProperty("v").GetString();
            case "b": return "b:" + e.GetProperty("v").GetBoolean();
            default: Errors.TryFromComCode(e.GetProperty("e").GetInt32(), out var ce); return "e:" + Errors.Text(ce);
        }
    }
}
