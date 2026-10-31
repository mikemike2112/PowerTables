using PowerTables.Core;
using PowerTables.Core.IO;
using System.Globalization;

namespace PowerTables.Tests;

public class OpsTests
{
    static (Workbook, Sheet) New() { var wb = new Workbook(); return (wb, wb.AddSheet("S")); }

    [Fact]
    public void PasteAdjustsRelativeReferences()
    {
        var (wb, s) = New();
        wb.Do("x", () => { wb.SetValue(s, 1, 1, 2.0); wb.SetValue(s, 2, 1, 3.0); wb.SetFormula(s, 1, 2, "A1*$A$1"); });
        var clip = wb.Copy(s, 1, 2, 1, 2);
        wb.Paste(clip, s, 2, 2);
        Assert.Equal("A2*$A$1", s.Get(2, 2).F!.Text);
        Assert.Equal(6.0, s.GetValue(2, 2));
    }

    [Fact]
    public void AutoFillSeriesAndText()
    {
        var (wb, s) = New();
        wb.Do("x", () => { wb.SetValue(s, 1, 1, 1.0); wb.SetValue(s, 2, 1, 3.0); wb.SetValue(s, 1, 2, "Item9"); wb.SetFormula(s, 1, 3, "A1+1"); });
        wb.AutoFill(s, 1, 1, 2, 1, 5, 1);
        Assert.Equal(9.0, s.GetValue(5, 1));
        wb.AutoFill(s, 1, 2, 1, 2, 3, 2);
        Assert.Equal("Item11", s.GetValue(3, 2));
        wb.AutoFill(s, 1, 3, 1, 3, 3, 3);
        Assert.Equal("A3+1", s.Get(3, 3).F!.Text);
        Assert.Equal(6.0, s.GetValue(3, 3));
    }

    [Fact]
    public void SortKeepsRowsAndAdjustsFormulas()
    {
        var (wb, s) = New();
        wb.Do("x", () =>
        {
            wb.SetValue(s, 1, 1, "n");
            double[] v = [3, 1, 2];
            for (int i = 0; i < 3; i++) { wb.SetValue(s, i + 2, 1, v[i]); wb.SetFormula(s, i + 2, 2, $"A{i + 2}*10"); }
        });
        wb.Sort(s, 1, 1, 4, 2, [(1, true)], true);
        Assert.Equal(1.0, s.GetValue(2, 1)); Assert.Equal(10.0, s.GetValue(2, 2)); Assert.Equal("A2*10", s.Get(2, 2).F!.Text);
        wb.Undo();
        Assert.Equal(3.0, s.GetValue(2, 1)); Assert.Equal(30.0, s.GetValue(2, 2));
    }

    [Fact]
    public void FilterHidesRows()
    {
        var (wb, s) = New();
        wb.Do("x", () => { wb.SetValue(s, 1, 1, "h"); foreach (var (i, t) in new[] { (2, "a"), (3, "b"), (4, "a") }) wb.SetValue(s, i, 1, t); });
        wb.SetAutoFilter(s, 1, 1, 4, 1);
        wb.SetColumnFilter(s, 1, ["a"], CultureInfo.InvariantCulture);
        Assert.False(s.Rows[2].Hidden); Assert.True(s.Rows[3].Hidden);
        Assert.Equal(0, s.RowHeight(3));
        wb.Undo();
        Assert.False(s.Rows[3].Hidden);
    }

    [Fact]
    public void CsvRoundTrip()
    {
        var (wb, s) = New();
        Csv.Import(wb, s, "a,b,c\r\n1,\"x,y\",2024-01-31\r\n=1+1,\"q\"\"r\",3.5\r\n");
        Assert.Equal("x,y", s.GetValue(2, 2));
        Assert.Equal(45322.0, s.GetValue(2, 3));
        Assert.Equal("=1+1", s.GetValue(3, 1)); // niente formule da CSV
        Assert.Equal(3.5, s.GetValue(3, 3));
        var p = Path.GetTempFileName();
        Csv.Export(wb, s, p);
        var t = File.ReadAllText(p);
        Assert.Contains("\"x,y\"", t); Assert.Contains("2024-01-31", t);
        File.Delete(p);
    }

    [Fact]
    public void StyleDeriveAndDisplay()
    {
        var (wb, s) = New();
        wb.Do("x", () => wb.SetValue(s, 1, 1, 1234.5));
        wb.FormatRange(s, 1, 1, 1, 1, x => x with { NumFmtId = 4 });
        Assert.Equal("1,234.50", wb.DisplayText(s, 1, 1, CultureInfo.InvariantCulture));
        wb.Do("x", () => wb.SetInput(s, 2, 1, "15/03/2024", CultureInfo.GetCultureInfo("it-IT")));
        Assert.Equal(45366.0, s.GetValue(2, 1));
        Assert.Equal("15/03/2024", wb.DisplayText(s, 2, 1, CultureInfo.GetCultureInfo("it-IT")));
        wb.Do("x", () => wb.SetInput(s, 3, 1, "12,5%", CultureInfo.GetCultureInfo("it-IT")));
        Assert.Equal(0.125, s.GetValue(3, 1));
    }
}
