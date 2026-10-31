using PowerTables.Core;
using PowerTables.Core.IO;
using PowerTables.Core.Formulas;
using System.Globalization;

namespace PowerTables.Tests;

public class EngineBasicTests
{
    static object? Calc(string f)
    {
        var wb = new Workbook(); var s = wb.AddSheet("S");
        wb.Do("t", () => { wb.SetValue(s, 1, 1, 10.0); wb.SetValue(s, 2, 1, 20.0); wb.SetValue(s, 3, 1, "x"); wb.SetFormula(s, 1, 3, f); });
        return s.GetValue(1, 3);
    }

    [Theory]
    [InlineData("1+2*3", 7.0)]
    [InlineData("-2^2", 4.0)]
    [InlineData("SUM(A1:A2)", 30.0)]
    [InlineData("IF(A1>5,\"big\",\"small\")", "big")]
    [InlineData("A1&A3", "10x")]
    [InlineData("VLOOKUP(20,A1:A2,1,FALSE)", 20.0)]
    [InlineData("DATE(1900,2,29)", 60.0)]
    [InlineData("DAY(60)", 29.0)]
    [InlineData("WEEKDAY(1)", 1.0)]
    [InlineData("ROUND(2.675,2)", 2.68)]
    [InlineData("COUNTIF(A1:A3,\">5\")", 2.0)]
    public void Basics(string f, object expected) => Assert.Equal(expected, Calc(f));

    [Fact] public void DivZero() => Assert.Equal(CellError.Div0, Calc("1/0"));

    [Fact]
    public void IncrementalRecalcAndCycle()
    {
        var wb = new Workbook(); var s = wb.AddSheet("S");
        wb.Do("a", () => { wb.SetValue(s, 1, 1, 1.0); wb.SetFormula(s, 1, 2, "A1*2"); wb.SetFormula(s, 1, 3, "B1+1"); });
        Assert.Equal(3.0, s.GetValue(1, 3));
        wb.Do("b", () => wb.SetValue(s, 1, 1, 5.0));
        Assert.Equal(11.0, s.GetValue(1, 3));
        wb.Do("c", () => wb.SetFormula(s, 1, 1, "C1"));
        Assert.NotEmpty(wb.Cycles);
        wb.Undo();
        Assert.Equal(11.0, s.GetValue(1, 3));
    }

    [Fact]
    public void SaveReopen()
    {
        var wb = XlsxWriter.CreateNew(); var s = wb.Sheets[0];
        wb.Do("a", () => { wb.SetValue(s, 1, 1, 7.0); wb.SetValue(s, 2, 1, "ciao"); wb.SetFormula(s, 1, 2, "A1*3"); wb.SetFormula(s, 3, 1, "XLOOKUP(7,A1:A1,A1:A1)"); });
        var p = Path.Combine(Path.GetTempPath(), "pt-" + Guid.NewGuid().ToString("N") + ".xlsx");
        XlsxWriter.Save(wb, p);
        var wb2 = XlsxReader.Open(p);
        var s2 = wb2.Sheets[0];
        Assert.Equal(21.0, s2.GetValue(1, 2));
        Assert.Equal("ciao", s2.GetValue(2, 1));
        Assert.Equal("_xlfn.XLOOKUP(7,A1:A1,A1:A1)", s2.Get(3, 1).F!.Text);
        File.Delete(p);
    }
}

public class Scratch
{
    [Fact]
    public void WriteSample()
    {
        var wb = XlsxWriter.CreateNew(); var s = wb.Sheets[0];
        wb.Do("a", () =>
        {
            wb.SetValue(s, 1, 1, 7.0); wb.SetValue(s, 2, 1, "ciao & <b>"); wb.SetFormula(s, 1, 2, "A1*3");
            wb.SetFormula(s, 3, 1, "XLOOKUP(7,A1:A1,A1:A1)"); wb.SetValue(s, 4, 1, true); wb.SetValue(s, 5, 1, CellError.NA);
            wb.SetFormula(s, 6, 1, "DATE(2024,2,29)"); wb.SetFormula(s, 7, 1, "TEXT(1234.5,\"0.00\")&\"x\"");
        });
        Directory.CreateDirectory(@"C:\dev\powertables\tests\work");
        XlsxWriter.Save(wb, @"C:\dev\powertables\tests\work\first.xlsx");
    }
}
