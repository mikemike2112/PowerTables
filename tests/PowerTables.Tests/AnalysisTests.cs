using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using PowerTables.Core;
using PowerTables.Core.IO;

namespace PowerTables.Tests;

public class AnalysisTests
{
    [Fact]
    public void GoalSeekFindsLinearAndQuadraticRootsUndoesAndRejects()
    {
        var wb = XlsxWriter.CreateNew(); var s = wb.Sheets[0];
        wb.Do("x", () => { wb.SetValue(s, 1, 1, 2.0); wb.SetFormula(s, 2, 1, "A1*3+1"); wb.SetFormula(s, 3, 1, "A1*A1"); wb.SetValue(s, 4, 1, "t"); });
        var r = wb.GoalSeek(s, 2, 1, 10, 1, 1);
        Assert.NotNull(r); Assert.Equal(3.0, r!.Value, 5); Assert.Equal(10.0, (double)s.GetValue(2, 1)!, 5);
        wb.Undo(); Assert.Equal(2.0, s.GetValue(1, 1));
        var q = wb.GoalSeek(s, 3, 1, 49, 1, 1);
        Assert.NotNull(q); Assert.Equal(49.0, (double)s.GetValue(3, 1)!, 4);
        wb.Undo();
        Assert.Null(wb.GoalSeek(s, 3, 1, -5, 1, 1));            // x^2 = -5: nessuna soluzione, cella invariata
        Assert.Equal(2.0, s.GetValue(1, 1));
        Assert.Throws<ArgumentException>(() => wb.GoalSeek(s, 1, 1, 5, 1, 2));    // la cella da impostare non è una formula
        Assert.Throws<ArgumentException>(() => wb.GoalSeek(s, 2, 1, 5, 2, 1));    // cella da cambiare con formula
        Assert.Throws<ArgumentException>(() => wb.GoalSeek(s, 2, 1, 5, 4, 1));    // cella da cambiare con testo
    }

    [Fact]
    public void ForecastSheetIsLinearHasChartValidatesAndUndoes()
    {
        string d = Path.Combine(TestUtil.Root, "tests", "work", "forecast"); Directory.CreateDirectory(d); string p = Path.Combine(d, "f.xlsx");
        var wb = XlsxWriter.CreateNew(); var s = wb.Sheets[0];
        wb.Do("x", () => { for (int i = 0; i < 6; i++) { wb.SetValue(s, i + 1, 1, (double)(i + 1)); wb.SetValue(s, i + 1, 2, 10.0 + 2 * (i + 1)); } });
        var f = wb.CreateForecastSheet(s, (1, 2, 6, 2), (1, 1, 6, 1), 3);
        Assert.Equal("Previsione", f.Name);
        Assert.Equal(7.0, f.GetValue(8, 1)); Assert.Equal(10.0 + 2 * 7, (double)f.GetValue(8, 3)!, 6);
        Assert.Equal((double)f.GetValue(8, 3)!, (double)f.GetValue(8, 4)!, 6);     // dati perfettamente lineari: limiti = previsione
        Assert.Single(f.Objects);
        XlsxWriter.Save(wb, p);
        using (var doc = SpreadsheetDocument.Open(p, false)) Assert.Empty(new OpenXmlValidator().Validate(doc));
        Assert.Equal(2, XlsxReader.Open(p).Sheets.Count);
        wb.Undo(); Assert.Single(wb.Sheets);
        Assert.Throws<ArgumentException>(() => wb.CreateForecastSheet(s, (1, 2, 2, 2), (1, 1, 2, 1), 3));
        wb.Do("x", () => wb.SetValue(s, 3, 1, 0.5));
        Assert.Throws<ArgumentException>(() => wb.CreateForecastSheet(s, (1, 2, 6, 2), (1, 1, 6, 1), 3));   // cronologia non crescente
    }
}
