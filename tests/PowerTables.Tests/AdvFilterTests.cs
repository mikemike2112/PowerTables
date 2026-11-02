using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using PowerTables.Core;
using PowerTables.Core.IO;

namespace PowerTables.Tests;

public class AdvFilterTests
{
    static readonly CultureInfo It = CultureInfo.GetCultureInfo("it-IT");

    static (Workbook, Sheet) Data()
    {
        var wb = XlsxWriter.CreateNew(); var s = wb.Sheets[0];
        wb.Do("x", () =>
        {
            string[] n = ["Anna", "Bruno", "Anna", "Carlo", "Alberto"]; double[] v = [10, 20, 10, 30, 40];
            wb.SetValue(s, 1, 1, "Nome"); wb.SetValue(s, 1, 2, "Valore");
            for (int i = 0; i < n.Length; i++) { wb.SetValue(s, i + 2, 1, n[i]); wb.SetValue(s, i + 2, 2, v[i]); }
            // criteri in E1:F3: (Nome inizia con A E Valore>15) O (Nome = Carlo)
            wb.SetValue(s, 1, 5, "Nome"); wb.SetValue(s, 1, 6, "Valore");
            wb.SetValue(s, 2, 5, "A"); wb.SetValue(s, 2, 6, ">15"); wb.SetValue(s, 3, 5, "=Carlo");
        });
        return (wb, s);
    }

    [Fact]
    public void CriteriaAndOrWildcardsUniqueAndErrors()
    {
        var (wb, s) = Data();
        var list = (1, 1, 6, 2); var crit = (1, 5, 3, 6);
        Assert.Equal([5, 6], wb.AdvancedFilterRows(s, list, crit, false, It).ToArray());   // Carlo (=Carlo) e Alberto (inizia con A e >15)
        var (wb2, s2) = Data(); wb2.Do("x", () => wb2.SetValue(s2, 2, 6, ">5"));
        Assert.Equal([2, 4, 5, 6], wb2.AdvancedFilterRows(s2, list, crit, false, It).ToArray());
        Assert.Equal([2, 5, 6], wb2.AdvancedFilterRows(s2, list, crit, true, It).ToArray());      // i due Anna uguali: ne resta uno
        Assert.Throws<ArgumentException>(() => wb.AdvancedFilterRows(s, list, (1, 5, 1, 6), false, It));        // senza righe criteri
        wb.Do("x", () => wb.SetValue(s, 1, 7, "Inesistente"));
        Assert.Throws<ArgumentException>(() => wb.AdvancedFilterRows(s, list, (1, 5, 3, 7), false, It));
    }

    [Fact]
    public void InPlaceIsOneUndoAndSurvivesSave_CopyDoesNotOverlap()
    {
        string d = Path.Combine(TestUtil.Root, "tests", "work", "adv"); Directory.CreateDirectory(d); string p = Path.Combine(d, "a.xlsx");
        var (wb, s) = Data();
        int n = wb.AdvancedFilterInPlace(s, (1, 1, 6, 2), (1, 5, 3, 6), false, It);
        Assert.Equal(2, n);
        Assert.True(s.Rows[2].Hidden && s.Rows[3].Hidden && s.Rows[4].Hidden);
        Assert.False(s.Rows.TryGetValue(5, out var r5) && r5.Hidden); Assert.False(s.Rows.TryGetValue(6, out var r6) && r6.Hidden);
        XlsxWriter.Save(wb, p);
        using (var doc = SpreadsheetDocument.Open(p, false)) Assert.Empty(new OpenXmlValidator().Validate(doc));
        var w2 = XlsxReader.Open(p);
        Assert.True(w2.Sheets[0].Rows[2].Hidden);
        wb.Undo(); Assert.False(s.Rows.TryGetValue(2, out var r2) && r2.Hidden);
        wb.Redo(); wb.ShowAllRows(s, 1, 6); Assert.False(s.Rows[2].Hidden); wb.Undo(); Assert.True(s.Rows[2].Hidden);
        // copia: sovrapposta all'elenco → errore; altrove → ok
        var (wb3, s3) = Data();
        Assert.Throws<ArgumentException>(() => wb3.AdvancedFilterCopy(s3, (1, 1, 6, 2), (1, 5, 3, 6), false, (3, 1), It));
        int c = wb3.AdvancedFilterCopy(s3, (1, 1, 6, 2), (1, 5, 3, 6), false, (10, 1), It);
        Assert.Equal(2, c); Assert.Equal("Nome", s3.GetValue(10, 1)); Assert.Equal("Alberto", s3.GetValue(12, 1)); Assert.Equal(40.0, s3.GetValue(12, 2));
        wb3.Undo(); Assert.Null(s3.GetValue(10, 1));
    }
}
