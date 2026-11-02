using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using PowerTables.Core;
using PowerTables.Core.IO;

namespace PowerTables.Tests;

public class ViewCoreTests
{
    [Fact]
    public void PagesSplitByPaperAndHeadingsGridlinesPersist()
    {
        string d = Path.Combine(TestUtil.Root, "tests", "work", "view"); Directory.CreateDirectory(d); string p = Path.Combine(d, "v.xlsx");
        var wb = XlsxWriter.CreateNew(); var s = wb.Sheets[0];
        wb.Do("x", () => { for (int r = 1; r <= 120; r++) wb.SetValue(s, r, 1, r); wb.SetValue(s, 1, 30, 1.0); });
        var (rows, cols, rg) = PdfExport.Pages(wb, s);
        Assert.True(rows.Count >= 2 && cols.Count >= 2); Assert.Equal((1, 1, 120, 30), rg);
        Assert.Equal(1, rows[0].From); Assert.Equal(rows[0].To + 1, rows[1].From); Assert.Equal(120, rows[^1].To);
        s.ShowHeadings = false; s.ShowGridLines = false; s.LayoutDirty = true;
        XlsxWriter.Save(wb, p);
        using (var doc = SpreadsheetDocument.Open(p, false)) Assert.Empty(new OpenXmlValidator().Validate(doc));
        var w2 = XlsxReader.Open(p);
        Assert.False(w2.Sheets[0].ShowHeadings); Assert.False(w2.Sheets[0].ShowGridLines);
    }
}
