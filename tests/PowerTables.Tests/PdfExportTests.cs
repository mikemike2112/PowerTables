using PdfSharp.Pdf.IO;
using PowerTables.Core;
using PowerTables.Core.IO;

namespace PowerTables.Tests;

public class PdfExportTests
{
    static string Tmp(string n) { string d = Path.Combine(TestUtil.Root, "tests", "work", "pdf"); Directory.CreateDirectory(d); return Path.Combine(d, n); }
    static int Pages(string p) { using var d = PdfReader.Open(p, PdfDocumentOpenMode.Import); return d.PageCount; }

    [Fact]
    public void ExportsValidPdfWithPagesFollowingPrintArea()
    {
        var wb = XlsxWriter.CreateNew(); var s = wb.Sheets[0];
        wb.Do("x", () =>
        {
            for (int r = 1; r <= 120; r++) { wb.SetValue(s, r, 1, "Riga " + r); wb.SetValue(s, r, 2, r * 1.5); }
            wb.SetValue(s, 1, 3, "Prezzo €");
        });
        wb.FormatRange(s, 1, 1, 1, 3, x => x with { Font = x.Font with { Bold = true }, FillColor = "FFFFC7CE", Border = new BorderInfo(new("thin", null), new("thin", null), new("thin", null), new("thin", null)) });
        wb.AddCondFormat(s, 2, 2, 120, 2, new CfSpec("cellIs", "greaterThan", "100", null, null, "FF9C0006", "FFFFC7CE", false));
        string p = Tmp("tutto.pdf");
        int n = PdfExport.Export(wb, s, p, new System.Globalization.CultureInfo("it-IT"));
        Assert.True(n >= 3, "120 righe devono occupare più pagine");
        Assert.Equal(n, Pages(p));
        Assert.Equal("%PDF", System.Text.Encoding.ASCII.GetString(File.ReadAllBytes(p), 0, 4));

        wb.SetPrintArea(s, (1, 1, 10, 3));
        Assert.Equal(1, PdfExport.Export(wb, s, p));
        wb.SetOrientation(s, true); wb.SetPaperSize(s, 8);
        Assert.Equal(1, PdfExport.Export(wb, s, p));
        using (var d = PdfReader.Open(p, PdfDocumentOpenMode.Import)) Assert.True(d.Pages[0].Width.Point > d.Pages[0].Height.Point);   // A3 orizzontale
    }

    [Fact]
    public void EmptySheetAndHugeMarginsAreHandled()
    {
        var wb = XlsxWriter.CreateNew(); var s = wb.Sheets[0];
        string p = Tmp("vuoto.pdf");
        Assert.Equal(1, PdfExport.Export(wb, s, p));
        wb.SetMargins(s, new Margins(5, 5, 5, 5, 0.3, 0.3));
        Assert.Throws<ArgumentException>(() => PdfExport.Export(wb, s, p));
    }

    [Fact]
    public void WideSheetSplitsAcrossColumns()
    {
        var wb = XlsxWriter.CreateNew(); var s = wb.Sheets[0];
        wb.Do("x", () => { for (int c = 1; c <= 30; c++) wb.SetValue(s, 1, c, "Colonna lunga numero " + c); });
        for (int c = 1; c <= 30; c++) wb.SetColWidth(s, c, 20);
        Assert.True(PdfExport.Export(wb, s, Tmp("largo.pdf")) >= 3);
    }
}
