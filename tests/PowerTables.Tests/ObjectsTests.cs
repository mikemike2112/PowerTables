using System.IO.Compression;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using PowerTables.Core;
using PowerTables.Core.IO;

namespace PowerTables.Tests;

public class ObjectsTests
{
    static string Tmp(string n) { string d = Path.Combine(TestUtil.Root, "tests", "work", "objects"); Directory.CreateDirectory(d); return Path.Combine(d, n); }
    static string[] Errors(string path) { using var d = SpreadsheetDocument.Open(path, false); return new OpenXmlValidator().Validate(d).Select(e => e.Path?.XPath + " " + e.Description).ToArray(); }
    static string[] Entries(string path) { using var z = ZipFile.OpenRead(path); return z.Entries.Select(e => e.FullName).ToArray(); }

    // PNG 1x1 valido
    static readonly byte[] Png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==");

    static (Workbook, Sheet) Data()
    {
        var wb = XlsxWriter.CreateNew(); var s = wb.Sheets[0];
        wb.Do("x", () =>
        {
            wb.SetValue(s, 1, 1, "Mese"); wb.SetValue(s, 1, 2, "Vendite"); wb.SetValue(s, 1, 3, "Costi");
            for (int r = 2; r <= 5; r++) { wb.SetValue(s, r, 1, "m" + r); wb.SetValue(s, r, 2, r * 10.0); wb.SetValue(s, r, 3, r * 4.0); }
        });
        return (wb, s);
    }

    [Fact]
    public void PictureShapeTextBoxChartRoundTripValidateAndUndo()
    {
        string p = Tmp("o.xlsx");
        var (wb, s) = Data();
        wb.AddPicture(s, 8, 1, Png, 64, 40);
        wb.AddShape(s, 8, 5, false, "ellipse", "Ciao\nmondo");
        wb.AddShape(s, 14, 1, true, "rect", "Nota");
        var ch = wb.AddChart(s, 1, 1, 5, 3, "col");
        wb.AddChart(s, 1, 1, 5, 3, "line");
        Assert.Equal(2, ch.Chart!.Series.Count); Assert.NotNull(ch.Chart.Series[0].Categories);
        Assert.Equal(["Vendite", "Costi"], ch.Chart.Series.Select(x => x.Name).ToArray());
        XlsxWriter.Save(wb, p);
        Assert.Empty(Errors(p));
        var en = Entries(p);
        Assert.Equal(2, en.Count(e => e.StartsWith("xl/drawings/charts/chart"))); Assert.Contains(en, e => e.StartsWith("xl/media/")); Assert.Contains("xl/drawings/drawing1.xml", en);

        var w2 = XlsxReader.Open(p); var s2 = w2.Sheets[0];
        Assert.Equal(5, s2.Objects.Count);
        Assert.Equal([ObjKind.Picture, ObjKind.Shape, ObjKind.TextBox, ObjKind.Chart, ObjKind.Chart], s2.Objects.Select(o => o.Kind).ToArray());
        Assert.Equal("Ciao\nmondo", s2.Objects[1].Text); Assert.Equal("ellipse", s2.Objects[1].Geom); Assert.Equal(Png, s2.Objects[0].Image);
        // elimina un grafico dopo la riapertura, salva: la parte sparisce e il resto resta
        w2.RemoveObject(s2, s2.Objects[3]);
        XlsxWriter.Save(w2, p);
        Assert.Empty(Errors(p)); Assert.Single(Entries(p), e => e.StartsWith("xl/drawings/charts/chart"));
        Assert.Equal(4, XlsxReader.Open(p).Sheets[0].Objects.Count);
        // annulla dopo il salvataggio: il grafico torna
        w2.Undo(); XlsxWriter.Save(w2, p);
        Assert.Empty(Errors(p)); Assert.Equal(2, Entries(p).Count(e => e.StartsWith("xl/drawings/charts/chart")));
        // annulla di un inserimento
        wb.Undo(); Assert.Equal(4, s.Objects.Count);
    }

    [Fact]
    public void PictureRejectsNonImages()
    {
        var (wb, s) = Data();
        Assert.Throws<ArgumentException>(() => wb.AddPicture(s, 1, 1, "MZ non un'immagine"u8.ToArray(), 10, 10));
        Assert.Throws<ArgumentException>(() => wb.AddPicture(s, 1, 1, [], 10, 10));
        Assert.Throws<ArgumentException>(() => wb.AddShape(s, 1, 1, false, "star", "x"));
        Assert.Throws<InvalidOperationException>(() => { var w = XlsxWriter.CreateNew(); w.AddChart(w.Sheets[0], 1, 1, 1, 1, "col"); });
        Assert.Empty(s.Objects);
    }

    [Fact]
    public void CommentsAddEditDeleteRoundTrip()
    {
        string p = Tmp("c.xlsx");
        var (wb, s) = Data();
        wb.SetComment(s, 2, 2, "prima"); wb.SetComment(s, 3, 3, "seconda", "Anna");
        XlsxWriter.Save(wb, p);
        Assert.Empty(Errors(p)); Assert.Contains(Entries(p), e => e.StartsWith("xl/comments")); Assert.Contains(Entries(p), e => e.EndsWith(".vml"));
        var w2 = XlsxReader.Open(p); var s2 = w2.Sheets[0];
        Assert.Equal("prima", s2.Comments[Addr.Key(2, 2)].Text); Assert.Equal("Anna", s2.Comments[Addr.Key(3, 3)].Author);
        w2.SetComment(s2, 2, 2, "modificata"); w2.SetComment(s2, 3, 3, ""); w2.SetComment(s2, 4, 1, "terza");
        XlsxWriter.Save(w2, p);
        Assert.Empty(Errors(p));
        var w3 = XlsxReader.Open(p); var s3 = w3.Sheets[0];
        Assert.Equal(["modificata", "terza"], s3.Comments.OrderBy(k => k.Key).Select(k => k.Value.Text).ToArray());
        w2.Undo(); w2.Undo(); w2.Undo();   // tutto annullato rispetto al file: le tre modifiche tornano com'erano
        Assert.Equal("prima", s2.Comments[Addr.Key(2, 2)].Text); Assert.True(s2.Comments.ContainsKey(Addr.Key(3, 3))); Assert.False(s2.Comments.ContainsKey(Addr.Key(4, 1)));
        XlsxWriter.Save(w2, p);
        Assert.Empty(Errors(p)); Assert.Equal(2, XlsxReader.Open(p).Sheets[0].Comments.Count);
    }

    [Fact]
    public void ExcelOpensObjectsAndCommentsWithoutRepair()
    {
        var dir = Path.Combine(TestUtil.Root, "tests", "work", "objects", "excel");
        if (Directory.Exists(dir)) Directory.Delete(dir, true);
        Directory.CreateDirectory(dir);
        var (wb, s) = Data();
        wb.AddPicture(s, 8, 1, Png, 64, 40); wb.AddShape(s, 8, 5, false, "ellipse", "Ciao"); wb.AddShape(s, 14, 1, true, "rect", "Nota");
        wb.AddChart(s, 1, 1, 5, 3, "col"); wb.AddChart(s, 1, 1, 5, 3, "line");
        wb.SetComment(s, 2, 2, "prima"); wb.SetComment(s, 3, 3, "seconda");
        XlsxWriter.Save(wb, Path.Combine(dir, "o.xlsx"));
        var (code, outp) = TestUtil.PsT("excel-batch.ps1", 300_000, "-InDir", dir, "-OutDir", Path.Combine(dir, "json"));
        Assert.True(code == 0, outp);
        var j = TestUtil.ReadJson(Path.Combine(dir, "json", "o.xlsx.json")).RootElement;
        Assert.True(j.GetProperty("ok").GetBoolean(), j.GetProperty("error").ToString());
        var sh = j.GetProperty("sheets")[0];
        Assert.Equal(2, sh.GetProperty("charts").GetInt32()); Assert.Equal(2, sh.GetProperty("comments").GetInt32()); Assert.True(sh.GetProperty("shapes").GetInt32() >= 5);
    }
}
