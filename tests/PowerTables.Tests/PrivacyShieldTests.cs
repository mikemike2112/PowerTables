using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using PowerTables.Core.IO;

namespace PowerTables.Tests;

[Collection("OfficeCom"), Trait("Category", "OfficeCom")]
public class PrivacyShieldTests
{
    static readonly string[] Secrets = ["Mario Rossi", "Anna Bianchi", "ACME Srl", "Luigi Verdi", "Mario.Rossi", "Cliente riservato"];

    static string Work(string name) { var d = Path.Combine(TestUtil.Root, "tests", "work", "privacy"); Directory.CreateDirectory(d); return Path.Combine(d, name); }

    static string Rd(ZipArchive z, string n) { using var r = new StreamReader(z.GetEntry(n)!.Open()); return r.ReadToEnd(); }
    static void Wr(ZipArchive z, string n, string t) { z.GetEntry(n)?.Delete(); using var w = new StreamWriter(z.CreateEntry(n).Open(), new UTF8Encoding(false)); w.Write(t); }

    /// <summary>File con dati personali NOTI (finti): parte da un file di Excel vero del corpus e ci inietta autori, società, proprietà, miniatura e percorso locale.</summary>
    static string MakeDirty()
    {
        string p = Work("sporco.xlsx"); File.Copy(TestUtil.Corpus("office-made", "comments-notes.xlsx"), p, true);
        using var z = ZipFile.Open(p, ZipArchiveMode.Update);
        Wr(z, "docProps/core.xml", Regex.Replace(Regex.Replace(Rd(z, "docProps/core.xml"), "<dc:creator>.*?</dc:creator>", "<dc:creator>Mario Rossi</dc:creator>"), "<cp:lastModifiedBy>.*?</cp:lastModifiedBy>", "<cp:lastModifiedBy>Anna Bianchi</cp:lastModifiedBy>"));
        string app = Rd(z, "docProps/app.xml");
        app = app.Contains("<Company>") ? Regex.Replace(app, "<Company>.*?</Company>", "<Company>ACME Srl</Company>") : app.Replace("</Properties>", "<Company>ACME Srl</Company></Properties>");
        app = app.Replace("</Properties>", "<Manager>Luigi Verdi</Manager></Properties>");
        Wr(z, "docProps/app.xml", app);
        Wr(z, "xl/comments1.xml", Rd(z, "xl/comments1.xml").Replace("<author>Corpus</author>", "<author>Mario Rossi</author>"));
        Wr(z, "docProps/custom.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"><property fmtid=\"{D5CDD505-2E9C-101B-9397-08002B2CF9AE}\" pid=\"2\" name=\"Cliente\"><vt:lpwstr>Cliente riservato</vt:lpwstr></property></Properties>");
        using (var w = z.CreateEntry("docProps/thumbnail.jpeg").Open()) w.Write(new byte[] { 0xFF, 0xD8, 0xFF, 0xD9 });
        string ct = Rd(z, "[Content_Types].xml");
        ct = ct.Replace("</Types>", "<Override PartName=\"/docProps/custom.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.custom-properties+xml\"/></Types>");
        if (!ct.Contains("Extension=\"jpeg\"")) { int i = ct.IndexOf("<Default ", StringComparison.Ordinal); ct = ct.Insert(i, "<Default Extension=\"jpeg\" ContentType=\"image/jpeg\"/>"); }
        Wr(z, "[Content_Types].xml", ct);
        Wr(z, "_rels/.rels", Rd(z, "_rels/.rels").Replace("</Relationships>", "<Relationship Id=\"rId90\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties\" Target=\"docProps/custom.xml\"/><Relationship Id=\"rId91\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail\" Target=\"docProps/thumbnail.jpeg\"/></Relationships>"));
        string wb = Rd(z, "xl/workbook.xml");
        wb = Regex.Replace(wb, "x15ac:absPath url=\"[^\"]*\"", @"x15ac:absPath url=""C:\Users\Mario.Rossi\Documents\""");   // Excel scrive già questo elemento: si cambia solo il percorso
        Wr(z, "xl/workbook.xml", wb);
        return p;
    }

    static string Hash(string p) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(p)));

    static string AllText(string p)
    {
        var sb = new StringBuilder(); using var z = ZipFile.OpenRead(p);
        foreach (var e in z.Entries) { using var s = e.Open(); var ms = new MemoryStream(); s.CopyTo(ms); sb.AppendLine(e.FullName).AppendLine(Encoding.UTF8.GetString(ms.ToArray())); }
        return sb.ToString();
    }

    [Fact]
    public void Analyze_FindsKnownData_And_CleanCopyContainsNone_OriginalUntouched_ExcelOpensIt()
    {
        string src = MakeDirty(); string before = Hash(src);
        var rep = PrivacyShield.Analyze(src);
        foreach (var s in new[] { "Mario Rossi", "Anna Bianchi", "ACME Srl", "Luigi Verdi", "Cliente riservato" }) Assert.Contains(rep.Findings, f => f.Value.Contains(s));
        Assert.Contains(rep.Findings, f => f.Kind == PrivacyShield.Kind.LocalPath && f.Value.Contains("Mario.Rossi"));
        Assert.Contains(rep.Findings, f => f.Kind == PrivacyShield.Kind.Thumbnail);
        Assert.Contains(rep.Findings, f => f.Kind == PrivacyShield.Kind.CommentAuthor);
        Assert.NotEmpty(rep.Warnings);                                    // riserve sempre esplicite
        Assert.Equal(before, Hash(src));                                 // l'analisi non tocca il file

        string dest = Work("sporco (pulito).xlsx");
        var after = PrivacyShield.CreateCleanCopy(src, dest, rep.Findings, PrivacyShield.CommentsMode.Anonymize);
        Assert.Equal(before, Hash(src));                                 // originale identico byte per byte
        string text = AllText(dest);
        foreach (var s in Secrets) Assert.DoesNotContain(s, text);       // nessuna parte contiene più i dati
        Assert.DoesNotContain("thumbnail.jpeg", text);
        Assert.True(after.Findings.Count == 0, string.Join("; ", after.Findings.Select(f => f.Kind + "=" + f.Value + "@" + f.Part)));
        Assert.Contains("Questa è una nota di prova", text);             // contenuto del commento invariato
        Assert.NotEmpty(after.Warnings);                                 // le riserve restano

        string dir = Work("excel-copia"); if (Directory.Exists(dir)) Directory.Delete(dir, true); Directory.CreateDirectory(dir);
        File.Copy(src, Path.Combine(dir, "a-orig.xlsx")); File.Copy(dest, Path.Combine(dir, "b-pulito.xlsx"));
        var (code, outp) = TestUtil.PsT("excel-batch.ps1", 300_000, "-InDir", dir, "-OutDir", Path.Combine(dir, "json"));
        Assert.True(code == 0, outp);
        var o = TestUtil.ReadJson(Path.Combine(dir, "json", "a-orig.xlsx.json")).RootElement; var c = TestUtil.ReadJson(Path.Combine(dir, "json", "b-pulito.xlsx.json")).RootElement;
        Assert.True(o.GetProperty("ok").GetBoolean(), o.ToString()); Assert.True(c.GetProperty("ok").GetBoolean(), "Excel segnala errore/ripristino sulla copia");
        Assert.Equal(o.GetProperty("sheets")[0].GetProperty("cells").ToString(), c.GetProperty("sheets")[0].GetProperty("cells").ToString());
        Assert.Equal(1, c.GetProperty("sheets")[0].GetProperty("comments").GetInt32());
    }

    [Fact]
    public void RemoveCommentsMode_DropsThem_AndExcelOpensCopy()
    {
        string src = MakeDirty(); string before = Hash(src);
        var rep = PrivacyShield.Analyze(src);
        string dest = Work("sporco (pulito, senza commenti).xlsx");
        var after = PrivacyShield.CreateCleanCopy(src, dest, rep.Findings, PrivacyShield.CommentsMode.Remove);
        Assert.Equal(before, Hash(src));
        string text = AllText(dest);
        foreach (var s in Secrets) Assert.DoesNotContain(s, text);
        Assert.DoesNotContain("comments1.xml", text); Assert.DoesNotContain("legacyDrawing", text);
        Assert.Empty(after.Findings);
        string dir = Work("excel-senza"); if (Directory.Exists(dir)) Directory.Delete(dir, true); Directory.CreateDirectory(dir);
        File.Copy(dest, Path.Combine(dir, "senza.xlsx"));
        var (code, outp) = TestUtil.PsT("excel-batch.ps1", 300_000, "-InDir", dir, "-OutDir", Path.Combine(dir, "json"));
        Assert.True(code == 0, outp);
        var c = TestUtil.ReadJson(Path.Combine(dir, "json", "senza.xlsx.json")).RootElement;
        Assert.True(c.GetProperty("ok").GetBoolean(), "Excel segnala errore/ripristino");
        Assert.Equal(0, c.GetProperty("sheets")[0].GetProperty("comments").GetInt32());
    }

    [Fact]
    public void AlreadyCleanFile_FindsNothing_AndRefusesSamePath()
    {
        string p = Work("pulito.xlsx"); File.Copy(TestUtil.Corpus("office-made", "formula-sum-avg.xlsx"), p, true);
        using (var z = ZipFile.Open(p, ZipArchiveMode.Update))
        {
            string t = Rd(z, "docProps/core.xml");
            t = Regex.Replace(t, "<dc:creator>.*?</dc:creator>", "<dc:creator></dc:creator>"); t = Regex.Replace(t, "<cp:lastModifiedBy>.*?</cp:lastModifiedBy>", "<cp:lastModifiedBy></cp:lastModifiedBy>");
            Wr(z, "docProps/core.xml", t);
        }
        var rep = PrivacyShield.Analyze(p);
        Assert.Empty(rep.Findings);
        Assert.Throws<InvalidOperationException>(() => PrivacyShield.CreateCleanCopy(p, p, rep.Findings, PrivacyShield.CommentsMode.Anonymize));
    }
}
