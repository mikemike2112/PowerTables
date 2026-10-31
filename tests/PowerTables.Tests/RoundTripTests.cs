using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using PowerTables.Core;
using PowerTables.Core.IO;

namespace PowerTables.Tests;

/// <summary>Esegue una sola volta: salva ogni file del corpus senza/con modifica col clone e fa aprire tutto a Excel vero.</summary>
public sealed class RtFixture
{
    public readonly string Root = TestUtil.Root;
    public string Orig => TestUtil.Corpus("office-made");
    public string NoMod => Path.Combine(Root, "tests", "work", "rt", "nomod");
    public string Mod => Path.Combine(Root, "tests", "work", "rt", "mod");
    public string Neg => Path.Combine(Root, "tests", "work", "rt", "neg");
    public string J(string kind) => Path.Combine(Root, "tests", "work", "rt", "json-" + kind);
    public readonly string[] Files;
    public string? ComError;

    public RtFixture()
    {
        Files = Directory.GetFiles(Orig, "*.xlsx").Select(Path.GetFileName).Order().ToArray()!;
        foreach (var d in new[] { NoMod, Mod, Neg, J("orig"), J("nomod"), J("mod"), J("neg") }) { if (Directory.Exists(d)) Directory.Delete(d, true); Directory.CreateDirectory(d); }
        foreach (var f in Files)
        {
            var wb = XlsxReader.Open(Path.Combine(Orig, f));
            XlsxWriter.Save(wb, Path.Combine(NoMod, f));
            var wb2 = XlsxReader.Open(Path.Combine(Orig, f));
            ApplyEdit(wb2);
            XlsxWriter.Save(wb2, Path.Combine(Mod, f));
        }
        MakeNegatives();
        if (Environment.GetEnvironmentVariable("PT_SKIP_COM") == "1") return;
        foreach (var (dir, kind) in new[] { (Orig, "orig"), (NoMod, "nomod"), (Mod, "mod"), (Neg, "neg") })
        {
            var (code, outp) = TestUtil.PsT("excel-batch.ps1", 900_000, "-InDir", dir, "-OutDir", J(kind));
            if (code != 0) ComError = (ComError ?? "") + $"[{kind}] {outp}\n";
        }
    }

    public static (int Row, int Col) EditCell(Workbook wb) => (Math.Max(1, wb.Sheets[0].MaxRow) + 2, 1);

    public static void ApplyEdit(Workbook wb)
    {
        var s = wb.Sheets[0]; var (r, _) = EditCell(wb);
        wb.Do("edit", () =>
        {
            wb.SetValue(s, r, 1, "PowerTables edit");
            wb.SetValue(s, r, 2, 12345.5);
            wb.SetFormula(s, r, 3, $"B{r}*2");
        });
    }

    static void Rewrite(ZipArchive z, string entry, Func<string, string> f)
    {
        var e = z.GetEntry(entry)!; string t;
        using (var r = new StreamReader(e.Open())) t = r.ReadToEnd();
        e.Delete();
        var ne = z.CreateEntry(entry);
        using var w = new StreamWriter(ne.Open()); w.Write(f(t));
    }

    // file volutamente corrotti: il rilevatore deve fallire su tutti
    void MakeNegatives()
    {
        string src = Path.Combine(Orig, "formula-sum-avg.xlsx");
        void Neg1(string name, Action<ZipArchive> corrupt)
        {
            var p = Path.Combine(Neg, name); File.Copy(src, p, true);
            using var z = ZipFile.Open(p, ZipArchiveMode.Update); corrupt(z);
        }
        Neg1("neg-broken-xml.xlsx", z => Rewrite(z, "xl/worksheets/sheet1.xml", t => t[..(t.Length / 2)]));
        Neg1("neg-missing-part.xlsx", z => z.GetEntry("xl/worksheets/sheet1.xml")!.Delete());
        Neg1("neg-bad-sst-index.xlsx", z => Rewrite(z, "xl/worksheets/sheet1.xml", t => Regex.Replace(t, "<c r=\"A1\"([^>]*)t=\"s\"><v>\\d+</v>", "<c r=\"A1\"$1t=\"s\"><v>99999</v>")));
        Neg1("neg-bad-rel.xlsx", z => Rewrite(z, "xl/_rels/workbook.xml.rels", t => t.Replace("worksheets/sheet1.xml", "worksheets/nonexistent.xml")));
        Neg1("neg-bad-formula.xlsx", z => Rewrite(z, "xl/worksheets/sheet1.xml", t => t.Replace("<f>SUM(C1:C8)</f>", "<f>SUM(C1:C8</f>")));
    }
}

[CollectionDefinition("OfficeCom", DisableParallelization = true)]
public class OfficeComCollection : ICollectionFixture<RtFixture> { }

[Collection("OfficeCom"), Trait("Category", "OfficeCom")]
public class RoundTripTests(RtFixture fx)
{
    public static IEnumerable<object[]> Files() => Directory.GetFiles(TestUtil.Corpus("office-made"), "*.xlsx").Select(f => new object[] { Path.GetFileName(f) }).OrderBy(x => (string)x[0]);

    static List<string> Validate(string path)
    {
        using var doc = SpreadsheetDocument.Open(path, false);
        var v = new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Microsoft365);
        return v.Validate(doc).Select(e => $"{e.Description} @ {e.Path?.XPath}").Order().ToList();
    }

    [Theory, MemberData(nameof(Files))]
    public void Validator_NoNewErrors(string f)
    {
        var before = Validate(Path.Combine(fx.Orig, f)); var after = Validate(Path.Combine(fx.NoMod, f)); var afterMod = Validate(Path.Combine(fx.Mod, f));
        Assert.Empty(after.Except(before));
        Assert.Empty(afterMod.Except(before));
    }

    /// <summary>Forma canonica semplice: elementi, attributi ordinati, testo (senza dichiarazione né ordine degli attributi).</summary>
    static string C14N(Stream s)
    {
        var doc = XDocument.Load(XmlReader.Create(s, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null }), LoadOptions.PreserveWhitespace);
        var sb = new System.Text.StringBuilder();
        void W(XNode n)
        {
            switch (n)
            {
                case XElement e:
                    sb.Append('<').Append(e.Name);
                    foreach (var a in e.Attributes().OrderBy(a => a.Name.ToString(), StringComparer.Ordinal)) if (!a.IsNamespaceDeclaration) sb.Append(' ').Append(a.Name).Append("=\"").Append(a.Value).Append('"');
                    sb.Append('>'); foreach (var c in e.Nodes()) W(c); sb.Append("</").Append(e.Name).Append('>'); break;
                case XText t: sb.Append(t.Value); break;
            }
        }
        W(doc.Root!);
        return sb.ToString();
    }

    [Theory, MemberData(nameof(Files))]
    public void NoEdit_AllPartsIdenticalAndSameRelationships(string f)
    {
        using var a = ZipFile.OpenRead(Path.Combine(fx.Orig, f)); using var b = ZipFile.OpenRead(Path.Combine(fx.NoMod, f));
        var an = a.Entries.Select(e => e.FullName).Order().ToList(); var bn = b.Entries.Select(e => e.FullName).Order().ToList();
        Assert.Equal(an, bn);
        Assert.Equal(ResolvedTypes(a), ResolvedTypes(b));
        foreach (var e in a.Entries)
        {
            var o = b.GetEntry(e.FullName)!;
            if (e.FullName == "[Content_Types].xml") continue;
            if (e.FullName.EndsWith(".rels"))
            {
                // il SDK riscrive i target relativi come assoluti: si confrontano le relazioni risolte
                Assert.True(Rels(e).SequenceEqual(Rels(o)), "relazioni diverse: " + e.FullName);
            }
            else if (e.FullName.EndsWith(".xml") || e.FullName.EndsWith(".vml"))
            {
                string x, y;
                using (var s = e.Open()) x = C14N(s);
                using (var s = o.Open()) y = C14N(s);
                Assert.True(x == y, "parte diversa: " + e.FullName);
            }
            else
            {
                using var s1 = e.Open(); using var s2 = o.Open();
                Assert.Equal(Convert.ToHexString(SHA256.HashData(s1)), Convert.ToHexString(SHA256.HashData(s2)));
            }
        }
    }

    static List<string> Rels(ZipArchiveEntry e)
    {
        XDocument d; using (var s = e.Open()) d = XDocument.Load(s);
        string baseDir = e.FullName.Contains("_rels/") ? e.FullName[..e.FullName.IndexOf("_rels/")] : "";
        return d.Root!.Elements().Select(r =>
        {
            string t = (string)r.Attribute("Target")!; bool ext = (string?)r.Attribute("TargetMode") == "External";
            if (!ext) t = t.StartsWith('/') ? t.TrimStart('/') : new Uri(new Uri("file:///" + baseDir), t).AbsolutePath.TrimStart((char)47);
            return $"{(string)r.Attribute("Id")!}|{(string)r.Attribute("Type")!}|{t}|{ext}";
        }).Order().ToList();
    }

    /// <summary>Tipo di contenuto effettivo di ogni parte (Override, altrimenti Default per estensione): il SDK può riorganizzare Default/Override senza cambiare il significato.</summary>
    static Dictionary<string, string> ResolvedTypes(ZipArchive z)
    {
        XDocument d; using (var s = z.GetEntry("[Content_Types].xml")!.Open()) d = XDocument.Load(s);
        var defs = d.Root!.Elements().Where(e => e.Name.LocalName == "Default").ToDictionary(e => ((string)e.Attribute("Extension")!).ToLowerInvariant(), e => (string)e.Attribute("ContentType")!);
        var ovr = d.Root.Elements().Where(e => e.Name.LocalName == "Override").ToDictionary(e => ((string)e.Attribute("PartName")!).TrimStart('/'), e => (string)e.Attribute("ContentType")!);
        return z.Entries.Where(e => e.FullName != "[Content_Types].xml").ToDictionary(e => e.FullName, e => ovr.TryGetValue(e.FullName, out var t) ? t : defs.GetValueOrDefault(Path.GetExtension(e.FullName).TrimStart('.').ToLowerInvariant(), "?"));
    }

    JsonDocument Read(string kind, string f)
    {
        Assert.True(fx.ComError == null, "Excel COM: " + fx.ComError);
        var p = Path.Combine(fx.J(kind), f + ".json");
        Assert.True(File.Exists(p), "manca il risultato COM per " + f + " (" + kind + ")");
        return TestUtil.ReadJson(p);
    }

    [Theory, MemberData(nameof(Files))]
    public void Original_OpensInExcel(string f)
    {
        var j = Read("orig", f).RootElement;
        Assert.True(j.GetProperty("ok").GetBoolean(), $"Excel non apre l'originale {f}: {j.GetProperty("error")}");
    }

    [Theory, MemberData(nameof(Files))]
    public void Saved_OpensInExcelWithoutRepair_AndSameContent(string f)
    {
        var o = Read("orig", f).RootElement; var s = Read("nomod", f).RootElement;
        Assert.True(s.GetProperty("ok").GetBoolean(), $"Excel segnala errore/ripristino su {f}: {s.GetProperty("error")} repaired={s.GetProperty("repaired")}");
        CompareContent(o, s, f);
    }

    [Theory, MemberData(nameof(Files))]
    public void Edited_OpensInExcel_EditPresent_RestIntact(string f)
    {
        var o = Read("orig", f).RootElement; var s = Read("mod", f).RootElement;
        Assert.True(s.GetProperty("ok").GetBoolean(), $"Excel segnala errore/ripristino su {f} dopo modifica: {s.GetProperty("error")}");
        var wb = XlsxReader.Open(Path.Combine(fx.Orig, f)); var (r, _) = RtFixture.EditCell(wb);
        CompareContent(o, s, f, true);
        var cells = s.GetProperty("sheets")[0].GetProperty("cells").EnumerateArray().Where(c => c.GetProperty("r").GetInt32() == r).ToList();
        Assert.Contains(cells, c => c.GetProperty("c").GetInt32() == 1 && c.GetProperty("v").GetString() == "PowerTables edit");
        Assert.Contains(cells, c => c.GetProperty("c").GetInt32() == 2 && c.GetProperty("v").GetDouble() == 12345.5);
        Assert.Contains(cells, c => c.GetProperty("c").GetInt32() == 3 && c.GetProperty("f").GetString() == $"=B{r}*2" && c.GetProperty("v").GetDouble() == 24691);
    }

    static bool Volatile(JsonElement c) => c.GetProperty("f").ValueKind == JsonValueKind.String && Regex.IsMatch(c.GetProperty("f").GetString()!, @"TODAY\(|NOW\(|RAND\(");

    static void CompareContent(JsonElement o, JsonElement s, string f, bool edited = false)
    {
        var os = o.GetProperty("sheets"); var ss = s.GetProperty("sheets");
        Assert.True(os.GetArrayLength() == ss.GetArrayLength(), $"{f}: numero fogli diverso");
        Assert.Equal(o.GetProperty("date1904").GetBoolean(), s.GetProperty("date1904").GetBoolean());
        Assert.Equal(o.GetProperty("names").GetArrayLength(), s.GetProperty("names").GetArrayLength());
        for (int i = 0; i < os.GetArrayLength(); i++)
        {
            var a = os[i]; var b = ss[i];
            Assert.Equal(a.GetProperty("name").GetString(), b.GetProperty("name").GetString());
            Assert.Equal(a.GetProperty("visible").GetInt32(), b.GetProperty("visible").GetInt32());
            foreach (var k in new[] { "shapes", "charts", "tables", "comments", "hyperlinks", "cf" }) Assert.True(a.GetProperty(k).GetInt32() == b.GetProperty(k).GetInt32(), $"{f}/{a.GetProperty("name")}: {k} {a.GetProperty(k)} -> {b.GetProperty(k)}");
            Assert.Equal(a.GetProperty("filter").GetBoolean(), b.GetProperty("filter").GetBoolean());
            Assert.Equal(a.GetProperty("protectedSheet").GetBoolean(), b.GetProperty("protectedSheet").GetBoolean());
            // la modifica può estendere l'area usata: si confrontano le posizioni già presenti nell'originale
            bool anyWrap = a.GetProperty("cells").EnumerateArray().Any(c => c.TryGetProperty("wrap", out var w) && w.ValueKind == JsonValueKind.True);
            foreach (var k in new[] { "colWidths", "rowHeights" })
            {
                if (k == "rowHeights" && anyWrap && edited) continue; // dopo la rimozione di calcChain Excel rimisura le righe con testo a capo non personalizzate
                var x = a.GetProperty(k).EnumerateArray().Select(v => v.GetDouble()).ToList(); var y = b.GetProperty(k).EnumerateArray().Select(v => v.GetDouble()).ToList();
                Assert.True(y.Count >= x.Count && x.SequenceEqual(y.Take(x.Count)), $"{f}/{a.GetProperty("name")}: {k} {string.Join(",", x)} -> {string.Join(",", y)}");
            }
            var bc = b.GetProperty("cells").EnumerateArray().ToDictionary(c => (c.GetProperty("r").GetInt32(), c.GetProperty("c").GetInt32()));
            foreach (var c in a.GetProperty("cells").EnumerateArray())
            {
                var key = (c.GetProperty("r").GetInt32(), c.GetProperty("c").GetInt32());
                Assert.True(bc.TryGetValue(key, out var d), $"{f}/{a.GetProperty("name")}: cella {key} persa");
                if (Volatile(c)) continue;
                foreach (var p in new[] { "t", "v", "e", "f", "nf", "b", "i", "fc", "fill", "ha", "wrap", "merge" })
                {
                    string x = c.TryGetProperty(p, out var xe) ? xe.ToString() : "", y = d.TryGetProperty(p, out var ye) ? ye.ToString() : "";
                    Assert.True(x == y, $"{f}/{a.GetProperty("name")} cella {key} campo {p}: '{x}' -> '{y}'");
                }
            }
        }
    }

    [Theory]
    [InlineData("neg-broken-xml.xlsx")]
    [InlineData("neg-missing-part.xlsx")]
    [InlineData("neg-bad-sst-index.xlsx")]
    [InlineData("neg-bad-rel.xlsx")]
    [InlineData("neg-bad-formula.xlsx")]
    public void NegativeControl_DetectorFailsOnCorruptFile(string f)
    {
        var j = Read("neg", f).RootElement;
        Assert.False(j.GetProperty("ok").GetBoolean(), $"il rilevatore NON ha colto il file corrotto {f}");
    }

    [Theory, MemberData(nameof(Files))]
    public void CorpusMetadata_AuthorIsCorpus(string f)
    {
        using var z = ZipFile.OpenRead(Path.Combine(fx.Orig, f));
        string core; using (var r = new StreamReader(z.GetEntry("docProps/core.xml")!.Open())) core = r.ReadToEnd();
        Assert.Contains("<dc:creator>Corpus</dc:creator>", core); Assert.Contains("<cp:lastModifiedBy>Corpus</cp:lastModifiedBy>", core);
    }
}
