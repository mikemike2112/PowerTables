using System.Globalization;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace PowerTables.Core.IO;

/// <summary>Immagini, forme, caselle di testo, grafici e commenti creati da PowerTables: lettura e scrittura. Gli oggetti del file non creati qui restano intatti.</summary>
static class XlsxObjects
{
    static readonly XNamespace Xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing", A = "http://schemas.openxmlformats.org/drawingml/2006/main",
        C = "http://schemas.openxmlformats.org/drawingml/2006/chart", R = XlsxReader.RelNs, V = "urn:schemas-microsoft-com:vml", O = "urn:schemas-microsoft-com:office:office", X = "urn:schemas-microsoft-com:office:excel";

    static readonly XNamespace Main = XlsxReader.MainNs;

    // ---------------- lettura ----------------
    public static void Load(Sheet sh, WorksheetPart wsp)
    {
        try { LoadDrawing(sh, wsp); } catch (Exception e) when (e is System.Xml.XmlException or JsonException or ArgumentException or InvalidOperationException or IOException or FormatException) { }
        try { LoadComments(sh, wsp); } catch (Exception e) when (e is System.Xml.XmlException or ArgumentException or InvalidOperationException or IOException or FormatException) { }
    }

    static void LoadDrawing(Sheet sh, WorksheetPart wsp)
    {
        var dp = wsp.DrawingsPart; if (dp == null) return;
        XDocument doc; using (var s = dp.GetStream(FileMode.Open, FileAccess.Read)) doc = XDocument.Load(s);
        foreach (var an in doc.Root!.Elements(Xdr + "twoCellAnchor"))
        {
            var nv = an.Descendants(Xdr + "cNvPr").FirstOrDefault(); string? name = (string?)nv?.Attribute("name");
            if (name == null || !name.StartsWith(SheetObject.Prefix)) continue;
            int P(XElement? e, string n) => (int?)e?.Element(Xdr + n) ?? 0;
            var f = an.Element(Xdr + "from"); var t = an.Element(Xdr + "to");
            var o = new SheetObject { Name = name, R1 = P(f, "row") + 1, C1 = P(f, "col") + 1, R2 = Math.Max(P(f, "row") + 1, P(t, "row")), C2 = Math.Max(P(f, "col") + 1, P(t, "col")), Written = true };
            string descr = (string?)nv!.Attribute("descr") ?? "";
            if (an.Element(Xdr + "pic") is { } pic)
            {
                o.Kind = ObjKind.Picture;
                var ip = dp.GetPartById((string)pic.Descendants(A + "blip").First().Attribute(R + "embed")!) as ImagePart ?? throw new InvalidOperationException();
                using var ms = new MemoryStream(); using (var s = ip.GetStream(FileMode.Open, FileAccess.Read)) s.CopyTo(ms);
                o.Image = ms.ToArray(); o.ImageExt = ip.ContentType.Contains("png") ? "png" : ip.ContentType.Contains("gif") ? "gif" : "jpeg";
            }
            else if (an.Element(Xdr + "graphicFrame") != null && descr.StartsWith("PTCHART:"))
            {
                o.Kind = ObjKind.Chart; o.Chart = JsonSerializer.Deserialize<ChartSpec>(descr[8..]);
                if (o.Chart == null) continue;
            }
            else if (an.Element(Xdr + "sp") is { } sp && descr.StartsWith("PTTEXT:"))
            {
                var ti = JsonSerializer.Deserialize<SheetObject.TextInfo>(descr[7..]);
                o.Kind = (string?)sp.Element(Xdr + "nvSpPr")?.Element(Xdr + "cNvSpPr")?.Attribute("txBox") == "1" ? ObjKind.TextBox : ObjKind.Shape;
                o.Text = ti?.Text; o.Geom = ti?.Geom ?? "rect";
            }
            else continue;
            sh.Objects.Add(o);
        }
    }

    static void LoadComments(Sheet sh, WorksheetPart wsp)
    {
        var cp = wsp.WorksheetCommentsPart; if (cp == null) return;
        XDocument doc; using (var s = cp.GetStream(FileMode.Open, FileAccess.Read)) doc = XDocument.Load(s);
        var m = doc.Root!.Name.Namespace;
        var authors = doc.Root.Element(m + "authors")?.Elements(m + "author").Select(a => a.Value).ToList() ?? [];
        foreach (var c in doc.Root.Element(m + "commentList")?.Elements(m + "comment") ?? [])
        {
            if (!Addr.TryParse((string?)c.Attribute("ref") ?? "", out var r, out var col)) continue;
            var text = string.Concat(c.Element(m + "text")?.Descendants(m + "t").Where(t => t.Parent?.Name.LocalName != "rPh").Select(t => t.Value) ?? []);
            int ai = (int?)c.Attribute("authorId") ?? 0;
            sh.Comments[Addr.Key(r, col)] = new CellComment { Author = ai >= 0 && ai < authors.Count ? authors[ai] : "", Text = text, Orig = text };
        }
    }

    // ---------------- scrittura ----------------
    static XElement Cell(string tag, int row0, int col0) => new(Xdr + tag, new XElement(Xdr + "col", col0), new XElement(Xdr + "colOff", 0), new XElement(Xdr + "row", row0), new XElement(Xdr + "rowOff", 0));

    public static void WriteObjects(Sheet sh, XElement root, XNamespace ns, WorksheetPart wsp)
    {
        var dEl = root.Element(ns + "drawing"); DrawingsPart? dp = null;
        if ((string?)dEl?.Attribute(R + "id") is { } rid) { try { dp = wsp.GetPartById(rid) as DrawingsPart; } catch (ArgumentOutOfRangeException) { } }
        if (dp == null && sh.Objects.Count == 0) return;
        XDocument doc;
        if (dp != null) { using var s = dp.GetStream(FileMode.Open, FileAccess.Read); doc = XDocument.Load(s); }
        else doc = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), new XElement(Xdr + "wsDr", new XAttribute(XNamespace.Xmlns + "xdr", Xdr), new XAttribute(XNamespace.Xmlns + "a", A)));
        var names = sh.Objects.Select(o => o.Name).ToHashSet();
        var present = new HashSet<string>(); bool changed = false;
        foreach (var an in doc.Root!.Elements().ToList())
        {
            string? n = (string?)an.Descendants(Xdr + "cNvPr").FirstOrDefault()?.Attribute("name");
            if (n == null || !n.StartsWith(SheetObject.Prefix)) continue;
            if (names.Contains(n)) { present.Add(n); continue; }
            foreach (var rr in an.Descendants().Attributes().Where(a => a.Name == R + "embed" || a.Name == R + "id").Select(a => a.Value).ToList())
                try { dp!.DeletePart(rr); } catch (ArgumentOutOfRangeException) { }
            an.Remove(); changed = true;
        }
        var pending = sh.Objects.Where(o => !present.Contains(o.Name)).ToList();
        if (pending.Count > 0 && dp == null) dp = wsp.AddNewPart<DrawingsPart>();
        int id = doc.Root.Descendants(Xdr + "cNvPr").Select(e => (int?)e.Attribute("id") ?? 0).DefaultIfEmpty(1).Max() + 1;
        foreach (var o in pending)
        {
            var an = new XElement(Xdr + "twoCellAnchor", Cell("from", o.R1 - 1, o.C1 - 1), Cell("to", o.R2, o.C2), Body(o, dp!, id++), new XElement(Xdr + "clientData"));
            doc.Root.Add(an); o.Written = true; changed = true;
        }
        if (dp == null) return;
        if (changed || dEl == null) { using var os = dp.GetStream(FileMode.Create); doc.Save(os); }
        if (dEl == null) { dEl = new XElement(ns + "drawing", new XAttribute(R + "id", wsp.GetIdOfPart(dp))); XlsxWriter.InsertOrdered(root, ns, dEl); }
    }

    static XElement Xfrm() => new(A + "xfrm", new XElement(A + "off", new XAttribute("x", 0), new XAttribute("y", 0)), new XElement(A + "ext", new XAttribute("cx", 0), new XAttribute("cy", 0)));

    static XElement Body(SheetObject o, DrawingsPart dp, int id)
    {
        var cnv = new XElement(Xdr + "cNvPr", new XAttribute("id", id), new XAttribute("name", o.Name), new XAttribute("descr", o.Descr));
        switch (o.Kind)
        {
            case ObjKind.Picture:
            {
                var ip = o.ImageExt switch { "png" => dp.AddImagePart(ImagePartType.Png), "gif" => dp.AddImagePart(ImagePartType.Gif), _ => dp.AddImagePart(ImagePartType.Jpeg) };
                using (var ms = new MemoryStream(o.Image!)) ip.FeedData(ms);
                return new XElement(Xdr + "pic", new XElement(Xdr + "nvPicPr", cnv, new XElement(Xdr + "cNvPicPr", new XElement(A + "picLocks", new XAttribute("noChangeAspect", 1)))),
                    new XElement(Xdr + "blipFill", new XElement(A + "blip", new XAttribute(XNamespace.Xmlns + "r", R), new XAttribute(R + "embed", dp.GetIdOfPart(ip))), new XElement(A + "stretch", new XElement(A + "fillRect"))),
                    new XElement(Xdr + "spPr", Xfrm(), new XElement(A + "prstGeom", new XAttribute("prst", "rect"), new XElement(A + "avLst"))));
            }
            case ObjKind.Chart:
            {
                var cp = dp.AddNewPart<ChartPart>();
                using (var os = cp.GetStream(FileMode.Create)) ChartXml(o.Chart!).Save(os);
                return new XElement(Xdr + "graphicFrame", new XAttribute("macro", ""), new XElement(Xdr + "nvGraphicFramePr", cnv, new XElement(Xdr + "cNvGraphicFramePr")),
                    new XElement(Xdr + "xfrm", new XElement(A + "off", new XAttribute("x", 0), new XAttribute("y", 0)), new XElement(A + "ext", new XAttribute("cx", 0), new XAttribute("cy", 0))),
                    new XElement(A + "graphic", new XElement(A + "graphicData", new XAttribute("uri", C.NamespaceName), new XElement(C + "chart", new XAttribute(XNamespace.Xmlns + "c", C), new XAttribute(XNamespace.Xmlns + "r", R), new XAttribute(R + "id", dp.GetIdOfPart(cp))))));
            }
            default:
            {
                bool tb = o.Kind == ObjKind.TextBox;
                var paras = (o.Text ?? "").Split('\n').Select(l => l.TrimEnd('\r')).Select(l =>
                {
                    var p = new XElement(A + "p", new XElement(A + "pPr", new XAttribute("algn", tb ? "l" : "ctr")));
                    if (l.Length > 0) p.Add(new XElement(A + "r", new XElement(A + "rPr", new XAttribute("lang", "it-IT"), new XAttribute("sz", 1100), new XElement(A + "solidFill", new XElement(A + "srgbClr", new XAttribute("val", tb ? "000000" : "FFFFFF")))), new XElement(A + "t", l)));
                    return p;
                });
                return new XElement(Xdr + "sp", new XAttribute("macro", ""), new XAttribute("textlink", ""),
                    new XElement(Xdr + "nvSpPr", cnv, new XElement(Xdr + "cNvSpPr", tb ? new XAttribute("txBox", 1) : null)),
                    new XElement(Xdr + "spPr", Xfrm(), new XElement(A + "prstGeom", new XAttribute("prst", tb ? "rect" : o.Geom), new XElement(A + "avLst")),
                        new XElement(A + "solidFill", new XElement(A + "srgbClr", new XAttribute("val", tb ? "FFFFFF" : "4472C4"))),
                        new XElement(A + "ln", new XAttribute("w", 9525), new XElement(A + "solidFill", new XElement(A + "srgbClr", new XAttribute("val", tb ? "808080" : "2F528F"))))),
                    new XElement(Xdr + "txBody", new XElement(A + "bodyPr", new XAttribute("vertOverflow", "clip"), new XAttribute("wrap", "square"), new XAttribute("rtlCol", 0), new XAttribute("anchor", tb ? "t" : "ctr")), new XElement(A + "lstStyle"), paras));
            }
        }
    }

    // ---------------- grafico ----------------
    static readonly string[] Palette = ["4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47"];

    /// <summary>Valori dell'intervallo (cache del grafico): letti dal foglio di destinazione. Foglio non trovato → vuoto.</summary>
    internal static Func<string, List<object?>>? RangeReader;

    static XDocument ChartXml(ChartSpec sp)
    {
        bool line = sp.Type == "line";
        var plot = new XElement(C + (line ? "lineChart" : "barChart"));
        if (!line) plot.Add(new XElement(C + "barDir", new XAttribute("val", "col")));
        plot.Add(new XElement(C + "grouping", new XAttribute("val", line ? "standard" : "clustered")), new XElement(C + "varyColors", new XAttribute("val", 0)));
        for (int i = 0; i < sp.Series.Count; i++)
        {
            var se = sp.Series[i]; string col = Palette[i % Palette.Length];
            var ser = new XElement(C + "ser", new XElement(C + "idx", new XAttribute("val", i)), new XElement(C + "order", new XAttribute("val", i)),
                new XElement(C + "tx", new XElement(C + "v", se.Name)));
            ser.Add(line
                ? new XElement(C + "spPr", new XElement(A + "ln", new XAttribute("w", 28575), new XElement(A + "solidFill", new XElement(A + "srgbClr", new XAttribute("val", col)))))
                : new XElement(C + "spPr", new XElement(A + "solidFill", new XElement(A + "srgbClr", new XAttribute("val", col)))));
            if (line) ser.Add(new XElement(C + "marker", new XElement(C + "symbol", new XAttribute("val", "none"))));
            if (se.Categories != null) ser.Add(new XElement(C + "cat", Ref("strRef", se.Categories, "strCache", null, false)));
            ser.Add(new XElement(C + "val", Ref("numRef", se.Values, "numCache", "General", true)));
            if (line) ser.Add(new XElement(C + "smooth", new XAttribute("val", 0)));
            plot.Add(ser);
        }
        if (line) plot.Add(new XElement(C + "marker", new XAttribute("val", 1)));
        else plot.Add(new XElement(C + "gapWidth", new XAttribute("val", 150)));
        plot.Add(new XElement(C + "axId", new XAttribute("val", 111111)), new XElement(C + "axId", new XAttribute("val", 222222)));
        var chart = new XElement(C + "chart");
        if (!string.IsNullOrEmpty(sp.Title)) chart.Add(new XElement(C + "title", new XElement(C + "tx", new XElement(C + "rich", new XElement(A + "bodyPr"), new XElement(A + "p", new XElement(A + "r", new XElement(A + "t", sp.Title))))), new XElement(C + "overlay", new XAttribute("val", 0))));
        chart.Add(new XElement(C + "autoTitleDeleted", new XAttribute("val", string.IsNullOrEmpty(sp.Title) ? 1 : 0)),
            new XElement(C + "plotArea", new XElement(C + "layout"), plot,
                new XElement(C + "catAx", new XElement(C + "axId", new XAttribute("val", 111111)), new XElement(C + "scaling", new XElement(C + "orientation", new XAttribute("val", "minMax"))), new XElement(C + "delete", new XAttribute("val", 0)), new XElement(C + "axPos", new XAttribute("val", "b")), new XElement(C + "numFmt", new XAttribute("formatCode", "General"), new XAttribute("sourceLinked", 1)), new XElement(C + "majorTickMark", new XAttribute("val", "out")), new XElement(C + "minorTickMark", new XAttribute("val", "none")), new XElement(C + "tickLblPos", new XAttribute("val", "nextTo")), new XElement(C + "crossAx", new XAttribute("val", 222222)), new XElement(C + "crosses", new XAttribute("val", "autoZero")), new XElement(C + "auto", new XAttribute("val", 1)), new XElement(C + "lblAlgn", new XAttribute("val", "ctr")), new XElement(C + "lblOffset", new XAttribute("val", 100)), new XElement(C + "noMultiLvlLbl", new XAttribute("val", 0))),
                new XElement(C + "valAx", new XElement(C + "axId", new XAttribute("val", 222222)), new XElement(C + "scaling", new XElement(C + "orientation", new XAttribute("val", "minMax"))), new XElement(C + "delete", new XAttribute("val", 0)), new XElement(C + "axPos", new XAttribute("val", "l")), new XElement(C + "majorGridlines"), new XElement(C + "numFmt", new XAttribute("formatCode", "General"), new XAttribute("sourceLinked", 1)), new XElement(C + "majorTickMark", new XAttribute("val", "out")), new XElement(C + "minorTickMark", new XAttribute("val", "none")), new XElement(C + "tickLblPos", new XAttribute("val", "nextTo")), new XElement(C + "crossAx", new XAttribute("val", 111111)), new XElement(C + "crosses", new XAttribute("val", "autoZero")), new XElement(C + "crossBetween", new XAttribute("val", "between")))),
            new XElement(C + "legend", new XElement(C + "legendPos", new XAttribute("val", "r")), new XElement(C + "overlay", new XAttribute("val", 0))),
            new XElement(C + "plotVisOnly", new XAttribute("val", 1)), new XElement(C + "dispBlanksAs", new XAttribute("val", "gap")));
        return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), new XElement(C + "chartSpace", new XAttribute(XNamespace.Xmlns + "c", C), new XAttribute(XNamespace.Xmlns + "a", A), new XAttribute(XNamespace.Xmlns + "r", R),
            new XElement(C + "roundedCorners", new XAttribute("val", 0)), chart));
    }

    static XElement Ref(string kind, string f, string cacheName, string? fmt, bool numeric)
    {
        var cache = new XElement(C + cacheName);
        if (fmt != null) cache.Add(new XElement(C + "formatCode", fmt));
        var vals = RangeReader?.Invoke(f) ?? [];
        cache.Add(new XElement(C + "ptCount", new XAttribute("val", vals.Count)));
        for (int i = 0; i < vals.Count; i++)
        {
            if (vals[i] == null) continue;
            string v = vals[i] is double d ? d.ToString("R", CultureInfo.InvariantCulture) : vals[i]!.ToString()!;
            if (numeric && vals[i] is not double) continue;
            cache.Add(new XElement(C + "pt", new XAttribute("idx", i), new XElement(C + "v", v)));
        }
        return new XElement(C + kind, new XElement(C + "f", f), cache);
    }

    // ---------------- commenti ----------------
    public static void WriteComments(Sheet sh, XElement root, XNamespace ns, WorksheetPart wsp)
    {
        if (!sh.CommentsDirty) return;
        var cp = wsp.WorksheetCommentsPart;
        XDocument doc;
        if (cp != null) { using var s = cp.GetStream(FileMode.Open, FileAccess.Read); doc = XDocument.Load(s); }
        else
        {
            if (sh.Comments.Count == 0) { sh.CommentsDirty = false; return; }
            doc = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), new XElement(Main + "comments", new XElement(Main + "authors"), new XElement(Main + "commentList")));
        }
        var m = doc.Root!.Name.Namespace;
        var authors = doc.Root.Element(m + "authors")!; var list = doc.Root.Element(m + "commentList")!;
        int AuthorId(string a)
        {
            var l = authors.Elements(m + "author").ToList(); int i = l.FindIndex(x => x.Value == a);
            if (i >= 0) return i; authors.Add(new XElement(m + "author", a)); return l.Count;
        }
        var removedRefs = new List<(int, int)>(); var added = new List<(int R, int C)>();
        foreach (var e in list.Elements(m + "comment").ToList())
        {
            if (!Addr.TryParse((string?)e.Attribute("ref") ?? "", out var r, out var c)) continue;
            if (!sh.Comments.TryGetValue(Addr.Key(r, c), out var cm)) { e.Remove(); removedRefs.Add((r, c)); continue; }
            if (cm.Text != cm.Orig) { e.Element(m + "text")?.Remove(); e.Add(TextEl(m, cm.Text)); cm.Orig = cm.Text; }
        }
        var have = list.Elements(m + "comment").Select(e => (string?)e.Attribute("ref")).ToHashSet();
        foreach (var kv in sh.Comments.OrderBy(k => Addr.Row(k.Key)).ThenBy(k => Addr.Col(k.Key)))
        {
            int r = Addr.Row(kv.Key), c = Addr.Col(kv.Key); string a1 = Addr.A1(r, c);
            if (have.Contains(a1)) continue;
            list.Add(new XElement(m + "comment", new XAttribute("ref", a1), new XAttribute("authorId", AuthorId(kv.Value.Author)), TextEl(m, kv.Value.Text)));
            kv.Value.Orig = kv.Value.Text; added.Add((r, c));
        }
        // le voci vanno in ordine di riga, poi colonna
        var sorted = list.Elements(m + "comment").OrderBy(e => { Addr.TryParse((string)e.Attribute("ref")!, out var r, out var c); return (long)r * 100000 + c; }).ToList();
        list.RemoveNodes(); foreach (var e in sorted) list.Add(e);
        if (cp == null && sh.Comments.Count > 0) cp = wsp.AddNewPart<WorksheetCommentsPart>();
        if (cp != null) { using var os = cp.GetStream(FileMode.Create); doc.Save(os); }
        WriteVml(sh, root, ns, wsp, removedRefs, added);
        sh.CommentsDirty = false;
    }

    static XElement TextEl(XNamespace m, string text) => new(m + "text", new XElement(m + "r", new XElement(m + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), text)));

    static void WriteVml(Sheet sh, XElement root, XNamespace ns, WorksheetPart wsp, List<(int, int)> removed, List<(int R, int C)> added)
    {
        VmlDrawingPart? vp = null;
        if ((string?)root.Element(ns + "legacyDrawing")?.Attribute(R + "id") is { } rid) { try { vp = wsp.GetPartById(rid) as VmlDrawingPart; } catch (ArgumentOutOfRangeException) { } }
        XDocument? doc = null;
        if (vp != null) { try { using var s = vp.GetStream(FileMode.Open, FileAccess.Read); doc = XDocument.Load(s); } catch (System.Xml.XmlException) { return; } }
        else if (added.Count > 0)
            doc = XDocument.Parse("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\"><o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"1\"/></o:shapelayout><v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" path=\"m,l,21600r21600,l21600,xe\"><v:stroke joinstyle=\"miter\"/><v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/></v:shapetype></xml>");
        if (doc == null) return;
        foreach (var (r, c) in removed)
            foreach (var sp in doc.Root!.Elements(V + "shape").Where(e => (int?)e.Element(X + "ClientData")?.Element(X + "Row") == r - 1 && (int?)e.Element(X + "ClientData")?.Element(X + "Column") == c - 1).ToList()) sp.Remove();
        int maxId = doc.Root!.Elements(V + "shape").Select(e => int.TryParse(((string?)e.Attribute("id") ?? "").Replace("_x0000_s", ""), out var n) ? n : 1024).DefaultIfEmpty(1024).Max();
        foreach (var (r, c) in added)
        {
            var sp = XElement.Parse($"<v:shape xmlns:v=\"{V}\" xmlns:o=\"{O}\" xmlns:x=\"{X}\" id=\"_x0000_s{++maxId}\" type=\"#_x0000_t202\" style=\"position:absolute;margin-left:80pt;margin-top:5pt;width:108pt;height:59pt;z-index:{maxId - 1024};visibility:hidden\" fillcolor=\"#ffffe1\" o:insetmode=\"auto\"><v:fill color2=\"#ffffe1\"/><v:shadow on=\"t\" color=\"black\" obscured=\"t\"/><v:path o:connecttype=\"none\"/><v:textbox style=\"mso-direction-alt:auto\"><div style=\"text-align:left\"></div></v:textbox><x:ClientData ObjectType=\"Note\"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>{c}, 15, {Math.Max(0, r - 2)}, 10, {c + 2}, 15, {r + 3}, 4</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>{r - 1}</x:Row><x:Column>{c - 1}</x:Column></x:ClientData></v:shape>");
            doc.Root.Add(sp);
        }
        if (vp == null) { vp = wsp.AddNewPart<VmlDrawingPart>(); XlsxWriter.InsertOrdered(root, ns, new XElement(ns + "legacyDrawing", new XAttribute(R + "id", wsp.GetIdOfPart(vp)))); }
        using var os = vp.GetStream(FileMode.Create); doc.Save(os);
    }
}
