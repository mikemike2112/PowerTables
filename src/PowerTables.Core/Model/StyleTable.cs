using System.Globalization;
using System.Xml.Linq;

namespace PowerTables.Core;

public sealed record FontInfo(string Name, double Size, bool Bold, bool Italic, bool Underline, bool Strike, string? Color);
public sealed record BorderSide(string? Style, string? Color);
public sealed record BorderInfo(BorderSide Left, BorderSide Right, BorderSide Top, BorderSide Bottom)
{
    public static readonly BorderInfo None = new(new(null, null), new(null, null), new(null, null), new(null, null));
}
public sealed record XfInfo(int NumFmtId, string? NumFmtCode, FontInfo Font, string? FillColor, BorderInfo Border, string? HAlign, string? VAlign, bool Wrap, int Indent);

/// <summary>Vista modificabile su styles.xml: legge gli xf, ne crea di nuovi senza toccare quelli esistenti.</summary>
public sealed class StyleTable
{
    static readonly XNamespace Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    public XDocument Doc { get; private set; }
    public bool Dirty { get; private set; }
    XElement _numFmts = null!, _fonts = null!, _fills = null!, _borders = null!, _xfs = null!;
    readonly Dictionary<uint, XfInfo> _cache = [];
    readonly Dictionary<XfInfo, uint> _reverse = [];
    readonly Dictionary<uint, string> _codes = [];

    public StyleTable(XDocument doc) { Doc = doc; Bind(); }

    public static StyleTable CreateDefault() => new(XDocument.Parse(
        """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>
        <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
        <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
        <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
        <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
        <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
        </styleSheet>
        """));

    void Bind()
    {
        var root = Doc.Root!;
        XElement Get(string n, string? before = null)
        {
            var e = root.Element(Ns + n);
            if (e != null) return e;
            e = new XElement(Ns + n, new XAttribute("count", 0));
            // ordine schema: numFmts, fonts, fills, borders, cellStyleXfs, cellXfs
            if (n == "numFmts") root.AddFirst(e);
            else root.Add(e);
            return e;
        }
        _numFmts = root.Element(Ns + "numFmts") ?? new XElement(Ns + "numFmts", new XAttribute("count", 0));
        _fonts = Get("fonts"); _fills = Get("fills"); _borders = Get("borders"); _xfs = Get("cellXfs");
    }

    public int XfCount => _xfs.Elements(Ns + "xf").Count();

    public string FormatCode(uint style)
    {
        var x = Get(style);
        return x.NumFmtCode ?? (Formulas.NumFmt.Builtin.TryGetValue(x.NumFmtId, out var c) ? c : "General");
    }

    public XfInfo Get(uint style)
    {
        if (_cache.TryGetValue(style, out var c)) return c;
        var xfs = _xfs.Elements(Ns + "xf").ToList();
        var xf = style < xfs.Count ? xfs[(int)style] : xfs[0];
        var info = ParseXf(xf);
        _cache[style] = info;
        _reverse.TryAdd(info, style);
        return info;
    }

    XfInfo ParseXf(XElement xf)
    {
        int nf = (int?)xf.Attribute("numFmtId") ?? 0;
        string? code = null;
        if (nf >= 164)
            code = _numFmts.Elements(Ns + "numFmt").Where(n => (int?)n.Attribute("numFmtId") == nf).Select(n => (string?)n.Attribute("formatCode")).FirstOrDefault();
        else if (!Formulas.NumFmt.Builtin.ContainsKey(nf)) code = null;
        var font = ParseFont(_fonts.Elements(Ns + "font").ElementAtOrDefault((int?)xf.Attribute("fontId") ?? 0));
        var fill = ParseFill(_fills.Elements(Ns + "fill").ElementAtOrDefault((int?)xf.Attribute("fillId") ?? 0));
        var border = ParseBorder(_borders.Elements(Ns + "border").ElementAtOrDefault((int?)xf.Attribute("borderId") ?? 0));
        var al = xf.Element(Ns + "alignment");
        return new XfInfo(nf, code, font, fill, border, (string?)al?.Attribute("horizontal"), (string?)al?.Attribute("vertical"),
            (bool?)al?.Attribute("wrapText") ?? false, (int?)al?.Attribute("indent") ?? 0);
    }

    static FontInfo ParseFont(XElement? f)
    {
        if (f == null) return new FontInfo("Calibri", 11, false, false, false, false, null);
        static bool Flag(XElement? e) => e != null && ((string?)e.Attribute("val") ?? "1") is not ("0" or "false");
        return new FontInfo((string?)f.Element(Ns + "name")?.Attribute("val") ?? "Calibri",
            (double?)f.Element(Ns + "sz")?.Attribute("val") ?? 11,
            Flag(f.Element(Ns + "b")), Flag(f.Element(Ns + "i")),
            f.Element(Ns + "u") is { } u && ((string?)u.Attribute("val") ?? "single") != "none", Flag(f.Element(Ns + "strike")),
            ParseColor(f.Element(Ns + "color")));
    }

    static string? ParseFill(XElement? f)
    {
        var p = f?.Element(Ns + "patternFill");
        if (p == null) return null;
        string type = (string?)p.Attribute("patternType") ?? "none";
        if (type == "none" || type == "gray125") return null;
        return ParseColor(p.Element(Ns + "fgColor")) ?? ParseColor(p.Element(Ns + "bgColor"));
    }

    static BorderInfo ParseBorder(XElement? b)
    {
        if (b == null) return BorderInfo.None;
        BorderSide Side(string n)
        {
            var e = b.Element(Ns + n);
            return new BorderSide((string?)e?.Attribute("style"), ParseColor(e?.Element(Ns + "color")));
        }
        return new BorderInfo(Side("left"), Side("right"), Side("top"), Side("bottom"));
    }

    static readonly string[] Theme = ["FFFFFF", "000000", "E7E6E6", "44546A", "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47", "0563C1", "954F72"];
    static readonly string[] Indexed =
    [
        "000000","FFFFFF","FF0000","00FF00","0000FF","FFFF00","FF00FF","00FFFF","000000","FFFFFF","FF0000","00FF00","0000FF","FFFF00","FF00FF","00FFFF",
        "800000","008000","000080","808000","800080","008080","C0C0C0","808080","9999FF","993366","FFFFCC","CCFFFF","660066","FF8080","0066CC","CCCCFF",
        "000080","FF00FF","FFFF00","00FFFF","800080","800000","008080","0000FF","00CCFF","CCFFFF","CCFFCC","FFFF99","99CCFF","FF99CC","CC99FF","FFCC99",
        "3366FF","33CCCC","99CC00","FFCC00","FF9900","FF6600","666699","969696","003366","339966","003300","333300","993300","993366","333399","333333",
    ];

    /// <summary>Colore → "AARRGGBB" (tema e indici risolti con la tavolozza Office predefinita).</summary>
    static string? ParseColor(XElement? c)
    {
        if (c == null) return null;
        string? rgb = (string?)c.Attribute("rgb");
        if (rgb != null) return rgb.Length == 6 ? "FF" + rgb : rgb.ToUpperInvariant();
        string? hex = null;
        if ((int?)c.Attribute("theme") is { } t && t >= 0 && t < Theme.Length) hex = Theme[t];
        else if ((int?)c.Attribute("indexed") is { } i) hex = i >= 0 && i < Indexed.Length ? Indexed[i] : (i == 64 ? "000000" : "FFFFFF");
        else if ((bool?)c.Attribute("auto") == true) return null;
        if (hex == null) return null;
        double tint = (double?)c.Attribute("tint") ?? 0;
        if (tint != 0) hex = Tint(hex, tint);
        return "FF" + hex;
    }

    static string Tint(string hex, double tint)
    {
        double r = Convert.ToInt32(hex[..2], 16) / 255.0, g = Convert.ToInt32(hex[2..4], 16) / 255.0, b = Convert.ToInt32(hex[4..6], 16) / 255.0;
        double max = Math.Max(r, Math.Max(g, b)), min = Math.Min(r, Math.Min(g, b)), l = (max + min) / 2, h = 0, s = 0;
        if (max != min)
        {
            double d = max - min;
            s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
            h = max == r ? (g - b) / d + (g < b ? 6 : 0) : max == g ? (b - r) / d + 2 : (r - g) / d + 4;
            h /= 6;
        }
        l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
        double q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
        static double Hue(double p, double q, double t) { if (t < 0) t += 1; if (t > 1) t -= 1; return t < 1.0 / 6 ? p + (q - p) * 6 * t : t < 0.5 ? q : t < 2.0 / 3 ? p + (q - p) * (2.0 / 3 - t) * 6 : p; }
        double rr = s == 0 ? l : Hue(p, q, h + 1.0 / 3), gg = s == 0 ? l : Hue(p, q, h), bb = s == 0 ? l : Hue(p, q, h - 1.0 / 3);
        return $"{(int)Math.Round(rr * 255):X2}{(int)Math.Round(gg * 255):X2}{(int)Math.Round(bb * 255):X2}";
    }

    // ------------- creazione di nuovi stili -------------
    public uint Derive(uint baseStyle, Func<XfInfo, XfInfo> change)
    {
        var bi = Get(baseStyle);
        var ni = change(bi);
        if (ni == bi) return baseStyle;
        if (_reverse.TryGetValue(ni, out var existing)) return existing;
        var xfs = _xfs.Elements(Ns + "xf").ToList();
        var baseXf = baseStyle < xfs.Count ? xfs[(int)baseStyle] : xfs[0];
        var xf = new XElement(baseXf.Name, baseXf.Attributes(), baseXf.Elements().Where(e => e.Name != Ns + "alignment"));

        // numFmt
        int nf = ni.NumFmtId;
        if (ni.NumFmtCode != null)
        {
            var bc = Formulas.NumFmt.Builtin.FirstOrDefault(kv => kv.Value == ni.NumFmtCode);
            if (bc.Value != null) nf = bc.Key;
            else
            {
                var ex = _numFmts.Elements(Ns + "numFmt").FirstOrDefault(n => (string?)n.Attribute("formatCode") == ni.NumFmtCode);
                if (ex != null) nf = (int)ex.Attribute("numFmtId")!;
                else
                {
                    nf = Math.Max(164, _numFmts.Elements(Ns + "numFmt").Select(n => (int)n.Attribute("numFmtId")!).DefaultIfEmpty(163).Max() + 1);
                    if (_numFmts.Parent == null) Doc.Root!.AddFirst(_numFmts);
                    _numFmts.Add(new XElement(Ns + "numFmt", new XAttribute("numFmtId", nf), new XAttribute("formatCode", ni.NumFmtCode)));
                    _numFmts.SetAttributeValue("count", _numFmts.Elements().Count());
                }
            }
        }
        xf.SetAttributeValue("numFmtId", nf);
        if (nf != 0) xf.SetAttributeValue("applyNumberFormat", 1);

        if (ni.Font != bi.Font) { xf.SetAttributeValue("fontId", AddFont(ni.Font, baseXf)); xf.SetAttributeValue("applyFont", 1); }
        if (ni.FillColor != bi.FillColor) { xf.SetAttributeValue("fillId", AddFill(ni.FillColor)); xf.SetAttributeValue("applyFill", 1); }
        if (ni.Border != bi.Border) { xf.SetAttributeValue("borderId", AddBorder(ni.Border)); xf.SetAttributeValue("applyBorder", 1); }

        if (ni.HAlign != null || ni.VAlign != null || ni.Wrap || ni.Indent != 0)
        {
            var al = new XElement(Ns + "alignment");
            if (ni.HAlign != null) al.SetAttributeValue("horizontal", ni.HAlign);
            if (ni.VAlign != null) al.SetAttributeValue("vertical", ni.VAlign);
            if (ni.Wrap) al.SetAttributeValue("wrapText", 1);
            if (ni.Indent != 0) al.SetAttributeValue("indent", ni.Indent);
            xf.AddFirst(al);
            xf.SetAttributeValue("applyAlignment", 1);
        }
        _xfs.Add(xf);
        uint idx = (uint)(xfs.Count);
        _xfs.SetAttributeValue("count", idx + 1);
        _cache[idx] = ni; _reverse[ni] = idx;
        Dirty = true;
        return idx;
    }

    int AddFont(FontInfo f, XElement baseXf)
    {
        var existing = _fonts.Elements(Ns + "font").Select((e, i) => (e, i)).FirstOrDefault(x => ParseFont(x.e) == f);
        if (existing.e != null) return existing.i;
        var el = new XElement(Ns + "font");
        if (f.Bold) el.Add(new XElement(Ns + "b"));
        if (f.Italic) el.Add(new XElement(Ns + "i"));
        if (f.Strike) el.Add(new XElement(Ns + "strike"));
        if (f.Underline) el.Add(new XElement(Ns + "u"));
        el.Add(new XElement(Ns + "sz", new XAttribute("val", f.Size.ToString(CultureInfo.InvariantCulture))));
        el.Add(f.Color != null ? new XElement(Ns + "color", new XAttribute("rgb", f.Color)) : new XElement(Ns + "color", new XAttribute("theme", 1)));
        el.Add(new XElement(Ns + "name", new XAttribute("val", f.Name)));
        el.Add(new XElement(Ns + "family", new XAttribute("val", 2)));
        _fonts.Add(el);
        int n = _fonts.Elements().Count();
        _fonts.SetAttributeValue("count", n);
        return n - 1;
    }

    int AddFill(string? color)
    {
        var existing = _fills.Elements(Ns + "fill").Select((e, i) => (e, i)).FirstOrDefault(x => ParseFill(x.e) == color && (color != null || (string?)x.e.Element(Ns + "patternFill")?.Attribute("patternType") == "none"));
        if (existing.e != null) return existing.i;
        var el = color == null
            ? new XElement(Ns + "fill", new XElement(Ns + "patternFill", new XAttribute("patternType", "none")))
            : new XElement(Ns + "fill", new XElement(Ns + "patternFill", new XAttribute("patternType", "solid"), new XElement(Ns + "fgColor", new XAttribute("rgb", color)), new XElement(Ns + "bgColor", new XAttribute("indexed", 64))));
        _fills.Add(el);
        int n = _fills.Elements().Count();
        _fills.SetAttributeValue("count", n);
        return n - 1;
    }

    int AddBorder(BorderInfo b)
    {
        var existing = _borders.Elements(Ns + "border").Select((e, i) => (e, i)).FirstOrDefault(x => ParseBorder(x.e) == b);
        if (existing.e != null) return existing.i;
        XElement Side(string n, BorderSide s)
        {
            var e = new XElement(Ns + n);
            if (s.Style != null)
            {
                e.SetAttributeValue("style", s.Style);
                e.Add(s.Color != null ? new XElement(Ns + "color", new XAttribute("rgb", s.Color)) : new XElement(Ns + "color", new XAttribute("auto", 1)));
            }
            return e;
        }
        _borders.Add(new XElement(Ns + "border", Side("left", b.Left), Side("right", b.Right), Side("top", b.Top), Side("bottom", b.Bottom), new XElement(Ns + "diagonal")));
        int n = _borders.Elements().Count();
        _borders.SetAttributeValue("count", n);
        return n - 1;
    }

    // ------------- stili differenziali (formattazione condizionale) -------------
    XElement? Dxfs => Doc.Root!.Element(Ns + "dxfs");

    /// <summary>Colore carattere, riempimento e grassetto di un dxf (indice in &lt;dxfs&gt;); null se assente.</summary>
    public (string? Font, string? Fill, bool Bold)? GetDxf(int id)
    {
        var d = Dxfs?.Elements(Ns + "dxf").ElementAtOrDefault(id);
        if (d == null) return null;
        string? fill = ParseColor(d.Element(Ns + "fill")?.Element(Ns + "patternFill")?.Element(Ns + "bgColor")) ?? ParseColor(d.Element(Ns + "fill")?.Element(Ns + "patternFill")?.Element(Ns + "fgColor"));
        bool bold = d.Element(Ns + "font")?.Element(Ns + "b") is { } b && ((string?)b.Attribute("val") ?? "1") is not ("0" or "false");
        return (ParseColor(d.Element(Ns + "font")?.Element(Ns + "color")), fill, bold);
    }

    /// <summary>Aggiunge (o riusa) un dxf con colore carattere, riempimento (AARRGGBB) e grassetto; restituisce l'indice.</summary>
    public int AddDxf(string? fontColor, string? fillColor, bool bold)
    {
        var root = Doc.Root!; var dxfs = Dxfs;
        if (dxfs == null)
        {
            dxfs = new XElement(Ns + "dxfs", new XAttribute("count", 0));
            var anchor = new[] { "cellStyles", "cellXfs" }.Select(n => root.Element(Ns + n)).FirstOrDefault(e => e != null);
            if (anchor != null) anchor.AddAfterSelf(dxfs); else root.Add(dxfs);
        }
        var existing = dxfs.Elements(Ns + "dxf").Select((e, i) => (e, i)).FirstOrDefault(x => GetDxf(x.i) == (fontColor, fillColor, bold) && x.e.Elements().All(c => c.Name.LocalName is "font" or "fill"));
        if (existing.e != null) return existing.i;
        var el = new XElement(Ns + "dxf");
        if (fontColor != null || bold)
        {
            var f = new XElement(Ns + "font");
            if (bold) f.Add(new XElement(Ns + "b"));
            if (fontColor != null) f.Add(new XElement(Ns + "color", new XAttribute("rgb", fontColor)));
            el.Add(f);
        }
        if (fillColor != null) el.Add(new XElement(Ns + "fill", new XElement(Ns + "patternFill", new XElement(Ns + "bgColor", new XAttribute("rgb", fillColor)))));
        dxfs.Add(el);
        int n = dxfs.Elements().Count();
        dxfs.SetAttributeValue("count", n);
        Dirty = true;
        return n - 1;
    }

    internal void MarkClean() => Dirty = false;
}
