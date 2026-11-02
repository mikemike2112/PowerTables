using System.Globalization;
using PdfSharp.Drawing;
using PdfSharp.Fonts;
using PdfSharp.Pdf;

namespace PowerTables.Core.IO;

/// <summary>Esportazione PDF del foglio (PDFsharp 6.2.x, licenza MIT). Usa area di stampa, orientamento, formato carta e margini del foglio.</summary>
public static class PdfExport
{
    const double Pt = 72;   // punti per pollice

    static PdfExport() { GlobalFontSettings.FontResolver = new SystemFonts(); }

    /// <summary>Risolutore minimo dei caratteri: file TrueType della cartella Fonts di sistema per le famiglie più comuni, altrimenti Arial.</summary>
    sealed class SystemFonts : IFontResolver
    {
        // famiglia → file normale, grassetto, corsivo, grassetto corsivo (senza estensione)
        static readonly Dictionary<string, string[]> Map = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Calibri"] = ["calibri", "calibrib", "calibrii", "calibriz"], ["Arial"] = ["arial", "arialbd", "ariali", "arialbi"],
            ["Times New Roman"] = ["times", "timesbd", "timesi", "timesbi"], ["Courier New"] = ["cour", "courbd", "couri", "courbi"],
            ["Verdana"] = ["verdana", "verdanab", "verdanai", "verdanaz"], ["Georgia"] = ["georgia", "georgiab", "georgiai", "georgiaz"],
            ["Segoe UI"] = ["segoeui", "segoeuib", "segoeuii", "segoeuiz"], ["Consolas"] = ["consola", "consolab", "consolai", "consolaz"],
        };
        static string Dir => Environment.GetFolderPath(Environment.SpecialFolder.Fonts);

        public FontResolverInfo? ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            if (!Map.TryGetValue(familyName.Trim(), out var files) || !File.Exists(Path.Combine(Dir, files[0] + ".ttf"))) files = Map["Arial"];
            int i = (isBold ? 1 : 0) + (isItalic ? 2 : 0);
            string f = File.Exists(Path.Combine(Dir, files[i] + ".ttf")) ? files[i] : files[0];
            return new FontResolverInfo(f);
        }

        public byte[]? GetFont(string faceName)
        {
            string path = Path.Combine(Dir, faceName + ".ttf");
            if (!File.Exists(path)) throw new IOException("Carattere non trovato: " + faceName);
            return File.ReadAllBytes(path);
        }
    }

    static double ColPt(Sheet s, int c) => s.ColWidth(c) <= 0 ? 0 : Math.Truncate(s.ColWidth(c) * 7 + 5) * 0.75;

    /// <summary>Area da stampare: area di stampa se presente, altrimenti l'area usata (almeno A1).</summary>
    public static (int R1, int C1, int R2, int C2) PrintRange(Sheet s)
    {
        if (s.PrintArea is { } pa)
        {
            string first = pa.Split(',')[0]; first = first[(first.LastIndexOf('!') + 1)..].Replace("$", "");
            if (Addr.TryParseRange(first, out var a, out var b, out var c, out var d)) return (a, b, c, d);
        }
        int mr = Math.Max(1, s.MaxRow), mc = Math.Max(1, s.MaxCol);
        // le righe/colonne con stile ma senza contenuto non estendono la pagina oltre ciò che è nello store
        return (1, 1, mr, mc);
    }

    static List<(int From, int To)> Split(int from, int to, Func<int, double> size, double avail)
    {
        var l = new List<(int, int)>(); int start = from; double acc = 0;
        for (int i = from; i <= to; i++)
        {
            double sz = size(i);
            if (acc + sz > avail && i > start) { l.Add((start, i - 1)); start = i; acc = 0; }
            acc += sz;
        }
        l.Add((start, to));
        return l;
    }

    /// <summary>Scrive il PDF e restituisce il numero di pagine.</summary>
    public static int Export(Workbook wb, Sheet s, string path, CultureInfo? ci = null)
    {
        ci ??= CultureInfo.CurrentCulture;
        var (r1, c1, r2, c2) = PrintRange(s);
        var paper = Workbook.PaperSizes.FirstOrDefault(p => p.Code == wb.PaperSize(s)); if (paper.Name == null) paper = Workbook.PaperSizes[0];
        bool land = wb.IsLandscape(s);
        double pw = (land ? paper.H : paper.W) / 25.4 * Pt, ph = (land ? paper.W : paper.H) / 25.4 * Pt;
        var m = wb.GetMargins(s);
        double left = m.Left * Pt, top = m.Top * Pt, cw = pw - (m.Left + m.Right) * Pt, ch = ph - (m.Top + m.Bottom) * Pt;
        if (cw < 36 || ch < 36) throw new ArgumentException("I margini lasciano troppo poco spazio nella pagina.");
        var rowPages = Split(r1, r2, r => s.RowHeight(r), ch);
        var colPages = Split(c1, c2, c => ColPt(s, c), cw);
        bool grid = wb.PrintGridLines(s);

        using var doc = new PdfDocument();
        doc.Info.Title = s.Name;
        int pages = 0;
        foreach (var rp in rowPages)
            foreach (var cp in colPages)
            {
                var page = doc.AddPage(); page.Width = XUnit.FromPoint(pw); page.Height = XUnit.FromPoint(ph);
                using var g = XGraphics.FromPdfPage(page);
                DrawPage(wb, s, g, rp.From, cp.From, rp.To, cp.To, left, top, grid, ci);
                pages++;
            }
        string dir = Path.GetDirectoryName(Path.GetFullPath(path))!;
        string tmp = Path.Combine(dir, "." + Path.GetFileName(path) + "." + Guid.NewGuid().ToString("N")[..8] + ".tmp");
        try { doc.Save(tmp); if (File.Exists(path)) File.Replace(tmp, path, null); else File.Move(tmp, path); }
        catch { try { if (File.Exists(tmp)) File.Delete(tmp); } catch { } throw; }
        return pages;
    }

    static XColor Col(string? argb, XColor def) => argb is { Length: 8 } ? XColor.FromArgb(Convert.ToInt32(argb[2..4], 16), Convert.ToInt32(argb[4..6], 16), Convert.ToInt32(argb[6..8], 16)) : def;

    static void DrawPage(Workbook wb, Sheet s, XGraphics g, int r1, int c1, int r2, int c2, double ox, double oy, bool grid, CultureInfo ci)
    {
        var xs = new double[c2 - c1 + 2]; var ys = new double[r2 - r1 + 2];
        for (int c = c1; c <= c2; c++) xs[c - c1 + 1] = xs[c - c1] + ColPt(s, c);
        for (int r = r1; r <= r2; r++) ys[r - r1 + 1] = ys[r - r1] + s.RowHeight(r);
        var gridPen = new XPen(XColors.LightGray, 0.5);
        var skip = new HashSet<long>();
        foreach (var mg in s.Merges) for (int r = mg.R1; r <= mg.R2; r++) for (int c = mg.C1; c <= mg.C2; c++) if (r != mg.R1 || c != mg.C1) skip.Add(Addr.Key(r, c));
        for (int r = r1; r <= r2; r++)
        {
            if (ys[r - r1 + 1] - ys[r - r1] <= 0) continue;
            for (int c = c1; c <= c2; c++)
            {
                if (xs[c - c1 + 1] - xs[c - c1] <= 0 || skip.Contains(Addr.Key(r, c))) continue;
                double x = ox + xs[c - c1], y = oy + ys[r - r1], w = xs[c - c1 + 1] - xs[c - c1], h = ys[r - r1 + 1] - ys[r - r1];
                var mg = s.MergeAt(r, c);
                if (mg is { } mm && mm.R1 == r && mm.C1 == c)
                {
                    w = 0; for (int k = mm.C1; k <= Math.Min(mm.C2, c2); k++) w += ColPt(s, k);
                    h = 0; for (int k = mm.R1; k <= Math.Min(mm.R2, r2); k++) h += s.RowHeight(k);
                }
                var cell = s.Get(r, c);
                var xf = wb.Styles.Get(cell.Style);
                var look = s.CondFmt.Count > 0 ? wb.CondFormatAt(s, r, c, ci) : null;
                string? fill = look?.Fill ?? xf.FillColor;
                if (fill != null) g.DrawRectangle(new XSolidBrush(Col(fill, XColors.White)), x, y, w, h);
                if (grid) g.DrawRectangle(gridPen, x, y, w, h);
                DrawBorders(g, xf.Border, x, y, w, h);
                if (cell.Kind == CellKind.Empty) continue;
                string text = wb.DisplayText(s, r, c, ci);
                if (text.Length == 0) continue;
                bool bold = xf.Font.Bold || look?.Bold == true;
                var style = (bold ? XFontStyleEx.Bold : XFontStyleEx.Regular) | (xf.Font.Italic ? XFontStyleEx.Italic : 0) | (xf.Font.Underline ? XFontStyleEx.Underline : 0) | (xf.Font.Strike ? XFontStyleEx.Strikeout : 0);
                var font = new XFont(xf.Font.Name, Math.Max(1, xf.Font.Size), style);
                var brush = new XSolidBrush(cell.Kind == CellKind.Err ? XColors.DarkRed : Col(look?.Font ?? xf.Font.Color, XColors.Black));
                string hal = xf.HAlign ?? (cell.Kind == CellKind.Num ? "right" : cell.Kind is CellKind.Bool or CellKind.Err ? "center" : "left");
                var lines = xf.Wrap ? WrapText(g, text, font, w - 4) : [text];
                double lh = font.GetHeight(), total = lines.Count * lh;
                double ty = (xf.VAlign ?? "bottom") switch { "top" => y + 1, "center" => y + (h - total) / 2, _ => y + h - total - 1 };
                g.Save();
                g.IntersectClip(new XRect(x, y, w, h));
                foreach (var ln in lines)
                {
                    double tw = g.MeasureString(ln, font).Width;
                    double tx = hal switch { "right" => x + w - 2 - tw, "center" or "centerContinuous" => x + (w - tw) / 2, _ => x + 2 + xf.Indent * 9 };
                    g.DrawString(ln, font, brush, new XRect(tx, ty, tw + 2, lh), XStringFormats.TopLeft);
                    ty += lh;
                }
                g.Restore();
            }
        }
    }

    static List<string> WrapText(XGraphics g, string text, XFont font, double width)
    {
        var res = new List<string>();
        foreach (var para in text.Replace("\r\n", "\n").Split('\n'))
        {
            string cur = "";
            foreach (var word in para.Split(' '))
            {
                string t = cur.Length == 0 ? word : cur + " " + word;
                if (cur.Length > 0 && g.MeasureString(t, font).Width > width) { res.Add(cur); cur = word; } else cur = t;
            }
            res.Add(cur);
        }
        return res;
    }

    static void DrawBorders(XGraphics g, BorderInfo b, double x, double y, double w, double h)
    {
        void Side(BorderSide s, double x1, double y1, double x2, double y2)
        {
            if (s.Style == null) return;
            double wd = s.Style switch { "medium" or "mediumDashed" => 1.25, "thick" => 2, "hair" => 0.25, _ => 0.6 };
            g.DrawLine(new XPen(Col(s.Color, XColors.Black), wd), x1, y1, x2, y2);
        }
        Side(b.Left, x, y, x, y + h); Side(b.Right, x + w, y, x + w, y + h); Side(b.Top, x, y, x + w, y); Side(b.Bottom, x, y + h, x + w, y + h);
    }
}
