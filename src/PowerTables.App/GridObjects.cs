using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using PowerTables.Core;

namespace PowerTables.App;

/// <summary>Disegno di oggetti (immagini, forme, caselle, grafici) e indicatori dei commenti sopra la griglia.</summary>
public sealed partial class GridCanvas
{
    static readonly Dictionary<SheetObject, ImageSource?> ImgCache = [];

    Rect ObjRect(SheetObject o) { double x = XOf(o.C1), y = YOf(o.R1); return new Rect(x, y, XOf(o.C2) + M.ColW(o.C2) - x, YOf(o.R2) + M.RowH(o.R2) - y); }

    /// <summary>Oggetto sotto il punto (l'ultimo disegnato vince).</summary>
    public SheetObject? ObjectAt(Point p) => Sheet?.Objects.LastOrDefault(o => ObjRect(o).Contains(p));

    void DrawObjects(DrawingContext dc, double W, double H)
    {
        var sh = Sheet;
        double fx = HeaderW + FrozenW, fy = HeaderH + FrozenH;
        // indicatore dei commenti: triangolo rosso nell'angolo in alto a destra
        foreach (var k in sh.Comments.Keys)
        {
            int r = Addr.Row(k), c = Addr.Col(k); double x = XOf(c) + M.ColW(c), y = YOf(r);
            if (x < HeaderW || y < HeaderH || x > W + 20 || y > H + 20 || M.ColW(c) <= 0 || M.RowH(r) <= 0) continue;
            var g = new StreamGeometry(); using (var cx = g.Open()) { cx.BeginFigure(new Point(x - 6, y), true, true); cx.LineTo(new Point(x, y), false, false); cx.LineTo(new Point(x, y + 6), false, false); }
            dc.PushClip(new RectangleGeometry(new Rect(HeaderW, HeaderH, W - HeaderW, H - HeaderH)));
            dc.DrawGeometry(Brushes.Red, null, g); dc.Pop();
        }
        if (ViewMode > 0) DrawPageBreaks(dc, W, H);
        foreach (var o in sh.Objects)
        {
            var rc = ObjRect(o);
            if (rc.Right < HeaderW || rc.Bottom < HeaderH || rc.Left > W || rc.Top > H || rc.Width < 2 || rc.Height < 2) continue;
            double cx0 = o.C1 <= FC ? HeaderW : fx, cy0 = o.R1 <= FR ? HeaderH : fy;
            dc.PushClip(new RectangleGeometry(new Rect(cx0, cy0, Math.Max(0, W - cx0), Math.Max(0, H - cy0))));
            switch (o.Kind)
            {
                case ObjKind.Picture:
                    if (!ImgCache.TryGetValue(o, out var img)) ImgCache[o] = img = Decode(o.Image!);
                    if (img != null) dc.DrawImage(img, rc); else dc.DrawRectangle(Brushes.LightGray, null, rc);
                    break;
                case ObjKind.Shape:
                case ObjKind.TextBox:
                    DrawShape(dc, o, rc);
                    break;
                case ObjKind.Chart:
                    DrawChart(dc, o, rc);
                    break;
            }
            dc.Pop();
        }
    }

    static ImageSource? Decode(byte[] data)
    {
        try { var b = new BitmapImage(); b.BeginInit(); b.CacheOption = BitmapCacheOption.OnLoad; b.StreamSource = new MemoryStream(data); b.EndInit(); b.Freeze(); return b; }
        catch (Exception e) when (e is NotSupportedException or InvalidOperationException or IOException or FileFormatException) { return null; }
    }

    void DrawShape(DrawingContext dc, SheetObject o, Rect rc)
    {
        bool tb = o.Kind == ObjKind.TextBox;
        var fill = new SolidColorBrush(tb ? Colors.White : Color.FromRgb(0x44, 0x72, 0xC4)); var pen = new Pen(new SolidColorBrush(tb ? Color.FromRgb(0x80, 0x80, 0x80) : Color.FromRgb(0x2F, 0x52, 0x8F)), 1);
        switch (o.Geom)
        {
            case "ellipse": dc.DrawEllipse(fill, pen, new Point(rc.X + rc.Width / 2, rc.Y + rc.Height / 2), rc.Width / 2, rc.Height / 2); break;
            case "roundRect": dc.DrawRoundedRectangle(fill, pen, rc, 10, 10); break;
            case "triangle":
            {
                var g = new StreamGeometry();
                using (var c = g.Open()) { c.BeginFigure(new Point(rc.X + rc.Width / 2, rc.Y), true, true); c.LineTo(new Point(rc.Right, rc.Bottom), true, false); c.LineTo(new Point(rc.X, rc.Bottom), true, false); }
                dc.DrawGeometry(fill, pen, g); break;
            }
            case "rightArrow":
            {
                var g = new StreamGeometry(); double h = rc.Height, t = rc.Y + h * .25, b = rc.Y + h * .75, nx = rc.Right - Math.Min(rc.Width * .4, h * .5);
                using (var c = g.Open()) { c.BeginFigure(new Point(rc.X, t), true, true); foreach (var q in new[] { new Point(nx, t), new Point(nx, rc.Y), new Point(rc.Right, rc.Y + h / 2), new Point(nx, rc.Bottom), new Point(nx, b), new Point(rc.X, b) }) c.LineTo(q, true, false); }
                dc.DrawGeometry(fill, pen, g); break;
            }
            default: dc.DrawRectangle(fill, pen, rc); break;
        }
        if (string.IsNullOrEmpty(o.Text)) return;
        var ft = new FormattedText(o.Text, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, new Typeface("Calibri"), 14.5, tb ? Brushes.Black : Brushes.White, VisualTreeHelper.GetDpi(this).PixelsPerDip)
        { MaxTextWidth = Math.Max(1, rc.Width - 8), MaxTextHeight = Math.Max(1, rc.Height - 4), TextAlignment = tb ? TextAlignment.Left : TextAlignment.Center, Trimming = TextTrimming.None };
        dc.DrawText(ft, new Point(rc.X + 4, tb ? rc.Y + 2 : rc.Y + Math.Max(0, (rc.Height - ft.Height) / 2)));
    }

    static readonly Color[] SerColors = [Color.FromRgb(0x44, 0x72, 0xC4), Color.FromRgb(0xED, 0x7D, 0x31), Color.FromRgb(0xA5, 0xA5, 0xA5), Color.FromRgb(0xFF, 0xC0, 0), Color.FromRgb(0x5B, 0x9B, 0xD5), Color.FromRgb(0x70, 0xAD, 0x47)];

    /// <summary>Valori numerici di un intervallo "Foglio!$B$2:$B$6" (NaN per le celle non numeriche).</summary>
    List<double> RangeNums(string f)
    {
        var res = new List<double>();
        EachCell(f, (sh, r, c) => res.Add(sh.GetValue(r, c) is double d ? d : double.NaN));
        return res;
    }

    /// <summary>Testo visualizzato delle celle di un intervallo: etichette di categoria del grafico.</summary>
    List<string> RangeLabels(string f)
    {
        var res = new List<string>();
        EachCell(f, (sh, r, c) => res.Add(Wb.DisplayText(sh, r, c, Culture)));
        return res;
    }

    void EachCell(string f, Action<Sheet, int, int> act)
    {
        int i = f.LastIndexOf('!'); if (i < 0) return;
        string sn = f[..i]; if (sn.Length > 1 && sn[0] == (char)39) sn = sn[1..^1].Replace("''", "'");
        var sh = Wb.Sheets.FirstOrDefault(x => string.Equals(x.Name, sn, StringComparison.OrdinalIgnoreCase));
        if (sh == null || !Addr.TryParseRange(f[(i + 1)..].Replace("$", ""), out var r1, out var c1, out var r2, out var c2)) return;
        for (int r = r1; r <= Math.Min(r2, r1 + 999); r++) for (int c = c1; c <= c2; c++) act(sh, r, c);
    }

    void DrawChart(DrawingContext dc, SheetObject o, Rect rc)
    {
        dc.DrawRectangle(Brushes.White, new Pen(new SolidColorBrush(Color.FromRgb(0xBF, 0xBF, 0xBF)), 1), rc);
        var sp = o.Chart!; var series = sp.Series.Select(s => RangeNums(s.Values)).ToList();
        int n = series.Max(s => s.Count); if (n == 0) return;
        var all = series.SelectMany(s => s).Where(v => !double.IsNaN(v)).ToList(); if (all.Count == 0) return;
        double max = Math.Max(0, all.Max()), min = Math.Min(0, all.Min()); if (max == min) max = min + 1;
        var plot = new Rect(rc.X + 36, rc.Y + 14, Math.Max(10, rc.Width - 50), Math.Max(10, rc.Height - 54));   // sotto: etichette di categoria e legenda
        double Y(double v) => plot.Bottom - (v - min) / (max - min) * plot.Height;
        var grid = new Pen(new SolidColorBrush(Color.FromRgb(0xD9, 0xD9, 0xD9)), 1);
        double ppd = VisualTreeHelper.GetDpi(this).PixelsPerDip;
        for (int g = 0; g <= 4; g++)
        {
            double v = min + (max - min) * g / 4, y = Y(v);
            dc.DrawLine(grid, new Point(plot.X, y), new Point(plot.Right, y));
            dc.DrawText(new FormattedText(v.ToString("0.##", CultureInfo.CurrentCulture), CultureInfo.CurrentCulture, FlowDirection.LeftToRight, new Typeface("Calibri"), 10, Brushes.Gray, ppd), new Point(rc.X + 3, y - 6));
        }
        double slot = plot.Width / n;
        for (int si = 0; si < series.Count; si++)
        {
            var br = new SolidColorBrush(SerColors[si % SerColors.Length]);
            if (sp.Type == "line")
            {
                var pen = new Pen(br, 2); Point? prev = null;
                for (int i = 0; i < series[si].Count; i++)
                {
                    double v = series[si][i]; if (double.IsNaN(v)) { prev = null; continue; }
                    var pt = new Point(plot.X + slot * (i + .5), Y(v));
                    if (prev != null) dc.DrawLine(pen, prev.Value, pt);
                    prev = pt;
                }
            }
            else
            {
                double bw = slot * .7 / series.Count;
                for (int i = 0; i < series[si].Count; i++)
                {
                    double v = series[si][i]; if (double.IsNaN(v)) continue;
                    double x = plot.X + slot * (i + .15) + bw * si, y0 = Y(0), y1 = Y(v);
                    dc.DrawRectangle(br, null, new Rect(x, Math.Min(y0, y1), Math.Max(1, bw), Math.Max(1, Math.Abs(y1 - y0))));
                }
            }
        }
        dc.DrawLine(new Pen(Brushes.Gray, 1), new Point(plot.X, Y(0)), new Point(plot.Right, Y(0)));
        if (sp.Series[0].Categories is { } cats)
        {
            var labels = RangeLabels(cats);
            for (int i = 0; i < Math.Min(n, labels.Count); i++)
            {
                var t = new FormattedText(labels[i], CultureInfo.CurrentCulture, FlowDirection.LeftToRight, new Typeface("Calibri"), 10, Brushes.Gray, ppd) { MaxTextWidth = Math.Max(1, slot), MaxLineCount = 1, Trimming = TextTrimming.CharacterEllipsis, TextAlignment = TextAlignment.Center };
                dc.DrawText(t, new Point(plot.X + slot * i, plot.Bottom + 2));
            }
        }
        double lx = plot.X;
        for (int si = 0; si < sp.Series.Count; si++)
        {
            var t = new FormattedText(sp.Series[si].Name, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, new Typeface("Calibri"), 10, Brushes.Black, ppd);
            dc.DrawRectangle(new SolidColorBrush(SerColors[si % SerColors.Length]), null, new Rect(lx, rc.Bottom - 15, 8, 8)); dc.DrawText(t, new Point(lx + 11, rc.Bottom - 17)); lx += t.Width + 24;
        }
    }

    /// <summary>Layout di pagina / Anteprima interruzioni: linee tratteggiate fra le pagine, numero di pagina e (solo anteprima) area fuori stampa in grigio.</summary>
    void DrawPageBreaks(DrawingContext dc, double W, double H)
    {
        List<(int From, int To)> rows, cols; (int R1, int C1, int R2, int C2) rg;
        try { (rows, cols, rg) = PowerTables.Core.IO.PdfExport.Pages(Wb, Sheet); } catch (ArgumentException) { return; }
        if (rg.R2 - rg.R1 > 20000 || rows.Count * cols.Count > 400) return;
        double ppd = VisualTreeHelper.GetDpi(this).PixelsPerDip;
        var area = new Rect(HeaderW, HeaderH, Math.Max(0, W - HeaderW), Math.Max(0, H - HeaderH));
        dc.PushClip(new RectangleGeometry(area));
        double x1 = XOf(rg.C1), y1 = YOf(rg.R1), x2 = XOf(rg.C2) + M.ColW(rg.C2), y2 = YOf(rg.R2) + M.RowH(rg.R2);
        if (ViewMode == 2)
        {
            var shade = new SolidColorBrush(Color.FromArgb(70, 0, 0, 0));
            dc.DrawRectangle(shade, null, new Rect(area.X, area.Y, area.Width, Math.Max(0, y1 - area.Y)));
            dc.DrawRectangle(shade, null, new Rect(area.X, y2, area.Width, Math.Max(0, area.Bottom - y2)));
            dc.DrawRectangle(shade, null, new Rect(area.X, y1, Math.Max(0, x1 - area.X), Math.Max(0, y2 - y1)));
            dc.DrawRectangle(shade, null, new Rect(x2, y1, Math.Max(0, area.Right - x2), Math.Max(0, y2 - y1)));
        }
        var pen = new Pen(new SolidColorBrush(Color.FromRgb(0x2F, 0x75, 0xB5)), 1.5) { DashStyle = DashStyles.Dash };
        dc.DrawRectangle(null, pen, new Rect(x1, y1, x2 - x1, y2 - y1));
        foreach (var rp in rows.Skip(1)) { double y = YOf(rp.From); dc.DrawLine(pen, new Point(x1, y), new Point(x2, y)); }
        foreach (var cp in cols.Skip(1)) { double x = XOf(cp.From); dc.DrawLine(pen, new Point(x, y1), new Point(x, y2)); }
        int n = 0;
        foreach (var rp in rows) foreach (var cp in cols)
        {
            n++;
            double px = XOf(cp.From), py = YOf(rp.From), pw = XOf(cp.To) + M.ColW(cp.To) - px, ph = YOf(rp.To) + M.RowH(rp.To) - py;
            var ft = new FormattedText("Pagina " + n, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, new Typeface("Segoe UI"), 28, new SolidColorBrush(Color.FromArgb(90, 0x80, 0x80, 0x80)), ppd);
            dc.DrawText(ft, new Point(px + (pw - ft.Width) / 2, py + (ph - ft.Height) / 2));
        }
        dc.Pop();
    }
}
