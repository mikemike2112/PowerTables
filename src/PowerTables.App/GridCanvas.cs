using System.Globalization;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using PowerTables.Core;
using PowerTables.Core.Formulas;

namespace PowerTables.App;

/// <summary>Griglia virtualizzata: disegna con DrawingContext solo le celle visibili. Nessun controllo per cella.</summary>
public sealed class GridCanvas : FrameworkElement
{
    public const double HeaderH = 22, HeaderW = 50;

    public Workbook Wb { get; private set; } = null!;
    public Sheet Sheet { get; private set; } = null!;
    public readonly GridMetrics M = new();
    public int ActiveRow = 1, ActiveCol = 1, AnchorRow = 1, AnchorCol = 1;
    public double ScrollX, ScrollY;
    public CultureInfo Culture = CultureInfo.CurrentCulture;

    public event Action? SelectionChanged, ViewChanged;
    public event Action<int, Point>? FilterMenuRequested;
    public event Action<char?>? EditRequested;      // tasto digitato o F2/doppio click
    public event Action? CopyRequested, CutRequested, PasteRequested;

    ClipData? _clipRect;     // formiche in marcia
    public ClipData? Marching { get => _clipRect; set { _clipRect = value; InvalidateVisual(); } }

    public (int R1, int C1, int R2, int C2) Selection =>
        (Math.Min(AnchorRow, ActiveRow), Math.Min(AnchorCol, ActiveCol), Math.Max(AnchorRow, ActiveRow), Math.Max(AnchorCol, ActiveCol));

    public GridCanvas()
    {
        Focusable = true; ClipToBounds = true; SnapsToDevicePixels = true;
        FocusVisualStyle = null;
        Cursor = Cursors.Cross;
    }

    public void Attach(Workbook wb, Sheet s)
    {
        if (Wb != null) Wb.Changed -= OnWbChanged;
        Wb = wb; Sheet = s; wb.Changed += OnWbChanged;
        M.Rebuild(s);
        ActiveRow = ActiveCol = AnchorRow = AnchorCol = 1; ScrollX = ScrollY = 0;
        _fmtCache.Clear();
        InvalidateVisual(); SelectionChanged?.Invoke(); ViewChanged?.Invoke();
    }

    public void SwitchSheet(Sheet s) => Attach(Wb, s);

    void OnWbChanged() { M.Rebuild(Sheet); InvalidateVisual(); ViewChanged?.Invoke(); }

    public void RelayoutAll() { M.Rebuild(Sheet); _fmtCache.Clear(); InvalidateVisual(); ViewChanged?.Invoke(); }

    // ---------------- geometria ----------------
    int FR => Sheet.FrozenRows; int FC => Sheet.FrozenCols;
    double FrozenW => M.ColLeft(FC + 1) - M.ColLeft(1);
    double FrozenH => M.RowTop(FR + 1) - M.RowTop(1);

    public double XOf(int c) => c <= FC ? HeaderW + M.ColLeft(c) : HeaderW + FrozenW + M.ColLeft(c) - M.ColLeft(FC + 1) - ScrollX;
    public double YOf(int r) => r <= FR ? HeaderH + M.RowTop(r) : HeaderH + FrozenH + M.RowTop(r) - M.RowTop(FR + 1) - ScrollY;

    public double ViewW => Math.Max(0, ActualWidth - HeaderW - FrozenW);
    public double ViewH => Math.Max(0, ActualHeight - HeaderH - FrozenH);
    public double MaxScrollX => Math.Max(0, M.TotalW - M.ColLeft(FC + 1) - ViewW + 40);
    public double MaxScrollY => Math.Max(0, M.TotalH - M.RowTop(FR + 1) - ViewH + 40);

    public int ColAtX(double x) => x < HeaderW + FrozenW ? M.ColAt(x - HeaderW + M.ColLeft(1)) : M.ColAt(ScrollX + M.ColLeft(FC + 1) + (x - HeaderW - FrozenW));
    public int RowAtY(double y) => y < HeaderH + FrozenH ? M.RowAt(y - HeaderH + M.RowTop(1)) : M.RowAt(ScrollY + M.RowTop(FR + 1) + (y - HeaderH - FrozenH));

    public void SetScroll(double x, double y)
    {
        x = Math.Clamp(x, 0, MaxScrollX); y = Math.Clamp(y, 0, MaxScrollY);
        if (x == ScrollX && y == ScrollY) return;
        ScrollX = x; ScrollY = y; InvalidateVisual(); ViewChanged?.Invoke();
    }

    public void EnsureVisible(int r, int c)
    {
        double nx = ScrollX, ny = ScrollY;
        if (c > FC)
        {
            double left = M.ColLeft(c) - M.ColLeft(FC + 1), right = left + M.ColW(c);
            if (left < ScrollX) nx = left; else if (right > ScrollX + ViewW) nx = right - ViewW;
        }
        if (r > FR)
        {
            double top = M.RowTop(r) - M.RowTop(FR + 1), bottom = top + M.RowH(r);
            if (top < ScrollY) ny = top; else if (bottom > ScrollY + ViewH) ny = bottom - ViewH;
        }
        SetScroll(nx, ny);
    }

    public Rect CellRect(int r, int c)
    {
        var m = Sheet.MergeAt(r, c);
        int r1 = m?.R1 ?? r, c1 = m?.C1 ?? c, r2 = m?.R2 ?? r, c2 = m?.C2 ?? c;
        double x = XOf(c1), y = YOf(r1);
        return new Rect(x, y, XOf(c2) + M.ColW(c2) - x, YOf(r2) + M.RowH(r2) - y);
    }

    // ---------------- rendering ----------------
    readonly Dictionary<string, FormattedText> _fmtCache = [];
    static readonly Typeface UiFont = new("Calibri");
    readonly Dictionary<(string, bool, bool), Typeface> _faces = [];
    Pen? _gridPen; Color _gridPenColor;

    Typeface Face(string name, bool b, bool i)
    {
        if (!_faces.TryGetValue((name, b, i), out var tf))
            _faces[(name, b, i)] = tf = new Typeface(new FontFamily(name), i ? FontStyles.Italic : FontStyles.Normal, b ? FontWeights.Bold : FontWeights.Normal, FontStretches.Normal);
        return tf;
    }

    static Color Argb(string s) => Color.FromArgb(Convert.ToByte(s[..2], 16), Convert.ToByte(s.Substring(2, 2), 16), Convert.ToByte(s.Substring(4, 2), 16), Convert.ToByte(s.Substring(6, 2), 16));
    static double Lum(Color c) => (0.299 * c.R + 0.587 * c.G + 0.114 * c.B) / 255;

    FormattedText Fmt(string text, XfInfo xf, Color color, double wrapWidth)
    {
        string key = text + "\u0001" + xf.Font.Name + xf.Font.Size + xf.Font.Bold + xf.Font.Italic + xf.Font.Underline + xf.Font.Strike + color + (wrapWidth > 0 ? Math.Round(wrapWidth) : 0);
        if (_fmtCache.TryGetValue(key, out var ft)) return ft;
        if (_fmtCache.Count > 6000) _fmtCache.Clear();
        double px = xf.Font.Size * 96 / 72;
        ft = new FormattedText(text, Culture, FlowDirection.LeftToRight, Face(xf.Font.Name, xf.Font.Bold, xf.Font.Italic), px, Theme.Brush(color), VisualTreeHelper.GetDpi(this).PixelsPerDip);
        if (wrapWidth > 0) ft.MaxTextWidth = Math.Max(1, wrapWidth); else ft.MaxLineCount = 1;
        if (xf.Font.Underline || xf.Font.Strike)
        {
            var td = new TextDecorationCollection();
            if (xf.Font.Underline) td.Add(TextDecorations.Underline[0]);
            if (xf.Font.Strike) td.Add(TextDecorations.Strikethrough[0]);
            ft.SetTextDecorations(td);
        }
        _fmtCache[key] = ft;
        return ft;
    }

    protected override void OnRender(DrawingContext dc)
    {
        if (Sheet == null) return;
        double W = ActualWidth, H = ActualHeight;
        dc.DrawRectangle(Theme.Brush(Theme.CellBg), null, new Rect(0, 0, W, H));

        var merged = new Dictionary<long, MergeRange>();
        var (sr1, sc1, sr2, sc2) = Selection;
        var quads = Quadrants(W, H);
        foreach (var q in quads) DrawQuadrant(dc, q, merged);
        DrawHeaders(dc, W, H, quads);
        // separatori dei riquadri bloccati
        var sep = new Pen(Theme.Brush(Theme.Dark ? Color.FromRgb(0x60, 0x60, 0x60) : Color.FromRgb(0x90, 0x90, 0x90)), 1.5);
        if (FR > 0) dc.DrawLine(sep, new Point(0, HeaderH + FrozenH), new Point(W, HeaderH + FrozenH));
        if (FC > 0) dc.DrawLine(sep, new Point(HeaderW + FrozenW, 0), new Point(HeaderW + FrozenW, H));
    }

    readonly record struct Quad(int R0, int R1, int C0, int C1, Rect Clip);

    List<Quad> Quadrants(double W, double H)
    {
        var res = new List<Quad>();
        int firstRow = M.RowAt(ScrollY + M.RowTop(FR + 1)), firstCol = M.ColAt(ScrollX + M.ColLeft(FC + 1));
        int lastRow = firstRow, lastCol = firstCol;
        while (lastRow < Addr.MaxRow && YOf(lastRow + 1) < H) lastRow++;
        while (lastCol < Addr.MaxCol && XOf(lastCol + 1) < W) lastCol++;
        double fx = HeaderW + FrozenW, fy = HeaderH + FrozenH;
        if (FR > 0 && FC > 0) res.Add(new Quad(1, FR, 1, FC, new Rect(HeaderW, HeaderH, FrozenW, FrozenH)));
        if (FR > 0) res.Add(new Quad(1, FR, firstCol, lastCol, new Rect(fx, HeaderH, Math.Max(0, W - fx), FrozenH)));
        if (FC > 0) res.Add(new Quad(firstRow, lastRow, 1, FC, new Rect(HeaderW, fy, FrozenW, Math.Max(0, H - fy))));
        res.Add(new Quad(firstRow, lastRow, firstCol, lastCol, new Rect(fx, fy, Math.Max(0, W - fx), Math.Max(0, H - fy))));
        return res;
    }

    void DrawQuadrant(DrawingContext dc, Quad q, Dictionary<long, MergeRange> merged)
    {
        if (q.Clip.Width <= 0 || q.Clip.Height <= 0) return;
        dc.PushClip(new RectangleGeometry(q.Clip));
        var wb = Wb; var sh = Sheet;
        // unioni visibili
        var mergeHere = new List<MergeRange>();
        foreach (var m in sh.Merges) if (m.Intersects(q.R0, q.C0, q.R1, q.C1)) mergeHere.Add(m);
        var skip = new HashSet<long>();
        foreach (var m in mergeHere)
            for (int r = Math.Max(m.R1, q.R0); r <= Math.Min(m.R2, q.R1); r++)
                for (int c = Math.Max(m.C1, q.C0); c <= Math.Min(m.C2, q.C1); c++) skip.Add(Addr.Key(r, c));

        // 1) riempimenti
        for (int r = q.R0; r <= q.R1; r++)
        {
            double rh = M.RowH(r); if (rh <= 0) continue;
            double y = YOf(r);
            for (int c = q.C0; c <= q.C1; c++)
            {
                double cw = M.ColW(c); if (cw <= 0) continue;
                if (skip.Contains(Addr.Key(r, c))) continue;
                if (!sh.CellMap.TryGetValue(Addr.Key(r, c), out var cell) || cell.Style == 0) continue;
                var xf = wb.Styles.Get(cell.Style);
                if (xf.FillColor != null) dc.DrawRectangle(Theme.Brush(Argb(xf.FillColor)), null, new Rect(XOf(c), y, cw, rh));
            }
        }
        foreach (var m in mergeHere)
        {
            var cell = sh.Get(m.R1, m.C1); var xf = wb.Styles.Get(cell.Style);
            var rect = CellRect(m.R1, m.C1);
            dc.DrawRectangle(xf.FillColor != null ? Theme.Brush(Argb(xf.FillColor)) : Theme.Brush(Theme.CellBg), null, rect);
        }

        // 2) griglia
        if (sh.ShowGridLines)
        {
            if (_gridPen == null || _gridPenColor != Theme.GridLine) { _gridPenColor = Theme.GridLine; _gridPen = new Pen(Theme.Brush(_gridPenColor), 1); _gridPen.Freeze(); }
            double x0 = XOf(q.C0), x1 = XOf(q.C1) + M.ColW(q.C1), y0 = YOf(q.R0), y1 = YOf(q.R1) + M.RowH(q.R1);
            for (int r = q.R0; r <= q.R1; r++)
            {
                double y = Math.Floor(YOf(r) + M.RowH(r)) + 0.5;
                if (mergeHere.Count == 0) dc.DrawLine(_gridPen, new Point(x0, y), new Point(x1, y));
                else DrawHSegments(dc, r, y, q, mergeHere);
            }
            for (int c = q.C0; c <= q.C1; c++)
            {
                double x = Math.Floor(XOf(c) + M.ColW(c)) + 0.5;
                if (mergeHere.Count == 0) dc.DrawLine(_gridPen, new Point(x, y0), new Point(x, y1));
                else DrawVSegments(dc, c, x, q, mergeHere);
            }
        }

        // 3) bordi e testo
        for (int r = q.R0; r <= q.R1; r++)
        {
            double rh = M.RowH(r); if (rh <= 0) continue;
            for (int c = q.C0; c <= q.C1; c++)
            {
                double cw = M.ColW(c); if (cw <= 0) continue;
                if (skip.Contains(Addr.Key(r, c))) continue;
                if (!sh.CellMap.TryGetValue(Addr.Key(r, c), out var cell)) continue;
                DrawCell(dc, r, c, cell, new Rect(XOf(c), YOf(r), cw, rh), q);
            }
        }
        foreach (var m in mergeHere)
        {
            var cell = sh.Get(m.R1, m.C1);
            DrawCell(dc, m.R1, m.C1, cell, CellRect(m.R1, m.C1), q);
        }

        // 4) filtro
        if (sh.Filter is { } f)
            for (int c = Math.Max(f.C1, q.C0); c <= Math.Min(f.C2, q.C1); c++)
                if (f.R1 >= q.R0 && f.R1 <= q.R1) DrawFilterButton(dc, c, f);

        // 5) selezione
        DrawSelection(dc, q);
        if (_clipRect != null && ReferenceEquals(_clipRect.Source, Sheet))
        {
            var pen = new Pen(Theme.Brush(Theme.Accent), 1.5) { DashStyle = DashStyles.Dash };
            var rr = RangeRect(_clipRect.Row, _clipRect.Col, _clipRect.Row + _clipRect.Rows - 1, _clipRect.Col + _clipRect.Cols - 1, q);
            if (rr.HasValue) dc.DrawRectangle(null, pen, rr.Value);
        }
        dc.Pop();
    }

    void DrawHSegments(DrawingContext dc, int r, double y, Quad q, List<MergeRange> merges)
    {
        int c = q.C0;
        while (c <= q.C1)
        {
            // salta segmenti dentro unioni che attraversano questa linea
            var m = merges.FirstOrDefault(mm => r >= mm.R1 && r < mm.R2 && c >= mm.C1 && c <= mm.C2);
            if (m.R2 != 0) { c = m.C2 + 1; continue; }
            dc.DrawLine(_gridPen!, new Point(XOf(c), y), new Point(XOf(c) + M.ColW(c), y));
            c++;
        }
    }

    void DrawVSegments(DrawingContext dc, int c, double x, Quad q, List<MergeRange> merges)
    {
        int r = q.R0;
        while (r <= q.R1)
        {
            var m = merges.FirstOrDefault(mm => c >= mm.C1 && c < mm.C2 && r >= mm.R1 && r <= mm.R2);
            if (m.R2 != 0) { r = m.R2 + 1; continue; }
            dc.DrawLine(_gridPen!, new Point(x, YOf(r)), new Point(x, YOf(r) + M.RowH(r)));
            r++;
        }
    }

    void DrawFilterButton(DrawingContext dc, int c, AutoFilterState f)
    {
        var rc = CellRect(f.R1, c);
        var b = new Rect(rc.Right - 18, rc.Bottom - 18, 16, 16);
        bool active = f.Allowed.ContainsKey(c);
        dc.DrawRectangle(Theme.Brush(Theme.PanelAlt), new Pen(Theme.Brush(active ? Theme.Accent : Theme.Border), 1), b);
        var g = new StreamGeometry();
        using (var ctx = g.Open()) { ctx.BeginFigure(new Point(b.X + 4, b.Y + 6), true, true); ctx.LineTo(new Point(b.X + 12, b.Y + 6), false, false); ctx.LineTo(new Point(b.X + 8, b.Y + 11), false, false); }
        g.Freeze();
        dc.DrawGeometry(Theme.Brush(active ? Theme.Accent : Theme.Text), null, g);
    }

    void DrawCell(DrawingContext dc, int r, int c, Cell cell, Rect rect, Quad q)
    {
        var xf = Wb.Styles.Get(cell.Style);
        // bordi
        if (cell.Style != 0 && xf.Border != BorderInfo.None) DrawBorders(dc, xf.Border, rect);
        if (cell.Kind == CellKind.Empty) return;
        string text = Wb.DisplayText(Sheet, r, c, Culture);
        if (text.Length == 0) return;
        // colore
        Color bg = xf.FillColor != null ? Argb(xf.FillColor) : Theme.CellBg;
        Color col;
        if (xf.Font.Color != null)
        {
            col = Argb(xf.Font.Color);
            if (xf.FillColor == null && Theme.Dark && Lum(col) < 0.2) col = Theme.CellText; // testo nero su sfondo scuro → colore automatico
        }
        else col = xf.FillColor != null ? (Lum(bg) > 0.55 ? Colors.Black : Colors.White) : Theme.CellText;
        if (cell.Kind == CellKind.Err) col = Color.FromRgb(0xe5, 0x6b, 0x6b);
        bool isNum = cell.Kind == CellKind.Num;
        string h = xf.HAlign ?? (isNum ? "right" : cell.Kind is CellKind.Bool or CellKind.Err ? "center" : "left");
        double pad = 3 + xf.Indent * 9;
        double avail = rect.Width - pad - 3;
        var ft = Fmt(text, xf, col, xf.Wrap ? avail : 0);
        if (isNum && !xf.Wrap && ft.Width > avail + 0.5)
        {
            var hash = Fmt("#", xf, col, 0);
            int n = Math.Max(1, (int)(avail / Math.Max(1, hash.Width)));
            ft = Fmt(new string('#', n), xf, col, 0);
        }
        double tx = h switch { "right" => rect.Right - 3 - ft.Width, "center" or "centerContinuous" => rect.X + (rect.Width - ft.Width) / 2, _ => rect.X + pad };
        double ty = (xf.VAlign ?? "bottom") switch { "top" => rect.Y + 1, "center" => rect.Y + (rect.Height - ft.Height) / 2, _ => rect.Bottom - ft.Height - 1 };
        if (xf.Wrap) ty = Math.Max(rect.Y + 1, ty);
        // sconfinamento nelle celle vuote a destra per testo allineato a sinistra
        var clip = rect;
        if (!xf.Wrap && h == "left" && cell.Kind == CellKind.Str && ft.Width > avail)
        {
            double extra = 0; int cc = c + 1;
            while (cc <= q.C1 + 6 && cc <= Addr.MaxCol && rect.Width + extra < ft.Width + pad + 3 && !Sheet.CellMap.ContainsKey(Addr.Key(r, cc)) && Sheet.MergeAt(r, cc) == null)
            { extra += M.ColW(cc); cc++; }
            clip = new Rect(rect.X, rect.Y, rect.Width + extra, rect.Height);
        }
        else if (!xf.Wrap && h == "right" && ft.Width > avail && !isNum) clip = new Rect(rect.X - 1000, rect.Y, rect.Width + 1000, rect.Height);
        dc.PushClip(new RectangleGeometry(clip));
        dc.DrawText(ft, new Point(tx, ty));
        dc.Pop();
    }

    void DrawBorders(DrawingContext dc, BorderInfo b, Rect r)
    {
        void Side(BorderSide s, Point a, Point z)
        {
            if (s.Style == null) return;
            double w = s.Style is "medium" or "mediumDashed" or "mediumDashDot" ? 2 : s.Style == "thick" ? 3 : 1;
            var color = s.Color != null ? Argb(s.Color) : (Theme.Dark ? Color.FromRgb(0xd0, 0xd0, 0xd0) : Colors.Black);
            if (s.Color != null && Theme.Dark && Lum(color) < 0.2) color = Color.FromRgb(0xd0, 0xd0, 0xd0);
            var pen = new Pen(Theme.Brush(color), w);
            if (s.Style is "dashed" or "mediumDashed") pen.DashStyle = DashStyles.Dash; else if (s.Style is "dotted" or "hair") pen.DashStyle = DashStyles.Dot;
            dc.DrawLine(pen, a, z);
        }
        double x0 = Math.Floor(r.X) + 0.5, x1 = Math.Floor(r.Right) + 0.5, y0 = Math.Floor(r.Y) + 0.5, y1 = Math.Floor(r.Bottom) + 0.5;
        Side(b.Left, new Point(x0, y0), new Point(x0, y1)); Side(b.Right, new Point(x1, y0), new Point(x1, y1));
        Side(b.Top, new Point(x0, y0), new Point(x1, y0)); Side(b.Bottom, new Point(x0, y1), new Point(x1, y1));
    }

    Rect? RangeRect(int r1, int c1, int r2, int c2, Quad q)
    {
        int rr1 = Math.Max(r1, q.R0), rr2 = Math.Min(r2, q.R1), cc1 = Math.Max(c1, q.C0), cc2 = Math.Min(c2, q.C1);
        if (rr1 > rr2 || cc1 > cc2) return null;
        double x = XOf(cc1), y = YOf(rr1);
        return new Rect(x, y, XOf(cc2) + M.ColW(cc2) - x, YOf(rr2) + M.RowH(rr2) - y);
    }

    void DrawSelection(DrawingContext dc, Quad q)
    {
        var (r1, c1, r2, c2) = Selection;
        // estende alle unioni coinvolte
        foreach (var m in Sheet.Merges) if (m.Intersects(r1, c1, r2, c2)) { r1 = Math.Min(r1, m.R1); c1 = Math.Min(c1, m.C1); r2 = Math.Max(r2, m.R2); c2 = Math.Max(c2, m.C2); }
        var rect = RangeRect(r1, c1, r2, c2, q);
        var accent = Theme.Accent;
        if (rect.HasValue)
        {
            var rc = rect.Value;
            bool single = r1 == r2 && c1 == c2;
            if (!single) dc.DrawRectangle(new SolidColorBrush(Color.FromArgb(0x28, accent.R, accent.G, accent.B)), null, rc);
            var pen = new Pen(Theme.Brush(accent), 2);
            dc.DrawRectangle(null, pen, new Rect(rc.X + 0.5, rc.Y + 0.5, Math.Max(0, rc.Width - 1), Math.Max(0, rc.Height - 1)));
            // cella attiva
            if (!single)
            {
                var ar = RangeRect(ActiveRow, ActiveCol, ActiveRow, ActiveCol, q);
                var am = Sheet.MergeAt(ActiveRow, ActiveCol);
                if (am != null) ar = RangeRect(am.Value.R1, am.Value.C1, am.Value.R2, am.Value.C2, q);
                if (ar.HasValue) dc.DrawRectangle(Theme.Brush(Theme.CellBg), null, new Rect(ar.Value.X + 1, ar.Value.Y + 1, Math.Max(0, ar.Value.Width - 2), Math.Max(0, ar.Value.Height - 2)));
                if (ar.HasValue) RedrawActive(dc, ActiveRow, ActiveCol, ar.Value, q);
            }
            // maniglia di riempimento
            if (r2 >= q.R0 && r2 <= q.R1 && c2 >= q.C0 && c2 <= q.C1)
                dc.DrawRectangle(Theme.Brush(accent), new Pen(Theme.Brush(Theme.CellBg), 1), new Rect(rc.Right - 4, rc.Bottom - 4, 7, 7));
        }
    }

    void RedrawActive(DrawingContext dc, int r, int c, Rect rect, Quad q)
    {
        var am = Sheet.MergeAt(r, c);
        int rr = am?.R1 ?? r, cc = am?.C1 ?? c;
        if (Sheet.CellMap.TryGetValue(Addr.Key(rr, cc), out var cell)) DrawCell(dc, rr, cc, cell, rect, q);
    }

    void DrawHeaders(DrawingContext dc, double W, double H, List<Quad> quads)
    {
        var bg = Theme.Brush(Theme.HeaderBg);
        var line = new Pen(Theme.Brush(Theme.Border), 1);
        var sel = Selection;
        dc.DrawRectangle(bg, null, new Rect(0, 0, W, HeaderH));
        dc.DrawRectangle(bg, null, new Rect(0, 0, HeaderW, H));
        var xf0 = new XfInfo(0, null, new FontInfo("Segoe UI", 9, false, false, false, false, null), null, BorderInfo.None, null, null, false, 0);
        var accent = Theme.Brush(Color.FromArgb(0x55, Theme.Accent.R, Theme.Accent.G, Theme.Accent.B));
        foreach (var q in quads)
        {
            // intestazioni di colonna
            dc.PushClip(new RectangleGeometry(new Rect(q.Clip.X, 0, q.Clip.Width, HeaderH)));
            for (int c = q.C0; c <= q.C1; c++)
            {
                double w = M.ColW(c); if (w <= 0) continue; double x = XOf(c);
                bool s = c >= sel.C1 && c <= sel.C2;
                if (s) dc.DrawRectangle(accent, null, new Rect(x, 0, w, HeaderH));
                dc.DrawLine(line, new Point(Math.Floor(x + w) + 0.5, 0), new Point(Math.Floor(x + w) + 0.5, HeaderH));
                var ft = Fmt(Addr.ColName(c), xf0, s ? Theme.Accent : Theme.TextDim, 0);
                dc.DrawText(ft, new Point(x + (w - ft.Width) / 2, (HeaderH - ft.Height) / 2));
            }
            dc.Pop();
            dc.PushClip(new RectangleGeometry(new Rect(0, q.Clip.Y, HeaderW, q.Clip.Height)));
            for (int r = q.R0; r <= q.R1; r++)
            {
                double h = M.RowH(r); if (h <= 0) continue; double y = YOf(r);
                bool s = r >= sel.R1 && r <= sel.R2;
                if (s) dc.DrawRectangle(accent, null, new Rect(0, y, HeaderW, h));
                dc.DrawLine(line, new Point(0, Math.Floor(y + h) + 0.5), new Point(HeaderW, Math.Floor(y + h) + 0.5));
                var ft = Fmt(r.ToString(), xf0, s ? Theme.Accent : Theme.TextDim, 0);
                dc.DrawText(ft, new Point(HeaderW - ft.Width - 6, y + (h - ft.Height) / 2));
            }
            dc.Pop();
        }
        dc.DrawLine(line, new Point(0, HeaderH + 0.5), new Point(W, HeaderH + 0.5));
        dc.DrawLine(line, new Point(HeaderW + 0.5, 0), new Point(HeaderW + 0.5, H));
    }

    // ---------------- selezione e navigazione ----------------
    public void Select(int r, int c, bool extend = false, bool ensure = true)
    {
        r = Math.Clamp(r, 1, Addr.MaxRow); c = Math.Clamp(c, 1, Addr.MaxCol);
        ActiveRow = r; ActiveCol = c;
        if (!extend) { AnchorRow = r; AnchorCol = c; }
        if (ensure) EnsureVisible(r, c);
        InvalidateVisual(); SelectionChanged?.Invoke();
    }

    public void SelectRange(int r1, int c1, int r2, int c2)
    {
        AnchorRow = r1; AnchorCol = c1; ActiveRow = r2; ActiveCol = c2;
        InvalidateVisual(); SelectionChanged?.Invoke();
    }

    bool HasData(int r, int c) => Sheet.CellMap.TryGetValue(Addr.Key(r, c), out var x) && x.Kind != CellKind.Empty;

    int Jump(int r, int c, int dr, int dc)
    {
        // comportamento Ctrl+freccia: salta al bordo del blocco di dati
        int rr = r, cc = c;
        bool Next(out int nr, out int nc) { nr = rr + dr; nc = cc + dc; return nr >= 1 && nr <= Addr.MaxRow && nc >= 1 && nc <= Addr.MaxCol; }
        if (!Next(out int r1, out int c1)) return 0;
        bool startFilled = HasData(r, c), nextFilled = HasData(r1, c1);
        if (startFilled && nextFilled)
        {
            while (Next(out r1, out c1) && HasData(r1, c1) && Math.Abs(r1 - r) + Math.Abs(c1 - c) < 2_000_000) { rr = r1; cc = c1; }
        }
        else
        {
            rr = r1; cc = c1;
            int limit = dr != 0 ? Math.Max(Sheet.MaxRow, r) : Math.Max(Sheet.MaxCol, c);
            while (!HasData(rr, cc))
            {
                if (!Next(out int a, out int b)) break;
                if ((dr != 0 && (dr > 0 ? a > limit : a < 1)) || (dc != 0 && (dc > 0 ? b > limit : b < 1))) { rr = dr > 0 ? Addr.MaxRow : dr < 0 ? 1 : rr; cc = dc > 0 ? Addr.MaxCol : dc < 0 ? 1 : cc; break; }
                rr = a; cc = b;
            }
        }
        _jr = rr; _jc = cc; return 1;
    }
    int _jr, _jc;

    // ---------------- input ----------------
    protected override void OnKeyDown(KeyEventArgs e)
    {
        base.OnKeyDown(e);
        if (Wb == null) return;
        var mods = Keyboard.Modifiers;
        bool ctrl = (mods & ModifierKeys.Control) != 0, shift = (mods & ModifierKeys.Shift) != 0;
        int dr = 0, dc = 0;
        var key = e.Key == Key.System ? e.SystemKey : e.Key;
        switch (key)
        {
            case Key.Up: dr = -1; break; case Key.Down: dr = 1; break;
            case Key.Left: dc = -1; break; case Key.Right: dc = 1; break;
            case Key.Tab: dc = shift ? -1 : 1; shift = false; break;
            case Key.Enter: dr = shift ? -1 : 1; shift = false; break;
            case Key.PageDown: Select(ActiveRow + (int)(ViewH / 20), ActiveCol, shift); e.Handled = true; return;
            case Key.PageUp: Select(ActiveRow - (int)(ViewH / 20), ActiveCol, shift); e.Handled = true; return;
            case Key.Home: Select(ctrl ? 1 : ActiveRow, 1, shift); e.Handled = true; return;
            case Key.End when ctrl: Select(Math.Max(1, Sheet.MaxRow), Math.Max(1, Sheet.MaxCol), shift); e.Handled = true; return;
            case Key.F2: EditRequested?.Invoke(null); e.Handled = true; return;
            case Key.Delete: case Key.Back: ClearSelection(); e.Handled = true; return;
            case Key.Escape: Marching = null; e.Handled = true; return;
            case Key.A when ctrl: SelectRange(1, 1, Math.Max(1, Sheet.MaxRow), Math.Max(1, Sheet.MaxCol)); e.Handled = true; return;
            case Key.C when ctrl: CopyRequested?.Invoke(); e.Handled = true; return;
            case Key.X when ctrl: CutRequested?.Invoke(); e.Handled = true; return;
            case Key.V when ctrl: PasteRequested?.Invoke(); e.Handled = true; return;
            case Key.Z when ctrl: Wb.Undo(); e.Handled = true; return;
            case Key.Y when ctrl: Wb.Redo(); e.Handled = true; return;
            default: return;
        }
        e.Handled = true;
        if (key is Key.Tab or Key.Enter)
        {
            var (r1, c1, r2, c2) = Selection;
            bool multi = r1 != r2 || c1 != c2;
            if (multi) { MoveWithin(dr, dc, r1, c1, r2, c2); return; }
        }
        if (ctrl && key is Key.Up or Key.Down or Key.Left or Key.Right && Jump(ActiveRow, ActiveCol, dr, dc) == 1) { Select(_jr, _jc, shift); return; }
        // salta le unioni
        int nr = ActiveRow, nc = ActiveCol;
        var mg = Sheet.MergeAt(nr, nc);
        if (mg != null) { if (dr > 0) nr = mg.Value.R2; if (dc > 0) nc = mg.Value.C2; if (dr < 0) nr = mg.Value.R1; if (dc < 0) nc = mg.Value.C1; }
        nr += dr; nc += dc;
        if (nr >= 1 && nc >= 1) while (nr >= 1 && nr < Addr.MaxRow && M.RowH(nr) <= 0) nr += dr == 0 ? 1 : dr;
        Select(nr, nc, shift);
    }

    void MoveWithin(int dr, int dc, int r1, int c1, int r2, int c2)
    {
        int r = ActiveRow + dr, c = ActiveCol + dc;
        if (c > c2) { c = c1; r++; } else if (c < c1) { c = c2; r--; }
        if (r > r2) { r = r1; c++; if (c > c2) c = c1; } else if (r < r1) { r = r2; c--; if (c < c1) c = c2; }
        ActiveRow = r; ActiveCol = c; EnsureVisible(r, c); InvalidateVisual(); SelectionChanged?.Invoke();
    }

    protected override void OnTextInput(TextCompositionEventArgs e)
    {
        base.OnTextInput(e);
        if (Wb == null || e.Text.Length == 0 || char.IsControl(e.Text[0]) || (Keyboard.Modifiers & ModifierKeys.Control) != 0) return;
        EditRequested?.Invoke(e.Text[0]);
        e.Handled = true;
    }

    public void ClearSelection()
    {
        var (r1, c1, r2, c2) = Selection;
        Wb.Do("Cancella", () =>
        {
            if ((long)(r2 - r1 + 1) * (c2 - c1 + 1) > 200_000)
            {
                foreach (var k in Sheet.CellMap.Keys.Where(k => Addr.Row(k) >= r1 && Addr.Row(k) <= r2 && Addr.Col(k) >= c1 && Addr.Col(k) <= c2).ToList()) Wb.ClearCell(Sheet, Addr.Row(k), Addr.Col(k));
            }
            else
                for (int r = r1; r <= r2; r++) for (int c = c1; c <= c2; c++) if (Sheet.CellMap.ContainsKey(Addr.Key(r, c))) Wb.ClearCell(Sheet, r, c);
        });
    }

    enum Drag { None, Select, ColHeader, RowHeader, ResizeCol, ResizeRow, Fill }
    Drag _drag; int _resizeIdx; double _resizeStart, _resizeOrig; int _fillR1, _fillC1, _fillR2, _fillC2;

    Rect? FillHandleRect()
    {
        var (r1, c1, r2, c2) = Selection;
        var rc = CellRect(r2, c2);
        var m = Sheet.MergeAt(r2, c2); if (m != null) rc = CellRect(m.Value.R1, m.Value.C1);
        return new Rect(rc.Right - 7, rc.Bottom - 7, 14, 14);
    }

    protected override void OnMouseDown(MouseButtonEventArgs e)
    {
        base.OnMouseDown(e);
        if (Wb == null) return;
        Focus();
        var p = e.GetPosition(this);
        bool shift = (Keyboard.Modifiers & ModifierKeys.Shift) != 0;
        if (e.ChangedButton != MouseButton.Left) { if (e.ChangedButton == MouseButton.Right) { var (r, c) = (RowAtY(p.Y), ColAtX(p.X)); var s = Selection; if (p.X > HeaderW && p.Y > HeaderH && !(r >= s.R1 && r <= s.R2 && c >= s.C1 && c <= s.C2)) Select(r, c); } return; }
        // maniglia di riempimento
        if (p.X > HeaderW && p.Y > HeaderH && FillHandleRect() is { } fh && fh.Contains(p))
        {
            (_fillR1, _fillC1, _fillR2, _fillC2) = Selection; _drag = Drag.Fill; CaptureMouse(); return;
        }
        if (p.Y < HeaderH && p.X > HeaderW)
        {
            int c = ColAtX(p.X);
            // bordo destro della colonna = ridimensiona
            double right = XOf(c) + M.ColW(c);
            if (Math.Abs(p.X - right) <= 4 && M.ColW(c) > 0) { _drag = Drag.ResizeCol; _resizeIdx = c; _resizeStart = p.X; _resizeOrig = M.ColW(c); CaptureMouse(); if (e.ClickCount == 2) { AutoFitColumn(c); _drag = Drag.None; ReleaseMouseCapture(); } return; }
            SelectRange(1, shift ? AnchorCol : c, Math.Max(1, Sheet.MaxRow == 0 ? 1 : Addr.MaxRow), c); ActiveRow = 1; _drag = Drag.ColHeader; CaptureMouse(); InvalidateVisual(); return;
        }
        if (p.X < HeaderW && p.Y > HeaderH)
        {
            int r = RowAtY(p.Y);
            double bottom = YOf(r) + M.RowH(r);
            if (Math.Abs(p.Y - bottom) <= 4) { _drag = Drag.ResizeRow; _resizeIdx = r; _resizeStart = p.Y; _resizeOrig = M.RowH(r); CaptureMouse(); return; }
            SelectRange(shift ? AnchorRow : r, 1, r, Addr.MaxCol); ActiveCol = 1; _drag = Drag.RowHeader; CaptureMouse(); InvalidateVisual(); return;
        }
        if (p.X < HeaderW && p.Y < HeaderH) { SelectRange(1, 1, Addr.MaxRow, Addr.MaxCol); return; }
        int row = RowAtY(p.Y), col = ColAtX(p.X);
        // pulsante filtro
        if (Sheet.Filter is { } f && row == f.R1 && col >= f.C1 && col <= f.C2)
        {
            var rc = CellRect(row, col);
            if (new Rect(rc.Right - 19, rc.Bottom - 19, 18, 18).Contains(p)) { FilterMenuRequested?.Invoke(col, PointToScreen(new Point(rc.Left, rc.Bottom))); return; }
        }
        if (e.ClickCount == 2) { EditRequested?.Invoke(null); return; }
        Select(row, col, shift, false);
        _drag = Drag.Select; CaptureMouse();
    }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        base.OnMouseMove(e);
        if (Wb == null) return;
        var p = e.GetPosition(this);
        switch (_drag)
        {
            case Drag.None:
                if (p.Y < HeaderH && p.X > HeaderW && Math.Abs(p.X - (XOf(ColAtX(p.X)) + M.ColW(ColAtX(p.X)))) <= 4) Cursor = Cursors.SizeWE;
                else if (p.X < HeaderW && p.Y > HeaderH && Math.Abs(p.Y - (YOf(RowAtY(p.Y)) + M.RowH(RowAtY(p.Y)))) <= 4) Cursor = Cursors.SizeNS;
                else if (p.X > HeaderW && p.Y > HeaderH && FillHandleRect() is { } fh && fh.Contains(p)) Cursor = Cursors.Cross;
                else Cursor = p.X < HeaderW || p.Y < HeaderH ? Cursors.Arrow : Cursors.Cross;
                break;
            case Drag.Select:
                {
                    AutoScroll(p);
                    int r = RowAtY(Math.Clamp(p.Y, HeaderH + 1, Math.Max(HeaderH + 2, ActualHeight - 1))), c = ColAtX(Math.Clamp(p.X, HeaderW + 1, Math.Max(HeaderW + 2, ActualWidth - 1)));
                    if (r != ActiveRow || c != ActiveCol) { ActiveRow = r; ActiveCol = c; InvalidateVisual(); SelectionChanged?.Invoke(); }
                    break;
                }
            case Drag.ColHeader: { int c = ColAtX(Math.Clamp(p.X, HeaderW + 1, ActualWidth - 1)); AutoScroll(p); if (c != ActiveCol) { ActiveCol = c; InvalidateVisual(); SelectionChanged?.Invoke(); } break; }
            case Drag.RowHeader: { int r = RowAtY(Math.Clamp(p.Y, HeaderH + 1, ActualHeight - 1)); AutoScroll(p); if (r != ActiveRow) { ActiveRow = r; InvalidateVisual(); SelectionChanged?.Invoke(); } break; }
            case Drag.ResizeCol: { double w = Math.Max(0, _resizeOrig + p.X - _resizeStart); Sheet.ColInfoFor(_resizeIdx).Width = Math.Max(0.1, w / GridMetrics.CharPx); Sheet.ColInfoFor(_resizeIdx).CustomWidth = true; M.Rebuild(Sheet); InvalidateVisual(); break; }
            case Drag.ResizeRow: { double h = Math.Max(2, _resizeOrig + p.Y - _resizeStart); var ri = Sheet.RowInfoFor(_resizeIdx); ri.Height = h * 72 / 96; ri.CustomHeight = true; M.Rebuild(Sheet); InvalidateVisual(); break; }
            case Drag.Fill:
                {
                    int r = RowAtY(Math.Clamp(p.Y, HeaderH + 1, ActualHeight - 1)), c = ColAtX(Math.Clamp(p.X, HeaderW + 1, ActualWidth - 1));
                    AutoScroll(p);
                    ActiveRow = Math.Max(r, _fillR1); ActiveCol = Math.Max(c, _fillC1);
                    // direzione prevalente
                    bool vert = r > _fillR2 || r < _fillR1 ? Math.Abs(r - (r < _fillR1 ? _fillR1 : _fillR2)) >= (c > _fillC2 || c < _fillC1 ? Math.Abs(c - (c < _fillC1 ? _fillC1 : _fillC2)) : 0) : false;
                    _fillTargetR = r; _fillTargetC = c; _fillVert = vert || (c >= _fillC1 && c <= _fillC2);
                    SelectRange(Math.Min(_fillR1, _fillVert ? r : _fillR1), Math.Min(_fillC1, _fillVert ? _fillC1 : c), Math.Max(_fillR2, _fillVert ? r : _fillR2), Math.Max(_fillC2, _fillVert ? _fillC2 : c));
                    break;
                }
        }
    }
    int _fillTargetR, _fillTargetC; bool _fillVert;

    DateTime _lastScroll;
    void AutoScroll(Point p)
    {
        if ((DateTime.Now - _lastScroll).TotalMilliseconds < 25) return;
        _lastScroll = DateTime.Now;
        double dx = p.X > ActualWidth ? 40 : p.X < HeaderW + FrozenW ? -40 : 0, dy = p.Y > ActualHeight ? 40 : p.Y < HeaderH + FrozenH ? -40 : 0;
        if (dx != 0 || dy != 0) SetScroll(ScrollX + dx, ScrollY + dy);
    }

    protected override void OnMouseUp(MouseButtonEventArgs e)
    {
        base.OnMouseUp(e);
        var d = _drag; _drag = Drag.None; ReleaseMouseCapture();
        if (Wb == null) return;
        switch (d)
        {
            case Drag.ResizeCol:
                { double w = Sheet.Cols[_resizeIdx].Width * GridMetrics.CharPx; Sheet.ColInfoFor(_resizeIdx).Width = _resizeOrig / GridMetrics.CharPx; Wb.SetColWidth(Sheet, _resizeIdx, (w - 5) / GridMetrics.CharPx); break; }
            case Drag.ResizeRow:
                { double h = Sheet.Rows[_resizeIdx].Height ?? 15; Sheet.RowInfoFor(_resizeIdx).Height = null; Sheet.RowInfoFor(_resizeIdx).CustomHeight = false; Wb.SetRowHeight(Sheet, _resizeIdx, h); break; }
            case Drag.Fill:
                {
                    int tr = _fillTargetR, tc = _fillTargetC;
                    try
                    {
                        if (_fillVert && (tr > _fillR2 || tr < _fillR1)) { Wb.AutoFill(Sheet, _fillR1, _fillC1, _fillR2, _fillC2, tr, _fillC2); SelectRange(Math.Min(_fillR1, tr), _fillC1, Math.Max(_fillR2, tr), _fillC2); }
                        else if (!_fillVert && (tc > _fillC2 || tc < _fillC1)) { Wb.AutoFill(Sheet, _fillR1, _fillC1, _fillR2, _fillC2, _fillR2, tc); SelectRange(_fillR1, Math.Min(_fillC1, tc), _fillR2, Math.Max(_fillC2, tc)); }
                        else SelectRange(_fillR1, _fillC1, _fillR2, _fillC2);
                    }
                    catch (Exception ex) when (ex is FormatException or InvalidOperationException) { MessageBox.Show(ex.Message); }
                    break;
                }
        }
    }

    protected override void OnMouseWheel(MouseWheelEventArgs e)
    {
        base.OnMouseWheel(e);
        double step = (Keyboard.Modifiers & ModifierKeys.Shift) != 0 ? 0 : 3 * 20;
        if (step == 0) SetScroll(ScrollX - e.Delta / 120.0 * 3 * 64, ScrollY); else SetScroll(ScrollX, ScrollY - e.Delta / 120.0 * step);
        e.Handled = true;
    }

    protected override void OnRenderSizeChanged(SizeChangedInfo sizeInfo) { base.OnRenderSizeChanged(sizeInfo); ViewChanged?.Invoke(); }

    public void AutoFitColumn(int c)
    {
        double max = 0; int n = 0;
        var xf0 = new XfInfo(0, null, new FontInfo("Calibri", 11, false, false, false, false, null), null, BorderInfo.None, null, null, false, 0);
        foreach (var kv in Sheet.CellMap)
        {
            if (Addr.Col(kv.Key) != c) continue;
            if (++n > 20000) break;
            string t = Wb.DisplayText(Sheet, Addr.Row(kv.Key), c, Culture);
            if (t.Length == 0) continue;
            var xf = Wb.Styles.Get(kv.Value.Style);
            max = Math.Max(max, Fmt(t, xf, Colors.Black, 0).Width);
        }
        if (max > 0) Wb.SetColWidth(Sheet, c, Math.Min(255, (max + 5) / GridMetrics.CharPx));
    }
}
