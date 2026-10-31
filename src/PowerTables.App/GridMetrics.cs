using PowerTables.Core;

namespace PowerTables.App;

/// <summary>Posizioni in pixel di righe e colonne di un foglio (store sparso: solo le righe con altezza personalizzata costano).</summary>
public sealed class GridMetrics
{
    double[] _colLeft = new double[Addr.MaxCol + 2];
    int[] _custom = [];
    double[] _cum = [];
    double _def = 20;
    int _maxRow = Addr.MaxRow;

    public const double CharPx = 7;

    public static double WidthPx(double w) => Math.Round(w * CharPx);
    public static double HeightPx(double pt) => Math.Round(pt * 96 / 72);

    public void Rebuild(Sheet s)
    {
        double x = 0;
        for (int c = 1; c <= Addr.MaxCol + 1; c++)
        {
            _colLeft[c] = x;
            if (c <= Addr.MaxCol) x += WidthPx(s.ColWidth(c));
        }
        _def = HeightPx(s.DefaultRowHeight);
        var keys = new List<int>();
        foreach (var kv in s.Rows)
        {
            double h = kv.Value.Hidden ? 0 : kv.Value.Height is { } ht ? HeightPx(ht) : _def;
            if (h != _def) keys.Add(kv.Key);
        }
        keys.Sort();
        _custom = [.. keys];
        _cum = new double[keys.Count];
        double acc = 0;
        for (int i = 0; i < keys.Count; i++) { acc += RowHRaw(s, keys[i]) - _def; _cum[i] = acc; }
        _s = s;
    }

    Sheet _s = null!;
    double RowHRaw(Sheet s, int r) => s.Rows.TryGetValue(r, out var ri) ? (ri.Hidden ? 0 : ri.Height is { } h ? HeightPx(h) : _def) : _def;

    public double ColLeft(int c) => _colLeft[Math.Clamp(c, 1, Addr.MaxCol + 1)];
    public double ColW(int c) => _colLeft[c + 1] - _colLeft[c];
    public double TotalW => _colLeft[Addr.MaxCol + 1];

    public double RowTop(int r)
    {
        int n = LowerBound(r); // quante righe personalizzate < r
        return (r - 1) * _def + (n > 0 ? _cum[n - 1] : 0);
    }

    public double RowH(int r) => _s == null ? _def : RowHRaw(_s, r);
    public double TotalH => RowTop(_maxRow + 1);

    int LowerBound(int r)
    {
        int lo = 0, hi = _custom.Length;
        while (lo < hi) { int mid = (lo + hi) / 2; if (_custom[mid] < r) lo = mid + 1; else hi = mid; }
        return lo;
    }

    public int RowAt(double y)
    {
        if (y <= 0) return 1;
        int lo = 1, hi = _maxRow;
        while (lo < hi)
        {
            int mid = (int)(((long)lo + hi + 1) / 2);
            if (RowTop(mid) <= y) lo = mid; else hi = mid - 1;
        }
        return lo;
    }

    public int ColAt(double x)
    {
        if (x <= 0) return 1;
        int lo = 1, hi = Addr.MaxCol;
        while (lo < hi) { int mid = (lo + hi + 1) / 2; if (_colLeft[mid] <= x) lo = mid; else hi = mid - 1; }
        return lo;
    }
}
