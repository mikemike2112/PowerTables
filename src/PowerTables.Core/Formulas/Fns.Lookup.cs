using System.Text;

namespace PowerTables.Core.Formulas;

public static partial class Fns
{
    static int TypeRank(object? v) => v switch { double => 0, string => 1, bool => 2, _ => 3 };

    static int CmpLookup(object? a, object? b)
    {
        int ra = TypeRank(a), rb = TypeRank(b);
        if (ra != rb) return ra.CompareTo(rb);
        return a switch
        {
            double x => x.CompareTo((double)b!),
            string s => string.Compare(s, (string)b!, StringComparison.OrdinalIgnoreCase),
            bool p => p.CompareTo((bool)b!),
            _ => 0,
        };
    }

    /// <summary>Indice (0-based) nel vettore, secondo match type: 0 esatto (con jolly), 1 ≤ ordinato crescente, -1 ≥ ordinato decrescente. -1 se non trovato.</summary>
    static int MatchIn(object? x, IReadOnlyList<object?> vec, int type)
    {
        if (type == 0)
        {
            Criteria? cr = x is string s && (s.Contains('*') || s.Contains('?') || s.Contains('~')) ? new Criteria("=" + s) : null;
            for (int i = 0; i < vec.Count; i++)
            {
                var v = vec[i];
                if (cr != null) { if (v is string && cr.Match(v)) return i; continue; }
                if (v == null) continue;
                if (TypeRank(v) == TypeRank(x) && CmpLookup(v, x) == 0) return i;
            }
            return -1;
        }
        // ricerca binaria sui soli elementi dello stesso tipo
        var idx = new List<int>();
        for (int i = 0; i < vec.Count; i++) if (vec[i] != null && TypeRank(vec[i]) == TypeRank(x)) idx.Add(i);
        int lo = 0, hi = idx.Count - 1, res = -1;
        while (lo <= hi)
        {
            int mid = (lo + hi) / 2;
            int c = CmpLookup(vec[idx[mid]], x);
            if (type == 1 ? c <= 0 : c >= 0) { res = idx[mid]; lo = mid + 1; } else hi = mid - 1;
        }
        return res;
    }

    static List<object?> VecOf(object? a, bool byRow)
    {
        var m = V.Matrix(a); var l = new List<object?>();
        if (byRow) for (int j = 0; j < m.GetLength(1); j++) l.Add(m[0, j]);
        else for (int i = 0; i < m.GetLength(0); i++) l.Add(m[i, 0]);
        return l;
    }

    static object? Cellv(object? tbl, int i, int j) => tbl switch
    {
        RangeRef r => i < r.Rows && j < r.Cols ? r.At(i, j) : CellError.Ref,
        Arr a => i < a.Rows && j < a.Cols ? a.V[i, j] : CellError.Ref,
        _ => tbl,
    };

    static object? Pass(object? v) => v;

    static void RegisterLookup()
    {
        R("MATCH", 2, 3, (cx, a) =>
        {
            var x = V.Scalar(a[0], cx); if (x is CellError e) return e;
            int t = a.Length > 2 ? Math.Sign(V.Num(V.Scalar(a[2], cx))) : 1;
            var m = V.Matrix(a[1]);
            if (m.GetLength(0) != 1 && m.GetLength(1) != 1) return CellError.NA;
            int i = MatchIn(x, VecOf(a[1], m.GetLength(0) == 1), t);
            return i < 0 ? CellError.NA : i + 1.0;
        });
        R("XMATCH", 2, 4, (cx, a) =>
        {
            var x = V.Scalar(a[0], cx); if (x is CellError e) return e;
            int mode = a.Length > 2 && a[2] is not Missing ? (int)V.Num(V.Scalar(a[2], cx)) : 0;
            var m = V.Matrix(a[1]);
            int i = XMatchIdx(x, VecOf(a[1], m.GetLength(0) == 1), mode, a.Length > 3 && a[3] is not Missing ? (int)V.Num(V.Scalar(a[3], cx)) : 1);
            return i < 0 ? CellError.NA : i + 1.0;
        });
        R("VLOOKUP", 3, 4, (cx, a) => VH(cx, a, false));
        R("HLOOKUP", 3, 4, (cx, a) => VH(cx, a, true));
        R("LOOKUP", 2, 3, (cx, a) =>
        {
            var x = V.Scalar(a[0], cx); if (x is CellError e) return e;
            var m = V.Matrix(a[1]);
            bool byRow = m.GetLength(0) == 1 || (a.Length < 3 && m.GetLength(1) > m.GetLength(0));
            var vec = a.Length > 2 ? VecOf(a[1], m.GetLength(0) == 1) : VecOf(a[1], byRow);
            int i = MatchIn(x, vec, 1);
            if (i < 0) return CellError.NA;
            if (a.Length > 2) { var rv = V.Matrix(a[2]); return rv.GetLength(0) == 1 ? rv[0, i] : rv[i, 0]; }
            return byRow ? m[m.GetLength(0) - 1, i] : m[i, m.GetLength(1) - 1];
        });
        R("XLOOKUP", 3, 6, (cx, a) =>
        {
            var x = V.Scalar(a[0], cx); if (x is CellError e) return e;
            var look = V.Matrix(a[1]);
            bool byRow = look.GetLength(0) == 1;
            if (!byRow && look.GetLength(1) != 1) return CellError.Value;
            int mode = a.Length > 4 && a[4] is not Missing ? (int)V.Num(V.Scalar(a[4], cx)) : 0;
            int sm = a.Length > 5 && a[5] is not Missing ? (int)V.Num(V.Scalar(a[5], cx)) : 1;
            int i = XMatchIdx(x, VecOf(a[1], byRow), mode, sm);
            if (i < 0) return a.Length > 3 && a[3] is not Missing ? a[3] : CellError.NA;
            if (a[2] is RangeRef rr)
                return byRow
                    ? (rr.Rows == 1 ? new RangeRef(rr.Sheet, rr.R1, rr.C1 + i, rr.R1, rr.C1 + i) : new RangeRef(rr.Sheet, rr.R1, rr.C1 + i, rr.R2, rr.C1 + i))
                    : (rr.Cols == 1 ? new RangeRef(rr.Sheet, rr.R1 + i, rr.C1, rr.R1 + i, rr.C1) : new RangeRef(rr.Sheet, rr.R1 + i, rr.C1, rr.R1 + i, rr.C2));
            var rm = V.Matrix(a[2]);
            return byRow ? rm[0, i] : rm[i, 0];
        });
        R("INDEX", 2, 4, (cx, a) =>
        {
            int rn = (int)V.Num(V.Scalar(a[1], cx)), cn = a.Length > 2 && a[2] is not Missing ? (int)V.Num(V.Scalar(a[2], cx)) : 0;
            if (a[0] is RangeRef r)
            {
                if (a.Length == 2 || a[2] is Missing) { if (r.Rows == 1 && a.Length == 2) { cn = rn; rn = r.Rows == 1 ? 1 : rn; } }
                if (rn < 0 || cn < 0 || rn > r.Rows || cn > r.Cols) return CellError.Ref;
                int r1 = rn == 0 ? r.R1 : r.R1 + rn - 1, r2 = rn == 0 ? r.R2 : r1, c1 = cn == 0 ? r.C1 : r.C1 + cn - 1, c2 = cn == 0 ? r.C2 : c1;
                return new RangeRef(r.Sheet, r1, c1, r2, c2);
            }
            var m = V.Matrix(a[0]);
            if (m.GetLength(0) == 1 && a.Length == 2) { cn = rn; rn = 1; }
            if (rn < 0 || cn < 0 || rn > m.GetLength(0) || cn > m.GetLength(1)) return CellError.Ref;
            if (rn == 0 || cn == 0)
            {
                int rows = rn == 0 ? m.GetLength(0) : 1, cols = cn == 0 ? m.GetLength(1) : 1;
                var res = new object?[rows, cols];
                for (int i = 0; i < rows; i++) for (int j = 0; j < cols; j++) res[i, j] = m[rn == 0 ? i : rn - 1, cn == 0 ? j : cn - 1];
                return new Arr(res);
            }
            return m[rn - 1, cn - 1];
        });
        R("OFFSET", 3, 5, (cx, a) =>
        {
            if (a[0] is not RangeRef r) return CellError.Value;
            int dr = (int)V.Num(V.Scalar(a[1], cx)), dc = (int)V.Num(V.Scalar(a[2], cx));
            int h = a.Length > 3 && a[3] is not Missing ? (int)V.Num(V.Scalar(a[3], cx)) : r.Rows, w = a.Length > 4 && a[4] is not Missing ? (int)V.Num(V.Scalar(a[4], cx)) : r.Cols;
            int r1 = r.R1 + dr, c1 = r.C1 + dc;
            if (h < 1 || w < 1 || r1 < 1 || c1 < 1 || r1 + h - 1 > Addr.MaxRow || c1 + w - 1 > Addr.MaxCol) return CellError.Ref;
            return new RangeRef(r.Sheet, r1, c1, r1 + h - 1, c1 + w - 1);
        });
        R("INDIRECT", 1, 2, (cx, a) =>
        {
            string t = St(V.Scalar(a[0], cx)).Trim();
            bool a1 = a.Length < 2 || V.Bool(V.Scalar(a[1], cx));
            if (!a1) return CellError.Ref;
            try
            {
                var toks = Lexer.Tokenize(t);
                if (toks.Count != 2 || toks[0].Kind != T.Ref) return CellError.Ref;
                return Evaluator.ResolveRef(toks[0].Ref!, cx);
            }
            catch (FormatException) { return CellError.Ref; }
        });
        R("ADDRESS", 2, 5, (cx, a) =>
        {
            int row = (int)V.Num(V.Scalar(a[0], cx)), col = (int)V.Num(V.Scalar(a[1], cx));
            int abs = a.Length > 2 && a[2] is not Missing ? (int)V.Num(V.Scalar(a[2], cx)) : 1;
            if (row < 1 || col < 1 || col > Addr.MaxCol || row > Addr.MaxRow || abs is < 1 or > 4) return CellError.Value;
            bool a1 = a.Length < 4 || a[3] is Missing || V.Bool(V.Scalar(a[3], cx));
            string s;
            if (a1) s = ((abs == 1 || abs == 3) ? "" : "") + (abs is 1 or 3 ? "$" : "") + Addr.ColName(col) + (abs is 1 or 2 ? "$" : "") + row;
            else s = "R" + (abs is 1 or 2 ? row.ToString() : $"[{row}]") + "C" + (abs is 1 or 3 ? col.ToString() : $"[{col}]");
            if (a.Length > 4 && a[4] is not Missing) s = Lexer.QuoteSheet(St(V.Scalar(a[4], cx))) + "!" + s;
            return s;
        });
        R("TRANSPOSE", 1, 1, (cx, a) =>
        {
            var m = V.Matrix(a[0]); var t = new object?[m.GetLength(1), m.GetLength(0)];
            for (int i = 0; i < m.GetLength(0); i++) for (int j = 0; j < m.GetLength(1); j++) t[j, i] = m[i, j];
            return new Arr(t);
        });
    }

    static int XMatchIdx(object? x, List<object?> vec, int mode, int search)
    {
        IEnumerable<int> order = search >= 0 ? Enumerable.Range(0, vec.Count) : Enumerable.Range(0, vec.Count).Reverse();
        if (mode == 2)
        {
            if (x is not string s) return MatchIn(x, vec, 0);
            var cr = new Criteria("=" + s);
            foreach (var i in order) if (vec[i] is string && cr.Match(vec[i])) return i;
            return -1;
        }
        int best = -1;
        foreach (var i in order)
        {
            var v = vec[i];
            if (v == null || TypeRank(v) != TypeRank(x)) continue;
            int c = CmpLookup(v, x);
            if (c == 0) return i;
            if (mode == -1 && c < 0 && (best < 0 || CmpLookup(v, vec[best]) > 0)) best = i;
            if (mode == 1 && c > 0 && (best < 0 || CmpLookup(v, vec[best]) < 0)) best = i;
        }
        return best;
    }

    static object? VH(EvalCtx cx, object?[] a, bool horizontal)
    {
        var x = V.Scalar(a[0], cx); if (x is CellError e) return e;
        int idx = (int)V.Num(V.Scalar(a[2], cx));
        bool approx = a.Length < 4 || a[3] is Missing || V.Bool(V.Scalar(a[3], cx));
        int rows, cols;
        Func<int, int, object?> at;
        if (a[1] is RangeRef r) { rows = r.Rows; cols = r.Cols; at = r.At; }
        else { var m = V.Matrix(a[1]); rows = m.GetLength(0); cols = m.GetLength(1); at = (i, j) => m[i, j]; }
        if (idx < 1) return CellError.Value;
        if (idx > (horizontal ? rows : cols)) return CellError.Ref;
        int n = horizontal ? cols : rows;
        if (a[1] is RangeRef rr) n = horizontal ? Math.Min(cols, Math.Max(rr.Sheet.MaxCol - rr.C1 + 1, 0)) : Math.Min(rows, Math.Max(rr.Sheet.MaxRow - rr.R1 + 1, 0));
        var vec = new List<object?>(n);
        for (int k = 0; k < n; k++) vec.Add(horizontal ? at(0, k) : at(k, 0));
        int i = MatchIn(x, vec, approx ? 1 : 0);
        if (i < 0) return CellError.NA;
        return horizontal ? at(idx - 1, i) : at(i, idx - 1);
    }
}
