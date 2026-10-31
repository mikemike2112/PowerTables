namespace PowerTables.Core.Formulas;

public delegate object? Fn(EvalCtx cx, Node[] a);

/// <summary>Registro delle funzioni: logiche, informative e infrastruttura di registrazione.</summary>
public static partial class Fns
{
    internal static readonly Dictionary<string, Fn> Table = new(StringComparer.Ordinal);
    public static IEnumerable<string> Names => Table.Keys;
    public static bool Exists(string n) => Table.ContainsKey(n);

    static Fns()
    {
        RegisterCore(); RegisterMath(); RegisterStats(); RegisterText(); RegisterLookup(); RegisterDate(); RegisterFinance();
    }

    // ---- helper di registrazione ----
    static void Check(int n, int min, int max) { if (n < min || n > max) throw new ExcelEx(CellError.Value); }

    static object? Ev(EvalCtx cx, Node n)
    {
        try { return Evaluator.Eval(n, cx); } catch (ExcelEx e) { return e.Err; }
    }

    static object? Sc(EvalCtx cx, Node n) => V.Scalar(Evaluator.Eval(n, cx), cx);

    /// <summary>Funzione con argomenti scalari (intersezione implicita), propaga gli errori.</summary>
    static void S(string name, int min, int max, Func<object?[], object?> f, bool lift = false)
    {
        Table[name] = (cx, a) =>
        {
            Check(a.Length, min, max);
            var vals = new object?[a.Length];
            bool arrays = false;
            for (int i = 0; i < a.Length; i++)
            {
                vals[i] = a[i] is MissingNode ? Missing.I : Evaluator.Eval(a[i], cx);
                if (vals[i] is Arr || vals[i] is RangeRef { IsSingle: false }) arrays = true;
            }
            if (lift && arrays && cx.ArrayMode > 0) return LiftCall(vals, f);
            for (int i = 0; i < vals.Length; i++)
            {
                if (vals[i] is Missing) continue;
                var v = V.Scalar(vals[i], cx);
                if (v is CellError e) return e;
                vals[i] = v;
            }
            return f(vals);
        };
    }

    static object? LiftCall(object?[] vals, Func<object?[], object?> f)
    {
        var ms = vals.Select(V.Matrix).ToArray();
        int rows = ms.Max(m => m.GetLength(0)), cols = ms.Max(m => m.GetLength(1));
        var res = new object?[rows, cols];
        for (int i = 0; i < rows; i++)
            for (int j = 0; j < cols; j++)
            {
                var args = new object?[vals.Length]; object? err = null;
                for (int k = 0; k < ms.Length; k++)
                {
                    var m = ms[k];
                    int ri = m.GetLength(0) == 1 ? 0 : i, cj = m.GetLength(1) == 1 ? 0 : j;
                    var v = ri < m.GetLength(0) && cj < m.GetLength(1) ? m[ri, cj] : CellError.NA;
                    if (v is CellError) err ??= v;
                    args[k] = vals[k] is Missing ? vals[k] : v;
                }
                try { res[i, j] = err ?? f(args); } catch (ExcelEx e) { res[i, j] = e.Err; }
            }
        return new Arr(res);
    }

    /// <summary>Funzione che riceve gli argomenti grezzi (intervalli inclusi).</summary>
    static void R(string name, int min, int max, Func<EvalCtx, object?[], object?> f)
    {
        Table[name] = (cx, a) =>
        {
            Check(a.Length, min, max);
            var vals = new object?[a.Length];
            for (int i = 0; i < a.Length; i++) vals[i] = a[i] is MissingNode ? Missing.I : Evaluator.Eval(a[i], cx);
            return f(cx, vals);
        };
    }

    static void L(string name, int min, int max, Fn f)
    {
        Table[name] = (cx, a) => { Check(a.Length, min, max); return f(cx, a); };
    }

    static bool Has(Node[] a, int i) => a.Length > i && a[i] is not MissingNode;

    // ---- funzioni logiche e informative ----
    static void RegisterCore()
    {
        L("IF", 2, 3, (cx, a) =>
        {
            bool c = V.Bool(Sc(cx, a[0]));
            if (c) return a[1] is MissingNode ? 0.0 : Evaluator.Eval(a[1], cx);
            if (a.Length < 3) return false;
            return a[2] is MissingNode ? 0.0 : Evaluator.Eval(a[2], cx);
        });
        L("IFS", 2, 254, (cx, a) =>
        {
            if (a.Length % 2 != 0) throw new ExcelEx(CellError.Value);
            for (int i = 0; i < a.Length; i += 2) if (V.Bool(Sc(cx, a[i]))) return Evaluator.Eval(a[i + 1], cx);
            return CellError.NA;
        });
        L("SWITCH", 3, 254, (cx, a) =>
        {
            var x = Sc(cx, a[0]);
            if (x is CellError xe) return xe;
            int i = 1;
            for (; i + 1 < a.Length; i += 2)
            {
                var c = Sc(cx, a[i]);
                if (c is CellError ce) return ce;
                if (Ops.Compare("=", x, c) is true) return Evaluator.Eval(a[i + 1], cx);
            }
            return i < a.Length ? Evaluator.Eval(a[i], cx) : CellError.NA;
        });
        L("CHOOSE", 2, 255, (cx, a) =>
        {
            int i = (int)Math.Floor(V.Num(Sc(cx, a[0])));
            if (i < 1 || i >= a.Length) return CellError.Value;
            return Evaluator.Eval(a[i], cx);
        });
        L("IFERROR", 2, 2, (cx, a) =>
        {
            var v = Ev(cx, a[0]);
            return V.Scalar(v, cx) is CellError ? Evaluator.Eval(a[1], cx) : v;
        });
        L("IFNA", 2, 2, (cx, a) =>
        {
            var v = Ev(cx, a[0]);
            return V.Scalar(v, cx) is CellError.NA ? Evaluator.Eval(a[1], cx) : v;
        });
        R("AND", 1, 255, (cx, a) => Logical(a, true, (x, y) => x && y));
        R("OR", 1, 255, (cx, a) => Logical(a, false, (x, y) => x || y));
        R("XOR", 1, 255, (cx, a) => Logical(a, false, (x, y) => x ^ y));
        S("NOT", 1, 1, a => !V.Bool(a[0]), true);
        L("TRUE", 0, 0, (_, _) => true);
        L("FALSE", 0, 0, (_, _) => false);

        // informative: ricevono errori come valori
        foreach (var (name, pred) in new (string, Func<object?, bool>)[]
        {
            ("ISNUMBER", v => v is double), ("ISTEXT", v => v is string), ("ISLOGICAL", v => v is bool),
            ("ISERROR", v => v is CellError), ("ISERR", v => v is CellError e && e != CellError.NA), ("ISNA", v => v is CellError.NA),
            ("ISBLANK", v => v == null), ("ISNONTEXT", v => v is not string),
        })
        {
            var p = pred;
            L(name, 1, 1, (cx, a) =>
            {
                var v = Ev(cx, a[0]);
                if (cx.ArrayMode > 0 && (v is Arr || v is RangeRef { IsSingle: false }))
                {
                    var m = V.Matrix(v); var r = new object?[m.GetLength(0), m.GetLength(1)];
                    for (int i = 0; i < r.GetLength(0); i++) for (int j = 0; j < r.GetLength(1); j++) r[i, j] = p(m[i, j]);
                    return new Arr(r);
                }
                return p(V.Scalar(v, cx));
            });
        }
        S("ISEVEN", 1, 1, a => Math.Floor(Math.Abs(V.Num(a[0]))) % 2 == 0, true);
        S("ISODD", 1, 1, a => Math.Floor(Math.Abs(V.Num(a[0]))) % 2 == 1, true);
        L("NA", 0, 0, (_, _) => CellError.NA);
        L("N", 1, 1, (cx, a) => Sc(cx, a[0]) switch { double d => d, bool b => b ? 1.0 : 0.0, CellError e => e, _ => 0.0 });
        L("T", 1, 1, (cx, a) => Sc(cx, a[0]) switch { string s => s, CellError e => e, _ => "" });
        L("TYPE", 1, 1, (cx, a) =>
        {
            var raw = Ev(cx, a[0]);
            if (raw is Arr { Rows: > 1 } || raw is Arr { Cols: > 1 }) return 64.0;
            return V.Scalar(raw, cx) switch { double => 1.0, string => 2.0, bool => 4.0, CellError => 16.0, null => 1.0, _ => 1.0 };
        });
        L("ERROR.TYPE", 1, 1, (cx, a) => Sc(cx, a[0]) is CellError e && (int)e <= 7 ? (double)(int)e : CellError.NA);
        L("ROW", 0, 1, (cx, a) => a.Length == 0 ? cx.Row : Evaluator.Eval(a[0], cx) is RangeRef r ? r.R1 : throw new ExcelEx(CellError.Value));
        L("COLUMN", 0, 1, (cx, a) => a.Length == 0 ? cx.Col : Evaluator.Eval(a[0], cx) is RangeRef r ? r.C1 : throw new ExcelEx(CellError.Value));
        L("ROWS", 1, 1, (cx, a) => Evaluator.Eval(a[0], cx) switch { RangeRef r => (double)r.Rows, Arr x => x.Rows, _ => 1.0 });
        L("COLUMNS", 1, 1, (cx, a) => Evaluator.Eval(a[0], cx) switch { RangeRef r => (double)r.Cols, Arr x => x.Cols, _ => 1.0 });
        L("SHEET", 0, 1, (cx, a) => a.Length == 0 ? cx.Sheet.Index + 1.0 : Evaluator.Eval(a[0], cx) is RangeRef r ? r.Sheet.Index + 1.0 : CellError.NA);
        L("SHEETS", 0, 1, (cx, a) => (double)cx.Wb.Sheets.Count);
        L("ISREF", 1, 1, (cx, a) => Ev(cx, a[0]) is RangeRef);
        L("ISFORMULA", 1, 1, (cx, a) => Evaluator.Eval(a[0], cx) is RangeRef r ? r.Sheet.Get(r.R1, r.C1).F != null : CellError.Value);
        L("FORMULATEXT", 1, 1, (cx, a) =>
        {
            if (Evaluator.Eval(a[0], cx) is not RangeRef r) return CellError.NA;
            var f = r.Sheet.Get(r.R1, r.C1).F;
            return f == null ? CellError.NA : "=" + FormulaText.ToDisplay(f.Text);
        });
    }

    static object? Logical(object?[] a, bool seed, Func<bool, bool, bool> op)
    {
        bool acc = seed, any = false;
        foreach (var x in a)
        {
            if (x is RangeRef or Arr)
            {
                foreach (var v in Flat(x))
                {
                    if (v is CellError e) return e;
                    if (v is bool b) { acc = any ? op(acc, b) : b; any = true; }
                    else if (v is double d) { acc = any ? op(acc, d != 0) : d != 0; any = true; }
                }
            }
            else
            {
                if (x is CellError e) return e;
                if (x is Missing) continue;
                bool b = x is string s ? (s.Equals("TRUE", StringComparison.OrdinalIgnoreCase) ? true : s.Equals("FALSE", StringComparison.OrdinalIgnoreCase) ? false : throw new ExcelEx(CellError.Value)) : V.Bool(x);
                acc = any ? op(acc, b) : b; any = true;
            }
        }
        return any ? acc : CellError.Value;
    }

    /// <summary>Scorre tutti i valori di un intervallo/matrice (righe poi colonne), limitato all'area usata.</summary>
    internal static IEnumerable<object?> Flat(object? x)
    {
        switch (x)
        {
            case RangeRef r:
                {
                    int rows = Math.Min(r.R2, Math.Max(r.Sheet.MaxRow, r.R1 - 1)) - r.R1 + 1, cols = Math.Min(r.C2, Math.Max(r.Sheet.MaxCol, r.C1 - 1)) - r.C1 + 1;
                    for (int i = 0; i < rows; i++) for (int j = 0; j < cols; j++) yield return r.At(i, j);
                    break;
                }
            case Arr a:
                for (int i = 0; i < a.Rows; i++) for (int j = 0; j < a.Cols; j++) yield return a.V[i, j];
                break;
            default: yield return x; break;
        }
    }

    /// <summary>Celle effettivamente presenti in un intervallo (ordine non garantito se l'area è grande).</summary>
    internal static IEnumerable<Cell> Stored(RangeRef r)
    {
        long area = (long)Math.Min(r.Rows, Math.Max(1, r.Sheet.MaxRow)) * Math.Min(r.Cols, Math.Max(1, r.Sheet.MaxCol));
        if (area > 2L * r.Sheet.CellCount)
        {
            foreach (var kv in r.Sheet.Cells)
            {
                int row = Addr.Row(kv.Key), col = Addr.Col(kv.Key);
                if (row >= r.R1 && row <= r.R2 && col >= r.C1 && col <= r.C2) yield return kv.Value;
            }
        }
        else
        {
            int rm = Math.Min(r.R2, r.Sheet.MaxRow), cm = Math.Min(r.C2, r.Sheet.MaxCol);
            for (int i = r.R1; i <= rm; i++)
                for (int j = r.C1; j <= cm; j++)
                    if (r.Sheet.Cells.TryGetValue(Addr.Key(i, j), out var c)) yield return c;
        }
    }
}
