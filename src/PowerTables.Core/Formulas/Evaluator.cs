using System.Globalization;
using System.Text;

namespace PowerTables.Core.Formulas;

public sealed class ExcelEx(CellError e) : Exception(Errors.Text(e))
{
    public readonly CellError Err = e;
}

/// <summary>Funzione non implementata: la cella conserva la formula e il valore in cache.</summary>
public sealed class UnsupportedEx(string name) : Exception("funzione non supportata: " + name)
{
    public readonly string FnName = name;
}

public sealed class Missing { public static readonly Missing I = new(); }

public sealed class EvalCtx(Workbook wb, Sheet sheet, int row, int col)
{
    public readonly Workbook Wb = wb;
    public readonly Sheet Sheet = sheet;
    public readonly int Row = row, Col = col;
    public int ArrayMode;
    public int NameDepth;
}

/// <summary>Coercizioni di tipo secondo le regole di Excel.</summary>
public static class V
{
    public static bool IsErr(object? v) => v is CellError;

    public static bool TryNum(object? v, out double d)
    {
        switch (v)
        {
            case double x: d = x; return true;
            case bool b: d = b ? 1 : 0; return true;
            case null: case Missing: d = 0; return true;
            case string s: return TryParseNum(s, out d);
            default: d = 0; return false;
        }
    }

    public static bool TryParseNum(string s, out double d)
    {
        s = s.Trim();
        if (s.Length == 0) { d = 0; return false; }
        bool pct = s.EndsWith('%');
        if (pct) s = s[..^1];
        if (double.TryParse(s, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out d) && !double.IsInfinity(d) && !s.Contains(',')) { if (pct) d /= 100; return true; }
        return false;
    }

    public static double Num(object? v)
    {
        if (v is CellError e) throw new ExcelEx(e);
        if (v is RangeRef or Arr) v = Scalar0(v);
        if (v is CellError e2) throw new ExcelEx(e2);
        if (TryNum(v, out var d)) return d;
        throw new ExcelEx(CellError.Value);
    }

    public static string Str(object? v)
    {
        if (v is RangeRef or Arr) v = Scalar0(v);
        return v switch
        {
            null or Missing => "",
            string s => s,
            double d => Fmt.NumToText(d),
            bool b => b ? "TRUE" : "FALSE",
            CellError e => throw new ExcelEx(e),
            _ => v.ToString() ?? "",
        };
    }

    public static bool Bool(object? v)
    {
        if (v is RangeRef or Arr) v = Scalar0(v);
        switch (v)
        {
            case bool b: return b;
            case double d: return d != 0;
            case null or Missing: return false;
            case string s:
                if (s.Equals("TRUE", StringComparison.OrdinalIgnoreCase)) return true;
                if (s.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) return false;
                throw new ExcelEx(CellError.Value);
            case CellError e: throw new ExcelEx(e);
            default: throw new ExcelEx(CellError.Value);
        }
    }

    /// <summary>Primo elemento (senza intersezione): per valori che devono essere scalari.</summary>
    static object? Scalar0(object? v) => v switch
    {
        RangeRef r => r.At(0, 0),
        Arr a => a.V[0, 0],
        _ => v,
    };

    /// <summary>Intersezione implicita di Excel legacy.</summary>
    public static object? Scalar(object? v, EvalCtx cx)
    {
        switch (v)
        {
            case RangeRef r:
                if (r.IsSingle) return r.At(0, 0);
                if (r.Cols == 1 && cx.Row >= r.R1 && cx.Row <= r.R2 && ReferenceEquals(r.Sheet, cx.Sheet)) return r.Sheet.GetValue(cx.Row, r.C1);
                if (r.Rows == 1 && cx.Col >= r.C1 && cx.Col <= r.C2 && ReferenceEquals(r.Sheet, cx.Sheet)) return r.Sheet.GetValue(r.R1, cx.Col);
                return CellError.Value;
            case Arr a: return a.Rows == 1 && a.Cols == 1 ? a.V[0, 0] : a.V[0, 0];
            case Missing: return null;
            default: return v;
        }
    }

    public static object?[,] Matrix(object? v)
    {
        switch (v)
        {
            case Arr a: return a.V;
            case RangeRef r:
                {
                    int rows = Math.Min(r.R2, Math.Max(r.Sheet.MaxRow, r.R1)) - r.R1 + 1;
                    int cols = Math.Min(r.C2, Math.Max(r.Sheet.MaxCol, r.C1)) - r.C1 + 1;
                    var m = new object?[rows, cols];
                    for (int i = 0; i < rows; i++) for (int j = 0; j < cols; j++) m[i, j] = r.At(i, j);
                    return m;
                }
            default: return new object?[1, 1] { { v is Missing ? null : v } };
        }
    }
}

public static class Evaluator
{
    /// <summary>Valuta una formula (senza '=') nel contesto di una cella e restituisce un valore scalare.</summary>
    public static object? EvalFormula(string formula, EvalCtx cx)
    {
        Node n;
        try { n = Parser.Parse(formula); }
        catch (FormatException) { throw new UnsupportedEx("sintassi non supportata (es. riferimenti strutturati di tabella)"); }
        try { return Finish(Eval(n, cx), cx); }
        catch (ExcelEx e) { return e.Err; }
        catch (Exception e) when (e is InvalidCastException or ArgumentException or IndexOutOfRangeException or OverflowException or InvalidOperationException or FormatException or NullReferenceException) { return CellError.Value; }
    }

    public static object? Finish(object? v, EvalCtx cx)
    {
        v = V.Scalar(v, cx);
        return v is Missing ? null : v;
    }

    public static object? Eval(Node n, EvalCtx cx)
    {
        switch (n)
        {
            case NumNode x: return x.V;
            case StrNode x: return x.V;
            case BoolNode x: return x.V;
            case ErrNode x: return x.V;
            case MissingNode: return Missing.I;
            case RefNode x: return ResolveRef(x.R, cx);
            case NameNode x: return ResolveName(x.Name, cx);
            case UnaryNode u:
                {
                    var v = Operand(Eval(u.X, cx), cx);
                    return Map1(v, cx, e => u.Op == '-' ? Ops.Neg(e) : Ops.Plus(e));
                }
            case PercentNode p:
                return Map1(Operand(Eval(p.X, cx), cx), cx, e => Ops.Arith("/", e, 100.0));
            case BinNode b:
                {
                    var l = Operand(Eval(b.L, cx), cx);
                    var r = Operand(Eval(b.R, cx), cx);
                    return Ops.Binary(b.Op, l, r);
                }
            case CallNode c: return Call(c, cx);
            case ArrayNode a:
                {
                    var m = new object?[a.Rows.Length, a.Rows[0].Length];
                    for (int i = 0; i < a.Rows.Length; i++)
                        for (int j = 0; j < a.Rows[i].Length; j++) m[i, j] = Eval(a.Rows[i][j], cx);
                    return new Arr(m);
                }
            default: throw new InvalidOperationException();
        }
    }

    static object? Operand(object? v, EvalCtx cx)
    {
        if (v is RangeRef r)
            return cx.ArrayMode > 0 ? (r.IsSingle ? r.At(0, 0) : new Arr(V.Matrix(r))) : V.Scalar(v, cx);
        if (v is Arr a) return cx.ArrayMode > 0 ? a : a.V[0, 0];
        return v is Missing ? null : v;
    }

    static object? Map1(object? v, EvalCtx cx, Func<object?, object?> f)
    {
        if (v is Arr a)
        {
            var m = new object?[a.Rows, a.Cols];
            for (int i = 0; i < a.Rows; i++) for (int j = 0; j < a.Cols; j++) m[i, j] = f(a.V[i, j]);
            return new Arr(m);
        }
        return f(v);
    }

    static object? Call(CallNode c, EvalCtx cx)
    {
        if (!Fns.Table.TryGetValue(c.Name, out var fn)) throw new UnsupportedEx(c.Name);
        return fn(cx, c.Args);
    }

    public static Sheet? FindSheet(EvalCtx cx, string? name)
    {
        if (name == null) return cx.Sheet;
        return cx.Wb.FindSheet(name);
    }

    public static RangeRef ResolveRef(RefToken r, EvalCtx cx)
    {
        var sh = FindSheet(cx, r.Sheet) ?? throw new ExcelEx(CellError.Ref);
        int r1 = r.A.Row, c1 = r.A.Col, r2 = r.IsRange ? r.B.Row : r.A.Row, c2 = r.IsRange ? r.B.Col : r.A.Col;
        if (r1 == 0) { r1 = 1; r2 = Addr.MaxRow; }
        if (c1 == 0) { c1 = 1; c2 = Addr.MaxCol; }
        if (r2 < r1) (r1, r2) = (r2, r1);
        if (c2 < c1) (c1, c2) = (c2, c1);
        return new RangeRef(sh, r1, c1, r2, c2);
    }

    static object? ResolveName(string name, EvalCtx cx)
    {
        if (cx.NameDepth > 8) throw new ExcelEx(CellError.Name);
        if (!cx.Wb.TryGetName(name, cx.Sheet, out var text)) throw new ExcelEx(CellError.Name);
        cx.NameDepth++;
        try
        {
            Node n;
            try { n = Parser.Parse(text); } catch (FormatException) { throw new ExcelEx(CellError.Name); }
            return Eval(n, cx);
        }
        finally { cx.NameDepth--; }
    }
}

/// <summary>Operatori binari con le regole di conversione/confronto di Excel.</summary>
public static class Ops
{
    public static object? Neg(object? v) => v is CellError ? v : TryN(v, out var d) ? -d : CellError.Value;
    public static object? Plus(object? v) => v;

    static bool TryN(object? v, out double d)
    {
        if (v is string s) return V.TryParseNum(s, out d);
        return V.TryNum(v, out d);
    }

    public static object? Binary(string op, object? l, object? r)
    {
        if (l is Arr || r is Arr) return Broadcast(op, l, r);
        return Scalar(op, l, r);
    }

    static object? Broadcast(string op, object? l, object? r)
    {
        var a = V.Matrix(l); var b = V.Matrix(r);
        int rows = Math.Max(a.GetLength(0), b.GetLength(0)), cols = Math.Max(a.GetLength(1), b.GetLength(1));
        var m = new object?[rows, cols];
        for (int i = 0; i < rows; i++)
            for (int j = 0; j < cols; j++)
            {
                object? x = Pick(a, i, j), y = Pick(b, i, j);
                m[i, j] = Scalar(op, x, y);
            }
        return new Arr(m);
    }

    static object? Pick(object?[,] m, int i, int j)
    {
        int ri = m.GetLength(0) == 1 ? 0 : i, cj = m.GetLength(1) == 1 ? 0 : j;
        if (ri >= m.GetLength(0) || cj >= m.GetLength(1)) return CellError.NA;
        return m[ri, cj];
    }

    public static object? Scalar(string op, object? l, object? r)
    {
        if (l is CellError le) return le;
        if (r is CellError re) return re;
        switch (op)
        {
            case "&": return V.Str(l) + V.Str(r);
            case "=" or "<>" or "<" or ">" or "<=" or ">=": return Compare(op, l, r);
            default: return Arith(op, l, r);
        }
    }

    public static object? Arith(string op, object? l, object? r)
    {
        if (l is CellError le) return le;
        if (r is CellError re) return re;
        if (!TryN(l, out var a) || !TryN(r, out var b)) return CellError.Value;
        double v;
        switch (op)
        {
            case "+": v = a + b; break;
            case "-": v = a - b; break;
            case "*": v = a * b; break;
            case "/": if (b == 0) return CellError.Div0; v = a / b; break;
            case "^":
                if (a == 0 && b == 0) return CellError.Num;
                if (a == 0 && b < 0) return CellError.Div0;
                if (a < 0 && b != Math.Floor(b)) return CellError.Num;
                v = Math.Pow(a, b); break;
            default: return CellError.Value;
        }
        return double.IsNaN(v) || double.IsInfinity(v) ? CellError.Num : v;
    }

    static int TypeRank(object? v) => v switch { string => 1, bool => 2, _ => 0 };

    public static object? Compare(string op, object? l, object? r)
    {
        if (l == null || l is Missing) l = r switch { string => "", bool => false, _ => 0.0 };
        if (r == null || r is Missing) r = l switch { string => "", bool => false, _ => 0.0 };
        int c;
        int tl = TypeRank(l), tr = TypeRank(r);
        if (tl != tr) c = tl.CompareTo(tr);
        else if (l is double a && r is double b) c = a.CompareTo(b);
        else if (l is string s1 && r is string s2) c = string.Compare(s1, s2, StringComparison.OrdinalIgnoreCase);
        else c = ((bool)l!).CompareTo((bool)r!);
        return op switch { "=" => c == 0, "<>" => c != 0, "<" => c < 0, ">" => c > 0, "<=" => c <= 0, _ => c >= 0 };
    }
}
