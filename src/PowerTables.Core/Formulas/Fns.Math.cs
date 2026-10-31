using System.Text.RegularExpressions;

namespace PowerTables.Core.Formulas;

/// <summary>Criterio di COUNTIF/SUMIF e simili.</summary>
public sealed class Criteria
{
    readonly string _op = "=";
    readonly object? _val;        // double, bool, string, CellError
    readonly Regex? _rx;

    public Criteria(object? crit)
    {
        if (crit is string s)
        {
            string rest = s;
            foreach (var op in new[] { "<=", ">=", "<>", "<", ">", "=" })
                if (s.StartsWith(op, StringComparison.Ordinal)) { _op = op; rest = s[op.Length..]; break; }
            if (V.TryParseNum(rest, out var d) && rest.Trim().Length > 0) _val = d;
            else if (rest.Equals("TRUE", StringComparison.OrdinalIgnoreCase)) _val = true;
            else if (rest.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) _val = false;
            else if (rest.StartsWith('#') && Errors.TryParse(rest, out var e)) _val = e;
            else
            {
                _val = rest;
                if (rest.Contains('*') || rest.Contains('?') || rest.Contains('~')) _rx = Wildcard(rest);
            }
        }
        else _val = crit is null or Missing ? 0.0 : crit;
    }

    public static Regex Wildcard(string pat, bool anchored = true)
    {
        var sb = new System.Text.StringBuilder(anchored ? "^" : "");
        for (int i = 0; i < pat.Length; i++)
        {
            char c = pat[i];
            if (c == '~' && i + 1 < pat.Length) sb.Append(Regex.Escape(pat[++i].ToString()));
            else if (c == '*') sb.Append("[\\s\\S]*");
            else if (c == '?') sb.Append("[\\s\\S]");
            else sb.Append(Regex.Escape(c.ToString()));
        }
        if (anchored) sb.Append('$');
        return new Regex(sb.ToString(), RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    }

    public bool Match(object? v)
    {
        bool r;
        switch (_val)
        {
            case double d:
                if (v is double x) r = Cmp(x.CompareTo(d));
                else if (v is string s && (_op is "=" or "<>") && V.TryParseNum(s, out var y)) r = Cmp(y.CompareTo(d));
                else r = _op == "<>";
                return r;
            case bool b:
                if (v is bool bv) r = Cmp(bv.CompareTo(b)); else r = _op == "<>";
                return r;
            case CellError e:
                r = v is CellError ve && ve == e; return _op == "<>" ? !r : r;
            case string t:
                if (t.Length == 0 && _op == "=") return v == null || (v is string s0 && s0.Length == 0);
                if (t.Length == 0 && _op == "<>") return v != null && !(v is string s1 && s1.Length == 0);
                if (v is string s2)
                {
                    if (_op is "=" or "<>")
                    {
                        r = _rx != null ? _rx.IsMatch(s2) : string.Equals(s2, t, StringComparison.OrdinalIgnoreCase);
                        return _op == "<>" ? !r : r;
                    }
                    return Cmp(string.Compare(s2, t, StringComparison.OrdinalIgnoreCase));
                }
                return _op == "<>";
            default: return false;
        }
    }

    bool Cmp(int c) => _op switch { "=" => c == 0, "<>" => c != 0, "<" => c < 0, ">" => c > 0, "<=" => c <= 0, _ => c >= 0 };
}

public static partial class Fns
{
    static IEnumerable<double> Nums(object?[] args, bool countBoolText = true)
    {
        foreach (var a in args)
        {
            if (a is RangeRef r)
            {
                foreach (var c in Stored(r))
                {
                    if (c.Kind == CellKind.Num) yield return c.Num;
                    else if (c.Kind == CellKind.Err) throw new ExcelEx((CellError)(int)c.Num);
                }
            }
            else if (a is Arr arr)
            {
                foreach (var v in Flat(arr))
                {
                    if (v is double d) yield return d;
                    else if (v is CellError e) throw new ExcelEx(e);
                }
            }
            else
            {
                if (a is CellError e) throw new ExcelEx(e);
                if (a is Missing) { yield return 0; continue; }
                if (a is string s) { if (V.TryParseNum(s, out var d)) yield return d; else throw new ExcelEx(CellError.Value); }
                else yield return V.Num(a);
            }
        }
    }

    static double RoundDec(double x, int n, MidpointRounding mode)
    {
        try
        {
            if (n >= 0 && n <= 15 && Math.Abs(x) < 7.9e27)
            {
                decimal d = (decimal)x;
                return (double)Math.Round(d, n, mode);
            }
            if (n < 0 && n >= -15 && Math.Abs(x) < 7.9e27)
            {
                decimal p = (decimal)Math.Pow(10, -n);
                return (double)(Math.Round((decimal)x / p, 0, mode) * p);
            }
        }
        catch (OverflowException) { }
        double f = Math.Pow(10, n);
        return Math.Round(x * f, mode) / f;
    }

    static double RoundUpDown(double x, int n, bool up)
    {
        if (!up) return RoundDec(x, n, MidpointRounding.ToZero);
        double t = RoundDec(x, n, MidpointRounding.ToZero);
        if (t == x) return x;
        double step = Math.Pow(10, -n);
        return t + (x > 0 ? step : -step);
    }

    static void RegisterMath()
    {
        R("SUM", 1, 255, (cx, a) => { double s = 0; foreach (var d in Nums(a)) s += d; return Chk(s); });
        R("PRODUCT", 1, 255, (cx, a) => { double s = 1; bool any = false; foreach (var d in Nums(a)) { s *= d; any = true; } return any ? Chk(s) : 0.0; });
        R("COUNT", 1, 255, (cx, a) =>
        {
            double n = 0;
            foreach (var x in a)
            {
                if (x is RangeRef r) { foreach (var c in Stored(r)) if (c.Kind == CellKind.Num) n++; }
                else if (x is Arr ar) { foreach (var v in Flat(ar)) if (v is double) n++; }
                else if (x is double || x is bool || (x is string s && V.TryParseNum(s, out _))) n++;
            }
            return n;
        });
        R("COUNTA", 1, 255, (cx, a) =>
        {
            double n = 0;
            foreach (var x in a)
            {
                if (x is RangeRef r) { foreach (var c in Stored(r)) if (c.Kind != CellKind.Empty) n++; }
                else if (x is Arr ar) { foreach (var v in Flat(ar)) if (v != null) n++; }
                else if (x is not Missing) n++;
            }
            return n;
        });
        R("COUNTBLANK", 1, 1, (cx, a) =>
        {
            if (a[0] is not RangeRef r) return CellError.Value;
            double n = (double)r.Rows * r.Cols;
            foreach (var c in Stored(r)) if (c.Kind != CellKind.Empty && !(c.Kind == CellKind.Str && ((string?)c.Obj)!.Length == 0)) n--;
            return n;
        });
        R("AVERAGE", 1, 255, (cx, a) => { double s = 0; int n = 0; foreach (var d in Nums(a)) { s += d; n++; } return n == 0 ? CellError.Div0 : s / n; });
        R("MAX", 1, 255, (cx, a) => { double m = double.NegativeInfinity; bool any = false; foreach (var d in Nums(a)) { if (d > m) m = d; any = true; } return any ? m : 0.0; });
        R("MIN", 1, 255, (cx, a) => { double m = double.PositiveInfinity; bool any = false; foreach (var d in Nums(a)) { if (d < m) m = d; any = true; } return any ? m : 0.0; });

        S("ABS", 1, 1, a => Math.Abs(V.Num(a[0])), true);
        S("SIGN", 1, 1, a => (double)Math.Sign(V.Num(a[0])), true);
        S("INT", 1, 1, a => Math.Floor(V.Num(a[0])), true);
        S("SQRT", 1, 1, a => { var x = V.Num(a[0]); return x < 0 ? CellError.Num : Math.Sqrt(x); }, true);
        S("EXP", 1, 1, a => Chk(Math.Exp(V.Num(a[0]))), true);
        S("LN", 1, 1, a => { var x = V.Num(a[0]); return x <= 0 ? CellError.Num : Math.Log(x); }, true);
        S("LOG10", 1, 1, a => { var x = V.Num(a[0]); return x <= 0 ? CellError.Num : Math.Log10(x); }, true);
        S("LOG", 1, 2, a =>
        {
            var x = V.Num(a[0]); double b = a.Length > 1 ? V.Num(a[1]) : 10;
            if (x <= 0 || b <= 0) return CellError.Num;
            if (b == 1) return CellError.Div0;
            return Math.Log(x) / Math.Log(b);
        }, true);
        L("PI", 0, 0, (_, _) => Math.PI);
        S("SIN", 1, 1, a => Math.Sin(V.Num(a[0])), true);
        S("COS", 1, 1, a => Math.Cos(V.Num(a[0])), true);
        S("TAN", 1, 1, a => Math.Tan(V.Num(a[0])), true);
        S("ASIN", 1, 1, a => { var x = V.Num(a[0]); return Math.Abs(x) > 1 ? CellError.Num : Math.Asin(x); }, true);
        S("ACOS", 1, 1, a => { var x = V.Num(a[0]); return Math.Abs(x) > 1 ? CellError.Num : Math.Acos(x); }, true);
        S("ATAN", 1, 1, a => Math.Atan(V.Num(a[0])), true);
        S("ATAN2", 2, 2, a => { double x = V.Num(a[0]), y = V.Num(a[1]); return x == 0 && y == 0 ? CellError.Div0 : Math.Atan2(y, x); }, true);
        S("DEGREES", 1, 1, a => V.Num(a[0]) * 180 / Math.PI, true);
        S("RADIANS", 1, 1, a => V.Num(a[0]) * Math.PI / 180, true);
        S("ROUND", 2, 2, a => RoundDec(V.Num(a[0]), (int)Math.Truncate(V.Num(a[1])), MidpointRounding.AwayFromZero), true);
        S("ROUNDUP", 2, 2, a => RoundUpDown(V.Num(a[0]), (int)Math.Truncate(V.Num(a[1])), true), true);
        S("ROUNDDOWN", 2, 2, a => RoundUpDown(V.Num(a[0]), (int)Math.Truncate(V.Num(a[1])), false), true);
        S("TRUNC", 1, 2, a => RoundUpDown(V.Num(a[0]), a.Length > 1 ? (int)Math.Truncate(V.Num(a[1])) : 0, false), true);
        S("MOD", 2, 2, a =>
        {
            double n = V.Num(a[0]), d = V.Num(a[1]);
            if (d == 0) return CellError.Div0;
            double r = n - d * Math.Floor(n / d);
            return Math.Abs(r) >= Math.Abs(d) ? 0.0 : r;
        }, true);
        S("QUOTIENT", 2, 2, a => { double d = V.Num(a[1]); return d == 0 ? CellError.Div0 : Math.Truncate(V.Num(a[0]) / d); }, true);
        S("POWER", 2, 2, a => Ops.Arith("^", V.Num(a[0]), V.Num(a[1])), true);
        S("CEILING", 1, 2, a =>
        {
            double x = V.Num(a[0]), s = a.Length > 1 ? V.Num(a[1]) : 1;
            if (s == 0) return 0.0;
            if (x > 0 && s < 0) return CellError.Num;
            return Math.Ceiling(x / s - 1e-15 * Math.Abs(x / s)) * s;
        }, true);
        S("FLOOR", 1, 2, a =>
        {
            double x = V.Num(a[0]), s = a.Length > 1 ? V.Num(a[1]) : 1;
            if (s == 0) return CellError.Div0;
            if (x > 0 && s < 0) return CellError.Num;
            return Math.Floor(x / s + 1e-15 * Math.Abs(x / s)) * s;
        }, true);
        S("MROUND", 2, 2, a =>
        {
            double x = V.Num(a[0]), m = V.Num(a[1]);
            if (m == 0) return 0.0;
            if (x * m < 0) return CellError.Num;
            return Math.Round(x / m, MidpointRounding.AwayFromZero) * m;
        }, true);
        S("EVEN", 1, 1, a => { double x = V.Num(a[0]); double r = 2 * Math.Ceiling(Math.Abs(x) / 2); return x < 0 ? -r : r; }, true);
        S("ODD", 1, 1, a => { double x = V.Num(a[0]); double r = 2 * Math.Ceiling((Math.Abs(x) - 1) / 2) + 1; return x < 0 ? -r : r; }, true);
        S("FACT", 1, 1, a =>
        {
            double x = Math.Floor(V.Num(a[0]));
            if (x < 0 || x > 170) return CellError.Num;
            double r = 1; for (int i = 2; i <= x; i++) r *= i; return r;
        }, true);
        S("GCD", 1, 255, a => { double g = 0; foreach (var x in a) { double v = Math.Floor(V.Num(x)); if (v < 0) return CellError.Num; g = Gcd(g, v); } return g; });
        S("LCM", 1, 255, a => { double l = 1; foreach (var x in a) { double v = Math.Floor(V.Num(x)); if (v < 0) return CellError.Num; if (v == 0) return 0.0; l = l / Gcd(l, v) * v; } return l; });
        L("RAND", 0, 0, (_, _) => Random.Shared.NextDouble());
        S("RANDBETWEEN", 2, 2, a =>
        {
            double lo = Math.Ceiling(V.Num(a[0])), hi = Math.Floor(V.Num(a[1]));
            return hi < lo ? CellError.Num : lo + Math.Floor(Random.Shared.NextDouble() * (hi - lo + 1));
        });

        // somme/conteggi condizionali
        R("COUNTIF", 2, 2, (cx, a) =>
        {
            if (a[0] is not RangeRef r) return CellError.Value;
            var cr = new Criteria(V.Scalar(a[1], cx));
            double n = 0;
            ForEachGrid(r, (i, j, v) => { if (cr.Match(v)) n++; });
            return n;
        });
        R("COUNTIFS", 2, 254, (cx, a) => CondAggregate(cx, a, 0, null, "COUNT"));
        R("SUMIF", 2, 3, (cx, a) =>
        {
            if (a[0] is not RangeRef r) return CellError.Value;
            var cr = new Criteria(V.Scalar(a[1], cx));
            var sr = a.Length > 2 && a[2] is RangeRef x ? new RangeRef(x.Sheet, x.R1, x.C1, x.R1 + r.Rows - 1, x.C1 + r.Cols - 1) : r;
            double s = 0;
            ForEachGrid(r, (i, j, v) => { if (cr.Match(v) && sr.At(i, j) is double d) s += d; });
            return s;
        });
        R("SUMIFS", 3, 255, (cx, a) => CondAggregate(cx, a, 1, a[0], "SUM"));
        R("AVERAGEIF", 2, 3, (cx, a) =>
        {
            if (a[0] is not RangeRef r) return CellError.Value;
            var cr = new Criteria(V.Scalar(a[1], cx));
            var sr = a.Length > 2 && a[2] is RangeRef x ? new RangeRef(x.Sheet, x.R1, x.C1, x.R1 + r.Rows - 1, x.C1 + r.Cols - 1) : r;
            double s = 0; int n = 0;
            ForEachGrid(r, (i, j, v) => { if (cr.Match(v) && sr.At(i, j) is double d) { s += d; n++; } });
            return n == 0 ? CellError.Div0 : s / n;
        });
        R("AVERAGEIFS", 3, 255, (cx, a) => CondAggregate(cx, a, 1, a[0], "AVERAGE"));
        R("MAXIFS", 3, 255, (cx, a) => CondAggregate(cx, a, 1, a[0], "MAX"));
        R("MINIFS", 3, 255, (cx, a) => CondAggregate(cx, a, 1, a[0], "MIN"));

        Table["SUMPRODUCT"] = (cx, a) =>
        {
            Check(a.Length, 1, 255);
            cx.ArrayMode++;
            try
            {
                var ms = a.Select(n => V.Matrix(Evaluator.Eval(n, cx))).ToArray();
                int rows = ms[0].GetLength(0), cols = ms[0].GetLength(1);
                if (ms.Any(m => m.GetLength(0) != rows || m.GetLength(1) != cols)) return CellError.Value;
                double s = 0;
                for (int i = 0; i < rows; i++)
                    for (int j = 0; j < cols; j++)
                    {
                        double p = 1;
                        foreach (var m in ms)
                        {
                            if (m[i, j] is CellError e) return e;
                            p *= m[i, j] is double d ? d : 0;
                        }
                        s += p;
                    }
                return s;
            }
            finally { cx.ArrayMode--; }
        };
        R("SUBTOTAL", 2, 255, (cx, a) =>
        {
            int code = (int)V.Num(V.Scalar(a[0], cx));
            bool ignoreManual = code > 100;
            if (ignoreManual) code -= 100;
            if (code is < 1 or > 11) return CellError.Value;
            var vals = new List<double>(); int counta = 0; CellError? firstErr = null;
            foreach (var x in a.Skip(1))
            {
                if (x is RangeRef r)
                {
                    int rm = Math.Min(r.R2, r.Sheet.MaxRow), cm = Math.Min(r.C2, r.Sheet.MaxCol);
                    for (int i = r.R1; i <= rm; i++)
                    {
                        if (r.Sheet.Rows.TryGetValue(i, out var ri) && (ri.Filtered || (ignoreManual && ri.Hidden))) continue;
                        for (int j = r.C1; j <= cm; j++)
                        {
                            var c = r.Sheet.Get(i, j);
                            if (c.Kind != CellKind.Empty) counta++;
                            if (c.Kind == CellKind.Num) vals.Add(c.Num);
                            else if (c.Kind == CellKind.Err) firstErr ??= (CellError)(int)c.Num;
                        }
                    }
                }
                else foreach (var d in Nums([x])) { vals.Add(d); counta++; }
            }
            if (firstErr != null && code is not (2 or 3)) return firstErr.Value;
            switch (code)
            {
                case 1: return vals.Count == 0 ? CellError.Div0 : vals.Sum() / vals.Count;
                case 2: return (double)vals.Count;
                case 3: return (double)counta;
                case 4: return vals.Count == 0 ? 0.0 : vals.Max();
                case 5: return vals.Count == 0 ? 0.0 : vals.Min();
                case 6: return vals.Aggregate(1.0, (p, d) => p * d);
                case 7: return vals.Count < 2 ? CellError.Div0 : Math.Sqrt(Var(vals, true));
                case 8: return vals.Count < 1 ? CellError.Div0 : Math.Sqrt(Var(vals, false));
                case 9: return vals.Sum();
                case 10: return vals.Count < 2 ? CellError.Div0 : Var(vals, true);
                default: return vals.Count < 1 ? CellError.Div0 : Var(vals, false);
            }
        });
    }

    static double Gcd(double a, double b) { while (b != 0) { (a, b) = (b, a % b); } return a; }

    static object Chk(double d) => double.IsNaN(d) || double.IsInfinity(d) ? CellError.Num : d;

    /// <summary>Scorre l'area usata di un intervallo con indici relativi (i,j).</summary>
    static void ForEachGrid(RangeRef r, Action<int, int, object?> f)
    {
        int rm = Math.Min(r.R2, r.Sheet.MaxRow) - r.R1, cm = Math.Min(r.C2, r.Sheet.MaxCol) - r.C1;
        for (int i = 0; i <= rm; i++) for (int j = 0; j <= cm; j++) f(i, j, r.At(i, j));
    }

    // a: [sumRange?] (range, crit)*  ; first = indice del primo criterio
    static object? CondAggregate(EvalCtx cx, object?[] a, int first, object? sumArg, string kind)
    {
        if ((a.Length - first) % 2 != 0) return CellError.Value;
        var ranges = new List<RangeRef>(); var crits = new List<Criteria>();
        for (int k = first; k < a.Length; k += 2)
        {
            if (a[k] is not RangeRef r) return CellError.Value;
            ranges.Add(r); crits.Add(new Criteria(V.Scalar(a[k + 1], cx)));
        }
        var shape = ranges[0];
        if (ranges.Any(r => r.Rows != shape.Rows || r.Cols != shape.Cols)) return CellError.Value;
        RangeRef? sum = null;
        if (sumArg != null) { sum = sumArg as RangeRef; if (sum == null) return CellError.Value; if (sum.Rows != shape.Rows || sum.Cols != shape.Cols) return CellError.Value; }
        int rm = Math.Min(shape.R2, Math.Max(shape.Sheet.MaxRow, sum?.Sheet.MaxRow ?? 0) ) - shape.R1;
        // l'area utile è limitata dall'area usata di tutti i fogli coinvolti
        int maxRel = 0, maxCol = 0;
        foreach (var r in ranges.Append(sum!).Where(x => x != null)) { maxRel = Math.Max(maxRel, Math.Min(r.R2, r.Sheet.MaxRow) - r.R1); maxCol = Math.Max(maxCol, Math.Min(r.C2, r.Sheet.MaxCol) - r.C1); }
        double s = 0, mx = double.NegativeInfinity, mn = double.PositiveInfinity; int n = 0;
        for (int i = 0; i <= maxRel; i++)
            for (int j = 0; j <= maxCol; j++)
            {
                bool ok = true;
                for (int k = 0; k < ranges.Count && ok; k++) ok = crits[k].Match(ranges[k].At(i, j));
                if (!ok) continue;
                if (kind == "COUNT") { n++; continue; }
                if (sum!.At(i, j) is double d) { s += d; n++; mx = Math.Max(mx, d); mn = Math.Min(mn, d); }
            }
        return kind switch
        {
            "COUNT" => (double)n,
            "SUM" => s,
            "AVERAGE" => n == 0 ? CellError.Div0 : s / n,
            "MAX" => n == 0 ? 0.0 : mx,
            _ => n == 0 ? 0.0 : mn,
        };
    }

    static double Var(List<double> v, bool sample)
    {
        double m = v.Average(), s = 0;
        foreach (var d in v) s += (d - m) * (d - m);
        return s / (sample ? v.Count - 1 : v.Count);
    }

    static void RegisterStats()
    {
        R("MEDIAN", 1, 255, (cx, a) =>
        {
            var l = Nums(a).OrderBy(d => d).ToList();
            if (l.Count == 0) return CellError.Num;
            return l.Count % 2 == 1 ? l[l.Count / 2] : (l[l.Count / 2 - 1] + l[l.Count / 2]) / 2;
        });
        R("LARGE", 2, 2, (cx, a) => KthOf(cx, a, true));
        R("SMALL", 2, 2, (cx, a) => KthOf(cx, a, false));
        R("MODE.SNGL", 1, 255, (cx, a) => Mode(a));
        R("MODE", 1, 255, (cx, a) => Mode(a));
        foreach (var (n, samp, sq) in new[] { ("STDEV", true, true), ("STDEV.S", true, true), ("STDEV.P", false, true), ("STDEVP", false, true), ("VAR", true, false), ("VAR.S", true, false), ("VAR.P", false, false), ("VARP", false, false) })
        {
            bool s = samp, q = sq;
            R(n, 1, 255, (cx, a) =>
            {
                var l = Nums(a).ToList();
                if (l.Count < (s ? 2 : 1)) return CellError.Div0;
                double v = Var(l, s); return q ? Math.Sqrt(v) : v;
            });
        }
        foreach (var n in new[] { "RANK", "RANK.EQ" })
            R(n, 2, 3, (cx, a) =>
            {
                double x = V.Num(V.Scalar(a[0], cx));
                if (a[1] is not RangeRef r) return CellError.Value;
                bool asc = a.Length > 2 && V.Num(V.Scalar(a[2], cx)) != 0;
                int rank = 1; bool found = false;
                foreach (var c in Stored(r)) if (c.Kind == CellKind.Num) { if (c.Num == x) found = true; if (asc ? c.Num < x : c.Num > x) rank++; }
                return found ? rank : CellError.NA;
            });
        R("PERCENTILE", 2, 2, (cx, a) => Percentile(cx, a));
        R("PERCENTILE.INC", 2, 2, (cx, a) => Percentile(cx, a));
        R("CORREL", 2, 2, (cx, a) =>
        {
            var x = Flat(a[0]).OfType<double>().ToList(); var y = Flat(a[1]).OfType<double>().ToList();
            if (x.Count != y.Count || x.Count < 2) return CellError.NA;
            double mx = x.Average(), my = y.Average(), sxy = 0, sxx = 0, syy = 0;
            for (int i = 0; i < x.Count; i++) { sxy += (x[i] - mx) * (y[i] - my); sxx += (x[i] - mx) * (x[i] - mx); syy += (y[i] - my) * (y[i] - my); }
            return sxx == 0 || syy == 0 ? CellError.Div0 : sxy / Math.Sqrt(sxx * syy);
        });
    }

    static object? Percentile(EvalCtx cx, object?[] a)
    {
        var l = Nums([a[0]]).OrderBy(d => d).ToList();
        double k = V.Num(V.Scalar(a[1], cx));
        if (l.Count == 0 || k < 0 || k > 1) return CellError.Num;
        double pos = k * (l.Count - 1); int lo = (int)Math.Floor(pos);
        return lo + 1 < l.Count ? l[lo] + (pos - lo) * (l[lo + 1] - l[lo]) : l[lo];
    }

    static object? KthOf(EvalCtx cx, object?[] a, bool large)
    {
        var l = Nums([a[0]]).ToList();
        double kd = V.Num(V.Scalar(a[1], cx));
        int k = (int)Math.Ceiling(kd);
        if (k < 1 || k > l.Count) return CellError.Num;
        l.Sort();
        return large ? l[l.Count - k] : l[k - 1];
    }

    static object? Mode(object?[] a)
    {
        var l = Nums(a).ToList();
        var g = l.GroupBy(d => d).Where(x => x.Count() > 1).ToList();
        if (g.Count == 0) return CellError.NA;
        int best = g.Max(x => x.Count());
        double first = l.First(d => g.Any(x => x.Key == d && x.Count() == best));
        return first;
    }
}
