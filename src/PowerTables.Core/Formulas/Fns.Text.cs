using System.Globalization;
using System.Text;

namespace PowerTables.Core.Formulas;

public static partial class Fns
{
    static string St(object? v) => V.Str(v);
    static int In(object? v) => (int)Math.Truncate(V.Num(v));

    static void RegisterText()
    {
        S("LEN", 1, 1, a => (double)St(a[0]).Length, true);
        S("LEFT", 1, 2, a => { var s = St(a[0]); int n = a.Length > 1 ? In(a[1]) : 1; if (n < 0) return CellError.Value; return s[..Math.Min(n, s.Length)]; }, true);
        S("RIGHT", 1, 2, a => { var s = St(a[0]); int n = a.Length > 1 ? In(a[1]) : 1; if (n < 0) return CellError.Value; return s[Math.Max(0, s.Length - n)..]; }, true);
        S("MID", 3, 3, a =>
        {
            var s = St(a[0]); int st = In(a[1]), n = In(a[2]);
            if (st < 1 || n < 0) return CellError.Value;
            if (st > s.Length) return "";
            return s.Substring(st - 1, Math.Min(n, s.Length - st + 1));
        }, true);
        S("UPPER", 1, 1, a => St(a[0]).ToUpperInvariant(), true);
        S("LOWER", 1, 1, a => St(a[0]).ToLowerInvariant(), true);
        S("PROPER", 1, 1, a =>
        {
            var sb = new StringBuilder(); bool up = true;
            foreach (var c in St(a[0])) { sb.Append(up ? char.ToUpperInvariant(c) : char.ToLowerInvariant(c)); up = !char.IsLetter(c); }
            return sb.ToString();
        }, true);
        S("TRIM", 1, 1, a => System.Text.RegularExpressions.Regex.Replace(St(a[0]).Trim(' '), " {2,}", " "), true);
        S("CLEAN", 1, 1, a => new string(St(a[0]).Where(c => c >= 32).ToArray()), true);
        S("EXACT", 2, 2, a => St(a[0]) == St(a[1]), true);
        S("REPT", 2, 2, a => { int n = In(a[1]); if (n < 0) return CellError.Value; var s = St(a[0]); if ((long)s.Length * n > 32767) return CellError.Value; return string.Concat(Enumerable.Repeat(s, n)); }, true);
        S("CHAR", 1, 1, a => { int n = In(a[0]); return n is < 1 or > 255 ? CellError.Value : ((char)n).ToString(); }, true);
        S("UNICHAR", 1, 1, a => { int n = In(a[0]); return n < 1 ? CellError.Value : char.ConvertFromUtf32(n); }, true);
        S("CODE", 1, 1, a => { var s = St(a[0]); return s.Length == 0 ? CellError.Value : (double)s[0]; }, true);
        S("UNICODE", 1, 1, a => { var s = St(a[0]); return s.Length == 0 ? CellError.Value : (double)char.ConvertToUtf32(s, 0); }, true);
        S("VALUE", 1, 1, a =>
        {
            if (a[0] is double) return a[0];
            var s = St(a[0]);
            return V.TryParseNum(s, out var d) ? d : CellError.Value;
        }, true);
        S("NUMBERVALUE", 1, 3, a =>
        {
            string s = St(a[0]).Replace(" ", ""); string dec = a.Length > 1 && a[1] is not Missing ? St(a[1]) : ".", grp = a.Length > 2 && a[2] is not Missing ? St(a[2]) : ",";
            if (s.Length == 0) return 0.0;
            if (dec.Length == 0) return CellError.Value;
            s = s.Replace(grp, "").Replace(dec, ".");
            return V.TryParseNum(s, out var d) ? d : CellError.Value;
        }, true);
        S("CONCATENATE", 1, 255, a => string.Concat(a.Select(St)));
        R("CONCAT", 1, 255, (cx, a) => string.Concat(a.SelectMany(Flat).Select(St)));
        R("TEXTJOIN", 3, 255, (cx, a) =>
        {
            string delim = St(V.Scalar(a[0], cx)); bool skip = V.Bool(V.Scalar(a[1], cx));
            var parts = a.Skip(2).SelectMany(Flat).Where(v => !(skip && (v == null || v is string { Length: 0 }))).Select(St);
            return string.Join(delim, parts);
        });
        S("FIND", 2, 3, a => FindImpl(a, false));
        S("SEARCH", 2, 3, a => FindImpl(a, true));
        S("SUBSTITUTE", 3, 4, a =>
        {
            string t = St(a[0]), o = St(a[1]), n = St(a[2]);
            if (o.Length == 0) return t;
            if (a.Length > 3)
            {
                int inst = In(a[3]); if (inst < 1) return CellError.Value;
                int pos = -1;
                for (int k = 0; k < inst; k++) { pos = t.IndexOf(o, pos + 1, StringComparison.Ordinal); if (pos < 0) return t; }
                return t[..pos] + n + t[(pos + o.Length)..];
            }
            return t.Replace(o, n, StringComparison.Ordinal);
        }, true);
        S("REPLACE", 4, 4, a =>
        {
            string t = St(a[0]); int st = In(a[1]), n = In(a[2]);
            if (st < 1 || n < 0) return CellError.Value;
            if (st > t.Length) return t + St(a[3]);
            n = Math.Min(n, t.Length - st + 1);
            return t[..(st - 1)] + St(a[3]) + t[(st - 1 + n)..];
        }, true);
        S("TEXT", 2, 2, a =>
        {
            var v = a[0];
            if (v is string s0 && V.TryParseNum(s0, out var pd)) v = pd;
            return NumFmt.Format(v, St(a[1]), CultureInfo.InvariantCulture);
        }, true);
        S("T", 1, 1, a => a[0] as string ?? "");
        S("TEXTBEFORE", 2, 2, a => { var t = St(a[0]); int p = t.IndexOf(St(a[1]), StringComparison.Ordinal); return p < 0 ? CellError.NA : t[..p]; });
        S("TEXTAFTER", 2, 2, a => { var t = St(a[0]); var d = St(a[1]); int p = t.IndexOf(d, StringComparison.Ordinal); return p < 0 ? CellError.NA : t[(p + d.Length)..]; });
        S("DOLLAR", 1, 2, a => NumFmt.Format(V.Num(a[0]), "$#,##0" + (a.Length > 1 && In(a[1]) > 0 ? "." + new string('0', In(a[1])) : a.Length > 1 ? "" : ".00"), CultureInfo.InvariantCulture));
        S("FIXED", 1, 3, a =>
        {
            int dec = a.Length > 1 ? In(a[1]) : 2; bool nocomma = a.Length > 2 && V.Bool(a[2]);
            double x = V.Num(a[0]);
            if (dec < 0) { double p = Math.Pow(10, -dec); x = Math.Round(x / p, MidpointRounding.AwayFromZero) * p; dec = 0; }
            return NumFmt.Format(x, (nocomma ? "0" : "#,##0") + (dec > 0 ? "." + new string('0', dec) : ""), CultureInfo.InvariantCulture);
        });
    }

    static object? FindImpl(object?[] a, bool search)
    {
        string find = St(a[0]), within = St(a[1]); int start = a.Length > 2 ? In(a[2]) : 1;
        if (start < 1 || start > within.Length + (find.Length == 0 ? 1 : 0)) return CellError.Value;
        int p;
        if (search && (find.Contains('*') || find.Contains('?') || find.Contains('~')))
        {
            var m = Criteria.Wildcard(find, false).Match(within, start - 1);
            p = m.Success ? m.Index : -1;
        }
        else p = within.IndexOf(find, start - 1, search ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal);
        return p < 0 ? CellError.Value : p + 1.0;
    }
}
