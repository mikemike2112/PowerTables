using System.Globalization;
using System.Text;

namespace PowerTables.Core.Formulas;

/// <summary>Formattazione secondo i codici di formato numerico di Excel.</summary>
public static class NumFmt
{
    public static readonly Dictionary<int, string> Builtin = new()
    {
        [0] = "General", [1] = "0", [2] = "0.00", [3] = "#,##0", [4] = "#,##0.00",
        [5] = "\"$\"#,##0_);\\(\"$\"#,##0\\)", [6] = "\"$\"#,##0_);[Red]\\(\"$\"#,##0\\)",
        [7] = "\"$\"#,##0.00_);\\(\"$\"#,##0.00\\)", [8] = "\"$\"#,##0.00_);[Red]\\(\"$\"#,##0.00\\)",
        [9] = "0%", [10] = "0.00%", [11] = "0.00E+00", [12] = "# ?/?", [13] = "# ??/??",
        [14] = "m/d/yyyy", [15] = "d-mmm-yy", [16] = "d-mmm", [17] = "mmm-yy", [18] = "h:mm AM/PM", [19] = "h:mm:ss AM/PM",
        [20] = "h:mm", [21] = "h:mm:ss", [22] = "m/d/yyyy h:mm",
        [37] = "#,##0_);(#,##0)", [38] = "#,##0_);[Red](#,##0)", [39] = "#,##0.00_);(#,##0.00)", [40] = "#,##0.00_);[Red](#,##0.00)",
        [41] = "_(* #,##0_);_(* \\(#,##0\\);_(* \"-\"_);_(@_)", [42] = "_(\"$\"* #,##0_);_(\"$\"* \\(#,##0\\);_(\"$\"* \"-\"_);_(@_)",
        [43] = "_(* #,##0.00_);_(* \\(#,##0.00\\);_(* \"-\"??_);_(@_)", [44] = "_(\"$\"* #,##0.00_);_(\"$\"* \\(#,##0.00\\);_(\"$\"* \"-\"??_);_(@_)",
        [45] = "mm:ss", [46] = "[h]:mm:ss", [47] = "mmss.0", [48] = "##0.0E+0", [49] = "@",
    };

    public static bool IsDateId(int id) => id is >= 14 and <= 22 or 45 or 46 or 47;

    public static bool IsDate(string code)
    {
        foreach (var sec in Split(code))
            if (Tokenize(sec).Any(t => t.K is 'y' or 'd' or 'h' or 's' or 'm' or 'A')) return true;
        return false;
    }

    sealed class Tok { public char K; public string S = ""; public int N; public bool Elapsed; }

    static List<string> Split(string code)
    {
        var res = new List<string>(); var sb = new StringBuilder(); bool q = false; int br = 0;
        for (int i = 0; i < code.Length; i++)
        {
            char c = code[i];
            if (c == '\\' && i + 1 < code.Length && !q) { sb.Append(c).Append(code[++i]); continue; }
            if (c == '"') q = !q;
            else if (!q && c == '[') br++;
            else if (!q && c == ']') br--;
            if (c == ';' && !q && br == 0) { res.Add(sb.ToString()); sb.Clear(); }
            else sb.Append(c);
        }
        res.Add(sb.ToString());
        return res;
    }

    static List<Tok> Tokenize(string sec)
    {
        var res = new List<Tok>();
        for (int i = 0; i < sec.Length;)
        {
            char c = sec[i];
            if (c == '"')
            {
                int e = sec.IndexOf('"', i + 1); if (e < 0) e = sec.Length;
                res.Add(new Tok { K = 'L', S = sec[(i + 1)..e] }); i = e + 1; continue;
            }
            if (c == '\\' && i + 1 < sec.Length) { res.Add(new Tok { K = 'L', S = sec[i + 1].ToString() }); i += 2; continue; }
            if (c == '_' && i + 1 < sec.Length) { res.Add(new Tok { K = 'L', S = " " }); i += 2; continue; }
            if (c == '*' && i + 1 < sec.Length) { i += 2; continue; }
            if (c == '[')
            {
                int e = sec.IndexOf(']', i); if (e < 0) e = sec.Length - 1;
                string inner = sec[(i + 1)..e];
                string low = inner.ToLowerInvariant();
                if (low is "h" or "hh" or "m" or "mm" or "s" or "ss") res.Add(new Tok { K = low[0], N = low.Length, Elapsed = true });
                else if (inner.StartsWith('$')) { int d = inner.IndexOf('-'); string sym = d > 1 ? inner[1..d] : inner[1..]; if (sym.Length > 0) res.Add(new Tok { K = 'L', S = sym }); }
                i = e + 1; continue;
            }
            char lc = char.ToLowerInvariant(c);
            if (lc is 'y' or 'd' or 'h' or 's' or 'm')
            {
                int j = i; while (j < sec.Length && char.ToLowerInvariant(sec[j]) == lc) j++;
                res.Add(new Tok { K = lc, N = j - i }); i = j; continue;
            }
            if (string.Compare(sec, i, "AM/PM", 0, 5, StringComparison.OrdinalIgnoreCase) == 0) { res.Add(new Tok { K = 'A', N = 5 }); i += 5; continue; }
            if (string.Compare(sec, i, "A/P", 0, 3, StringComparison.OrdinalIgnoreCase) == 0) { res.Add(new Tok { K = 'A', N = 3 }); i += 3; continue; }
            if (string.Compare(sec, i, "General", 0, 7, StringComparison.OrdinalIgnoreCase) == 0) { res.Add(new Tok { K = 'G' }); i += 7; continue; }
            if ((c == 'E' || c == 'e') && i + 1 < sec.Length && (sec[i + 1] == '+' || sec[i + 1] == '-'))
            {
                int j = i + 2; while (j < sec.Length && sec[j] == '0') j++;
                res.Add(new Tok { K = 'E', S = sec[i + 1].ToString(), N = j - i - 2 }); i = j; continue;
            }
            if (c is '0' or '#' or '?') { res.Add(new Tok { K = c }); i++; continue; }
            if (c is '.' or ',' or '%' or '@') { res.Add(new Tok { K = c }); i++; continue; }
            res.Add(new Tok { K = 'L', S = c.ToString() }); i++;
        }
        // 'm' = minuti se dopo h o prima di s
        for (int i = 0; i < res.Count; i++)
        {
            if (res[i].K != 'm') continue;
            int p = i - 1; while (p >= 0 && res[p].K == 'L') p--;
            int n = i + 1; while (n < res.Count && res[n].K == 'L') n++;
            if ((p >= 0 && res[p].K == 'h') || (n < res.Count && res[n].K == 's')) res[i].K = 'M';
        }
        return res;
    }

    public static string Format(object? value, string code, CultureInfo? ci = null)
    {
        ci ??= CultureInfo.CurrentCulture;
        switch (value)
        {
            case null: return "";
            case CellError e: return Errors.Text(e);
            case bool b: return b ? "TRUE" : "FALSE";
            case string s: return FormatText(s, code, ci);
            case double d: return FormatNumber(d, code, ci);
            default: return value.ToString() ?? "";
        }
    }

    public static string Format(double v, int id, string? custom, CultureInfo? ci = null)
    {
        ci ??= CultureInfo.CurrentCulture;
        if (custom != null) return FormatNumber(v, custom, ci);
        if (id == 14) return FormatNumber(v, ShortDate(ci), ci);
        return Builtin.TryGetValue(id, out var c) ? FormatNumber(v, c, ci) : FormatNumber(v, "General", ci);
    }

    public static string ShortDate(CultureInfo ci) => ci.DateTimeFormat.ShortDatePattern.Replace('M', 'm').Replace("yyyy", "yyyy").Replace("yy", "yyyy").Replace("yyyyyyyy", "yyyy");

    static string FormatText(string s, string code, CultureInfo ci)
    {
        var secs = Split(code);
        var sec = secs.Count >= 4 ? secs[3] : secs.Count == 1 ? secs[0] : null;
        if (sec == null) return s;
        var toks = Tokenize(sec);
        if (!toks.Any(t => t.K == '@')) return s;
        var sb = new StringBuilder();
        foreach (var t in toks) sb.Append(t.K == '@' ? s : t.K == 'L' ? t.S : "");
        return sb.ToString();
    }

    public static string General(double v)
    {
        if (v == 0) return "0";
        double a = Math.Abs(v);
        if (a >= 1e11 || a < 1e-9)
        {
            string s = v.ToString("0.#####E+00", CultureInfo.InvariantCulture);
            return s;
        }
        int intLen = a >= 1 ? (int)Math.Floor(Math.Log10(a)) + 1 : 1;
        int dec = Math.Max(0, 10 - intLen);
        decimal dv;
        try { dv = Math.Round((decimal)v, Math.Min(dec, 15), MidpointRounding.AwayFromZero); } catch (OverflowException) { return v.ToString("G10", CultureInfo.InvariantCulture); }
        string r = dv.ToString("0.###############", CultureInfo.InvariantCulture);
        return r;
    }

    static string FormatNumber(double v, string code, CultureInfo ci)
    {
        if (double.IsNaN(v) || double.IsInfinity(v)) return "#NUM!";
        var secs = Split(code);
        int pick = 0; bool negSection = false;
        // condizioni
        var cond = new (string Op, double Val)?[secs.Count];
        for (int i = 0; i < secs.Count; i++) cond[i] = ParseCond(secs[i]);
        if (cond.Any(c => c != null))
        {
            pick = -1;
            for (int i = 0; i < secs.Count; i++)
            {
                var c = cond[i];
                bool ok = c == null ? i == secs.Count - 1 || i == 1 && cond[0] != null : c.Value.Op switch
                {
                    ">" => v > c.Value.Val, "<" => v < c.Value.Val, ">=" => v >= c.Value.Val, "<=" => v <= c.Value.Val, "<>" => v != c.Value.Val, _ => v == c.Value.Val,
                };
                if (ok) { pick = i; break; }
            }
            if (pick < 0) pick = 0;
        }
        else if (secs.Count >= 2)
        {
            if (v < 0) { pick = 1; negSection = true; }
            else if (v == 0 && secs.Count >= 3) pick = 2;
        }
        var toks = Tokenize(secs[pick]);
        double val = negSection ? -v : v;
        bool autoMinus = !negSection && v < 0;
        if (toks.Count == 0) return "";
        if (toks.Any(t => t.K is 'y' or 'd' or 'h' or 's' or 'm' or 'M' or 'A')) return FormatDate(toks, v, ci);
        return FormatNum(toks, val, autoMinus, ci);
    }

    static (string, double)? ParseCond(string sec)
    {
        if (!sec.StartsWith('[')) return null;
        int e = sec.IndexOf(']'); if (e < 0) return null;
        string inner = sec[1..e];
        foreach (var op in new[] { "<=", ">=", "<>", "<", ">", "=" })
            if (inner.StartsWith(op) && double.TryParse(inner[op.Length..], NumberStyles.Float, CultureInfo.InvariantCulture, out var d)) return (op, d);
        return null;
    }

    static string FormatNum(List<Tok> toks, double v, bool autoMinus, CultureInfo ci)
    {
        if (toks.Count == 1 && toks[0].K == 'G' || toks.All(t => t.K == 'G' || t.K == 'L') && toks.Any(t => t.K == 'G'))
        {
            var sbg = new StringBuilder();
            foreach (var t in toks) sbg.Append(t.K == 'G' ? General(v).Replace('.', ci.NumberFormat.NumberDecimalSeparator[0]) : t.S);
            return sbg.ToString();
        }
        int pct = toks.Count(t => t.K == '%');
        for (int i = 0; i < pct; i++) v *= 100;
        int firstPh = toks.FindIndex(t => t.K is '0' or '#' or '?');
        if (firstPh < 0)
        {
            // nessun segnaposto: solo letterali (es. testo fisso)
            var sb0 = new StringBuilder(autoMinus ? "-" : "");
            foreach (var t in toks) sb0.Append(t.K == 'L' ? t.S : t.K == '%' ? "%" : t.K == '@' ? "" : "");
            return sb0.ToString();
        }
        int lastPh = toks.FindLastIndex(t => t.K is '0' or '#' or '?');
        int dot = toks.FindIndex(t => t.K == '.');
        if (dot > lastPh || (dot >= 0 && dot < firstPh && false)) dot = -1;
        bool sci = toks.Any(t => t.K == 'E');
        // virgole di scala (subito dopo l'ultimo segnaposto intero)
        int intEnd = dot >= 0 ? dot : (sci ? toks.FindIndex(t => t.K == 'E') : toks.Count);
        int scale = 0; bool thousands = false;
        {
            int lastIntPh = -1;
            for (int i = 0; i < intEnd; i++) if (toks[i].K is '0' or '#' or '?') lastIntPh = i;
            for (int i = 0; i < intEnd; i++)
                if (toks[i].K == ',')
                {
                    if (i > lastIntPh && lastIntPh >= 0) scale++;
                    else if (i > firstPh) thousands = true;
                }
        }
        for (int i = 0; i < scale; i++) v /= 1000;
        int decimals = 0;
        if (dot >= 0) for (int i = dot + 1; i < toks.Count; i++) if (toks[i].K is '0' or '#' or '?') decimals++;
        double av = Math.Abs(v);
        string intPart, fracPart, expPart = "";
        if (sci)
        {
            var et = toks.First(t => t.K == 'E');
            int exp = av == 0 ? 0 : (int)Math.Floor(Math.Log10(av));
            double mant = av == 0 ? 0 : av / Math.Pow(10, exp);
            decimal rm = Math.Round((decimal)mant, decimals, MidpointRounding.AwayFromZero);
            if (rm >= 10) { rm /= 10; exp++; }
            var parts = rm.ToString("F" + decimals, CultureInfo.InvariantCulture).Split('.');
            intPart = parts[0]; fracPart = parts.Length > 1 ? parts[1] : "";
            expPart = "E" + (exp < 0 ? "-" : et.S == "+" ? "+" : "") + Math.Abs(exp).ToString(new string('0', Math.Max(1, et.N)), CultureInfo.InvariantCulture);
        }
        else
        {
            string f;
            try { f = Math.Round((decimal)av, Math.Min(decimals, 20), MidpointRounding.AwayFromZero).ToString("F" + decimals, CultureInfo.InvariantCulture); }
            catch (OverflowException) { f = av.ToString("F" + decimals, CultureInfo.InvariantCulture); }
            var parts = f.Split('.');
            intPart = parts[0]; fracPart = parts.Length > 1 ? parts[1] : "";
        }
        bool isZero = intPart.Trim('0').Length == 0 && fracPart.Trim('0').Length == 0;
        string digits = intPart.TrimStart('0');
        // segnaposto interi
        var intPhs = new List<int>();
        for (int i = firstPh; i < intEnd && i <= lastPh; i++) if (toks[i].K is '0' or '#' or '?') intPhs.Add(i);
        if (sci) { intPhs.Clear(); for (int i = firstPh; i < (dot >= 0 ? dot : toks.FindIndex(t => t.K == 'E')); i++) if (toks[i].K is '0' or '#' or '?') intPhs.Add(i); }
        var sb = new StringBuilder();
        if (autoMinus && !(isZero)) sb.Append('-');
        for (int i = 0; i < firstPh; i++) AppendNonDigit(sb, toks[i], ci);
        int n = intPhs.Count;
        string grouped = digits;
        if (thousands)
        {
            int minLen = intPhs.Count(i => toks[i].K == '0');
            string padded = digits.PadLeft(minLen, '0');
            var g = new StringBuilder();
            for (int i = 0; i < padded.Length; i++)
            {
                if (i > 0 && (padded.Length - i) % 3 == 0) g.Append(ci.NumberFormat.NumberGroupSeparator);
                g.Append(padded[i]);
            }
            grouped = g.ToString();
            sb.Append(grouped);
            // eventuali letterali tra i segnaposto interi vengono omessi con il raggruppamento
        }
        else
        {
            int idx = 0;
            for (int i = firstPh; i < intEnd && i <= lastPh; i++)
            {
                var t = toks[i];
                if (t.K is '0' or '#' or '?')
                {
                    int di = idx - (n - digits.Length);
                    if (idx == 0 && digits.Length > n) { sb.Append(digits, 0, digits.Length - n + 1); }
                    else if (di >= 0 && di < digits.Length) sb.Append(digits[di]);
                    else if (t.K == '0') sb.Append('0');
                    else if (t.K == '?') sb.Append(' ');
                    idx++;
                }
                else if (t.K != ',') AppendNonDigit(sb, t, ci);
            }
        }
        if (dot >= 0 && decimals > 0)
        {
            sb.Append(ci.NumberFormat.NumberDecimalSeparator);
            int fi = 0;
            string fp = fracPart;
            for (int i = dot + 1; i < toks.Count; i++)
            {
                var t = toks[i];
                if (t.K is '0' or '#' or '?')
                {
                    char d = fi < fp.Length ? fp[fi] : '0';
                    bool trailingZeros = fp[fi..].All(ch => ch == '0');
                    if (t.K == '0') sb.Append(d);
                    else if (!trailingZeros) sb.Append(d);
                    else if (t.K == '?') sb.Append(' ');
                    fi++;
                }
                else if (t.K == 'E') { sb.Append(expPart); }
                else if (t.K != ',') AppendNonDigit(sb, t, ci);
                if (i >= lastPh && t.K != 'E' && toks.Skip(i + 1).All(x => x.K != '0' && x.K != '#' && x.K != '?')) { lastPh = i; }
            }
        }
        else if (dot >= 0)
        {
            // "0." : mostra il punto
            sb.Append(ci.NumberFormat.NumberDecimalSeparator);
        }
        if (sci && (dot < 0 || decimals == 0) ) sb.Append(expPart);
        int tail = Math.Max(lastPh + 1, dot >= 0 ? dot + 1 + 0 : 0);
        // coda: tutti i token dopo l'ultimo segnaposto (e dopo l'eventuale esponente già emesso)
        for (int i = lastPh + 1; i < toks.Count; i++)
        {
            var t = toks[i];
            if (t.K == 'E') { if (dot >= 0 && decimals > 0) continue; continue; }
            AppendNonDigit(sb, t, ci);
        }
        // se E è tra i token dopo i decimali ed è già stato emesso, ok; altrimenti emesso sopra
        return sb.ToString();
    }

    static void AppendNonDigit(StringBuilder sb, Tok t, CultureInfo ci)
    {
        switch (t.K)
        {
            case 'L': sb.Append(t.S); break;
            case '%': sb.Append('%'); break;
            case '.': break;
            case '@': break;
        }
    }

    static string FormatDate(List<Tok> toks, double v, CultureInfo ci)
    {
        if (v < 0) return "########";
        long secs = XlDate.SecondsOfDay(v);
        double serial = Math.Floor(v);
        if (secs >= 86400) { secs -= 86400; serial++; }
        var (y, m, d) = XlDate.FromSerial(serial);
        int dow = XlDate.DayOfWeek(serial);
        int hh = (int)(secs / 3600), mi = (int)(secs % 3600 / 60), ss = (int)(secs % 60);
        bool ampm = toks.Any(t => t.K == 'A');
        var dt = ci.DateTimeFormat;
        var sb = new StringBuilder();
        for (int i = 0; i < toks.Count; i++)
        {
            var t = toks[i];
            switch (t.K)
            {
                case 'y': sb.Append(t.N <= 2 ? (y % 100).ToString("00") : y.ToString("0000")); break;
                case 'm':
                    sb.Append(t.N switch
                    {
                        1 => m.ToString(), 2 => m.ToString("00"),
                        3 => dt.GetAbbreviatedMonthName(m), 4 => dt.GetMonthName(m), _ => dt.GetMonthName(m)[..1],
                    }); break;
                case 'd':
                    sb.Append(t.N switch
                    {
                        1 => d.ToString(), 2 => d.ToString("00"),
                        3 => dt.GetAbbreviatedDayName((DayOfWeek)((dow - 1) % 7)), _ => dt.GetDayName((DayOfWeek)((dow - 1) % 7)),
                    }); break;
                case 'h':
                    {
                        int h = t.Elapsed ? (int)(Math.Floor(v) * 24 + hh) : hh;
                        if (ampm && !t.Elapsed) h = h % 12 == 0 ? 12 : h % 12;
                        sb.Append(t.N >= 2 ? h.ToString("00") : h.ToString()); break;
                    }
                case 'M':
                    {
                        long mm = t.Elapsed ? (long)(Math.Floor(v) * 1440 + hh * 60 + mi) : mi;
                        sb.Append(t.N >= 2 ? mm.ToString("00") : mm.ToString()); break;
                    }
                case 's':
                    {
                        long s2 = t.Elapsed ? (long)(Math.Floor(v) * 86400 + secs) : ss;
                        sb.Append(t.N >= 2 ? s2.ToString("00") : s2.ToString());
                        // secondi frazionari ".0"
                        if (i + 2 < toks.Count && toks[i + 1].K == '.' && toks[i + 2].K == '0')
                        {
                            int fd = 0; int j = i + 2; while (j < toks.Count && toks[j].K == '0') { fd++; j++; }
                            double fr = Math.Round((v - Math.Floor(v)) * 86400 % 1, fd);
                            sb.Append(ci.NumberFormat.NumberDecimalSeparator).Append(((long)Math.Round(fr * Math.Pow(10, fd))).ToString(new string('0', fd)));
                            i = j - 1;
                        }
                        break;
                    }
                case 'A': sb.Append(t.N == 5 ? (hh < 12 ? "AM" : "PM") : (hh < 12 ? "A" : "P")); break;
                case 'L': sb.Append(t.S); break;
                case '@': break;
                default: break;
            }
        }
        return sb.ToString();
    }
}
