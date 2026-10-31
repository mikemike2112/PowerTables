using System.Globalization;
using System.Text;

namespace PowerTables.Core.Formulas;

public enum T { Num, Str, Ref, Name, Func, Bool, Err, Op, LP, RP, Comma, Semi, LB, RB, Colon, End }

/// <summary>Una componente di riferimento: cella, solo colonna, o solo riga.</summary>
public struct RefPart
{
    public int Row, Col;            // 0 = assente
    public bool AbsRow, AbsCol;

    public readonly string Text()
    {
        var sb = new StringBuilder();
        if (Col > 0) { if (AbsCol) sb.Append('$'); sb.Append(Addr.ColName(Col)); }
        if (Row > 0) { if (AbsRow) sb.Append('$'); sb.Append(Row); }
        return sb.ToString();
    }
}

public sealed class RefToken
{
    public string? Sheet;           // nome foglio senza apici; null se assente
    public RefPart A, B;
    public bool IsRange;

    public bool WholeCol => A.Row == 0 && A.Col > 0;
    public bool WholeRow => A.Col == 0 && A.Row > 0;

    public string Text()
    {
        var sb = new StringBuilder();
        if (Sheet != null) sb.Append(Lexer.QuoteSheet(Sheet)).Append('!');
        sb.Append(A.Text());
        if (IsRange) sb.Append(':').Append(B.Text());
        return sb.ToString();
    }

    /// <summary>Sposta le parti relative; null se il risultato esce dal foglio (#REF!).</summary>
    public RefToken? Shifted(int dr, int dc)
    {
        var r = new RefToken { Sheet = Sheet, IsRange = IsRange, A = A, B = B };
        if (!ShiftPart(ref r.A, dr, dc)) return null;
        if (IsRange && !ShiftPart(ref r.B, dr, dc)) return null;
        return r;
    }

    static bool ShiftPart(ref RefPart p, int dr, int dc)
    {
        if (p.Row > 0 && !p.AbsRow) { p.Row += dr; if (p.Row < 1 || p.Row > Addr.MaxRow) return false; }
        if (p.Col > 0 && !p.AbsCol) { p.Col += dc; if (p.Col < 1 || p.Col > Addr.MaxCol) return false; }
        return true;
    }
}

public struct Token
{
    public T Kind;
    public string Text;
    public int Pos, Len;
    public double Num;
    public RefToken? Ref;
    public CellError Err;
}

public static class Lexer
{
    static readonly string[] ErrNames = ["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#SPILL!", "#CALC!"];

    public static string QuoteSheet(string s)
    {
        bool plain = s.Length > 0 && (char.IsLetter(s[0]) || s[0] == '_');
        foreach (var ch in s) if (!(char.IsLetterOrDigit(ch) || ch == '_' || ch == '.')) plain = false;
        if (plain && Addr.TryParse(s, out _, out _)) plain = false;
        return plain ? s : "'" + s.Replace("'", "''") + "'";
    }

    /// <summary>Scompone la formula (senza '='). Lancia FormatException su sintassi non riconoscibile.</summary>
    public static List<Token> Tokenize(string s)
    {
        var res = new List<Token>();
        int i = 0;
        while (i < s.Length)
        {
            char ch = s[i];
            if (char.IsWhiteSpace(ch)) { i++; continue; }
            int start = i;
            if (ch == '"')
            {
                var sb = new StringBuilder(); i++;
                while (true)
                {
                    if (i >= s.Length) throw new FormatException("stringa non chiusa");
                    if (s[i] == '"') { if (i + 1 < s.Length && s[i + 1] == '"') { sb.Append('"'); i += 2; continue; } i++; break; }
                    sb.Append(s[i++]);
                }
                res.Add(new Token { Kind = T.Str, Text = sb.ToString(), Pos = start, Len = i - start });
            }
            else if (ch == '#')
            {
                string? m = ErrNames.FirstOrDefault(e => string.Compare(s, i, e, 0, e.Length, StringComparison.OrdinalIgnoreCase) == 0);
                if (m == null) throw new FormatException("errore sconosciuto");
                i += m.Length; Errors.TryParse(m, out var ce);
                res.Add(new Token { Kind = T.Err, Text = m, Err = ce, Pos = start, Len = i - start });
            }
            else if (char.IsAsciiDigit(ch) || (ch == '.' && i + 1 < s.Length && char.IsAsciiDigit(s[i + 1])))
            {
                // riga intera "1:3"?
                int j = i; while (j < s.Length && char.IsAsciiDigit(s[j])) j++;
                if (j < s.Length && s[j] == ':' && j + 1 < s.Length && (char.IsAsciiDigit(s[j + 1]) || s[j + 1] == '$') && TryReadRef(s, ref i, null, out var rt))
                    res.Add(new Token { Kind = T.Ref, Ref = rt, Text = s[start..i], Pos = start, Len = i - start });
                else
                {
                    i = start;
                    while (i < s.Length && (char.IsAsciiDigit(s[i]) || s[i] == '.')) i++;
                    if (i < s.Length && (s[i] == 'e' || s[i] == 'E'))
                    {
                        int k = i + 1;
                        if (k < s.Length && (s[k] == '+' || s[k] == '-')) k++;
                        if (k < s.Length && char.IsAsciiDigit(s[k])) { while (k < s.Length && char.IsAsciiDigit(s[k])) k++; i = k; }
                    }
                    if (!double.TryParse(s.AsSpan(start, i - start), NumberStyles.Float, CultureInfo.InvariantCulture, out var d)) throw new FormatException("numero non valido");
                    res.Add(new Token { Kind = T.Num, Num = d, Text = s[start..i], Pos = start, Len = i - start });
                }
            }
            else if (ch == '\'' || ch == '$' || char.IsLetter(ch) || ch == '_' || ch == '\\')
            {
                string? sheet = null;
                if (ch == '\'')
                {
                    var sb = new StringBuilder(); i++;
                    while (true)
                    {
                        if (i >= s.Length) throw new FormatException("nome foglio non chiuso");
                        if (s[i] == '\'') { if (i + 1 < s.Length && s[i + 1] == '\'') { sb.Append('\''); i += 2; continue; } i++; break; }
                        sb.Append(s[i++]);
                    }
                    if (i >= s.Length || s[i] != '!') throw new FormatException("atteso '!'");
                    i++; sheet = sb.ToString();
                    if (!TryReadRef(s, ref i, sheet, out var rt0)) throw new FormatException("riferimento non valido");
                    res.Add(new Token { Kind = T.Ref, Ref = rt0, Text = s[start..i], Pos = start, Len = i - start });
                    continue;
                }
                int j = i;
                while (j < s.Length && (char.IsLetterOrDigit(s[j]) || s[j] is '_' or '.' or '$' or '\\')) j++;
                string word = s[i..j];
                if (j < s.Length && s[j] == '!')
                {
                    i = j + 1; sheet = word;
                    if (!TryReadRef(s, ref i, sheet, out var rt1)) throw new FormatException("riferimento non valido");
                    res.Add(new Token { Kind = T.Ref, Ref = rt1, Text = s[start..i], Pos = start, Len = i - start });
                }
                else if (j < s.Length && s[j] == '(')
                {
                    i = j; res.Add(new Token { Kind = T.Func, Text = word, Pos = start, Len = j - start });
                }
                else if (TryReadRef(s, ref i, null, out var rt2))
                {
                    res.Add(new Token { Kind = T.Ref, Ref = rt2, Text = s[start..i], Pos = start, Len = i - start });
                }
                else
                {
                    i = j;
                    if (word.Equals("TRUE", StringComparison.OrdinalIgnoreCase) || word.Equals("FALSE", StringComparison.OrdinalIgnoreCase))
                        res.Add(new Token { Kind = T.Bool, Num = word.Length == 4 ? 1 : 0, Text = word, Pos = start, Len = j - start });
                    else res.Add(new Token { Kind = T.Name, Text = word, Pos = start, Len = j - start });
                }
            }
            else
            {
                i++;
                switch (ch)
                {
                    case '(': res.Add(Mk(T.LP, "(", start, 1)); break;
                    case ')': res.Add(Mk(T.RP, ")", start, 1)); break;
                    case ',': res.Add(Mk(T.Comma, ",", start, 1)); break;
                    case ';': res.Add(Mk(T.Semi, ";", start, 1)); break;
                    case '{': res.Add(Mk(T.LB, "{", start, 1)); break;
                    case '}': res.Add(Mk(T.RB, "}", start, 1)); break;
                    case ':': res.Add(Mk(T.Colon, ":", start, 1)); break;
                    case '+' or '-' or '*' or '/' or '^' or '&' or '=' or '%': res.Add(Mk(T.Op, ch.ToString(), start, 1)); break;
                    case '<':
                        if (i < s.Length && (s[i] == '=' || s[i] == '>')) { res.Add(Mk(T.Op, s.Substring(i - 1, 2), start, 2)); i++; }
                        else res.Add(Mk(T.Op, "<", start, 1));
                        break;
                    case '>':
                        if (i < s.Length && s[i] == '=') { res.Add(Mk(T.Op, ">=", start, 2)); i++; }
                        else res.Add(Mk(T.Op, ">", start, 1));
                        break;
                    default: throw new FormatException($"carattere inatteso '{ch}'");
                }
            }
        }
        res.Add(new Token { Kind = T.End, Text = "", Pos = s.Length });
        return res;
    }

    static Token Mk(T k, string t, int p, int l) => new() { Kind = k, Text = t, Pos = p, Len = l };

    // Legge una parte di riferimento a partire da i (cella, colonna, riga).
    static bool ReadPart(string s, ref int i, out RefPart p)
    {
        p = default; int k = i;
        if (k < s.Length && s[k] == '$') { k++; p.AbsCol = true; }
        int a = k; while (k < s.Length && char.IsAsciiLetter(s[k])) k++;
        int col = a == k ? 0 : Addr.ColIndex(s.AsSpan(a, k - a));
        if (a != k && col == 0) return false;
        p.Col = col;
        if (col == 0) p.AbsCol = false;
        int ra = k; bool absR = false;
        if (k < s.Length && s[k] == '$') { k++; absR = true; ra = k; }
        int d = k; while (k < s.Length && char.IsAsciiDigit(s[k])) k++;
        if (k > d)
        {
            if (k - d > 7) return false;
            p.Row = int.Parse(s.AsSpan(d, k - d));
            if (p.Row < 1 || p.Row > Addr.MaxRow) return false;
            p.AbsRow = absR;
        }
        else if (absR) return false;
        if (p.Col == 0 && p.Row == 0) return false;
        // non deve proseguire con caratteri di identificatore
        if (k < s.Length && (char.IsLetterOrDigit(s[k]) || s[k] == '_' || s[k] == '.' || s[k] == '(')) return false;
        i = k; return true;
    }

    static bool TryReadRef(string s, ref int i, string? sheet, out RefToken rt)
    {
        rt = null!; int k = i;
        if (!ReadPart(s, ref k, out var a)) return false;
        var tok = new RefToken { Sheet = sheet, A = a };
        if (k < s.Length && s[k] == ':')
        {
            int k2 = k + 1;
            if (ReadPart(s, ref k2, out var b) && (a.Row == 0) == (b.Row == 0) && (a.Col == 0) == (b.Col == 0)) { tok.B = b; tok.IsRange = true; k = k2; }
            else if (a.Row == 0 || a.Col == 0) return false;
        }
        else if (a.Row == 0 || a.Col == 0) return false; // colonna sola o riga sola senza ':' = nome
        rt = tok; i = k; return true;
    }
}
