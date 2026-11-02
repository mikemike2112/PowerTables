using System.Text;

namespace PowerTables.Core.Formulas;

/// <summary>Trasformazioni testuali di formule (spostamento riferimenti, prefissi _xlfn., rinomina fogli).</summary>
public static class FormulaText
{
    /// <summary>Funzioni introdotte dopo Excel 2007: nel file vogliono il prefisso _xlfn.</summary>
    static readonly HashSet<string> Future = new(StringComparer.OrdinalIgnoreCase)
    {
        "IFNA","IFS","SWITCH","CONCAT","TEXTJOIN","MAXIFS","MINIFS","XLOOKUP","XMATCH","XOR","DAYS","STDEV.S","STDEV.P","VAR.S","VAR.P",
        "MODE.SNGL","RANK.EQ","RANK.AVG","NUMBERVALUE","UNICHAR","UNICODE","ISOWEEKNUM","CEILING.MATH","FLOOR.MATH","FILTER","SORT","SORTBY","UNIQUE","SEQUENCE",
        "TEXTBEFORE","TEXTAFTER","TEXTSPLIT","LET","LAMBDA","RANDARRAY","ISFORMULA","FORMULATEXT","DAYS","PERCENTILE.INC","QUARTILE.INC","COUNTBLANK_",
        "BITAND","BITOR","BITXOR","CONCAT","MEDIAN_","PMT_","ERF.PRECISE","GAMMA","STDEV_","SHEET","SHEETS","XOR","ARRAYTOTEXT","VALUETOTEXT","TAKE","DROP","VSTACK","HSTACK","TOCOL","TOROW","CHOOSECOLS","CHOOSEROWS","EXPAND","WRAPCOLS","WRAPROWS"
    };

    static string Rebuild(string f, List<Token> toks, Func<Token, string?> map)
    {
        var sb = new StringBuilder(); int last = 0;
        foreach (var t in toks)
        {
            if (t.Kind == T.End) break;
            string? r = map(t);
            if (r == null) continue;
            sb.Append(f, last, t.Pos - last).Append(r);
            last = t.Pos + t.Len;
        }
        sb.Append(f, last, f.Length - last);
        return sb.ToString();
    }

    /// <summary>Sposta i riferimenti relativi (copia/incolla, riempimento). Fuori foglio → #REF!.</summary>
    public static string Shift(string formula, int dRow, int dCol)
    {
        if (dRow == 0 && dCol == 0) return formula;
        List<Token> toks;
        try { toks = Lexer.Tokenize(formula); } catch (FormatException) { return formula; }
        return Rebuild(formula, toks, t => t.Kind == T.Ref ? (t.Ref!.Shifted(dRow, dCol)?.Text() ?? "#REF!") : null);
    }

    /// <summary>
    /// Testo digitato dall'utente → forma neutra (nomi inglesi, virgola): accetta anche nomi italiani (SOMMA, SE...) e il separatore ';'
    /// con la virgola decimale (=SOMMA(A1;A2;1,5)). Nel file restano sempre nomi inglesi e virgola.
    /// </summary>
    public static string FromUser(string text)
    {
        bool semi = HasListSemicolon(text);
        string s = semi ? ItalianDecimals(text) : text;
        List<Token> toks;
        try { toks = Lexer.Tokenize(s); } catch (FormatException) { return text; }
        int brace = 0;
        var depthAt = new Dictionary<int, bool>();
        foreach (var t in toks) { if (t.Kind == T.LB) brace++; else if (t.Kind == T.RB) brace--; else if (t.Kind == T.Semi) depthAt[t.Pos] = brace > 0; }
        return Rebuild(s, toks, t =>
        {
            switch (t.Kind)
            {
                case T.Semi when semi && !depthAt[t.Pos]: return ",";
                case T.Func:
                    {
                        string n = Parser.Normalize(t.Text);
                        if (!Fns.Exists(n) && FunctionCatalog.EnglishOf(n) is { } en) return en;
                        return null;
                    }
                case T.Name when t.Text.Equals("VERO", StringComparison.OrdinalIgnoreCase): return "TRUE";
                case T.Name when t.Text.Equals("FALSO", StringComparison.OrdinalIgnoreCase): return "FALSE";
                default: return null;
            }
        });
    }

    static bool HasListSemicolon(string s)
    {
        bool q = false; int brace = 0;
        foreach (var c in s) { if (c == '"') q = !q; else if (!q) { if (c == '{') brace++; else if (c == '}') brace--; else if (c == ';' && brace == 0) return true; } }
        return false;
    }

    static string ItalianDecimals(string s)
    {
        var sb = new StringBuilder(s.Length); bool q = false;
        for (int i = 0; i < s.Length; i++)
        {
            char c = s[i];
            if (c == '"') q = !q;
            if (!q && c == ',' && i > 0 && i + 1 < s.Length && char.IsAsciiDigit(s[i - 1]) && char.IsAsciiDigit(s[i + 1])) sb.Append('.'); else sb.Append(c);
        }
        return sb.ToString();
    }

    /// <summary>Forma utente → forma file: aggiunge _xlfn. alle funzioni moderne, maiuscolo sulle funzioni.</summary>
    public static string ToFile(string formula)
    {
        List<Token> toks;
        try { toks = Lexer.Tokenize(formula); } catch (FormatException) { return formula; }
        return Rebuild(formula, toks, t =>
        {
            if (t.Kind != T.Func) return null;
            string n = Parser.Normalize(t.Text);
            return Future.Contains(n) ? "_xlfn." + n : n;
        });
    }

    /// <summary>Forma file → forma di visualizzazione (toglie _xlfn.).</summary>
    public static string ToDisplay(string formula)
    {
        if (!formula.Contains("_xl", StringComparison.OrdinalIgnoreCase)) return formula;
        List<Token> toks;
        try { toks = Lexer.Tokenize(formula); } catch (FormatException) { return formula; }
        return Rebuild(formula, toks, t => t.Kind == T.Func ? Parser.Normalize(t.Text) : null);
    }

    /// <summary>Aggiorna i riferimenti quando un foglio viene rinominato.</summary>
    public static string RenameSheet(string formula, string oldName, string newName)
    {
        List<Token> toks;
        try { toks = Lexer.Tokenize(formula); } catch (FormatException) { return formula; }
        return Rebuild(formula, toks, t =>
        {
            if (t.Kind != T.Ref || t.Ref!.Sheet == null || !string.Equals(t.Ref.Sheet, oldName, StringComparison.OrdinalIgnoreCase)) return null;
            t.Ref.Sheet = newName; return t.Ref.Text();
        });
    }

    /// <summary>I riferimenti a un foglio eliminato diventano #REF!.</summary>
    public static string DeleteSheet(string formula, string name)
    {
        List<Token> toks;
        try { toks = Lexer.Tokenize(formula); } catch (FormatException) { return formula; }
        return Rebuild(formula, toks, t => t.Kind == T.Ref && t.Ref!.Sheet != null && string.Equals(t.Ref.Sheet, name, StringComparison.OrdinalIgnoreCase) ? "#REF!" : null);
    }

    /// <summary>
    /// Mappa un intervallo di indici [lo,hi] dopo inserimento (count&gt;0) o eliminazione (count&lt;0) di |count| righe/colonne da <paramref name="at"/>.
    /// null = intervallo interamente eliminato.
    /// </summary>
    public static (int Lo, int Hi)? MapSpan(int lo, int hi, int at, int count)
    {
        if (count > 0) return (lo >= at ? lo + count : lo, hi >= at ? hi + count : hi);
        int n = -count, end = at + n - 1;
        if (lo >= at && hi <= end) return null;
        int nl = lo < at ? lo : lo <= end ? at : lo - n;
        int nh = hi < at ? hi : hi <= end ? at - 1 : hi - n;
        return (nl, nh);
    }

    /// <summary>Aggiorna i riferimenti a <paramref name="target"/> (foglio <paramref name="ctx"/> = foglio della formula) dopo inserimento/eliminazione di righe o colonne.</summary>
    public static string AdjustLines(string formula, string ctx, string target, bool rows, int at, int count)
    {
        List<Token> toks;
        try { toks = Lexer.Tokenize(formula); } catch (FormatException) { return formula; }
        return Rebuild(formula, toks, t =>
        {
            if (t.Kind != T.Ref) return null;
            var r = t.Ref!;
            if (!string.Equals(r.Sheet ?? ctx, target, StringComparison.OrdinalIgnoreCase)) return null;
            int a = rows ? r.A.Row : r.A.Col, b = r.IsRange ? (rows ? r.B.Row : r.B.Col) : a;
            if (a == 0) return null; // l'asse non è interessato (es. colonna intera con righe)
            var m = MapSpan(Math.Min(a, b), Math.Max(a, b), at, count);
            if (m == null) return "#REF!";
            var (lo, hi) = m.Value;
            if (hi < lo || hi > (rows ? Addr.MaxRow : Addr.MaxCol)) return "#REF!";
            // mantiene l'ordine originale degli estremi
            int na = a <= b ? lo : hi, nb = a <= b ? hi : lo;
            var n = new RefToken { Sheet = r.Sheet, IsRange = r.IsRange, A = r.A, B = r.B };
            if (rows) { n.A.Row = na; n.B.Row = nb; } else { n.A.Col = na; n.B.Col = nb; }
            return n.Text();
        });
    }

    /// <summary>Aggiorna un elenco di intervalli separati da spazio (sqref); vuoto se tutto eliminato.</summary>
    public static string AdjustSqref(string sqref, bool rows, int at, int count) =>
        string.Join(' ', sqref.Split(' ', StringSplitOptions.RemoveEmptyEntries).Select(x => AdjustLines(x, "", "", rows, at, count)).Where(x => !x.Contains("#REF!")));
}
