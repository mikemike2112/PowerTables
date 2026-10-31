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
}
