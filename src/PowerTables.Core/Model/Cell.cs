using System.Globalization;

namespace PowerTables.Core;

/// <summary>Valori d'errore di Excel.</summary>
public enum CellError : byte { Null = 1, Div0, Value, Ref, Name, Num, NA, Spill, Calc }

public static class Errors
{
    static readonly string[] Texts = ["", "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#SPILL!", "#CALC!"];
    // Codici restituiti da Excel COM in Value2 per gli errori.
    static readonly int[] Codes = [0, -2146826288, -2146826281, -2146826273, -2146826265, -2146826259, -2146826252, -2146826246, -2146826180, -2146826177];

    public static string Text(CellError e) => Texts[(int)e];
    public static int ComCode(CellError e) => Codes[(int)e];
    public static bool TryParse(string s, out CellError e)
    {
        for (int i = 1; i < Texts.Length; i++)
            if (string.Equals(Texts[i], s, StringComparison.OrdinalIgnoreCase)) { e = (CellError)i; return true; }
        e = 0; return false;
    }
    public static bool TryFromComCode(int code, out CellError e)
    {
        for (int i = 1; i < Codes.Length; i++) if (Codes[i] == code) { e = (CellError)i; return true; }
        e = 0; return false;
    }
}

public enum CellKind : byte { Empty, Num, Str, Bool, Err }

/// <summary>Formula di una cella. Text è in forma OOXML invariante (virgole, nomi inglesi), senza '='.</summary>
public sealed class FormulaInfo
{
    public string Text = "";
    /// <summary>Elemento &lt;f&gt; originale da riscrivere identico (array, dataTable, dinamiche): la cella non viene ricalcolata.</summary>
    public string? RawF;
    /// <summary>Attributi extra dell'elemento &lt;c&gt; (cm, vm...) da preservare.</summary>
    public string? CellAttrs;
    public bool Opaque => RawF != null;
}

/// <summary>Cella: struct compatta per reggere milioni di celle in uno store sparso.</summary>
public struct Cell
{
    public double Num;          // numero, bool (0/1) o codice errore
    public object? Obj;         // stringa per Kind=Str
    public FormulaInfo? F;
    public uint Style;
    public CellKind Kind;

    public readonly bool IsEmpty => Kind == CellKind.Empty && F == null && Style == 0;

    public readonly object? Value => Kind switch
    {
        CellKind.Num => Num,
        CellKind.Str => (string?)Obj ?? "",
        CellKind.Bool => Num != 0,
        CellKind.Err => (CellError)(int)Num,
        _ => null,
    };

    public void SetValue(object? v)
    {
        switch (v)
        {
            case null: Kind = CellKind.Empty; Obj = null; Num = 0; break;
            case double d: Kind = CellKind.Num; Num = d; Obj = null; break;
            case string s: Kind = CellKind.Str; Obj = s; Num = 0; break;
            case bool b: Kind = CellKind.Bool; Num = b ? 1 : 0; Obj = null; break;
            case CellError e: Kind = CellKind.Err; Num = (int)e; Obj = null; break;
            case int i: Kind = CellKind.Num; Num = i; Obj = null; break;
            default: throw new ArgumentException("tipo valore non supportato: " + v.GetType());
        }
    }

    public static Cell FromValue(object? v, uint style = 0) { var c = new Cell { Style = style }; c.SetValue(v); return c; }

    public readonly bool SameValue(in Cell o) => Kind == o.Kind && Num.Equals(o.Num) && Equals(Obj, o.Obj);
}

/// <summary>Riferimento a un intervallo di una cartella (valore "per riferimento" nel motore formule).</summary>
public sealed class RangeRef(Sheet sheet, int r1, int c1, int r2, int c2)
{
    public readonly Sheet Sheet = sheet;
    public readonly int R1 = r1, C1 = c1, R2 = r2, C2 = c2;
    public int Rows => R2 - R1 + 1;
    public int Cols => C2 - C1 + 1;
    public bool IsSingle => R1 == R2 && C1 == C2;
    public object? At(int i, int j) => Sheet.GetValue(R1 + i, C1 + j);
}

/// <summary>Matrice di valori (risultato di operazioni array o costanti {..}).</summary>
public sealed class Arr(object?[,] v)
{
    public readonly object?[,] V = v;
    public int Rows => V.GetLength(0);
    public int Cols => V.GetLength(1);
}

public static class Fmt
{
    /// <summary>Numero → testo come la conversione "General" di Excel in concatenazione (15 cifre significative).</summary>
    public static string NumToText(double d)
    {
        if (d == 0) return "0";
        if (double.IsNaN(d) || double.IsInfinity(d)) return "#NUM!";
        string r = d.ToString("E14", CultureInfo.InvariantCulture); // d.ddddddddddddddE+xxx
        int e = r.IndexOf('E');
        int exp = int.Parse(r.AsSpan(e + 1), CultureInfo.InvariantCulture);
        string mant = r[..e].TrimEnd('0').TrimEnd('.');
        bool neg = mant.StartsWith('-');
        string digits = mant.Replace("-", "").Replace(".", "");
        string res;
        if (exp >= 15 || exp < -9)
        {
            string m = digits.Length > 1 ? digits[0] + "." + digits[1..] : digits;
            res = m + "E" + (exp < 0 ? "-" : "+") + Math.Abs(exp).ToString("00", CultureInfo.InvariantCulture);
        }
        else if (exp >= 0)
        {
            res = digits.Length <= exp + 1 ? digits.PadRight(exp + 1, '0') : digits[..(exp + 1)] + "." + digits[(exp + 1)..];
        }
        else res = "0." + new string('0', -exp - 1) + digits;
        return neg ? "-" + res : res;
    }
}
