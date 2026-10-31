using System.Text;

namespace PowerTables.Core;

/// <summary>Utility per indirizzi di cella (righe e colonne a base 1).</summary>
public static class Addr
{
    public const int MaxRow = 1048576, MaxCol = 16384;

    public static long Key(int row, int col) => ((long)row << 15) | (uint)col;
    public static int Row(long key) => (int)(key >> 15);
    public static int Col(long key) => (int)(key & 0x7FFF);

    public static string ColName(int col)
    {
        var sb = new StringBuilder(3);
        while (col > 0) { col--; sb.Insert(0, (char)('A' + col % 26)); col /= 26; }
        return sb.ToString();
    }

    /// <summary>Indice colonna da lettere ("AB" → 28); 0 se non valido.</summary>
    public static int ColIndex(ReadOnlySpan<char> s)
    {
        if (s.Length is 0 or > 3) return 0;
        int n = 0;
        foreach (var ch in s)
        {
            char u = char.ToUpperInvariant(ch);
            if (u < 'A' || u > 'Z') return 0;
            n = n * 26 + (u - 'A' + 1);
        }
        return n <= MaxCol ? n : 0;
    }

    public static string A1(int row, int col) => ColName(col) + row;

    /// <summary>Interpreta "A1" / "$B$2". Non accetta intervalli.</summary>
    public static bool TryParse(ReadOnlySpan<char> a, out int row, out int col)
    {
        row = col = 0;
        int i = 0;
        if (i < a.Length && a[i] == '$') i++;
        int s = i;
        while (i < a.Length && char.IsAsciiLetter(a[i])) i++;
        col = ColIndex(a[s..i]);
        if (col == 0) return false;
        if (i < a.Length && a[i] == '$') i++;
        int d = i;
        while (i < a.Length && char.IsAsciiDigit(a[i])) i++;
        if (i != a.Length || i == d || i - d > 7) return false;
        row = int.Parse(a[d..i]);
        return row is >= 1 and <= MaxRow;
    }

    /// <summary>Interpreta "A1:B3" oppure "A1".</summary>
    public static bool TryParseRange(string a, out int r1, out int c1, out int r2, out int c2)
    {
        r2 = c2 = 0;
        int p = a.IndexOf(':');
        if (p < 0) { var ok = TryParse(a, out r1, out c1); r2 = r1; c2 = c1; return ok; }
        if (!TryParse(a.AsSpan(0, p), out r1, out c1) || !TryParse(a.AsSpan(p + 1), out r2, out c2)) return false;
        if (r2 < r1) (r1, r2) = (r2, r1);
        if (c2 < c1) (c1, c2) = (c2, c1);
        return true;
    }
}
