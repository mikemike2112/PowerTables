using System.Globalization;
using PowerTables.Core.Formulas;

namespace PowerTables.Core;

public sealed partial class Workbook
{
    /// <summary>Copia formato: applica lo stile delle celle sorgente (ripetuto a piastrelle) all'area di destinazione, senza toccare i contenuti.</summary>
    public void PasteFormats(Sheet src, int sr1, int sc1, int sr2, int sc2, Sheet dst, int r, int c, int r2, int c2)
    {
        int h = sr2 - sr1 + 1, w = sc2 - sc1 + 1;
        Do("Copia formato", () =>
        {
            for (int i = r; i <= r2; i++)
                for (int j = c; j <= c2; j++)
                {
                    uint st = src.Get(sr1 + (i - r) % h, sc1 + (j - c) % w).Style;
                    var cell = dst.Get(i, j);
                    if (cell.Style == st) continue;
                    cell.Style = st; Write(dst, i, j, cell);
                }
        });
    }

    /// <summary>Rimuove le righe duplicate (confronto sulle colonne indicate, testo senza distinzione di maiuscole come Excel).
    /// Le righe rimaste salgono, in fondo all'area le celle si svuotano. Restituisce il numero di righe rimosse.</summary>
    public int RemoveDuplicates(Sheet s, int r1, int c1, int r2, int c2, IReadOnlyList<int> keyCols, bool hasHeader)
    {
        if (keyCols.Count == 0) throw new ArgumentException("Selezionare almeno una colonna.");
        int first = hasHeader ? r1 + 1 : r1;
        if (first >= r2) return 0;
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var keep = new List<int>();
        for (int r = first; r <= r2; r++)
        {
            string key = string.Join('\u0001', keyCols.Select(c => KeyOf(s.Get(r, c))));
            if (seen.Add(key)) keep.Add(r);
        }
        int removed = (r2 - first + 1) - keep.Count;
        if (removed == 0) return 0;
        Do("Rimuovi duplicati", () =>
        {
            var snap = keep.Select(r => Enumerable.Range(c1, c2 - c1 + 1).Select(c => s.Get(r, c)).ToArray()).ToList();
            for (int i = 0; i < r2 - first + 1; i++)
                for (int c = c1; c <= c2; c++)
                {
                    int dst = first + i;
                    if (i >= keep.Count) { Write(s, dst, c, new Cell { Style = s.Get(dst, c).Style }); continue; }
                    var nc = snap[i][c - c1];
                    if (nc.F is { Opaque: false } f && keep[i] != dst) nc.F = new FormulaInfo { Text = FormulaText.Shift(f.Text, dst - keep[i], 0) };
                    else if (nc.F is { Opaque: true } && keep[i] != dst) nc.F = null;
                    Write(s, dst, c, nc);
                }
        });
        return removed;
    }

    static string KeyOf(Cell c) => c.Kind switch
    {
        CellKind.Num => "n" + c.Num.ToString("R", CultureInfo.InvariantCulture),
        CellKind.Str => "s" + ((string?)c.Obj ?? "").ToUpperInvariant(),
        CellKind.Bool => "b" + c.Num,
        CellKind.Err => "e" + c.Num,
        _ => "",
    };

    /// <summary>Testo in colonne (delimitato): divide il testo della colonna <paramref name="col"/> nelle colonne a destra.
    /// Restituisce il numero massimo di parti. I valori sono interpretati come se digitati.</summary>
    public int TextToColumns(Sheet s, int r1, int col, int r2, string delimiters, bool mergeConsecutive = false, char qualifier = '"', CultureInfo? ci = null)
    {
        if (delimiters.Length == 0) throw new ArgumentException("Indicare almeno un delimitatore.");
        var rowsParts = new List<List<string>?>();
        int max = 1;
        for (int r = r1; r <= r2; r++)
        {
            var cell = s.Get(r, col);
            if (cell.Kind == CellKind.Empty || cell.F != null) { rowsParts.Add(null); continue; }
            var parts = SplitText(DisplayText(s, r, col, ci), delimiters, mergeConsecutive, qualifier);
            rowsParts.Add(parts); max = Math.Max(max, parts.Count);
        }
        if (col + max - 1 > Addr.MaxCol) throw new InvalidOperationException("Non ci sono abbastanza colonne a destra.");
        Do("Testo in colonne", () =>
        {
            for (int i = 0; i < rowsParts.Count; i++)
            {
                var p = rowsParts[i]; if (p == null) continue;
                int r = r1 + i;
                for (int j = 0; j < p.Count; j++)
                {
                    if (p[j].Length == 0 && j > 0) ClearCell(s, r, col + j);
                    else SetInputSafe(s, r, col + j, p[j], ci);
                }
            }
        });
        return max;
    }

    public static List<string> SplitText(string text, string delimiters, bool mergeConsecutive, char qualifier)
    {
        var res = new List<string>(); var sb = new System.Text.StringBuilder(); bool inQ = false, lastWasDelim = false;
        for (int i = 0; i < text.Length; i++)
        {
            char ch = text[i];
            if (inQ)
            {
                if (ch == qualifier) { if (i + 1 < text.Length && text[i + 1] == qualifier) { sb.Append(ch); i++; } else inQ = false; }
                else sb.Append(ch);
                continue;
            }
            if (ch == qualifier && qualifier != '\0' && sb.Length == 0) { inQ = true; lastWasDelim = false; continue; }
            if (delimiters.Contains(ch))
            {
                if (!(mergeConsecutive && lastWasDelim)) { res.Add(sb.ToString()); sb.Clear(); }
                lastWasDelim = true; continue;
            }
            lastWasDelim = false; sb.Append(ch);
        }
        res.Add(sb.ToString());
        return res;
    }
}
