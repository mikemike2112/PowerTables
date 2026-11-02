using System.Globalization;

namespace PowerTables.Core;

public sealed partial class Workbook
{
    // ponytail: ricerca nei valori visualizzati e nel testo delle formule (non nei commenti); scansione lineare dello store sparso → indice se serve
    string FindText(Sheet s, int r, int c, CultureInfo? ci) => s.Get(r, c).F != null ? EditText(s, r, c, ci) : DisplayText(s, r, c, ci);

    static bool Matches(string text, string find, bool matchCase, bool whole)
    {
        var cmp = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        return whole ? string.Equals(text, find, cmp) : text.Contains(find, cmp);
    }

    /// <summary>Prossima cella che contiene il testo, in ordine di riga, dopo (row,col) e con giro completo; null se nessuna.</summary>
    public (int Row, int Col)? Find(Sheet s, string find, int row, int col, bool forward = true, bool matchCase = false, bool whole = false, CultureInfo? ci = null)
    {
        if (find.Length == 0) return null;
        var hits = new List<long>();
        foreach (var kv in s.CellMap)
            if (kv.Value.Kind != CellKind.Empty || kv.Value.F != null)
            {
                int r = Addr.Row(kv.Key), c = Addr.Col(kv.Key);
                if (Matches(FindText(s, r, c, ci), find, matchCase, whole)) hits.Add(kv.Key);
            }
        if (hits.Count == 0) return null;
        // ordine riga poi colonna (Key non lo garantisce)
        hits.Sort((a, b) => Addr.Row(a) != Addr.Row(b) ? Addr.Row(a).CompareTo(Addr.Row(b)) : Addr.Col(a).CompareTo(Addr.Col(b)));
        long cur = (long)row * 100000 + col;
        long Ord(long k) => (long)Addr.Row(k) * 100000 + Addr.Col(k);
        long pick = forward ? hits.FirstOrDefault(k => Ord(k) > cur, hits[0]) : hits.LastOrDefault(k => Ord(k) < cur, hits[^1]);
        return (Addr.Row(pick), Addr.Col(pick));
    }

    /// <summary>Sostituisce in tutto il foglio (o solo nell'intervallo dato); restituisce il numero di celle modificate. Un solo Ctrl+Z.</summary>
    public int ReplaceAll(Sheet s, string find, string repl, bool matchCase = false, bool whole = false, (int R1, int C1, int R2, int C2)? area = null, CultureInfo? ci = null)
    {
        if (find.Length == 0) return 0;
        var cmp = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        var targets = new List<(int R, int C, string New)>();
        foreach (var kv in s.CellMap)
        {
            int r = Addr.Row(kv.Key), c = Addr.Col(kv.Key);
            if (area is { } a && (r < a.R1 || r > a.R2 || c < a.C1 || c > a.C2)) continue;
            if (kv.Value.Kind == CellKind.Empty && kv.Value.F == null) continue;
            var cell = kv.Value;
            // le formule opache (matriciali ecc.) non si toccano
            if (cell.F is { Opaque: true }) continue;
            string t = FindText(s, r, c, ci);
            if (!Matches(t, find, matchCase, whole)) continue;
            string nt = whole ? repl : t.Replace(find, repl, cmp);
            if (nt != t) targets.Add((r, c, cell.Kind == CellKind.Str && cell.F == null ? "'" + nt : nt));
        }
        if (targets.Count == 0) return 0;
        Do("Sostituisci", () =>
        {
            foreach (var (r, c, nt) in targets)
            {
                // il prefisso ' forza il testo per le stringhe; vale anche per testo che sembra un numero
                try { SetInput(s, r, c, nt, ci); } catch (FormatException) { SetValue(s, r, c, nt.StartsWith('\'') ? nt[1..] : nt); }
            }
        });
        return targets.Count;
    }
}
