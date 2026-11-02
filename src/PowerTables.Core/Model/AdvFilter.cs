using System.Globalization;
using System.Text.RegularExpressions;

namespace PowerTables.Core;

public sealed partial class Workbook
{
    /// <summary>Un criterio del filtro avanzato: operatore opzionale (=, &lt;&gt;, &gt;, &lt;, &gt;=, &lt;=) e valore. Senza operatore: numero = uguale, testo = «inizia con» (come Excel); * e ? sono caratteri jolly.</summary>
    static bool CriterionMatches(string crit, object? v, CultureInfo ci)
    {
        crit = crit.Trim();
        string op = "";
        foreach (var o in new[] { ">=", "<=", "<>", ">", "<", "=" }) if (crit.StartsWith(o, StringComparison.Ordinal)) { op = o; crit = crit[o.Length..].Trim(); break; }
        bool isNum = double.TryParse(crit, NumberStyles.Float, ci, out var cn);
        if (crit.Length == 0) return op == "<>" ? v != null : op == "=" ? v == null : true;   // "=" vuoto: celle vuote; "<>" vuoto: non vuote; solo criterio vuoto: nessun vincolo
        if (isNum && v is double d)
            return op switch { ">" => d > cn, "<" => d < cn, ">=" => d >= cn, "<=" => d <= cn, "<>" => d != cn, _ => d == cn };
        string t = v switch { null => "", double dv => dv.ToString(ci), bool b => b ? "TRUE" : "FALSE", _ => v.ToString() ?? "" };
        if (op is ">" or "<" or ">=" or "<=")
        {
            int cmp = string.Compare(t, crit, ci, CompareOptions.IgnoreCase);
            return op switch { ">" => cmp > 0, "<" => cmp < 0, ">=" => cmp >= 0, _ => cmp <= 0 };
        }
        // testo: confronto senza maiuscole; con "=" o "<>" è esatto, senza operatore è «inizia con»
        string rx = "^" + Regex.Escape(crit).Replace("\\*", ".*").Replace("\\?", ".") + (op.Length == 0 ? "" : "$");
        bool m = Regex.IsMatch(t, rx, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));
        return op == "<>" ? !m : m;
    }

    /// <summary>Righe dell'elenco (escluse intestazioni) che soddisfano i criteri: le righe sotto l'intestazione dei criteri sono alternative (O), le colonne con la stessa intestazione si combinano (E).</summary>
    public List<int> AdvancedFilterRows(Sheet s, (int R1, int C1, int R2, int C2) list, (int R1, int C1, int R2, int C2) crit, bool unique, CultureInfo? ci = null)
    {
        ci ??= CultureInfo.CurrentCulture;
        if (list.R2 <= list.R1) throw new ArgumentException("L'elenco deve avere un'intestazione e almeno una riga di dati.");
        if (crit.R2 < crit.R1 || crit.C2 < crit.C1) throw new ArgumentException("Intervallo criteri non valido.");
        if (crit.R2 == crit.R1) throw new ArgumentException("L'intervallo criteri deve avere l'intestazione e almeno una riga di criteri.");
        // colonna dell'elenco per ogni colonna dei criteri
        var heads = new Dictionary<string, int>(StringComparer.CurrentCultureIgnoreCase);
        for (int c = list.C1; c <= list.C2; c++) if (DisplayText(s, list.R1, c, ci) is { Length: > 0 } h) heads.TryAdd(h, c);
        var map = new List<(int CritCol, int ListCol)>();
        for (int c = crit.C1; c <= crit.C2; c++)
        {
            string h = DisplayText(s, crit.R1, c, ci); if (h.Length == 0) continue;
            if (!heads.TryGetValue(h, out var lc)) throw new ArgumentException($"L'intestazione dei criteri «{h}» non esiste nell'elenco.");
            map.Add((c, lc));
        }
        if (map.Count == 0) throw new ArgumentException("Nessuna intestazione nei criteri.");
        var rows = new List<int>(); var seen = new HashSet<string>();
        for (int r = list.R1 + 1; r <= list.R2; r++)
        {
            bool any = false, hasCrit = false;
            for (int cr = crit.R1 + 1; cr <= crit.R2 && !any; cr++)
            {
                bool all = true, rowHas = false;
                foreach (var (cc, lc) in map)
                {
                    string ct = DisplayText(s, cr, cc, ci); if (ct.Length == 0) continue;
                    rowHas = true;
                    if (!CriterionMatches(ct, s.GetValue(r, lc), ci)) { all = false; break; }
                }
                if (rowHas) hasCrit = true;
                if (all && rowHas) any = true;
            }
            if (!hasCrit) any = true;   // riga di criteri tutta vuota: passa tutto
            if (!any) continue;
            if (unique && !seen.Add(string.Join("\u0001", Enumerable.Range(list.C1, list.C2 - list.C1 + 1).Select(c => DisplayText(s, r, c, ci).ToLowerInvariant())))) continue;
            rows.Add(r);
        }
        return rows;
    }

    /// <summary>Filtro avanzato sul posto: nasconde le righe che non soddisfano i criteri (un solo Annulla). Il filtro automatico sullo stesso elenco va tolto prima.</summary>
    public int AdvancedFilterInPlace(Sheet s, (int R1, int C1, int R2, int C2) list, (int R1, int C1, int R2, int C2) crit, bool unique, CultureInfo? ci = null)
    {
        if (s.Filter != null) throw new InvalidOperationException("C'è già un filtro automatico: rimuoverlo prima.");
        var keep = AdvancedFilterRows(s, list, crit, unique, ci).ToHashSet();
        var oldHidden = SnapshotHidden(s);
        Do("Filtro avanzato", () =>
        {
            void Apply(IReadOnlySet<int> show)
            {
                for (int r = list.R1 + 1; r <= list.R2; r++)
                {
                    bool hide = !show.Contains(r);
                    if (!hide && !s.Rows.ContainsKey(r)) continue;
                    var ri = s.RowInfoFor(r);
                    if (ri.Hidden != hide || ri.Filtered != hide) { ri.Hidden = hide; ri.Filtered = hide; s.MarkRowDirty(r); }
                }
                s.LayoutDirty = true;
            }
            Apply(keep);
            AddUndo(() => { RestoreHidden(s, oldHidden); s.LayoutDirty = true; }, () => Apply(keep));
        });
        return keep.Count;
    }

    /// <summary>Filtro avanzato con copia: scrive intestazioni e righe trovate (valori e formati) a partire da <paramref name="dest"/>. Non può sovrapporsi all'elenco.</summary>
    public int AdvancedFilterCopy(Sheet s, (int R1, int C1, int R2, int C2) list, (int R1, int C1, int R2, int C2) crit, bool unique, (int R, int C) dest, CultureInfo? ci = null)
    {
        var rows = AdvancedFilterRows(s, list, crit, unique, ci);
        int nc = list.C2 - list.C1 + 1, lastR = dest.R + rows.Count, lastC = dest.C + nc - 1;
        if (lastR > Addr.MaxRow || lastC > Addr.MaxCol) throw new ArgumentException("La destinazione esce dal foglio.");
        if (dest.R <= list.R2 && lastR >= list.R1 && dest.C <= list.C2 && lastC >= list.C1) throw new ArgumentException("La destinazione si sovrappone all'elenco.");
        if (dest.R <= crit.R2 && lastR >= crit.R1 && dest.C <= crit.C2 && lastC >= crit.C1) throw new ArgumentException("La destinazione si sovrappone ai criteri.");
        Do("Filtro avanzato (copia)", () =>
        {
            void Put(int dr, int sr)
            {
                for (int c = 0; c < nc; c++)
                {
                    var src = s.Get(sr, list.C1 + c); var nw = new Cell { Style = src.Style }; nw.SetValue(src.Value);
                    Write(s, dr, dest.C + c, nw);
                }
            }
            Put(dest.R, list.R1);
            for (int i = 0; i < rows.Count; i++) Put(dest.R + 1 + i, rows[i]);
        });
        return rows.Count;
    }

    /// <summary>Mostra di nuovo le righe nascoste dal filtro avanzato nell'intervallo (un solo Annulla); le righe nascoste a mano restano.</summary>
    public void ShowAllRows(Sheet s, int r1, int r2)
    {
        var rows = s.Rows.Where(kv => kv.Key >= r1 && kv.Key <= r2 && kv.Value.Filtered && kv.Value.Hidden).Select(kv => kv.Key).ToList();
        if (rows.Count == 0) return;
        Do("Mostra tutto", () =>
        {
            void Set(bool hide) { foreach (var r in rows) { var ri = s.RowInfoFor(r); ri.Hidden = hide; ri.Filtered = hide; s.MarkRowDirty(r); } s.LayoutDirty = true; }
            Set(false); AddUndo(() => Set(true), () => Set(false));
        });
    }
}
