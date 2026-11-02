using PowerTables.Core.Formulas;

namespace PowerTables.Core;

public sealed partial class Workbook
{
    /// <summary>Inserisce <paramref name="count"/> righe (o colonne) prima di <paramref name="at"/>, aggiornando riferimenti delle formule.</summary>
    public void InsertLines(Sheet s, bool rows, int at, int count = 1)
    {
        if (count < 1 || at < 1) return;
        int lim = rows ? Addr.MaxRow : Addr.MaxCol, used = rows ? s.MaxRow : s.MaxCol;
        if (at <= used && used + count > lim) throw new InvalidOperationException("Impossibile inserire: ci sono celle in fondo al foglio che andrebbero perse.");
        Lines(s, rows, at, count, rows ? "Inserisci righe" : "Inserisci colonne");
    }

    /// <summary>Elimina <paramref name="count"/> righe (o colonne) da <paramref name="at"/>; i riferimenti eliminati diventano #REF!.</summary>
    public void DeleteLines(Sheet s, bool rows, int at, int count = 1)
    {
        if (count < 1 || at < 1) return;
        Lines(s, rows, at, -count, rows ? "Elimina righe" : "Elimina colonne");
    }

    void Lines(Sheet s, bool rows, int at, int count, string name)
    {
        // istantanea per l'annullamento
        var oldCells = new Dictionary<long, Cell>(s.Cells);
        var oldRows = new Dictionary<int, RowInfo>(s.Rows);
        var oldCols = new Dictionary<int, ColInfo>(s.Cols);
        var oldMerges = new List<MergeRange>(s.Merges);
        var oldRich = new Dictionary<long, (string, string)>(s.Rich);
        var oldFilter = s.Filter;
        var oldFilterCopy = oldFilter == null ? null : new AutoFilterState { R1 = oldFilter.R1, C1 = oldFilter.C1, R2 = oldFilter.R2, C2 = oldFilter.C2, Allowed = new(oldFilter.Allowed), Dirty = true };
        int oldMaxR = s.MaxRow, oldMaxC = s.MaxCol;
        var oldCf = CloneList(s.CondFmt); var oldDv = CloneList(s.DataVal); var oldLk = CloneList(s.Links);

        int Map(int v) { var m = FormulaText.MapSpan(v, v, at, count); return m == null ? 0 : m.Value.Lo; }
        long MapKey(long k, out bool gone)
        {
            int r = Addr.Row(k), c = Addr.Col(k);
            int n = Map(rows ? r : c);
            gone = n == 0;
            return gone ? 0 : rows ? Addr.Key(n, c) : Addr.Key(r, n);
        }

        var op = new LineOp(s.Name, rows, at, count);
        var inv = new LineOp(s.Name, rows, at, -count);
        var oldNames = new Dictionary<string, string>(Names); var oldLocal = new Dictionary<(int, string), string>(LocalNames);
        var oldPrint = Sheets.Select(x => x.PrintArea).ToArray();
        void AdjustNames()
        {
            NamesDirty = true;
            foreach (var sh in Sheets) if (sh.PrintArea != null) sh.PrintArea = FormulaText.AdjustLines(sh.PrintArea, "", s.Name, rows, at, count);
            foreach (var k in Names.Keys.ToList()) Names[k] = FormulaText.AdjustLines(Names[k], "", s.Name, rows, at, count);
            foreach (var k in LocalNames.Keys.ToList()) LocalNames[k] = FormulaText.AdjustLines(LocalNames[k], "", s.Name, rows, at, count);
        }
        void RestoreNames() { Replace(Names, oldNames); Replace(LocalNames, oldLocal); for (int i = 0; i < oldPrint.Length && i < Sheets.Count; i++) Sheets[i].PrintArea = oldPrint[i]; }
        Begin(name);
        try
        {
            // formule di tutti i fogli
            var fchanges = new List<(Sheet S, long K, FormulaInfo Old, FormulaInfo New)>();
            foreach (var sh in Sheets)
                foreach (var kv in sh.Cells)
                    if (kv.Value.F is { Opaque: false } f)
                    {
                        string nt = FormulaText.AdjustLines(f.Text, sh.Name, s.Name, rows, at, count);
                        if (nt != f.Text) fchanges.Add((sh, kv.Key, f, new FormulaInfo { Text = nt, CellAttrs = f.CellAttrs }));
                    }
            void ApplyFormulas(bool redo)
            {
                foreach (var (sh, k, o, n) in fchanges)
                {
                    var c = sh.Get(Addr.Row(k), Addr.Col(k)); c.F = redo ? n : o; sh.Cells[k] = c;
                }
            }
            // ricostruzione del foglio
            void Remap()
            {
                var curCells = new Dictionary<long, Cell>(s.Cells); var curRich = new Dictionary<long, (string, string)>(s.Rich);
                var curRows = new Dictionary<int, RowInfo>(s.Rows); var curCols = new Dictionary<int, ColInfo>(s.Cols);
                var curMerges = new List<MergeRange>(s.Merges); var curFilter = s.Filter;
                var cells = new Dictionary<long, Cell>(curCells.Count);
                foreach (var kv in curCells) { long nk = MapKey(kv.Key, out bool gone); if (!gone) cells[nk] = kv.Value; }
                var rich = new Dictionary<long, (string, string)>();
                foreach (var kv in curRich) { long nk = MapKey(kv.Key, out bool gone); if (!gone) rich[nk] = kv.Value; }
                Replace(s.Cells, cells);
                Replace(s.Rich, rich);
                if (rows) { var d = new Dictionary<int, RowInfo>(); foreach (var kv in curRows) { int n = Map(kv.Key); if (n != 0) d[n] = kv.Value; } Replace(s.Rows, d); }
                else { var d = new Dictionary<int, ColInfo>(); foreach (var kv in curCols) { int n = Map(kv.Key); if (n != 0) d[n] = kv.Value; } Replace(s.Cols, d); }
                s.Merges.Clear();
                foreach (var m in curMerges)
                {
                    var sp = FormulaText.MapSpan(rows ? m.R1 : m.C1, rows ? m.R2 : m.C2, at, count);
                    if (sp == null) continue;
                    var nm = rows ? new MergeRange(sp.Value.Lo, m.C1, sp.Value.Hi, m.C2) : new MergeRange(m.R1, sp.Value.Lo, m.R2, sp.Value.Hi);
                    if (nm.R1 != nm.R2 || nm.C1 != nm.C2) s.Merges.Add(nm);
                }
                s.Filter = null;
                if (curFilter is { } f && FormulaText.MapSpan(rows ? f.R1 : f.C1, rows ? f.R2 : f.C2, at, count) is { } fs)
                {
                    var nf = new AutoFilterState { R1 = f.R1, C1 = f.C1, R2 = f.R2, C2 = f.C2, Dirty = true };
                    if (rows) { nf.R1 = fs.Lo; nf.R2 = fs.Hi; nf.Allowed = new(f.Allowed); }
                    else
                    {
                        nf.C1 = fs.Lo; nf.C2 = fs.Hi;
                        foreach (var kv in f.Allowed) { int n = Map(kv.Key); if (n != 0) nf.Allowed[n] = kv.Value; }
                    }
                    s.Filter = nf;
                }
                s.RecomputeExtent();
                AdjustSections(s, rows, at, count);
                Touch();
            }
            void Restore()
            {
                Replace(s.Cells, oldCells); Replace(s.Rich, oldRich); Replace(s.Rows, oldRows); Replace(s.Cols, oldCols);
                s.Merges.Clear(); s.Merges.AddRange(oldMerges);
                s.Filter = oldFilterCopy == null ? null : new AutoFilterState { R1 = oldFilterCopy.R1, C1 = oldFilterCopy.C1, R2 = oldFilterCopy.R2, C2 = oldFilterCopy.C2, Allowed = new(oldFilterCopy.Allowed), Dirty = true };
                s.MaxRow = oldMaxR; s.MaxCol = oldMaxC;
                SetList(s.CondFmt, oldCf); SetList(s.DataVal, oldDv); SetList(s.Links, oldLk); s.CfVersion++;
                Touch();
            }
            // righe da riscrivere nel file + ricalcolo di tutte le formule
            void Touch()
            {
                s.LayoutDirty = true;
                foreach (var k in oldCells.Keys) s.MarkRowDirty(Addr.Row(k));
                foreach (var k in s.Cells.Keys) s.MarkRowDirty(Addr.Row(k));
                foreach (var r in oldRows.Keys) s.MarkRowDirty(r);
                foreach (var r in s.Rows.Keys) s.MarkRowDirty(r);
                foreach (var (sh, k, _, _) in fchanges) sh.MarkRowDirty(Addr.Row(k));
                InvalidateGraph();
                _formulasChanged = true; FormulasChangedSinceOpen = true;
                foreach (var sh in Sheets) foreach (var kv in sh.Cells) if (kv.Value.F != null) _changed.Add((sh, kv.Key));
            }
            ApplyFormulas(true); Remap(); AdjustNames(); s.LineOps.Add(op);
            AddUndo(() =>
            {
                ApplyFormulas(false); Restore(); RestoreNames();
                if (s.LineOps.Count > 0 && ReferenceEquals(s.LineOps[^1], op)) s.LineOps.RemoveAt(s.LineOps.Count - 1); else s.LineOps.Add(inv); // ponytail: dopo un salvataggio l'annullo di un'eliminazione non riespande CF/convalida → memorizzare la regola originale
            }, () => { ApplyFormulas(true); Remap(); AdjustNames(); s.LineOps.Add(op); });
        }
        finally { End(); }
    }

    static void Replace<K, V>(Dictionary<K, V> d, Dictionary<K, V> src) where K : notnull
    {
        if (ReferenceEquals(d, src)) return;
        d.Clear(); foreach (var kv in src) d[kv.Key] = kv.Value;
    }
}
