using PowerTables.Core.Formulas;

namespace PowerTables.Core;

/// <summary>Nuovo foglio, elimina foglio, sposta foglio: tutto annullabile con un solo Ctrl+Z.</summary>
public sealed partial class Workbook
{
    /// <summary>Applica un nuovo ordine ai fogli rimappando i nomi locali (e scartando quelli dei fogli tolti).</summary>
    void ApplyOrder(List<Sheet> order, Dictionary<(int, string), string> locals, Dictionary<(int, string), string> localCase)
    {
        Sheets.Clear(); Sheets.AddRange(order);
        for (int i = 0; i < Sheets.Count; i++) Sheets[i].Index = i;
        Replace(LocalNames, locals); Replace(LocalNameCase, localCase);
        NamesDirty = true; InvalidateGraph(); _formulasChanged = true; FormulasChangedSinceOpen = true;
        foreach (var sh in Sheets) foreach (var kv in sh.Cells) if (kv.Value.F != null) _changed.Add((sh, kv.Key));
    }

    /// <summary>Calcola i nomi locali dopo un riordino: la chiave segue il foglio.</summary>
    (Dictionary<(int, string), string> L, Dictionary<(int, string), string> C) RemapLocals(List<Sheet> order)
    {
        var byOld = Sheets.ToDictionary(x => x.Index);
        var l = new Dictionary<(int, string), string>(); var c = new Dictionary<(int, string), string>();
        foreach (var kv in LocalNames)
            if (byOld.TryGetValue(kv.Key.Item1, out var sh) && order.IndexOf(sh) is var ni && ni >= 0)
            {
                l[(ni, kv.Key.Item2)] = kv.Value;
                if (LocalNameCase.TryGetValue(kv.Key, out var cs)) c[(ni, kv.Key.Item2)] = cs;
            }
        return (l, c);
    }

    void ReorderUndoable(string label, List<Sheet> newOrder, Action? extra = null, Action? extraUndo = null)
    {
        var oldOrder = Sheets.ToList();
        var oldL = new Dictionary<(int, string), string>(LocalNames); var oldC = new Dictionary<(int, string), string>(LocalNameCase);
        var (nl, nc) = RemapLocals(newOrder);
        Do(label, () =>
        {
            extra?.Invoke();
            ApplyOrder(newOrder, nl, nc);
            AddUndo(() => { ApplyOrder(oldOrder, oldL, oldC); extraUndo?.Invoke(); }, () => { extra?.Invoke(); ApplyOrder(newOrder, nl, nc); });
        });
    }

    /// <summary>Aggiunge un foglio (alla fine o prima di <paramref name="at"/>) con un solo passo di annullamento.</summary>
    public Sheet InsertSheet(string? name = null, int? at = null)
    {
        name = name?.Trim();
        if (name != null) ValidateSheetName(name, null);
        var s = new Sheet(name ?? NextSheetName()) { Book = this };
        var order = Sheets.ToList(); order.Insert(Math.Clamp(at ?? order.Count, 0, order.Count), s);
        ReorderUndoable("Nuovo foglio", order);
        return s;
    }

    /// <summary>Elimina il foglio: i riferimenti a esso nelle formule e nei nomi diventano #REF!.</summary>
    public void DeleteSheet(Sheet s)
    {
        if (Sheets.Count <= 1) throw new InvalidOperationException("Non si può eliminare l'unico foglio della cartella di lavoro.");
        var fch = new List<(Sheet S, long K, FormulaInfo Old, FormulaInfo New)>();
        foreach (var sh in Sheets)
        {
            if (ReferenceEquals(sh, s)) continue;
            foreach (var kv in sh.Cells)
                if (kv.Value.F is { Opaque: false } f)
                {
                    string nt = FormulaText.DeleteSheet(f.Text, s.Name);
                    if (nt != f.Text) fch.Add((sh, kv.Key, f, new FormulaInfo { Text = nt, CellAttrs = f.CellAttrs }));
                }
        }
        var oldN = new Dictionary<string, string>(Names);
        void Formulas(bool redo)
        {
            foreach (var (sh, k, o, n) in fch) { var c = sh.Get(Addr.Row(k), Addr.Col(k)); c.F = redo ? n : o; sh.Cells[k] = c; sh.MarkRowDirty(Addr.Row(k)); }
            if (redo) foreach (var k in Names.Keys.ToList()) Names[k] = FormulaText.DeleteSheet(Names[k], s.Name);
            else Replace(Names, oldN);
        }
        var order = Sheets.Where(x => !ReferenceEquals(x, s)).ToList();
        ReorderUndoable("Elimina foglio", order, () => Formulas(true), () => Formulas(false));
    }

    /// <summary>Sposta il foglio di <paramref name="delta"/> posizioni (negativo = verso sinistra).</summary>
    public void MoveSheet(Sheet s, int delta)
    {
        int i = Sheets.IndexOf(s), j = Math.Clamp(i + delta, 0, Sheets.Count - 1);
        if (i < 0 || i == j) return;
        var order = Sheets.ToList(); order.RemoveAt(i); order.Insert(j, s);
        ReorderUndoable("Sposta foglio", order);
    }

    void ValidateSheetName(string name, Sheet? self)
    {
        if (name.Length == 0 || name.Length > 31 || name.IndexOfAny([':', '\\', '/', '?', '*', '[', ']']) >= 0 || name.StartsWith('\'') || name.EndsWith('\'')) throw new ArgumentException("Nome foglio non valido.");
        if (Sheets.Any(x => x != self && string.Equals(x.Name, name, StringComparison.OrdinalIgnoreCase))) throw new ArgumentException("Esiste già un foglio con questo nome.");
    }
}
