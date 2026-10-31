using System.Globalization;
using PowerTables.Core.Formulas;

namespace PowerTables.Core;

/// <summary>Appunti interni: istantanea rettangolare di celle.</summary>
public sealed class ClipData
{
    public Cell[,] Cells = new Cell[0, 0];
    public Sheet Source = null!;
    public int Row, Col;
    public bool Cut;
    public int Rows => Cells.GetLength(0);
    public int Cols => Cells.GetLength(1);
}

public sealed partial class Workbook
{
    // ---------------- testo visualizzato ----------------
    public string CodeFor(uint style, CultureInfo ci) => Styles.Get(style) is { NumFmtId: 14, NumFmtCode: null } ? NumFmt.ShortDate(ci) : Styles.FormatCode(style);

    public string DisplayText(Sheet s, int r, int c, CultureInfo? ci = null)
    {
        ci ??= CultureInfo.CurrentCulture;
        var cell = s.Get(r, c);
        switch (cell.Kind)
        {
            case CellKind.Empty: return "";
            case CellKind.Num:
                {
                    if (cell.Style == 0) return NumFmt.General(cell.Num).Replace(".", ci.NumberFormat.NumberDecimalSeparator);
                    return NumFmt.Format(cell.Num, CodeFor(cell.Style, ci), ci);
                }
            case CellKind.Str:
                return cell.Style == 0 ? (string)cell.Obj! : NumFmt.Format((string)cell.Obj!, CodeFor(cell.Style, ci), ci);
            case CellKind.Bool: return cell.Num != 0 ? "TRUE" : "FALSE";
            default: return Errors.Text((CellError)(int)cell.Num);
        }
    }

    /// <summary>Testo da mostrare nella barra della formula / in modifica.</summary>
    public string EditText(Sheet s, int r, int c, CultureInfo? ci = null)
    {
        ci ??= CultureInfo.CurrentCulture;
        var cell = s.Get(r, c);
        if (cell.F != null) return "=" + FormulaText.ToDisplay(cell.F.Text);
        switch (cell.Kind)
        {
            case CellKind.Num:
                if (NumFmt.IsDate(Styles.FormatCode(cell.Style)) && cell.Style != 0) return NumFmt.Format(cell.Num, Styles.FormatCode(cell.Style), ci);
                return Fmt.NumToText(cell.Num).Replace(".", ci.NumberFormat.NumberDecimalSeparator);
            case CellKind.Str:
                {
                    var str = (string)cell.Obj!;
                    return str.Length > 0 && (str[0] is '=' or '\'' || ParseInput(str, ci).Value is not string) ? "'" + str : str;
                }
            case CellKind.Bool: return cell.Num != 0 ? "TRUE" : "FALSE";
            case CellKind.Err: return Errors.Text((CellError)(int)cell.Num);
            default: return "";
        }
    }

    // ---------------- copia / incolla ----------------
    public ClipData Copy(Sheet s, int r1, int c1, int r2, int c2, bool cut = false)
    {
        var clip = new ClipData { Source = s, Row = r1, Col = c1, Cut = cut, Cells = new Cell[r2 - r1 + 1, c2 - c1 + 1] };
        for (int i = r1; i <= r2; i++) for (int j = c1; j <= c2; j++) clip.Cells[i - r1, j - c1] = s.Get(i, j);
        return clip;
    }

    /// <summary>Incolla adattando i riferimenti relativi delle formule. Restituisce l'intervallo scritto.</summary>
    public (int R2, int C2) Paste(ClipData clip, Sheet s, int r, int c, bool valuesOnly = false)
    {
        int dr = r - clip.Row, dc = c - clip.Col;
        int r2 = Math.Min(Addr.MaxRow, r + clip.Rows - 1), c2 = Math.Min(Addr.MaxCol, c + clip.Cols - 1);
        Do("Incolla", () =>
        {
            for (int i = 0; i <= r2 - r; i++)
                for (int j = 0; j <= c2 - c; j++)
                {
                    var src = clip.Cells[i, j];
                    var nc = src;
                    if (src.F != null)
                    {
                        if (valuesOnly || src.F.Opaque) nc.F = null;
                        else nc.F = new FormulaInfo { Text = FormulaText.Shift(src.F.Text, dr, dc) };
                    }
                    if (valuesOnly) nc.Style = s.Get(r + i, c + j).Style;
                    Write(s, r + i, c + j, nc);
                }
            if (clip.Cut && ReferenceEquals(clip.Source, s) is var same)
            {
                for (int i = 0; i < clip.Rows; i++)
                    for (int j = 0; j < clip.Cols; j++)
                    {
                        int sr = clip.Row + i, sc = clip.Col + j;
                        if (sr >= r && sr <= r2 && sc >= c && sc <= c2 && same) continue;
                        Write(clip.Source, sr, sc, default);
                    }
            }
        });
        return (r2, c2);
    }

    /// <summary>Incolla testo tabulato (da appunti di sistema).</summary>
    public (int R2, int C2) PasteText(Sheet s, int r, int c, string text, CultureInfo? ci = null)
    {
        var lines = text.Replace("\r\n", "\n").TrimEnd('\n').Split('\n');
        int maxC = 0;
        Do("Incolla testo", () =>
        {
            for (int i = 0; i < lines.Length && r + i <= Addr.MaxRow; i++)
            {
                var parts = lines[i].Split('\t');
                maxC = Math.Max(maxC, parts.Length);
                for (int j = 0; j < parts.Length && c + j <= Addr.MaxCol; j++) SetInputSafe(s, r + i, c + j, parts[j], ci);
            }
        });
        return (r + lines.Length - 1, c + Math.Max(1, maxC) - 1);
    }

    void SetInputSafe(Sheet s, int r, int c, string text, CultureInfo? ci)
    {
        try { SetInput(s, r, c, text, ci); } catch (FormatException) { SetValue(s, r, c, text); }
    }

    // ---------------- riempimento automatico ----------------
    /// <summary>Estende l'intervallo sorgente fino a (toRow,toCol) nella direzione indicata.</summary>
    public void AutoFill(Sheet s, int r1, int c1, int r2, int c2, int toRow, int toCol)
    {
        bool vertical = toRow > r2 || toRow < r1;
        bool forward = vertical ? toRow > r2 : toCol > c2;
        Do("Riempi", () =>
        {
            if (vertical)
                for (int c = c1; c <= c2; c++) FillVector(s, Enumerable.Range(r1, r2 - r1 + 1).Select(r => (r, c)).ToList(), forward, forward ? toRow - r2 : r1 - toRow, true);
            else
                for (int r = r1; r <= r2; r++) FillVector(s, Enumerable.Range(c1, c2 - c1 + 1).Select(c => (r, c)).ToList(), forward, forward ? toCol - c2 : c1 - toCol, false);
        });
    }

    void FillVector(Sheet s, List<(int r, int c)> src, bool forward, int count, bool vertical)
    {
        var cells = src.Select(p => s.Get(p.r, p.c)).ToList();
        int n = cells.Count;
        bool allNum = cells.All(x => x.Kind == CellKind.Num && x.F == null);
        double step = 0;
        if (allNum && n >= 2) step = (cells[^1].Num - cells[0].Num) / (n - 1);
        var last = src[^1]; var first = src[0];
        for (int k = 1; k <= count; k++)
        {
            int tr = vertical ? (forward ? last.r + k : first.r - k) : first.r;
            int tc = vertical ? first.c : (forward ? last.c + k : first.c - k);
            if (tr < 1 || tc < 1 || tr > Addr.MaxRow || tc > Addr.MaxCol) break;
            // elemento sorgente corrispondente (ciclico)
            int idx = forward ? (k - 1) % n : (n - 1 - (k - 1) % n);
            var sc = cells[idx];
            var nc = new Cell { Style = sc.Style };
            int dr = tr - src[idx].r, dc = tc - src[idx].c;
            if (sc.F != null && !sc.F.Opaque) nc.F = new FormulaInfo { Text = FormulaText.Shift(sc.F.Text, dr, dc) };
            else if (allNum && n >= 2) nc.SetValue(forward ? cells[^1].Num + step * k : cells[0].Num - step * k);
            else if (allNum && n == 1 && s.Get(first.r, first.c) is var one && Styles.FormatCode(one.Style) is var code && NumFmt.IsDate(code) && one.Style != 0)
                nc.SetValue(one.Num + (forward ? k : -k));
            else if (sc.Kind == CellKind.Str && n == 1 && SplitTrailingNumber((string)sc.Obj!) is var (pre, num, width)) nc.SetValue(pre + (num + (forward ? k : -k)).ToString(new string('0', width), CultureInfo.InvariantCulture));
            else if (sc.F == null) { nc.Kind = sc.Kind; nc.Num = sc.Num; nc.Obj = sc.Obj; }
            Write(s, tr, tc, nc);
        }
    }

    static (string, int, int)? SplitTrailingNumber(string t)
    {
        int i = t.Length;
        while (i > 0 && char.IsAsciiDigit(t[i - 1])) i--;
        if (i == t.Length || t.Length - i > 9) return null;
        return (t[..i], int.Parse(t[i..], CultureInfo.InvariantCulture), t.Length - i);
    }

    // ---------------- ordinamento ----------------
    public void Sort(Sheet s, int r1, int c1, int r2, int c2, IReadOnlyList<(int Col, bool Asc)> keys, bool hasHeader)
    {
        int first = r1 + (hasHeader ? 1 : 0);
        if (first >= r2) return;
        int n = r2 - first + 1, w = c2 - c1 + 1;
        var rows = new Cell[n][];
        for (int i = 0; i < n; i++) { rows[i] = new Cell[w]; for (int j = 0; j < w; j++) rows[i][j] = s.Get(first + i, c1 + j); }
        var order = Enumerable.Range(0, n).ToArray();
        int Cmp(int a, int b)
        {
            foreach (var (col, asc) in keys)
            {
                var x = rows[a][col - c1]; var y = rows[b][col - c1];
                bool bx = x.Kind == CellKind.Empty, by = y.Kind == CellKind.Empty;
                if (bx || by) { if (bx && by) continue; return bx ? 1 : -1; } // vuote sempre in fondo
                int c = SortCompare(x, y);
                if (c != 0) return asc ? c : -c;
            }
            return a.CompareTo(b); // stabile
        }
        Array.Sort(order, Cmp);
        Do("Ordina", () =>
        {
            for (int i = 0; i < n; i++)
            {
                int from = order[i];
                for (int j = 0; j < w; j++)
                {
                    var cell = rows[from][j];
                    if (cell.F is { Opaque: false } f && from != i) cell.F = new FormulaInfo { Text = FormulaText.Shift(f.Text, i - from, 0) };
                    Write(s, first + i, c1 + j, cell);
                }
            }
        });
    }

    static int SortCompare(in Cell x, in Cell y)
    {
        static int Rank(in Cell c) => c.Kind switch { CellKind.Num => 0, CellKind.Str => 1, CellKind.Bool => 2, _ => 3 };
        int rx = Rank(x), ry = Rank(y);
        if (rx != ry) return rx.CompareTo(ry);
        return x.Kind switch
        {
            CellKind.Num => x.Num.CompareTo(y.Num),
            CellKind.Str => string.Compare((string)x.Obj!, (string)y.Obj!, StringComparison.CurrentCultureIgnoreCase),
            _ => x.Num.CompareTo(y.Num),
        };
    }

    // ---------------- filtri ----------------
    public void SetAutoFilter(Sheet s, int r1, int c1, int r2, int c2)
    {
        var old = s.Filter;
        var oldHidden = SnapshotHidden(s);
        Do("Filtro", () =>
        {
            s.Filter = new AutoFilterState { R1 = r1, C1 = c1, R2 = r2, C2 = c2, Dirty = true };
            s.LayoutDirty = true;
            AddUndo(() => { s.Filter = old; s.LayoutDirty = true; RestoreHidden(s, oldHidden); }, () => { s.Filter = new AutoFilterState { R1 = r1, C1 = c1, R2 = r2, C2 = c2, Dirty = true }; s.LayoutDirty = true; });
        });
    }

    public void ClearAutoFilter(Sheet s)
    {
        var old = s.Filter; var oldHidden = SnapshotHidden(s);
        if (old == null) return;
        Do("Rimuovi filtro", () =>
        {
            for (int r = old.R1 + 1; r <= old.R2; r++) if (s.Rows.TryGetValue(r, out var ri) && ri.Filtered) { ri.Filtered = false; ri.Hidden = false; s.MarkRowDirty(r); }
            s.Filter = null; s.LayoutDirty = true;
            AddUndo(() => { s.Filter = old; s.LayoutDirty = true; RestoreHidden(s, oldHidden); }, () => { s.Filter = null; s.LayoutDirty = true; });
        });
    }

    /// <summary>Imposta i valori ammessi per una colonna (null = nessun filtro) e rivaluta le righe.</summary>
    public void SetColumnFilter(Sheet s, int col, HashSet<string>? allowed, CultureInfo? ci = null)
    {
        var f = s.Filter ?? throw new InvalidOperationException("nessun filtro automatico");
        var oldHidden = SnapshotHidden(s);
        var oldAllowed = f.Allowed.ToDictionary(k => k.Key, k => k.Value);
        Do("Filtro colonna", () =>
        {
            if (allowed == null) f.Allowed.Remove(col); else f.Allowed[col] = allowed;
            f.Dirty = true; s.LayoutDirty = true;
            ReapplyFilter(s, ci);
            AddUndo(() => { f.Allowed = oldAllowed; f.Dirty = true; s.LayoutDirty = true; RestoreHidden(s, oldHidden); },
                    () => { if (allowed == null) f.Allowed.Remove(col); else f.Allowed[col] = allowed; f.Dirty = true; s.LayoutDirty = true; ReapplyFilter(s, ci); });
        });
    }

    public void ReapplyFilter(Sheet s, CultureInfo? ci = null)
    {
        var f = s.Filter; if (f == null) return;
        for (int r = f.R1 + 1; r <= f.R2; r++)
        {
            bool show = true;
            foreach (var (col, allowed) in f.Allowed)
                if (!allowed.Contains(DisplayText(s, r, col, ci))) { show = false; break; }
            var ri = s.RowInfoFor(r);
            if (ri.Filtered != !show || ri.Hidden != !show && ri.Filtered)
            {
                ri.Filtered = !show; ri.Hidden = !show; s.MarkRowDirty(r);
            }
        }
    }

    static List<int> SnapshotHidden(Sheet s) => s.Rows.Where(kv => kv.Value.Hidden || kv.Value.Filtered).Select(kv => kv.Key).ToList();

    static void RestoreHidden(Sheet s, List<int> hidden)
    {
        var set = hidden.ToHashSet();
        foreach (var kv in s.Rows)
        {
            bool h = set.Contains(kv.Key);
            if (kv.Value.Hidden != h || kv.Value.Filtered && !h) { kv.Value.Hidden = h; kv.Value.Filtered = h; s.MarkRowDirty(kv.Key); }
        }
    }

    /// <summary>Valori distinti (testo visualizzato) di una colonna del filtro.</summary>
    public List<string> DistinctValues(Sheet s, int col, CultureInfo? ci = null)
    {
        var f = s.Filter!; var set = new SortedSet<string>(StringComparer.CurrentCultureIgnoreCase);
        for (int r = f.R1 + 1; r <= f.R2; r++) set.Add(DisplayText(s, r, col, ci));
        return [.. set];
    }

    // ---------------- layout ----------------
    public void SetColWidth(Sheet s, int col, double width)
    {
        var old = s.Cols.TryGetValue(col, out var oi) ? new ColInfo { Width = oi.Width, Hidden = oi.Hidden, CustomWidth = oi.CustomWidth, Style = oi.Style } : null;
        Do("Larghezza colonna", () =>
        {
            var ci = s.ColInfoFor(col); ci.Width = width; ci.CustomWidth = true; s.LayoutDirty = true;
            AddUndo(() => { if (old == null) s.Cols.Remove(col); else s.Cols[col] = old; s.LayoutDirty = true; },
                    () => { var x = s.ColInfoFor(col); x.Width = width; x.CustomWidth = true; s.LayoutDirty = true; });
        });
    }

    public void SetRowHeight(Sheet s, int row, double height)
    {
        double? old = s.Rows.TryGetValue(row, out var ri) ? ri.Height : null; bool oc = ri?.CustomHeight ?? false;
        Do("Altezza riga", () =>
        {
            var x = s.RowInfoFor(row); x.Height = height; x.CustomHeight = true; s.MarkRowDirty(row);
            AddUndo(() => { var y = s.RowInfoFor(row); y.Height = old; y.CustomHeight = oc; s.MarkRowDirty(row); },
                    () => { var y = s.RowInfoFor(row); y.Height = height; y.CustomHeight = true; s.MarkRowDirty(row); });
        });
    }

    public void SetFreeze(Sheet s, int rows, int cols)
    {
        int or = s.FrozenRows, oc = s.FrozenCols;
        Do("Blocca riquadri", () =>
        {
            s.FrozenRows = rows; s.FrozenCols = cols; s.LayoutDirty = true;
            AddUndo(() => { s.FrozenRows = or; s.FrozenCols = oc; s.LayoutDirty = true; }, () => { s.FrozenRows = rows; s.FrozenCols = cols; s.LayoutDirty = true; });
        });
    }

    public void Merge(Sheet s, int r1, int c1, int r2, int c2)
    {
        if (r1 == r2 && c1 == c2) return;
        var removed = s.Merges.Where(m => m.Intersects(r1, c1, r2, c2)).ToList();
        var nm = new MergeRange(r1, c1, r2, c2);
        Do("Unisci celle", () =>
        {
            foreach (var m in removed) s.Merges.Remove(m);
            s.Merges.Add(nm);
            for (int r = r1; r <= r2; r++) for (int c = c1; c <= c2; c++) if (r != r1 || c != c1) { var x = s.Get(r, c); if (x.Kind != CellKind.Empty || x.F != null) Write(s, r, c, new Cell { Style = x.Style }); }
            s.LayoutDirty = true;
            AddUndo(() => { s.Merges.Remove(nm); s.Merges.AddRange(removed); s.LayoutDirty = true; }, () => { foreach (var m in removed) s.Merges.Remove(m); s.Merges.Add(nm); s.LayoutDirty = true; });
        });
    }

    public void Unmerge(Sheet s, int r1, int c1, int r2, int c2)
    {
        var removed = s.Merges.Where(m => m.Intersects(r1, c1, r2, c2)).ToList();
        if (removed.Count == 0) return;
        Do("Dividi celle", () =>
        {
            foreach (var m in removed) s.Merges.Remove(m);
            s.LayoutDirty = true;
            AddUndo(() => { s.Merges.AddRange(removed); s.LayoutDirty = true; }, () => { foreach (var m in removed) s.Merges.Remove(m); s.LayoutDirty = true; });
        });
    }

    public void RenameSheet(Sheet s, string newName)
    {
        newName = newName.Trim();
        if (newName.Length == 0 || newName.Length > 31 || newName.IndexOfAny([':', '\\', '/', '?', '*', '[', ']']) >= 0) throw new ArgumentException("Nome foglio non valido.");
        if (Sheets.Any(x => x != s && string.Equals(x.Name, newName, StringComparison.OrdinalIgnoreCase))) throw new ArgumentException("Esiste già un foglio con questo nome.");
        string old = s.Name;
        foreach (var sh in Sheets)
            foreach (var key in sh.Cells.Where(kv => kv.Value.F is { Opaque: false }).Select(kv => kv.Key).ToList())
            {
                var cell = sh.Cells[key];
                string nt = FormulaText.RenameSheet(cell.F!.Text, old, newName);
                if (nt != cell.F.Text) { cell.F = new FormulaInfo { Text = nt, CellAttrs = cell.F.CellAttrs }; sh.Cells[key] = cell; sh.MarkRowDirty(Addr.Row(key)); }
            }
        s.Name = newName;
        InvalidateGraph();
        FormulasChangedSinceOpen = true;
        Modified = true;
        Changed?.Invoke();
    }

    /// <summary>Applica una modifica di stile a un intervallo.</summary>
    public void FormatRange(Sheet s, int r1, int c1, int r2, int c2, Func<XfInfo, XfInfo> change)
    {
        Do("Formato", () =>
        {
            var memo = new Dictionary<uint, uint>();
            for (int r = r1; r <= r2; r++)
                for (int c = c1; c <= c2; c++)
                {
                    var cell = s.Get(r, c);
                    if (!memo.TryGetValue(cell.Style, out var ns)) memo[cell.Style] = ns = Styles.Derive(cell.Style, change);
                    if (ns != cell.Style) { cell.Style = ns; Write(s, r, c, cell); }
                }
        });
    }
}
