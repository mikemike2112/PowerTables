using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using PowerTables.Core.Formulas;

namespace PowerTables.Core;

/// <summary>Grafo delle dipendenze e ricalcolo incrementale con rilevamento dei cicli.</summary>
public sealed partial class Workbook
{
    struct DepList
    {
        public long One; public List<long>? More; public int Count;
        public void Add(long id) { if (Count == 0) One = id; else (More ??= []).Add(id); Count++; }
        public void Remove(long id)
        {
            if (Count == 0) return;
            if (More != null && More.Remove(id)) { Count--; return; }
            if (One == id)
            {
                if (More is { Count: > 0 }) { One = More[^1]; More.RemoveAt(More.Count - 1); }
                Count--;
            }
        }
        public readonly IEnumerable<long> Items()
        {
            if (Count == 0) yield break;
            yield return One;
            if (More != null) foreach (var m in More) yield return m;
        }
    }

    readonly record struct RangeDep(int R1, int C1, int R2, int C2, long Formula);

    readonly Dictionary<long, DepList> _dependents = [];
    readonly Dictionary<int, List<RangeDep>> _rangeDeps = [];
    readonly HashSet<long> _volatile = [];
    bool _graphReady;

    /// <summary>Celle coinvolte in cicli nell'ultimo ricalcolo.</summary>
    public List<(Sheet Sheet, int Row, int Col)> Cycles { get; } = [];
    public int LastRecalcCount { get; private set; }

    static long Id(Sheet s, long key) => ((long)s.Index << 36) | key;
    Sheet SheetOf(long id) => Sheets[(int)(id >> 36)];
    static long KeyOf(long id) => id & ((1L << 36) - 1);

    public void InvalidateGraph()
    {
        _dependents.Clear(); _rangeDeps.Clear(); _volatile.Clear(); _graphReady = false;
    }

    static readonly HashSet<string> VolatileFns = ["NOW", "TODAY", "RAND", "RANDBETWEEN", "OFFSET", "INDIRECT"];

    /// <summary>Elenca i precedenti di una formula (fogli non trovati ignorati).</summary>
    void Collect(Node n, Sheet home, List<(Sheet S, int R1, int C1, int R2, int C2)> acc, ref bool vol, int depth = 0)
    {
        switch (n)
        {
            case RefNode r:
                {
                    var sh = r.R.Sheet == null ? home : FindSheet(r.R.Sheet);
                    if (sh == null) return;
                    int r1 = r.R.A.Row, c1 = r.R.A.Col, r2 = r.R.IsRange ? r.R.B.Row : r1, c2 = r.R.IsRange ? r.R.B.Col : c1;
                    if (r1 == 0) { r1 = 1; r2 = Addr.MaxRow; }
                    if (c1 == 0) { c1 = 1; c2 = Addr.MaxCol; }
                    acc.Add((sh, Math.Min(r1, r2), Math.Min(c1, c2), Math.Max(r1, r2), Math.Max(c1, c2)));
                    break;
                }
            case NameNode nm:
                if (depth < 8 && TryGetName(nm.Name, home, out var text))
                {
                    try { Collect(Parser.Parse(text), home, acc, ref vol, depth + 1); } catch (FormatException) { }
                }
                break;
            case UnaryNode u: Collect(u.X, home, acc, ref vol, depth); break;
            case PercentNode p: Collect(p.X, home, acc, ref vol, depth); break;
            case BinNode b: Collect(b.L, home, acc, ref vol, depth); Collect(b.R, home, acc, ref vol, depth); break;
            case CallNode c:
                if (VolatileFns.Contains(c.Name)) vol = true;
                foreach (var a in c.Args) Collect(a, home, acc, ref vol, depth);
                break;
            case ArrayNode ar: break;
        }
    }

    List<(Sheet S, int R1, int C1, int R2, int C2)>? Precedents(Sheet s, FormulaInfo f, out bool vol)
    {
        vol = false;
        try
        {
            var acc = new List<(Sheet, int, int, int, int)>();
            Collect(Parser.Parse(f.Text), s, acc, ref vol);
            return acc;
        }
        catch (FormatException) { return null; }
    }

    void Register(Sheet s, long key, FormulaInfo f)
    {
        var pre = Precedents(s, f, out bool vol);
        long id = Id(s, key);
        if (vol) _volatile.Add(id);
        if (pre == null) return;
        foreach (var (ps, r1, c1, r2, c2) in pre)
        {
            if (r1 == r2 && c1 == c2)
            {
                ref var dl = ref CollectionsMarshal.GetValueRefOrAddDefault(_dependents, Id(ps, Addr.Key(r1, c1)), out _);
                dl.Add(id);
            }
            else
            {
                if (!_rangeDeps.TryGetValue(ps.Index, out var l)) _rangeDeps[ps.Index] = l = [];
                l.Add(new RangeDep(r1, c1, r2, c2, id));
            }
        }
    }

    void Unregister(Sheet s, long key, FormulaInfo f)
    {
        var pre = Precedents(s, f, out _);
        long id = Id(s, key);
        _volatile.Remove(id);
        if (pre == null) return;
        foreach (var (ps, r1, c1, r2, c2) in pre)
        {
            if (r1 == r2 && c1 == c2)
            {
                var k = Id(ps, Addr.Key(r1, c1));
                if (_dependents.TryGetValue(k, out var dl)) { dl.Remove(id); if (dl.Count == 0) _dependents.Remove(k); else _dependents[k] = dl; }
            }
            else if (_rangeDeps.TryGetValue(ps.Index, out var l))
            {
                int ix = l.FindIndex(d => d.Formula == id && d.R1 == r1 && d.C1 == c1 && d.R2 == r2 && d.C2 == c2);
                if (ix >= 0) l.RemoveAt(ix);
            }
        }
    }

    void EnsureGraph()
    {
        if (_graphReady) return;
        _dependents.Clear(); _rangeDeps.Clear(); _volatile.Clear();
        for (int i = 0; i < Sheets.Count; i++) Sheets[i].Index = i;
        foreach (var s in Sheets)
            foreach (var kv in s.Cells)
                if (kv.Value.F is { Opaque: false } f) Register(s, kv.Key, f);
        _graphReady = true;
    }

    IEnumerable<long> DependentsOf(long id)
    {
        if (_dependents.TryGetValue(id, out var dl)) foreach (var x in dl.Items()) yield return x;
        int si = (int)(id >> 36);
        if (_rangeDeps.TryGetValue(si, out var l))
        {
            long key = KeyOf(id); int r = Addr.Row(key), c = Addr.Col(key);
            for (int i = 0; i < l.Count; i++)
            {
                var d = l[i];
                if (r >= d.R1 && r <= d.R2 && c >= d.C1 && c <= d.C2) yield return d.Formula;
            }
        }
    }

    /// <summary>Ricalcola tutte le formule (carica il grafo se serve).</summary>
    public void RecalcAll()
    {
        EnsureGraph();
        var seeds = new List<long>();
        foreach (var s in Sheets) foreach (var kv in s.Cells) if (kv.Value.F is { Opaque: false }) seeds.Add(Id(s, kv.Key));
        RecalcCore(seeds);
        _changed.Clear();
    }

    void Recalc()
    {
        _changed.RemoveWhere(x => x.Item1.Index < 0 || x.Item1.Index >= Sheets.Count || !ReferenceEquals(Sheets[x.Item1.Index], x.Item1));   // fogli tolti (Annulla di un nuovo foglio)
        if (_changed.Count == 0) { return; }
        if (_formulasChanged || _graphReady || true)
        {
            EnsureGraph();
            var seeds = new List<long>(_changed.Count);
            foreach (var (s, k) in _changed) seeds.Add(Id(s, k));
            _changed.Clear();
            RecalcCore(seeds);
        }
        _formulasChanged = false;
    }

    void RecalcCore(List<long> seeds)
    {
        Cycles.Clear();
        var dirty = new HashSet<long>();
        var adj = new Dictionary<long, List<long>>();
        var queue = new Queue<long>();
        void AddSeed(long id)
        {
            var s = SheetOf(id); long k = KeyOf(id);
            var cell = s.Get(Addr.Row(k), Addr.Col(k));
            if (cell.F != null && !cell.F.Opaque && dirty.Add(id)) queue.Enqueue(id);
            else if (cell.F == null || cell.F.Opaque) queue.Enqueue(id);
        }
        foreach (var id in seeds) AddSeed(id);
        foreach (var id in _volatile) if (dirty.Add(id)) queue.Enqueue(id);
        var visitedValue = new HashSet<long>();
        while (queue.Count > 0)
        {
            long p = queue.Dequeue();
            bool pIsFormula = dirty.Contains(p);
            if (!pIsFormula && !visitedValue.Add(p)) continue;
            foreach (var f in DependentsOf(p))
            {
                if (pIsFormula)
                {
                    if (!adj.TryGetValue(p, out var l)) adj[p] = l = [];
                    l.Add(f);
                }
                if (dirty.Add(f)) queue.Enqueue(f);
            }
        }
        var indeg = new Dictionary<long, int>(dirty.Count);
        foreach (var l in adj.Values) foreach (var f in l) indeg[f] = indeg.GetValueOrDefault(f) + 1;
        var ready = new Queue<long>(dirty.Where(d => !indeg.ContainsKey(d)));
        int done = 0;
        while (ready.Count > 0)
        {
            long id = ready.Dequeue();
            EvalCell(id); done++;
            if (adj.TryGetValue(id, out var l))
                foreach (var f in l)
                    if (--indeg[f] == 0) ready.Enqueue(f);
        }
        LastRecalcCount = dirty.Count;
        if (done < dirty.Count)
        {
            foreach (var id in dirty)
                if (indeg.TryGetValue(id, out var d) && d > 0)
                {
                    var s = SheetOf(id); long k = KeyOf(id);
                    Cycles.Add((s, Addr.Row(k), Addr.Col(k)));
                    SetComputed(s, k, 0.0);
                }
        }
    }

    void EvalCell(long id)
    {
        var s = SheetOf(id); long k = KeyOf(id);
        int r = Addr.Row(k), c = Addr.Col(k);
        var cell = s.Get(r, c);
        if (cell.F == null || cell.F.Opaque) return;
        object? v;
        try
        {
            v = Evaluator.EvalFormula(cell.F.Text, new EvalCtx(this, s, r, c));
        }
        catch (UnsupportedEx)
        {
            if (cell.Kind != CellKind.Empty) return; // mantiene il valore in cache del file
            v = CellError.Name;
        }
        SetComputed(s, k, v);
    }

    void SetComputed(Sheet s, long k, object? v)
    {
        ref var cell = ref CollectionsMarshal.GetValueRefOrNullRef(s.Cells, k);
        if (Unsafe.IsNullRef(ref cell)) return;
        var before = cell;
        cell.SetValue(v is double d && (double.IsNaN(d) || double.IsInfinity(d)) ? CellError.Num : v);
        if (!before.SameValue(cell)) s.MarkRowDirty(Addr.Row(k));
    }
}
