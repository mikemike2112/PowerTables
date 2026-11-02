using System.Runtime.InteropServices;
using System.Xml.Linq;

namespace PowerTables.Core;

public sealed class RowInfo
{
    public double? Height;          // punti; null = default
    public bool Hidden;
    /// <summary>Nascosta da un filtro automatico (non salvata come hidden manuale ambigua: salvata hidden="1").</summary>
    public bool Filtered;
    public bool CustomHeight;
    /// <summary>Altri attributi XML della riga (s, customFormat, ht gestito a parte, x14ac:dyDescent...) preservati.</summary>
    public List<(string Name, string Prefix, string Ns, string Value)> Extra = [];
    public uint Style;
    public bool CustomFormat;
}

public sealed class ColInfo
{
    public double Width = 9.140625;
    public bool Hidden, CustomWidth;
    public uint Style;
}

/// <summary>Inserimento (Count&gt;0) o eliminazione (Count&lt;0) di righe/colonne non ancora riportata su CF, convalida, nomi nel file.</summary>
public sealed record LineOp(string Sheet, bool Rows, int At, int Count);

public readonly record struct MergeRange(int R1, int C1, int R2, int C2)
{
    public bool Contains(int r, int c) => r >= R1 && r <= R2 && c >= C1 && c <= C2;
    public bool Intersects(int r1, int c1, int r2, int c2) => r1 <= R2 && r2 >= R1 && c1 <= C2 && c2 >= C1;
}

public sealed class AutoFilterState
{
    public int R1, C1, R2, C2;                          // area incluse intestazioni
    public Dictionary<int, HashSet<string>> Allowed = []; // colonna assoluta → valori ammessi (testo visualizzato)
    public bool Dirty;
}

/// <summary>Foglio: store sparso di celle + metadati di layout.</summary>
public sealed class Sheet
{
    public string Name;
    public Workbook Book { get; internal set; } = null!;
    public int Index;
    /// <summary>Posizione del foglio nel file all'ultimo caricamento/salvataggio (-1 = nuovo); serve a rimappare i nomi riservati al riordino.</summary>
    internal int FileIndex = -1;
    internal readonly Dictionary<long, Cell> Cells = new();
    public IReadOnlyDictionary<long, Cell> CellMap => Cells;
    public readonly Dictionary<int, RowInfo> Rows = [];
    public readonly Dictionary<int, ColInfo> Cols = [];
    public readonly List<MergeRange> Merges = [];
    /// <summary>Testo ricco originale (contenuto XML di si/is) per cella: riscritto al salvataggio finché il testo della cella non cambia.</summary>
    internal readonly Dictionary<long, (string Plain, string Xml)> Rich = [];
    internal readonly List<LineOp> LineOps = [];
    public int FrozenRows, FrozenCols;
    public double DefaultRowHeight = 15, DefaultColWidth = 9.140625;
    public AutoFilterState? Filter;
    public bool ShowGridLines = true;
    /// <summary>Area di stampa (riferimento in forma di file, es. 'Foglio1'!$A$1:$C$9) o null.</summary>
    public string? PrintArea;
    /// <summary>Sezioni del foglio modellate come XML (caricate dal file, riscritte al salvataggio quando il layout è sporco).</summary>
    public readonly List<XElement> CondFmt = [], DataVal = [], Links = [];
    public readonly List<TableDef> Tables = [];
    /// <summary>pageMargins / pageSetup / printOptions originali o modificati.</summary>
    public XElement? PageMargins, PageSetup, PrintOptions;
    /// <summary>Incrementato a ogni modifica di CondFmt (invalida la cache delle regole).</summary>
    public int CfVersion;

    public int MaxRow { get; internal set; }
    public int MaxCol { get; internal set; }

    // Persistenza (round-trip)
    internal HashSet<int> DirtyRows = [];
    public bool LayoutDirty;      // cols/merge/pane/filtro modificati
    internal string? PartId;
    internal string? OrigName;        // id relazione nel workbook
    internal bool Loaded = true;

    public Sheet(string name) { Name = name; }

    public int CellCount => Cells.Count;

    public Cell Get(int row, int col) => Cells.TryGetValue(Addr.Key(row, col), out var c) ? c : default;
    public object? GetValue(int row, int col) => Cells.TryGetValue(Addr.Key(row, col), out var c) ? c.Value : null;

    /// <summary>Scrittura grezza (nessun ricalcolo, nessun undo): usare Workbook.* per le modifiche utente.</summary>
    internal void Put(int row, int col, in Cell c)
    {
        long k = Addr.Key(row, col);
        if (c.IsEmpty) { Cells.Remove(k); return; }
        CollectionsMarshal.GetValueRefOrAddDefault(Cells, k, out _) = c;
        if (row > MaxRow) MaxRow = row;
        if (col > MaxCol) MaxCol = col;
    }

    public double RowHeight(int row)
    {
        if (Rows.TryGetValue(row, out var ri)) return ri.Hidden ? 0 : ri.Height ?? DefaultRowHeight;
        return DefaultRowHeight;
    }

    public double ColWidth(int col)
    {
        if (Cols.TryGetValue(col, out var ci)) return ci.Hidden ? 0 : ci.Width;
        return DefaultColWidth;
    }

    public RowInfo RowInfoFor(int row)
    {
        if (!Rows.TryGetValue(row, out var ri)) Rows[row] = ri = new RowInfo();
        return ri;
    }

    public ColInfo ColInfoFor(int col)
    {
        if (!Cols.TryGetValue(col, out var ci)) Cols[col] = ci = new ColInfo { Width = DefaultColWidth };
        return ci;
    }

    public MergeRange? MergeAt(int r, int c)
    {
        foreach (var m in Merges) if (m.Contains(r, c)) return m;
        return null;
    }

    /// <summary>Ricalcola gli estremi di MaxRow/MaxCol (dopo cancellazioni).</summary>
    public void RecomputeExtent()
    {
        int mr = 0, mc = 0;
        foreach (var k in Cells.Keys) { mr = Math.Max(mr, Addr.Row(k)); mc = Math.Max(mc, Addr.Col(k)); }
        MaxRow = mr; MaxCol = mc;
    }

    public void MarkRowDirty(int row) => DirtyRows.Add(row);
}
