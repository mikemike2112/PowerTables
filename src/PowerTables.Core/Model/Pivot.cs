using System.Globalization;

namespace PowerTables.Core;

/// <summary>Tabella pivot minima creata da PowerTables: un campo riga, un campo valori con somma o conteggio. Excel la ricalcola all'apertura (refreshOnLoad).</summary>
public sealed class PivotDef
{
    public string Name = "";
    public string SrcSheet = "";
    public int SR1, SC1, SR2, SC2;          // origine con intestazioni
    public int RowCol, DataCol;             // colonne assolute nel foglio origine
    public bool Count;                      // false = somma
    public int R1, C1, R2, C2;              // posizione dell'output nel foglio di destinazione
    public List<object> RowItems = [];      // valori distinti (testo o numeri) in ordine
    public string[] FieldNames = [];
    public string DataCaption => (Count ? "Conteggio di " : "Somma di ") + FieldNames[DataCol - SC1];
}

public sealed partial class Workbook
{
    /// <summary>Crea la tabella pivot: scrive le celle dell'output (valori) nel foglio di destinazione e registra la definizione per il file. Un solo Annulla.</summary>
    public PivotDef CreatePivot(Sheet src, (int R1, int C1, int R2, int C2) area, int rowCol, int dataCol, bool count, Sheet dest, int destR, int destC, CultureInfo? ci = null)
    {
        ci ??= CultureInfo.CurrentCulture;
        if (area.R2 <= area.R1) throw new ArgumentException("L'origine deve avere un'intestazione e almeno una riga di dati.");
        if (rowCol < area.C1 || rowCol > area.C2 || dataCol < area.C1 || dataCol > area.C2) throw new ArgumentException("Le colonne scelte non sono nell'origine.");
        if (rowCol == dataCol) throw new ArgumentException("Il campo riga e il campo valori devono essere colonne diverse.");
        var names = new string[area.C2 - area.C1 + 1]; var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (int c = area.C1; c <= area.C2; c++)
        {
            if (src.Get(area.R1, c).Obj as string is not { Length: > 0 } h || src.Get(area.R1, c).Kind != CellKind.Str) throw new ArgumentException("La prima riga dell'origine deve contenere intestazioni di testo in tutte le colonne.");
            if (!seen.Add(h)) throw new ArgumentException($"Intestazione duplicata: «{h}».");
            names[c - area.C1] = h;
        }
        bool text = src.Get(area.R1 + 1, rowCol).Kind == CellKind.Str;
        var groups = new SortedDictionary<object, double>(Comparer<object>.Create((a, b) => a is double x && b is double y ? x.CompareTo(y) : string.Compare(a.ToString(), b.ToString(), ci, CompareOptions.IgnoreCase)));
        var cnt = new Dictionary<object, int>();
        double total = 0; int totalCnt = 0;
        for (int r = area.R1 + 1; r <= area.R2; r++)
        {
            var kc = src.Get(r, rowCol); var dc = src.Get(r, dataCol);
            if (text ? kc.Kind != CellKind.Str : kc.Kind != CellKind.Num) throw new ArgumentException("Il campo riga deve contenere solo testo o solo numeri, senza celle vuote.");
            if (!count && dc.Kind != CellKind.Num) throw new ArgumentException("Il campo valori della somma deve contenere solo numeri, senza celle vuote.");
            if (count && dc.Kind == CellKind.Empty) throw new ArgumentException("Il campo valori non deve avere celle vuote.");
            object key = text ? (string)kc.Obj! : kc.Num; double v = count ? 1 : dc.Num;
            if (text) { string ks = (string)key; var existing = groups.Keys.FirstOrDefault(k => string.Equals((string)k, ks, StringComparison.OrdinalIgnoreCase)); if (existing != null) key = existing; }
            groups[key] = groups.GetValueOrDefault(key) + v; total += v; totalCnt++;
        }
        if (groups.Count == 0 || groups.Count > 10000) throw new ArgumentException("Numero di gruppi non valido.");
        int rows = groups.Count + 2;                      // intestazione + gruppi + totale
        if (destR + rows - 1 > Addr.MaxRow || destC + 1 > Addr.MaxCol) throw new ArgumentException("L'output esce dal foglio.");
        if (ReferenceEquals(src, dest) && destR <= area.R2 && destR + rows - 1 >= area.R1 && destC <= area.C2 && destC + 1 >= area.C1) throw new ArgumentException("L'output si sovrappone ai dati di origine.");
        for (int r = destR; r < destR + rows; r++) for (int c = destC; c <= destC + 1; c++) if (dest.CellMap.TryGetValue(Addr.Key(r, c), out var ex) && ex.Kind != CellKind.Empty) throw new ArgumentException("La destinazione non è vuota.");
        string pname = "TabellaPivot1"; for (int i = 1; Sheets.Any(x => x.Pivots.Any(p => p.Name.Equals(pname, StringComparison.OrdinalIgnoreCase))); ) pname = "TabellaPivot" + ++i;
        var def = new PivotDef { Name = pname, SrcSheet = src.Name, SR1 = area.R1, SC1 = area.C1, SR2 = area.R2, SC2 = area.C2, RowCol = rowCol, DataCol = dataCol, Count = count, R1 = destR, C1 = destC, R2 = destR + rows - 1, C2 = destC + 1, RowItems = [.. groups.Keys], FieldNames = names };
        Do("Tabella pivot", () =>
        {
            SetValue(dest, destR, destC, "Etichette di riga"); SetValue(dest, destR, destC + 1, def.DataCaption);
            int r = destR + 1;
            foreach (var kv in groups) { SetValue(dest, r, destC, kv.Key); SetValue(dest, r, destC + 1, kv.Value); r++; }
            SetValue(dest, r, destC, "Totale complessivo"); SetValue(dest, r, destC + 1, total);
            dest.Pivots.Add(def); dest.PivotsDirty = true; dest.LayoutDirty = true;
            AddUndo(() => { dest.Pivots.Remove(def); dest.PivotsDirty = true; dest.LayoutDirty = true; }, () => { dest.Pivots.Add(def); dest.PivotsDirty = true; dest.LayoutDirty = true; });
        });
        return def;
    }
}
