namespace PowerTables.Core;

/// <summary>Tabella (xl/tables/tableN.xml): area, nome e stile. Le tabelle già presenti nel file sono lette e conservate.</summary>
public sealed class TableDef
{
    public string Name = "";
    public int R1, C1, R2, C2;
    public string Style = "TableStyleMedium2";
    /// <summary>Id della relazione verso la parte tabella del foglio (null = non ancora scritta).</summary>
    internal string? PartRel;
    internal bool Dirty;      // area cambiata rispetto al file
    public bool NoHeader;     // letta da un file con headerRowCount=0
    public bool Pending => PartRel == null;

    public TableDef Clone() => (TableDef)MemberwiseClone();
}

public sealed partial class Workbook
{
    /// <summary>Stili offerti: (nome nel file, etichetta, colore intestazione, colore righe alternate), colori come la tavolozza Office.</summary>
    public static readonly (string Style, string Label, string Header, string Band)[] TableStyles =
    [
        ("TableStyleMedium2", "Blu", "FF4472C4", "FFD9E1F2"), ("TableStyleMedium3", "Arancione", "FFED7D31", "FFFCE4D6"), ("TableStyleMedium4", "Grigio", "FFA5A5A5", "FFEDEDED"),
        ("TableStyleMedium5", "Oro", "FFFFC000", "FFFFF2CC"), ("TableStyleMedium7", "Verde", "FF70AD47", "FFE2EFDA"),
    ];

    /// <summary>Tabella che contiene la cella, o null.</summary>
    public TableDef? TableAt(Sheet s, int r, int c) => s.Tables.FirstOrDefault(t => r >= t.R1 && r <= t.R2 && c >= t.C1 && c <= t.C2);

    /// <summary>Aspetto dato dallo stile della tabella (sotto formato di cella e formattazione condizionale).</summary>
    public CfLook? TableLookAt(Sheet s, int r, int c)
    {
        if (s.Tables.Count == 0 || TableAt(s, r, c) is not { } t) return null;
        var st = TableStyles.FirstOrDefault(x => x.Style == t.Style); if (st.Style == null) return null;   // stile non gestito: lo disegna solo Excel
        if (!t.NoHeader && r == t.R1) return new CfLook("FFFFFFFF", st.Header, true);
        return (r - t.R1) % 2 == (t.NoHeader ? 0 : 1) ? new CfLook(null, st.Band, false) : null;
    }

    /// <summary>Crea una tabella sull'area. Senza intestazioni inserisce una riga di intestazione sopra l'area (come Excel).
    /// Le intestazioni vuote, non testuali o duplicate diventano «Colonna N».</summary>
    public TableDef CreateTable(Sheet s, int r1, int c1, int r2, int c2, bool hasHeader, string? style = null, string? name = null)
    {
        style ??= TableStyles[0].Style;
        if (!TableStyles.Any(x => x.Style == style)) throw new ArgumentException("Stile di tabella non supportato.");
        if (r2 < r1 || c2 < c1) throw new ArgumentException("Area non valida.");
        if (hasHeader && r2 == r1) throw new ArgumentException("Una tabella con intestazioni richiede almeno una riga di dati.");
        if (s.Tables.Any(t => r1 <= t.R2 && r2 >= t.R1 && c1 <= t.C2 && c2 >= t.C1)) throw new ArgumentException("L'area si sovrappone a una tabella esistente.");
        if (s.Merges.Any(m => m.Intersects(r1, c1, r2, c2))) throw new ArgumentException("Le celle unite non sono ammesse in una tabella.");
        if (s.Filter is { } f && r1 <= f.R2 && r2 >= f.R1 && c1 <= f.C2 && c2 >= f.C1) throw new ArgumentException("L'area contiene un filtro automatico: rimuoverlo prima.");
        name ??= NextTableName();
        if (ValidateName(name) is { } err) throw new ArgumentException(err.Replace("Il nome", "Il nome della tabella"));
        if (Sheets.Any(x => x.Tables.Any(t => string.Equals(t.Name, name, StringComparison.OrdinalIgnoreCase))) || Names.ContainsKey(name)) throw new ArgumentException("Esiste già una tabella o un nome con questo nome.");
        var def = new TableDef { Name = name, R1 = r1, C1 = c1, R2 = r2, C2 = c2, Style = style };
        ChangeSections(s, "Tabella", () =>
        {
            if (!hasHeader)
            {
                InsertLines(s, true, r1, 1);
                def.R2 = r2 + 1;   // l'area sale di una riga: la nuova riga di intestazione è r1
            }
            var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int c = c1; c <= c2; c++)
            {
                var cell = s.Get(def.R1, c);
                string? h = cell.Kind == CellKind.Str && cell.F == null ? ((string?)cell.Obj)?.Trim() : null;
                if (string.IsNullOrEmpty(h) || !used.Add(h))
                {
                    int n = c - c1 + 1; string cand;
                    do cand = "Colonna" + n++; while (!used.Add(cand));
                    h = cand;
                    var nc = new Cell { Style = cell.Style }; nc.SetValue(h); Write(s, def.R1, c, nc);
                }
            }
            s.Tables.Add(def);
        });
        return def;
    }

    string NextTableName()
    {
        for (int i = 1; ; i++) { var n = "Tabella" + i; if (!Sheets.Any(x => x.Tables.Any(t => t.Name.Equals(n, StringComparison.OrdinalIgnoreCase))) && !Names.ContainsKey(n)) return n; }
    }

    /// <summary>Converte in intervallo: l'area resta con i dati, senza lo stile della tabella.</summary>
    public void RemoveTable(Sheet s, TableDef t)
    {
        if (!s.Tables.Contains(t)) return;
        ChangeSections(s, "Converti in intervallo", () => s.Tables.RemoveAll(x => x.Name == t.Name));
    }

    void AdjustTables(Sheet s, bool rows, int at, int count)
    {
        foreach (var t in s.Tables.ToList())
        {
            var sp = Formulas.FormulaText.MapSpan(rows ? t.R1 : t.C1, rows ? t.R2 : t.C2, at, count);
            if (sp == null) { s.Tables.Remove(t); continue; }
            if (rows) { if (t.R1 != sp.Value.Lo || t.R2 != sp.Value.Hi) { t.R1 = sp.Value.Lo; t.R2 = sp.Value.Hi; t.Dirty = true; } }
            else if (t.C1 != sp.Value.Lo || t.C2 != sp.Value.Hi) { t.C1 = sp.Value.Lo; t.C2 = sp.Value.Hi; t.Dirty = true; }
        }
    }
}
