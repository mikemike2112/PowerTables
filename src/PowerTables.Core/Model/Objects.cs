using System.Text.Json;

namespace PowerTables.Core;

public enum ObjKind { Picture, Shape, TextBox, Chart }

/// <summary>Serie di un grafico: nome, intervallo dei valori e (facoltativo) delle categorie, in forma di file (es. Foglio1!$B$2:$B$6).</summary>
public sealed record ChartSeries(string Name, string Values, string? Categories);

/// <summary>Grafico creato da PowerTables: tipo "col" (istogramma) o "line".</summary>
public sealed record ChartSpec(string Type, string? Title, List<ChartSeries> Series);

/// <summary>Oggetto di disegno creato da PowerTables (immagine, forma, casella di testo, grafico). Gli oggetti già nel file restano intatti e non sono elencati qui.</summary>
public sealed class SheetObject
{
    public string Name = "";                      // "PT-xxxxxxxx": identifica l'oggetto nel file
    public ObjKind Kind;
    public int R1, C1, R2, C2;                    // celle di ancoraggio (R2/C2 incluse)
    public byte[]? Image; public string? ImageExt;  // png, jpeg, gif
    public string? Text; public string Geom = "rect";
    public ChartSpec? Chart;
    internal bool Written;                        // già nel file
    public const string Prefix = "PT-";
    public string Descr => Kind == ObjKind.Chart ? "PTCHART:" + JsonSerializer.Serialize(Chart) : Kind is ObjKind.Shape or ObjKind.TextBox ? "PTTEXT:" + JsonSerializer.Serialize(new TextInfo(Text, Geom)) : "Immagine";
    public sealed record TextInfo(string? Text, string Geom);
}

public sealed class CellComment { public string Author = "PowerTables"; public string Text = ""; internal string? Orig; }

public sealed partial class Workbook
{
    const long MaxImageBytes = 10_000_000;

    static string NewObjName() => SheetObject.Prefix + Guid.NewGuid().ToString("N")[..8];

    /// <summary>Aggiunge un oggetto con un solo passo di annullamento.</summary>
    public SheetObject AddObject(Sheet s, SheetObject o)
    {
        if (o.R2 < o.R1 || o.C2 < o.C1 || o.R1 < 1 || o.C1 < 1 || o.R2 > Addr.MaxRow || o.C2 > Addr.MaxCol) throw new ArgumentException("Posizione non valida.");
        if (o.Name.Length == 0) o.Name = NewObjName();
        Do("Inserisci " + o.Kind switch { ObjKind.Picture => "immagine", ObjKind.Chart => "grafico", ObjKind.TextBox => "casella di testo", _ => "forma" }, () =>
        {
            s.Objects.Add(o); s.LayoutDirty = true;
            AddUndo(() => { s.Objects.Remove(o); s.LayoutDirty = true; }, () => { s.Objects.Add(o); s.LayoutDirty = true; });
        });
        return o;
    }

    public void RemoveObject(Sheet s, SheetObject o)
    {
        if (!s.Objects.Contains(o)) return;
        Do("Elimina oggetto", () =>
        {
            s.Objects.Remove(o); s.LayoutDirty = true;
            AddUndo(() => { s.Objects.Add(o); s.LayoutDirty = true; }, () => { s.Objects.Remove(o); s.LayoutDirty = true; });
        });
    }

    /// <summary>Immagine dal file: solo PNG, JPEG, GIF riconosciuti dall'intestazione (non dall'estensione), fino a 10 MB.</summary>
    public SheetObject AddPicture(Sheet s, int r, int c, byte[] data, int pxW, int pxH)
    {
        if (data.Length == 0 || data.Length > MaxImageBytes) throw new ArgumentException("Immagine vuota o più grande di 10 MB.");
        string? ext = data.Length > 8 && data[0] == 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47 ? "png"
            : data.Length > 3 && data[0] == 0xFF && data[1] == 0xD8 && data[2] == 0xFF ? "jpeg"
            : data.Length > 6 && data[0] == 'G' && data[1] == 'I' && data[2] == 'F' && data[3] == '8' ? "gif" : null;
        if (ext == null) throw new ArgumentException("Formato immagine non supportato (PNG, JPEG, GIF).");
        int cols = Math.Clamp((int)Math.Ceiling(Math.Max(1, pxW) / 64.0), 1, 20), rows = Math.Clamp((int)Math.Ceiling(Math.Max(1, pxH) / 20.0), 1, 30);
        return AddObject(s, new SheetObject { Kind = ObjKind.Picture, R1 = r, C1 = c, R2 = Math.Min(Addr.MaxRow, r + rows - 1), C2 = Math.Min(Addr.MaxCol, c + cols - 1), Image = data, ImageExt = ext });
    }

    /// <summary>Forma (rect, ellipse, rightArrow, triangle, roundRect) o casella di testo.</summary>
    public SheetObject AddShape(Sheet s, int r, int c, bool textBox, string geom, string text)
    {
        if (!(geom is "rect" or "ellipse" or "rightArrow" or "triangle" or "roundRect")) throw new ArgumentException("Forma non supportata.");
        if (text.Length > 32000) throw new ArgumentException("Testo troppo lungo.");
        return AddObject(s, new SheetObject { Kind = textBox ? ObjKind.TextBox : ObjKind.Shape, Geom = geom, Text = text, R1 = r, C1 = c, R2 = Math.Min(Addr.MaxRow, r + (textBox ? 3 : 5)), C2 = Math.Min(Addr.MaxCol, c + (textBox ? 3 : 2)) });
    }

    /// <summary>Grafico "col" o "line" sull'area: se una cella della prima riga è testo, quella riga è l'intestazione delle serie; se la prima colonna è testo sono le categorie. Ogni colonna numerica è una serie.</summary>
    public SheetObject AddChart(Sheet s, int r1, int c1, int r2, int c2, string type)
    {
        if (type is not ("col" or "line")) throw new ArgumentException("Tipo di grafico non supportato.");
        var sp = BuildChartSpec(s, r1, c1, r2, c2, type);
        int top = Math.Min(Math.Min(r2 + 2, Addr.MaxRow), Addr.MaxRow - 14), left = Math.Max(1, Math.Min(c2 + 2, Addr.MaxCol - 7));
        return AddObject(s, new SheetObject { Kind = ObjKind.Chart, Chart = sp, R1 = top, C1 = left, R2 = top + 14, C2 = left + 7 });
    }

    static bool IsNum(Sheet s, int r, int c) => s.Get(r, c).Kind == CellKind.Num;
    static bool IsText(Sheet s, int r, int c) => s.Get(r, c).Kind == CellKind.Str;

    ChartSpec BuildChartSpec(Sheet s, int r1, int c1, int r2, int c2, string type)
    {
        if (r1 == r2 && c1 == c2) (r1, c1, r2, c2) = CurrentRegionOf(s, r1, c1);
        bool hdr = r2 > r1 && Enumerable.Range(c1, c2 - c1 + 1).Any(c => IsText(s, r1, c));
        int d1 = hdr ? r1 + 1 : r1;
        bool cats = IsText(s, d1, c1);
        int v1 = cats ? c1 + 1 : c1;
        string q = Formulas.Lexer.QuoteSheet(s.Name);
        var ser = new List<ChartSeries>();
        for (int c = v1; c <= c2; c++)
        {
            if (!Enumerable.Range(d1, r2 - d1 + 1).Any(r => IsNum(s, r, c))) continue;
            string name = hdr && IsText(s, r1, c) ? s.Get(r1, c).Obj as string ?? "" : "Serie" + (ser.Count + 1);
            string cn = Addr.ColName(c), cc = Addr.ColName(c1);
            ser.Add(new ChartSeries(name, $"{q}!${cn}${d1}:${cn}${r2}", cats ? $"{q}!${cc}${d1}:${cc}${r2}" : null));
        }
        if (ser.Count == 0) throw new InvalidOperationException("Nessun dato numerico nella selezione.");
        return new ChartSpec(type, null, ser);
    }

    /// <summary>Area contigua non vuota che contiene la cella (come Ctrl+*).</summary>
    public static (int R1, int C1, int R2, int C2) CurrentRegionOf(Sheet s, int r, int c)
    {
        int r1 = r, r2 = r, c1 = c, c2 = c; bool grew = true;
        bool Any(int ra, int ca, int rb, int cb)
        {
            for (int i = Math.Max(1, ra); i <= rb; i++) for (int j = Math.Max(1, ca); j <= cb; j++) if (s.CellMap.TryGetValue(Addr.Key(i, j), out var x) && x.Kind != CellKind.Empty) return true;
            return false;
        }
        while (grew)
        {
            grew = false;
            if (r1 > 1 && Any(r1 - 1, c1 - 1, r1 - 1, c2 + 1)) { r1--; grew = true; }
            if (r2 < s.MaxRow && Any(r2 + 1, c1 - 1, r2 + 1, c2 + 1)) { r2++; grew = true; }
            if (c1 > 1 && Any(r1 - 1, c1 - 1, r2 + 1, c1 - 1)) { c1--; grew = true; }
            if (c2 < s.MaxCol && Any(r1 - 1, c2 + 1, r2 + 1, c2 + 1)) { c2++; grew = true; }
        }
        return (r1, c1, r2, c2);
    }

    // ---------------- commenti ----------------
    /// <summary>Imposta (o con testo vuoto elimina) il commento della cella; un solo passo di annullamento.</summary>
    public void SetComment(Sheet s, int r, int c, string text, string author = "PowerTables")
    {
        text = text.Trim();
        if (text.Length > 32000) throw new ArgumentException("Commento troppo lungo.");
        long k = Addr.Key(r, c); s.Comments.TryGetValue(k, out var old);
        if (text.Length == 0 && old == null) return;
        var nw = text.Length == 0 ? null : new CellComment { Author = old?.Author ?? author, Text = text };
        void Put(CellComment? v) { if (v == null) s.Comments.Remove(k); else s.Comments[k] = v; s.LayoutDirty = true; s.CommentsDirty = true; }
        Do(text.Length == 0 ? "Elimina commento" : "Commento", () => { Put(nw); AddUndo(() => Put(old), () => Put(nw)); });
    }
}
