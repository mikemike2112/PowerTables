using System.Globalization;
using System.Runtime.InteropServices;
using PowerTables.Core.Formulas;

namespace PowerTables.Core;

public readonly record struct CellChange(Sheet Sheet, long Key, Cell Old, Cell New);

/// <summary>Un gruppo di modifiche annullabile con un solo Ctrl+Z.</summary>
public sealed class Tx
{
    public string Name = "";
    public List<CellChange> Cells = [];
    public List<(Action Undo, Action Redo)> Other = [];
}

/// <summary>Cartella di lavoro: fogli, nomi, modifica con undo e ricalcolo incrementale.</summary>
public sealed partial class Workbook
{
    public readonly List<Sheet> Sheets = [];
    public readonly Dictionary<string, string> Names = new(StringComparer.OrdinalIgnoreCase);
    public readonly Dictionary<(int, string), string> LocalNames = [];
    /// <summary>Maiuscole/minuscole originali dei nomi locali (la chiave di LocalNames è in maiuscolo).</summary>
    public readonly Dictionary<(int, string), string> LocalNameCase = [];
    /// <summary>Nomi o aree di stampa modificati: al salvataggio definedNames viene riscritto.</summary>
    public bool NamesDirty;
    public bool Date1904;
    public string? FilePath;
    public StyleTable Styles = StyleTable.CreateDefault();
    public bool Modified { get; set; }
    internal IO.PackageState? Pkg;
    public bool FormulasChangedSinceOpen;
    public event Action? Changed;

    public Sheet? FindSheet(string name) => Sheets.FirstOrDefault(s => string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase));

    public Sheet AddSheet(string? name = null)
    {
        name ??= NextSheetName();
        var s = new Sheet(name) { Book = this, Index = Sheets.Count };
        Sheets.Add(s);
        InvalidateGraph();
        Modified = true;
        return s;
    }

    string NextSheetName()
    {
        for (int i = Sheets.Count + 1; ; i++) { var n = "Foglio" + i; if (FindSheet(n) == null) return n; }
    }

    public bool TryGetName(string name, Sheet ctx, out string text)
    {
        if (LocalNames.TryGetValue((ctx.Index, name.ToUpperInvariant()), out text!)) return true;
        return Names.TryGetValue(name, out text!);
    }

    // ---------------- transazioni e undo ----------------
    Tx? _tx;
    int _txDepth;
    readonly Stack<Tx> _undo = new(), _redo = new();
    readonly HashSet<(Sheet, long)> _changed = [];
    public bool CanUndo => _undo.Count > 0;
    public bool CanRedo => _redo.Count > 0;

    public void Begin(string name = "")
    {
        if (_txDepth++ == 0) _tx = new Tx { Name = name };
    }

    public void End()
    {
        if (--_txDepth > 0) return;
        var tx = _tx!; _tx = null;
        Recalc();
        if (tx.Cells.Count > 0 || tx.Other.Count > 0) { _undo.Push(tx); _redo.Clear(); Modified = true; }
        Changed?.Invoke();
    }

    public void Do(string name, Action a)
    {
        Begin(name);
        try { a(); } finally { End(); }
    }

    public void AddUndo(Action undo, Action redo)
    {
        if (_tx == null) throw new InvalidOperationException("fuori transazione");
        _tx.Other.Add((undo, redo));
    }

    public void Undo()
    {
        if (_undo.Count == 0) return;
        var tx = _undo.Pop();
        ApplyTx(tx, true);
        _redo.Push(tx);
    }

    public void Redo()
    {
        if (_redo.Count == 0) return;
        var tx = _redo.Pop();
        ApplyTx(tx, false);
        _undo.Push(tx);
    }

    void ApplyTx(Tx tx, bool undo)
    {
        _txDepth++; var saved = _tx; _tx = null;
        try
        {
            if (undo) { for (int i = tx.Cells.Count - 1; i >= 0; i--) RawWrite(tx.Cells[i].Sheet, tx.Cells[i].Key, tx.Cells[i].Old); for (int i = tx.Other.Count - 1; i >= 0; i--) tx.Other[i].Undo(); }
            else { foreach (var c in tx.Cells) RawWrite(c.Sheet, c.Key, c.New); foreach (var o in tx.Other) o.Redo(); }
        }
        finally { _txDepth--; _tx = saved; }
        Recalc();
        Modified = true;
        Changed?.Invoke();
    }

    // ---------------- scrittura celle ----------------
    /// <summary>Scrive una cella registrando undo e dipendenze.</summary>
    public void Write(Sheet s, int row, int col, Cell nc)
    {
        long k = Addr.Key(row, col);
        var old = s.Get(row, col);
        if (old.Kind == nc.Kind && old.Num.Equals(nc.Num) && Equals(old.Obj, nc.Obj) && old.Style == nc.Style && ReferenceEquals(old.F, nc.F)) return;
        _tx?.Cells.Add(new CellChange(s, k, old, nc));
        RawWrite(s, k, nc);
    }

    void RawWrite(Sheet s, long k, Cell nc)
    {
        int row = Addr.Row(k), col = Addr.Col(k);
        var old = s.Get(row, col);
        if (_graphReady)
        {
            if (old.F != null && !old.F.Opaque && !ReferenceEquals(old.F, nc.F)) Unregister(s, k, old.F);
            if (nc.F != null && !nc.F.Opaque && !ReferenceEquals(old.F, nc.F)) Register(s, k, nc.F);
        }
        s.Put(row, col, nc);
        s.MarkRowDirty(row);
        if (old.F != null || nc.F != null) { _formulasChanged = true; FormulasChangedSinceOpen = true; }
        _changed.Add((s, k));
    }

    bool _formulasChanged;

    public void SetValue(Sheet s, int r, int c, object? v)
    {
        var old = s.Get(r, c);
        var nc = new Cell { Style = old.Style }; nc.SetValue(v);
        Write(s, r, c, nc);
    }

    /// <summary>Formula in forma utente (senza '='); viene convertita nella forma di file (_xlfn.).</summary>
    public void SetFormula(Sheet s, int r, int c, string userFormula)
    {
        var text = FormulaText.ToFile(FormulaText.FromUser(userFormula.Trim()));
        try { Parser.Parse(text); } catch (FormatException ex) { throw new FormatException("Formula non valida: " + ex.Message); }
        var old = s.Get(r, c);
        Write(s, r, c, new Cell { Style = old.Style, F = new FormulaInfo { Text = text } });
    }

    public void SetFormulaRaw(Sheet s, int r, int c, string fileFormula, double? cached = null)
    {
        var old = s.Get(r, c);
        Write(s, r, c, new Cell { Style = old.Style, F = new FormulaInfo { Text = fileFormula } });
    }

    public void SetStyle(Sheet s, int r, int c, uint style)
    {
        var old = s.Get(r, c);
        old.Style = style;
        Write(s, r, c, old);
    }

    public void ClearCell(Sheet s, int r, int c, bool keepStyle = true)
    {
        var old = s.Get(r, c);
        Write(s, r, c, new Cell { Style = keepStyle ? old.Style : 0 });
    }

    /// <summary>Interpreta il testo digitato come farebbe Excel (numero, data, percentuale, booleano, formula, testo).</summary>
    public void SetInput(Sheet s, int r, int c, string text, CultureInfo? ci = null)
    {
        ci ??= CultureInfo.CurrentCulture;
        if (text.Length > 1 && text[0] == '=') { SetFormula(s, r, c, text[1..]); return; }
        var (v, fmtId) = ParseInput(text, ci);
        var old = s.Get(r, c);
        uint st = old.Style;
        if (fmtId != 0 && Styles.Get(st).NumFmtId == 0) st = Styles.Derive(st, x => x with { NumFmtId = fmtId, NumFmtCode = null });
        var nc = new Cell { Style = st }; nc.SetValue(v);
        Write(s, r, c, nc);
    }

    public static (object? Value, int FmtId) ParseInput(string text, CultureInfo ci)
    {
        if (text.Length == 0) return (null, 0);
        if (text[0] == '\'') return (text[1..], 0);
        string t = text.Trim();
        if (t.Equals("TRUE", StringComparison.OrdinalIgnoreCase)) return (true, 0);
        if (t.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) return (false, 0);
        if (t.StartsWith('#') && Errors.TryParse(t, out var e)) return (e, 0);
        const NumberStyles ns = NumberStyles.Float | NumberStyles.AllowThousands;
        bool pct = t.EndsWith('%');
        string n = pct ? t[..^1].TrimEnd() : t;
        if (n.Length > 0 && (char.IsDigit(n[0]) || n[0] is '-' or '+' or '.' or ',' ) && (n.Any(char.IsDigit)))
        {
            if (double.TryParse(n, ns, ci, out var d) && !double.IsInfinity(d)) return pct ? (d / 100, n.Contains(ci.NumberFormat.NumberDecimalSeparator) ? 10 : 9) : (d, 0);
            if (double.TryParse(n, ns, CultureInfo.InvariantCulture, out d) && !double.IsInfinity(d)) return pct ? (d / 100, 9) : (d, 0);
        }
        if (t.Length >= 6 && t.Any(char.IsDigit) && !pct)
        {
            if (DateTime.TryParse(t, ci, DateTimeStyles.NoCurrentDateDefault | DateTimeStyles.AllowWhiteSpaces, out var dt) && dt.Year >= 1900)
            {
                double serial = XlDate.ToSerial(dt.Year, dt.Month, dt.Day) + dt.TimeOfDay.TotalDays;
                bool hasDate = dt.Year > 1 && (t.Contains('/') || t.Contains('-') || t.Contains('.') || char.IsLetter(t[0]) || t.Count(char.IsLetter) > 0);
                if (!double.IsNaN(serial)) return (serial, dt.TimeOfDay.Ticks != 0 ? 22 : 14);
            }
        }
        if (t.Contains(':') && t.Length <= 12 && TimeSpan.TryParse(t, ci, out var ts) && ts.TotalHours >= 0 && ts.TotalHours < 24 * 100)
            return (ts.TotalDays, t.Count(ch => ch == ':') >= 2 ? 21 : 20);
        return (text, 0);
    }
}
