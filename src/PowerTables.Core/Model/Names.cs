using System.Text.RegularExpressions;
using PowerTables.Core.Formulas;

namespace PowerTables.Core;

public sealed partial class Workbook
{
    static readonly Regex NameRx = new(@"^[\p{L}_\\][\p{L}\p{N}_.\\]*$", RegexOptions.Compiled);
    static readonly Regex R1C1Rx = new(@"^[Rr](\d*)([Cc](\d*))?$|^[Cc]\d*$", RegexOptions.Compiled);

    /// <summary>Verifica il nome come farebbe Excel; restituisce il messaggio d'errore o null se valido.</summary>
    public static string? ValidateName(string name)
    {
        if (name.Length == 0) return "Il nome non può essere vuoto.";
        if (name.Length > 255) return "Il nome è troppo lungo.";
        if (!NameRx.IsMatch(name)) return "Il nome non è valido: deve iniziare con una lettera o '_' e non può contenere spazi.";
        if (Addr.TryParse(name, out _, out _) || R1C1Rx.IsMatch(name)) return "Il nome non può coincidere con un riferimento di cella.";
        if (name.Equals("TRUE", StringComparison.OrdinalIgnoreCase) || name.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) return "Il nome non è valido.";
        return null;
    }

    /// <summary>Il nome esiste nell'ambito indicato (null = cartella di lavoro)?</summary>
    public bool HasName(string name, Sheet? scope) => scope == null ? Names.ContainsKey(name) : LocalNames.ContainsKey((scope.Index, name.ToUpperInvariant()));

    /// <summary>Tutti i nomi definiti: (nome, ambito o null, riferimento in forma utente senza '=').</summary>
    public List<(string Name, Sheet? Scope, string RefersTo)> ListNames()
    {
        var l = new List<(string, Sheet?, string)>();
        foreach (var kv in Names) l.Add((kv.Key, null, FormulaText.ToDisplay(kv.Value)));
        foreach (var kv in LocalNames)
            if (kv.Key.Item1 >= 0 && kv.Key.Item1 < Sheets.Count)
                l.Add((LocalNameCase.GetValueOrDefault(kv.Key, kv.Key.Item2), Sheets[kv.Key.Item1], FormulaText.ToDisplay(kv.Value)));
        return [.. l.OrderBy(x => x.Item1, StringComparer.OrdinalIgnoreCase)];
    }

    /// <summary>Definisce o ridefinisce un nome. <paramref name="refersTo"/> è in forma utente (senza '=').</summary>
    public void DefineName(string name, string refersTo, Sheet? scope = null)
    {
        name = name.Trim();
        if (ValidateName(name) is { } err) throw new ArgumentException(err);
        string text = FormulaText.ToFile(FormulaText.FromUser(refersTo.Trim().TrimStart('=')));
        if (text.Length == 0) throw new ArgumentException("Indicare a cosa si riferisce il nome.");
        try { Parser.Parse(text); } catch (FormatException ex) { throw new ArgumentException("Riferimento non valido: " + ex.Message); }
        ChangeNames("Definisci nome", () =>
        {
            if (scope == null) { foreach (var k in Names.Keys.Where(k => k.Equals(name, StringComparison.OrdinalIgnoreCase)).ToList()) Names.Remove(k); Names[name] = text; }
            else { var key = (scope.Index, name.ToUpperInvariant()); LocalNames[key] = text; LocalNameCase[key] = name; }
        });
    }

    public void DeleteName(string name, Sheet? scope = null)
    {
        if (!HasName(name, scope)) return;
        ChangeNames("Elimina nome", () =>
        {
            if (scope == null) Names.Remove(name);
            else { var key = (scope.Index, name.ToUpperInvariant()); LocalNames.Remove(key); LocalNameCase.Remove(key); }
        });
    }

    /// <summary>Applica una modifica ai nomi con istantanea per annulla/ripeti e ricalcolo di tutte le formule.</summary>
    void ChangeNames(string label, Action mutate)
    {
        var oldN = new Dictionary<string, string>(Names); var oldL = new Dictionary<(int, string), string>(LocalNames); var oldC = new Dictionary<(int, string), string>(LocalNameCase);
        Do(label, () =>
        {
            mutate();
            var newN = new Dictionary<string, string>(Names); var newL = new Dictionary<(int, string), string>(LocalNames); var newC = new Dictionary<(int, string), string>(LocalNameCase);
            void Set(Dictionary<string, string> n, Dictionary<(int, string), string> l, Dictionary<(int, string), string> c)
            {
                Replace(Names, n); Replace(LocalNames, l); Replace(LocalNameCase, c);
                NamesDirty = true; InvalidateGraph(); _formulasChanged = true; FormulasChangedSinceOpen = true;
                foreach (var sh in Sheets) foreach (var kv in sh.Cells) if (kv.Value.F != null) _changed.Add((sh, kv.Key));
            }
            Set(newN, newL, newC);
            AddUndo(() => Set(oldN, oldL, oldC), () => Set(newN, newL, newC));
        });
    }

    /// <summary>Imposta (o toglie, con null) l'area di stampa del foglio; <paramref name="area"/> è un intervallo A1 sul foglio stesso.</summary>
    public void SetPrintArea(Sheet s, (int R1, int C1, int R2, int C2)? area)
    {
        string? text = area is { } a ? QuoteSheet(s.Name) + "!" + Abs(a.R1, a.C1) + ":" + Abs(a.R2, a.C2) : null;
        string? old = s.PrintArea;
        if (old == text) return;
        Do("Area di stampa", () =>
        {
            s.PrintArea = text; NamesDirty = true;
            AddUndo(() => { s.PrintArea = old; NamesDirty = true; }, () => { s.PrintArea = text; NamesDirty = true; });
        });
    }

    public static string QuoteSheet(string n) => n.All(ch => char.IsLetterOrDigit(ch) || ch == '_') && !char.IsDigit(n[0]) ? n : "'" + n.Replace("'", "''") + "'";
    static string Abs(int r, int c) => "$" + Addr.ColName(c) + "$" + r;
}
