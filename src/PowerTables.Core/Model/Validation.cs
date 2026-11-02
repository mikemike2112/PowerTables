using System.Globalization;
using System.Xml.Linq;
using PowerTables.Core.Formulas;

namespace PowerTables.Core;

/// <summary>Regola di convalida dati (tipi: list, whole, decimal, textLength, date). Le formule sono in forma di file (virgole, nomi inglesi).</summary>
public sealed record DvSpec(string Type, string Op, string F1, string? F2 = null, bool AllowBlank = true, bool ShowError = true, string? ErrTitle = null, string? Err = null, bool Dropdown = true);

public sealed partial class Workbook
{
    // ---------------- utilità per le sezioni ----------------
    public static List<(int R1, int C1, int R2, int C2)> ParseSqref(string sqref)
    {
        var l = new List<(int, int, int, int)>();
        foreach (var t in sqref.Split(' ', StringSplitOptions.RemoveEmptyEntries))
        {
            var x = t.Replace("$", "");
            if (Addr.TryParseRange(x, out var a, out var b, out var c, out var d)) l.Add((a, b, c, d));
            else if (Addr.TryParse(x, out var r, out var cc)) l.Add((r, cc, r, cc));
        }
        return l;
    }

    public static string FormatSqref(IEnumerable<(int R1, int C1, int R2, int C2)> rects) =>
        string.Join(' ', rects.Select(r => r.R1 == r.R2 && r.C1 == r.C2 ? Addr.A1(r.R1, r.C1) : Addr.A1(r.R1, r.C1) + ":" + Addr.A1(r.R2, r.C2)));

    /// <summary>Toglie il rettangolo <paramref name="b"/> dal rettangolo <paramref name="a"/>: restano fino a 4 rettangoli.</summary>
    static IEnumerable<(int R1, int C1, int R2, int C2)> Subtract((int R1, int C1, int R2, int C2) a, (int R1, int C1, int R2, int C2) b)
    {
        if (b.R1 > a.R2 || b.R2 < a.R1 || b.C1 > a.C2 || b.C2 < a.C1) { yield return a; yield break; }
        if (b.R1 > a.R1) yield return (a.R1, a.C1, b.R1 - 1, a.C2);
        if (b.R2 < a.R2) yield return (b.R2 + 1, a.C1, a.R2, a.C2);
        int top = Math.Max(a.R1, b.R1), bot = Math.Min(a.R2, b.R2);
        if (b.C1 > a.C1) yield return (top, a.C1, bot, b.C1 - 1);
        if (b.C2 < a.C2) yield return (top, b.C2 + 1, bot, a.C2);
    }

    /// <summary>Toglie l'area dallo sqref degli elementi della lista; gli elementi che restano senza celle vengono eliminati.</summary>
    static void CutSqref(List<XElement> list, string attr, (int, int, int, int) area)
    {
        foreach (var e in list.ToList())
        {
            var rects = ParseSqref((string?)e.Attribute(attr) ?? "");
            var left = rects.SelectMany(r => Subtract(r, area)).ToList();
            if (left.Count == 0) list.Remove(e); else e.SetAttributeValue(attr, FormatSqref(left));
        }
    }

    /// <summary>Esegue una modifica alle sezioni XML del foglio con annulla/ripeti (istantanea di tutte le sezioni).</summary>
    void ChangeSections(Sheet s, string label, Action mutate)
    {
        (List<XElement> Cf, List<XElement> Dv, List<XElement> Lk, XElement? Pm, XElement? Ps, XElement? Po) Snap() =>
            (CloneList(s.CondFmt), CloneList(s.DataVal), CloneList(s.Links), s.PageMargins == null ? null : new XElement(s.PageMargins), s.PageSetup == null ? null : new XElement(s.PageSetup), s.PrintOptions == null ? null : new XElement(s.PrintOptions));
        void Put((List<XElement> Cf, List<XElement> Dv, List<XElement> Lk, XElement? Pm, XElement? Ps, XElement? Po) v)
        {
            SetList(s.CondFmt, v.Cf); SetList(s.DataVal, v.Dv); SetList(s.Links, v.Lk);
            s.PageMargins = v.Pm == null ? null : new XElement(v.Pm); s.PageSetup = v.Ps == null ? null : new XElement(v.Ps); s.PrintOptions = v.Po == null ? null : new XElement(v.Po);
            s.LayoutDirty = true; s.CfVersion++;
        }
        var before = Snap();
        Do(label, () =>
        {
            mutate();
            s.LayoutDirty = true; s.CfVersion++;
            var after = Snap();
            AddUndo(() => Put(before), () => Put(after));
        });
    }

    // ---------------- convalida dati ----------------
    static DvSpec? ToSpec(XElement dv) => new(
        (string?)dv.Attribute("type") ?? "none", (string?)dv.Attribute("operator") ?? "between",
        dv.Element(MainX + "formula1")?.Value ?? "", dv.Element(MainX + "formula2")?.Value,
        (bool?)dv.Attribute("allowBlank") ?? false, (bool?)dv.Attribute("showErrorMessage") ?? false,
        (string?)dv.Attribute("errorTitle"), (string?)dv.Attribute("error"), !((bool?)dv.Attribute("showDropDown") ?? false));

    XElement? FindValidationEl(Sheet s, int r, int c) =>
        s.DataVal.FirstOrDefault(dv => ParseSqref((string?)dv.Attribute("sqref") ?? "").Any(x => r >= x.R1 && r <= x.R2 && c >= x.C1 && c <= x.C2));

    public DvSpec? GetValidation(Sheet s, int r, int c) => FindValidationEl(s, r, c) is { } e ? ToSpec(e) : null;

    /// <summary>Imposta (o con null toglie) la convalida sull'area; le regole preesistenti sovrapposte vengono ristrette.</summary>
    public void SetValidation(Sheet s, int r1, int c1, int r2, int c2, DvSpec? spec)
    {
        if (spec != null)
        {
            if (spec.Type is not ("list" or "whole" or "decimal" or "textLength" or "date")) throw new ArgumentException("Tipo di convalida non supportato.");
            if (spec.F1.Trim().Length == 0) throw new ArgumentException("Indicare il valore o l'origine della convalida.");
            bool two = spec.Op is "between" or "notBetween" && spec.Type != "list";
            if (two && string.IsNullOrWhiteSpace(spec.F2)) throw new ArgumentException("Indicare il valore massimo.");
            foreach (var f in new[] { spec.F1, two ? spec.F2 : null })
                if (f != null) try { Parser.Parse(f); } catch (FormatException ex) { throw new ArgumentException("Formula non valida: " + ex.Message); }
        }
        ChangeSections(s, "Convalida dati", () =>
        {
            CutSqref(s.DataVal, "sqref", (r1, c1, r2, c2));
            if (spec == null) return;
            var e = new XElement(MainX + "dataValidation", new XAttribute("type", spec.Type));
            if (spec.Type != "list" && spec.Op != "between") e.SetAttributeValue("operator", spec.Op);
            if (spec.AllowBlank) e.SetAttributeValue("allowBlank", 1);
            if (spec.Type == "list" && !spec.Dropdown) e.SetAttributeValue("showDropDown", 1);   // sì: in OOXML "showDropDown=1" NASCONDE la freccia
            if (spec.ShowError) e.SetAttributeValue("showErrorMessage", 1);
            if (spec.ErrTitle is { Length: > 0 } et) e.SetAttributeValue("errorTitle", et);
            if (spec.Err is { Length: > 0 } er) e.SetAttributeValue("error", er);
            e.SetAttributeValue("sqref", FormatSqref([(r1, c1, r2, c2)]));
            e.Add(new XElement(MainX + "formula1", spec.F1));
            if (spec.Type != "list" && spec.Op is "between" or "notBetween") e.Add(new XElement(MainX + "formula2", spec.F2));
            s.DataVal.Add(e);
        });
    }

    /// <summary>Voci dell'elenco a discesa della cella (convalida di tipo elenco) o null.</summary>
    public List<string>? ValidationList(Sheet s, int r, int c, CultureInfo? ci = null)
    {
        var sp = GetValidation(s, r, c);
        if (sp is not { Type: "list" }) return null;
        string f = sp.F1.Trim();
        if (f.Length >= 2 && f[0] == '"' && f[^1] == '"') return [.. f[1..^1].Replace("\"\"", "\"").Split(',').Select(x => x.Trim())];
        try
        {
            var cx = new EvalCtx(this, s, r, c);
            var v = Evaluator.Eval(Parser.Parse(f), cx);
            if (v is RangeRef rr) { var l = new List<string>(); for (int i = 0; i < rr.Rows; i++) for (int j = 0; j < rr.Cols; j++) { var t = DisplayText(rr.Sheet, rr.R1 + i, rr.C1 + j, ci); if (t.Length > 0) l.Add(t); } return l; }
            if (v is Arr a) return [.. a.V.Cast<object?>().Select(x => V.Str(x))];
        }
        catch (Exception ex) when (ex is ExcelEx or FormatException or UnsupportedEx) { }
        return [];
    }

    /// <summary>Controlla un valore digitato: restituisce il messaggio di errore se la convalida lo rifiuta, altrimenti null.</summary>
    public string? CheckInput(Sheet s, int r, int c, string text, CultureInfo? ci = null)
    {
        var sp = GetValidation(s, r, c);
        if (sp == null || !sp.ShowError) return null;
        string msg = sp.Err is { Length: > 0 } m ? m : "Il valore immesso non è valido. Un utente ha limitato i valori che è possibile immettere in questa cella.";
        if (text.Length == 0) return sp.AllowBlank ? null : msg;
        ci ??= CultureInfo.CurrentCulture;
        if (text[0] == '=') return null;   // ponytail: le formule digitate non vengono controllate → valutare il risultato dopo il calcolo
        if (sp.Type == "list")
        {
            var l = ValidationList(s, r, c, ci);
            return l == null || l.Any(x => string.Equals(x, text.Trim(), StringComparison.OrdinalIgnoreCase)) ? null : msg;
        }
        var cx = new EvalCtx(this, s, r, c);
        double? Ev(string? f)
        {
            if (f == null) return null;
            try { var v = V.Scalar(Evaluator.EvalFormula(f, cx), cx); return V.TryNum(v, out var d) ? d : null; }
            catch (Exception ex) when (ex is ExcelEx or FormatException or UnsupportedEx) { return null; }
        }
        double x;
        if (sp.Type == "textLength") x = text.Length;
        else
        {
            var (v, _) = ParseInput(text, ci);
            if (v is not double d) return msg;
            if (sp.Type == "whole" && d != Math.Floor(d)) return msg;
            x = d;
        }
        double? a = Ev(sp.F1), b = Ev(sp.F2);
        if (a == null) return null;   // origine non calcolabile: non blocca
        bool ok = sp.Op switch
        {
            "between" => b != null && x >= a && x <= b,
            "notBetween" => b != null && (x < a || x > b),
            "equal" => x == a,
            "notEqual" => x != a,
            "greaterThan" => x > a,
            "lessThan" => x < a,
            "greaterThanOrEqual" => x >= a,
            "lessThanOrEqual" => x <= a,
            _ => true,
        };
        return ok ? null : msg;
    }
}
