using System.Globalization;
using System.Xml.Linq;
using PowerTables.Core.Formulas;

namespace PowerTables.Core;

/// <summary>Regola base di formattazione condizionale. Kind: cellIs (Op + F1[/F2]) o containsText (Text).</summary>
public sealed record CfSpec(string Kind, string Op, string F1, string? F2, string? Text, string? FontColor, string? FillColor, bool Bold);

/// <summary>Risultato della valutazione per una cella: colori da sovrapporre allo stile (AARRGGBB) e grassetto.</summary>
public readonly record struct CfLook(string? Font, string? Fill, bool Bold);

public sealed partial class Workbook
{
    sealed record CfRuleC(List<(int R1, int C1, int R2, int C2)> Area, string Type, string Op, string F1, string? F2, string? Text, int Priority, int DxfId, bool Stop);
    readonly Dictionary<Sheet, (int Ver, List<CfRuleC> Rules)> _cfCache = [];

    List<CfRuleC> CfRules(Sheet s)
    {
        if (_cfCache.TryGetValue(s, out var c) && c.Ver == s.CfVersion) return c.Rules;
        var l = new List<CfRuleC>();
        foreach (var cf in s.CondFmt)
        {
            var area = ParseSqref((string?)cf.Attribute("sqref") ?? "");
            foreach (var r in cf.Elements(MainX + "cfRule"))
            {
                string type = (string?)r.Attribute("type") ?? "";
                if (type is not ("cellIs" or "containsText")) continue;
                var fs = r.Elements(MainX + "formula").Select(f => f.Value).ToList();
                if (type == "cellIs" && fs.Count == 0) continue;
                l.Add(new CfRuleC(area, type, (string?)r.Attribute("operator") ?? "equal", fs.ElementAtOrDefault(0) ?? "", fs.ElementAtOrDefault(1), (string?)r.Attribute("text"),
                    (int?)r.Attribute("priority") ?? 1, (int?)r.Attribute("dxfId") ?? -1, (bool?)r.Attribute("stopIfTrue") ?? false));
            }
        }
        l.Sort((a, b) => a.Priority.CompareTo(b.Priority));
        _cfCache[s] = (s.CfVersion, l);
        return l;
    }

    /// <summary>Formato condizionale che si applica alla cella (prima regola valida per priorità, poi le altre per le proprietà non ancora impostate).</summary>
    public CfLook? CondFormatAt(Sheet s, int r, int c, CultureInfo? ci = null)
    {
        if (s.CondFmt.Count == 0) return null;
        string? font = null, fill = null; bool bold = false, any = false;
        foreach (var rule in CfRules(s))
        {
            if (!rule.Area.Any(a => r >= a.R1 && r <= a.R2 && c >= a.C1 && c <= a.C2)) continue;
            if (!RuleMatches(rule, s, r, c, ci)) continue;
            if (Styles.GetDxf(rule.DxfId) is { } d)
            {
                any = true; font ??= d.Font; fill ??= d.Fill; bold |= d.Bold;
            }
            if (rule.Stop) break;
        }
        return any ? new CfLook(font, fill, bold) : null;
    }

    bool RuleMatches(CfRuleC rule, Sheet s, int r, int c, CultureInfo? ci)
    {
        var cell = s.Get(r, c);
        if (rule.Type == "containsText")
        {
            string t = rule.Text ?? ""; if (t.Length == 0) return false;
            return DisplayText(s, r, c, ci).Contains(t, StringComparison.OrdinalIgnoreCase);
        }
        var cx = new EvalCtx(this, s, r, c);
        object? Ev(string? f)
        {
            if (f == null) return null;
            try { return V.Scalar(Evaluator.EvalFormula(f, cx), cx); } catch (Exception ex) when (ex is ExcelEx or FormatException or UnsupportedEx) { return null; }
        }
        var a = Ev(rule.F1); var b = Ev(rule.F2);
        if (a == null) return false;
        object? v = cell.Value;
        if (cell.Kind == CellKind.Empty) return false;
        if (cell.Kind == CellKind.Err) return false;
        int Cmp(object? x, object? y)
        {
            if (V.TryNum(x, out var dx) && x is not string && V.TryNum(y, out var dy) && y is not string) return dx.CompareTo(dy);
            if (x is string sx && y is string sy) return string.Compare(sx, sy, StringComparison.OrdinalIgnoreCase);
            return x is string ? 1 : -1;   // come Excel: testo > numeri
        }
        return rule.Op switch
        {
            "equal" => Cmp(v, a) == 0,
            "notEqual" => Cmp(v, a) != 0,
            "greaterThan" => Cmp(v, a) > 0,
            "lessThan" => Cmp(v, a) < 0,
            "greaterThanOrEqual" => Cmp(v, a) >= 0,
            "lessThanOrEqual" => Cmp(v, a) <= 0,
            "between" => b != null && Cmp(v, a) >= 0 && Cmp(v, b) <= 0,
            "notBetween" => b != null && (Cmp(v, a) < 0 || Cmp(v, b) > 0),
            _ => false,
        };
    }

    /// <summary>Aggiunge una regola all'area (priorità sopra quelle esistenti).</summary>
    public void AddCondFormat(Sheet s, int r1, int c1, int r2, int c2, CfSpec spec)
    {
        if (spec.Kind == "cellIs")
        {
            if (spec.Op is not ("greaterThan" or "lessThan" or "equal" or "notEqual" or "greaterThanOrEqual" or "lessThanOrEqual" or "between" or "notBetween")) throw new ArgumentException("Operatore non valido.");
            if (string.IsNullOrWhiteSpace(spec.F1)) throw new ArgumentException("Indicare il valore.");
            if (spec.Op is "between" or "notBetween" && string.IsNullOrWhiteSpace(spec.F2)) throw new ArgumentException("Indicare il secondo valore.");
            foreach (var f in new[] { spec.F1, spec.F2 }) if (!string.IsNullOrWhiteSpace(f)) try { Parser.Parse(f); } catch (FormatException ex) { throw new ArgumentException("Formula non valida: " + ex.Message); }
        }
        else if (spec.Kind == "containsText") { if (string.IsNullOrEmpty(spec.Text)) throw new ArgumentException("Indicare il testo."); }
        else throw new ArgumentException("Tipo di regola non supportato.");
        if (spec.FontColor == null && spec.FillColor == null && !spec.Bold) throw new ArgumentException("Scegliere un formato.");
        ChangeSections(s, "Formattazione condizionale", () =>
        {
            int dxf = Styles.AddDxf(spec.FontColor, spec.FillColor, spec.Bold);
            int prio = 1;
            foreach (var cf in s.CondFmt) foreach (var r in cf.Elements(MainX + "cfRule")) { int p = (int?)r.Attribute("priority") ?? 0; r.SetAttributeValue("priority", p + 1); }
            var rule = new XElement(MainX + "cfRule", new XAttribute("type", spec.Kind), new XAttribute("dxfId", dxf), new XAttribute("priority", prio));
            string tl = Addr.A1(r1, c1);
            if (spec.Kind == "cellIs")
            {
                rule.SetAttributeValue("operator", spec.Op);
                rule.Add(new XElement(MainX + "formula", spec.F1));
                if (spec.Op is "between" or "notBetween") rule.Add(new XElement(MainX + "formula", spec.F2));
            }
            else
            {
                string q = spec.Text!.Replace("\"", "\"\"");
                rule.SetAttributeValue("operator", "containsText"); rule.SetAttributeValue("text", spec.Text);
                rule.Add(new XElement(MainX + "formula", $"NOT(ISERROR(SEARCH(\"{q}\",{tl})))"));
            }
            s.CondFmt.Insert(0, new XElement(MainX + "conditionalFormatting", new XAttribute("sqref", FormatSqref([(r1, c1, r2, c2)])), rule));
        });
    }

    /// <summary>Cancella le regole dall'area (le regole più ampie vengono ristrette; formule relative non riadattate).</summary>
    public void ClearCondFormat(Sheet s, int r1, int c1, int r2, int c2)
    {
        if (s.CondFmt.Count == 0) return;
        ChangeSections(s, "Cancella regole", () => CutSqref(s.CondFmt, "sqref", (r1, c1, r2, c2)));
    }

    /// <summary>Descrizioni delle regole che toccano la cella (per la finestra di gestione).</summary>
    public int CondFormatCount(Sheet s) => s.CondFmt.Count;
}
