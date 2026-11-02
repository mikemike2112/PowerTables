using System.Xml.Linq;
using PowerTables.Core.Formulas;

namespace PowerTables.Core;

/// <summary>Sezioni XML del foglio (formattazione condizionale, convalida, collegamenti): copia e adattamento a inserimento/eliminazione di righe e colonne.</summary>
public sealed partial class Workbook
{
    static readonly XNamespace MainX = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    static List<XElement> CloneList(List<XElement> l) => l.Select(e => new XElement(e)).ToList();

    static void SetList(List<XElement> dst, List<XElement> src) { dst.Clear(); dst.AddRange(src.Select(e => new XElement(e))); }

    void AdjustSections(Sheet s, bool rows, int at, int count)
    {
        s.CfVersion++;
        string Adj(string f) => FormulaText.AdjustLines(f, s.Name, s.Name, rows, at, count);
        foreach (var cf in s.CondFmt.ToList())
        {
            string sq = FormulaText.AdjustSqref((string?)cf.Attribute("sqref") ?? "", rows, at, count);
            if (sq.Length == 0) { s.CondFmt.Remove(cf); continue; }
            cf.SetAttributeValue("sqref", sq);
            foreach (var f in cf.Descendants(MainX + "formula")) f.Value = Adj(f.Value);
        }
        foreach (var dv in s.DataVal.ToList())
        {
            string sq = FormulaText.AdjustSqref((string?)dv.Attribute("sqref") ?? "", rows, at, count);
            if (sq.Length == 0) { s.DataVal.Remove(dv); continue; }
            dv.SetAttributeValue("sqref", sq);
            foreach (var f in dv.Elements().Where(e => e.Name.LocalName is "formula1" or "formula2")) f.Value = Adj(f.Value);
        }
        foreach (var h in s.Links.ToList())
        {
            string sq = FormulaText.AdjustSqref((string?)h.Attribute("ref") ?? "", rows, at, count);
            if (sq.Length == 0) s.Links.Remove(h); else h.SetAttributeValue("ref", sq);
        }
    }
}
