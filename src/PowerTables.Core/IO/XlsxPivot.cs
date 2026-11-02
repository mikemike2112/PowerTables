using System.Globalization;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace PowerTables.Core.IO;

/// <summary>Scrive (o toglie, dopo Annulla) le tabelle pivot create da PowerTables. Le pivot già nel file non vengono toccate. Excel le ricalcola all'apertura (refreshOnLoad, nessun record in cache).</summary>
static class XlsxPivot
{
    static readonly XNamespace M = XlsxReader.MainNs, R = XlsxReader.RelNs;

    static string? NameOf(PivotTablePart p)
    {
        try { using var s = p.GetStream(FileMode.Open, FileAccess.Read); return (string?)XDocument.Load(s).Root?.Attribute("name"); }
        catch (System.Xml.XmlException) { return null; }
    }

    public static void Apply(Workbook wb, WorkbookPart wbp, DocumentFormat.OpenXml.Spreadsheet.Workbook wbRoot, ref bool wbChanged)
    {
        foreach (var sh in wb.Sheets.Where(x => x.PivotsDirty && x.PartId != null))
        {
            var wsp = (WorksheetPart)wbp.GetPartById(sh.PartId!);
            var existing = wsp.PivotTableParts.Select(p => (Part: p, Name: NameOf(p))).ToList();
            var wanted = sh.Pivots.Select(p => p.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);
            foreach (var (part, name) in existing)
            {
                if (name == null || !sh.PivotWritten.Contains(name) || wanted.Contains(name)) continue;
                var cache = part.PivotTableCacheDefinitionPart;
                if (cache != null)
                {
                    string rid = wbp.GetIdOfPart(cache);
                    wbRoot.PivotCaches?.Elements<PivotCache>().FirstOrDefault(c => c.Id?.Value == rid)?.Remove();
                    if (wbRoot.PivotCaches != null && !wbRoot.PivotCaches.HasChildren) wbRoot.PivotCaches.Remove();
                    wsp.DeletePart(part); wbp.DeletePart(cache);
                }
                else wsp.DeletePart(part);
                sh.PivotWritten.Remove(name); wbChanged = true;
            }
            var have = existing.Select(e => e.Name).Where(n => n != null).ToHashSet(StringComparer.OrdinalIgnoreCase);
            foreach (var def in sh.Pivots.Where(d => !have.Contains(d.Name)))
            {
                var src = wb.Sheets.First(x => x.Name == def.SrcSheet);
                var cdp = wbp.AddNewPart<PivotTableCacheDefinitionPart>();
                using (var os = cdp.GetStream(FileMode.Create)) CacheXml(def, src).Save(os);
                uint cacheId = (wbRoot.PivotCaches?.Elements<PivotCache>().Select(c => c.CacheId?.Value ?? 0u).DefaultIfEmpty(0u).Max() ?? 0u) + 1;
                wbRoot.PivotCaches ??= new PivotCaches();
                wbRoot.PivotCaches.Append(new PivotCache { CacheId = cacheId, Id = wbp.GetIdOfPart(cdp) });
                var pp = wsp.AddNewPart<PivotTablePart>(); pp.AddPart(cdp);
                using (var os = pp.GetStream(FileMode.Create)) TableXml(def, cacheId).Save(os);
                sh.PivotWritten.Add(def.Name); wbChanged = true;
            }
            sh.PivotsDirty = false;
        }
    }

    static string Inv(double d) => d.ToString("R", CultureInfo.InvariantCulture);

    static XDocument CacheXml(PivotDef d, Sheet src)
    {
        int recs = d.SR2 - d.SR1;
        var fields = new XElement(M + "cacheFields", new XAttribute("count", d.FieldNames.Length));
        for (int c = d.SC1; c <= d.SC2; c++)
        {
            var si = new XElement(M + "sharedItems");
            var nums = Enumerable.Range(d.SR1 + 1, recs).Select(r => src.Get(r, c)).Where(x => x.Kind == CellKind.Num).Select(x => x.Num).ToList();
            bool allNum = nums.Count == recs && recs > 0;
            if (c == d.RowCol)
            {
                if (d.RowItems.Count > 0 && d.RowItems[0] is double)
                {
                    si.Add(new XAttribute("containsSemiMixedTypes", 0), new XAttribute("containsString", 0), new XAttribute("containsNumber", 1), new XAttribute("minValue", Inv(nums.Min())), new XAttribute("maxValue", Inv(nums.Max())), new XAttribute("count", d.RowItems.Count));
                    foreach (var it in d.RowItems) si.Add(new XElement(M + "n", new XAttribute("v", Inv((double)it))));
                }
                else
                {
                    si.Add(new XAttribute("count", d.RowItems.Count));
                    foreach (var it in d.RowItems) si.Add(new XElement(M + "s", new XAttribute("v", (string)it)));
                }
            }
            else if (allNum)
            {
                si.Add(new XAttribute("containsSemiMixedTypes", 0), new XAttribute("containsString", 0), new XAttribute("containsNumber", 1));
                if (nums.All(x => x == Math.Floor(x))) si.Add(new XAttribute("containsInteger", 1));
                si.Add(new XAttribute("minValue", Inv(nums.Min())), new XAttribute("maxValue", Inv(nums.Max())));
            }
            fields.Add(new XElement(M + "cacheField", new XAttribute("name", d.FieldNames[c - d.SC1]), new XAttribute("numFmtId", 0), si));
        }
        string aref = Addr.A1(d.SR1, d.SC1) + ":" + Addr.A1(d.SR2, d.SC2);
        return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), new XElement(M + "pivotCacheDefinition", new XAttribute("saveData", 0), new XAttribute("refreshOnLoad", 1), new XAttribute("refreshedBy", "PowerTables"), new XAttribute("refreshedDate", 45000), new XAttribute("createdVersion", 6), new XAttribute("refreshedVersion", 6), new XAttribute("minRefreshableVersion", 3), new XAttribute("recordCount", recs),
            new XElement(M + "cacheSource", new XAttribute("type", "worksheet"), new XElement(M + "worksheetSource", new XAttribute("ref", aref), new XAttribute("sheet", d.SrcSheet))), fields));
    }

    static XDocument TableXml(PivotDef d, uint cacheId)
    {
        int n = d.FieldNames.Length, k = d.RowItems.Count;
        var pf = new XElement(M + "pivotFields", new XAttribute("count", n));
        for (int c = d.SC1; c <= d.SC2; c++)
        {
            if (c == d.RowCol)
            {
                var items = new XElement(M + "items", new XAttribute("count", k + 1));
                for (int i = 0; i < k; i++) items.Add(new XElement(M + "item", new XAttribute("x", i)));
                items.Add(new XElement(M + "item", new XAttribute("t", "default")));
                pf.Add(new XElement(M + "pivotField", new XAttribute("axis", "axisRow"), new XAttribute("showAll", 0), items));
            }
            else if (c == d.DataCol) pf.Add(new XElement(M + "pivotField", new XAttribute("dataField", 1), new XAttribute("showAll", 0)));
            else pf.Add(new XElement(M + "pivotField", new XAttribute("showAll", 0)));
        }
        var rowItems = new XElement(M + "rowItems", new XAttribute("count", k + 1));
        for (int i = 0; i < k; i++) rowItems.Add(new XElement(M + "i", i == 0 ? new XElement(M + "x") : new XElement(M + "x", new XAttribute("v", i))));
        rowItems.Add(new XElement(M + "i", new XAttribute("t", "grand"), new XElement(M + "x")));
        var df = new XElement(M + "dataField", new XAttribute("name", d.DataCaption), new XAttribute("fld", d.DataCol - d.SC1), new XAttribute("baseField", 0), new XAttribute("baseItem", 0));
        if (d.Count) df.SetAttributeValue("subtotal", "count");
        return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), new XElement(M + "pivotTableDefinition", new XAttribute("name", d.Name), new XAttribute("cacheId", cacheId), new XAttribute("applyNumberFormats", 0), new XAttribute("applyBorderFormats", 0), new XAttribute("applyFontFormats", 0), new XAttribute("applyPatternFormats", 0), new XAttribute("applyAlignmentFormats", 0), new XAttribute("applyWidthHeightFormats", 1), new XAttribute("dataCaption", "Valori"), new XAttribute("updatedVersion", 6), new XAttribute("minRefreshableVersion", 3), new XAttribute("useAutoFormatting", 1), new XAttribute("itemPrintTitles", 1), new XAttribute("createdVersion", 6), new XAttribute("indent", 0), new XAttribute("outline", 1), new XAttribute("outlineData", 1), new XAttribute("multipleFieldFilters", 0),
            new XElement(M + "location", new XAttribute("ref", Addr.A1(d.R1, d.C1) + ":" + Addr.A1(d.R2, d.C2)), new XAttribute("firstHeaderRow", 1), new XAttribute("firstDataRow", 1), new XAttribute("firstDataCol", 1)),
            pf, new XElement(M + "rowFields", new XAttribute("count", 1), new XElement(M + "field", new XAttribute("x", d.RowCol - d.SC1))), rowItems,
            new XElement(M + "colItems", new XAttribute("count", 1), new XElement(M + "i")), new XElement(M + "dataFields", new XAttribute("count", 1), df),
            new XElement(M + "pivotTableStyleInfo", new XAttribute("name", "PivotStyleLight16"), new XAttribute("showRowHeaders", 1), new XAttribute("showColHeaders", 1), new XAttribute("showRowStripes", 0), new XAttribute("showColStripes", 0), new XAttribute("showLastColumn", 1))));
    }
}
