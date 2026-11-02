using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace PowerTables.Core.IO;

/// <summary>Salvataggio: riscrive solo le righe/parti toccate, il resto del package resta invariato.</summary>
public static class XlsxWriter
{
    static readonly XNamespace M = XlsxReader.MainNs;

    static readonly string[] SheetOrder =
    [
        "sheetPr","dimension","sheetViews","sheetFormatPr","cols","sheetData","sheetCalcPr","sheetProtection","protectedRanges","scenarios",
        "autoFilter","sortState","dataConsolidate","customSheetViews","mergeCells","phoneticPr","conditionalFormatting","dataValidations",
        "hyperlinks","printOptions","pageMargins","pageSetup","headerFooter","rowBreaks","colBreaks","customProperties","cellWatches",
        "ignoredErrors","smartTags","drawing","legacyDrawing","legacyDrawingHF","drawingHF","picture","oleObjects","controls","webPublishItems",
        "tableParts","extLst",
    ];

    /// <summary>Nuova cartella con un foglio vuoto (package creato da zero: solo per documenti nuovi).</summary>
    public static Workbook CreateNew()
    {
        var st = new PackageState { Ms = new MemoryStream() };
        st.Doc = SpreadsheetDocument.Create(st.Ms, SpreadsheetDocumentType.Workbook, true);
        var wbp = st.Doc.AddWorkbookPart();
        wbp.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook(new Sheets());
        var stp = wbp.AddNewPart<WorkbookStylesPart>();
        var wb = new Workbook { Pkg = st };
        using (var s = stp.GetStream(FileMode.Create)) wb.Styles.Doc.Save(s);
        wb.Styles.MarkClean();
        wb.AddSheetPartFor(wb.AddSheet());
        wb.Modified = false;
        return wb;
    }

    internal static void AddSheetPartFor(this Workbook wb, Sheet sh)
    {
        var st = wb.Pkg!; var wbp = st.Doc.WorkbookPart!;
        var wsp = wbp.AddNewPart<WorksheetPart>();
        using (var s = wsp.GetStream(FileMode.Create))
        {
            var bytes = Encoding.UTF8.GetBytes("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<worksheet xmlns=\"" + XlsxReader.MainNs + "\"><dimension ref=\"A1\"/><sheetData/></worksheet>");
            s.Write(bytes);
        }
        sh.PartId = wbp.GetIdOfPart(wsp);
        sh.OrigName = sh.Name;
        var wbRoot = wbp.Workbook ?? throw new XlsxException("workbook.xml assente o illeggibile");
        var sheets = wbRoot.Sheets ??= new Sheets();
        uint id = sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Select(x => x.SheetId?.Value ?? 0).DefaultIfEmpty(0u).Max() + 1;
        sheets.Append(new DocumentFormat.OpenXml.Spreadsheet.Sheet { Name = sh.Name, SheetId = id, Id = sh.PartId });
        sh.DirtyRows.Add(1);
    }

    /// <summary>Salvataggio atomico: scrive su un file temporaneo e poi sostituisce la destinazione.</summary>
    public static void Save(Workbook wb, string path)
    {
        var st = wb.Pkg ?? throw new InvalidOperationException("nessun package");
        var wbp = st.Doc.WorkbookPart!;
        // nuovi fogli e rinomine
        var wbRoot = wbp.Workbook ?? throw new XlsxException("workbook.xml assente o illeggibile");
        bool wbChanged = false;
        foreach (var sh in wb.Sheets)
        {
            if (sh.PartId == null) { wb.AddSheetPartFor(sh); wbChanged = true; }
            if (sh.OrigName == sh.Name) continue;
            var el = wbRoot.Sheets?.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().FirstOrDefault(x => x.Id?.Value == sh.PartId);
            if (el != null && el.Name?.Value != sh.Name) { el.Name = sh.Name; wbChanged = true; }
            sh.OrigName = sh.Name;
        }
        if (wb.Sheets.Any(x => x.LineOps.Count > 0) && wbRoot.DefinedNames != null)
            foreach (var dn in wbRoot.DefinedNames.Elements<DefinedName>())
            {
                string t = dn.Text ?? "";
                foreach (var sh in wb.Sheets) foreach (var op in sh.LineOps) t = Formulas.FormulaText.AdjustLines(t, "", op.Sheet, op.Rows, op.At, op.Count);
                if (t != dn.Text) { dn.Text = t; wbChanged = true; }
            }
        if (wb.NamesDirty) { RewriteNames(wb, wbRoot); wbChanged = true; wb.NamesDirty = false; }
        if (wbChanged) wbRoot.Save();

        foreach (var sh in wb.Sheets)
        {
            if (!sh.Loaded) continue;
            if (sh.DirtyRows.Count == 0 && !sh.LayoutDirty) continue;
            var wsp = (WorksheetPart)wbp.GetPartById(sh.PartId!);
            byte[] orig;
            using (var rs = wsp.GetStream(FileMode.Open, FileAccess.Read)) { var ms = new MemoryStream(); rs.CopyTo(ms); orig = ms.ToArray(); }
            var bytes = RewriteSheet(sh, orig);
            using (var s = wsp.GetStream(FileMode.Create)) s.Write(bytes);
            sh.DirtyRows.Clear(); sh.LayoutDirty = false;
        }
        foreach (var sh in wb.Sheets) sh.LineOps.Clear();
        if (wb.Styles.Dirty)
        {
            var sp = wbp.WorkbookStylesPart ?? wbp.AddNewPart<WorkbookStylesPart>();
            using var s = sp.GetStream(FileMode.Create);
            wb.Styles.Doc.Save(s);
            wb.Styles.MarkClean();
        }
        if (wb.FormulasChangedSinceOpen && wbp.CalculationChainPart != null) wbp.DeletePart(wbp.CalculationChainPart); // Excel la ricostruisce

        string dir = Path.GetDirectoryName(Path.GetFullPath(path))!;
        string tmp = Path.Combine(dir, "." + Path.GetFileName(path) + "." + Guid.NewGuid().ToString("N")[..8] + ".tmp");
        try
        {
            using (var clone = st.Doc.Clone(tmp, false)) { }
            if (File.Exists(path)) File.Replace(tmp, path, null);
            else File.Move(tmp, path);
        }
        catch
        {
            try { if (File.Exists(tmp)) File.Delete(tmp); } catch { }
            throw;
        }
        wb.FilePath = path;
        wb.Modified = false;
    }

    /// <summary>Riscrive definedNames dal modello: conserva i nomi riservati/nascosti, aggiorna i testi, aggiunge i nuovi, toglie gli eliminati.</summary>
    static void RewriteNames(Workbook wb, DocumentFormat.OpenXml.Spreadsheet.Workbook wbRoot)
    {
        var dns = wbRoot.DefinedNames ?? new DefinedNames();
        var done = new HashSet<(int, string)>(); var doneGlobal = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var e in dns.Elements<DefinedName>().ToList())
        {
            string n = e.Name?.Value ?? ""; int? ls = e.LocalSheetId?.Value is { } v ? (int)v : null;
            if (n == "_xlnm.Print_Area") { e.Remove(); continue; }
            if (n.StartsWith("_xlnm.") || (e.Hidden?.Value == true && n.StartsWith("_"))) continue;
            if (ls is { } l)
            {
                var key = (l, n.ToUpperInvariant());
                if (wb.LocalNames.TryGetValue(key, out var t)) { e.Text = t; done.Add(key); } else e.Remove();
            }
            else if (wb.Names.TryGetValue(n, out var t)) { e.Text = t; doneGlobal.Add(n); }
            else e.Remove();
        }
        foreach (var kv in wb.Names.Where(k => !doneGlobal.Contains(k.Key)).OrderBy(k => k.Key, StringComparer.OrdinalIgnoreCase)) dns.Append(new DefinedName { Name = kv.Key, Text = kv.Value });
        foreach (var kv in wb.LocalNames.Where(k => !done.Contains(k.Key)))
            dns.Append(new DefinedName { Name = wb.LocalNameCase.GetValueOrDefault(kv.Key, kv.Key.Item2), LocalSheetId = (uint)kv.Key.Item1, Text = kv.Value });
        foreach (var sh in wb.Sheets.Where(x => x.PrintArea != null))
            dns.Append(new DefinedName { Name = "_xlnm.Print_Area", LocalSheetId = (uint)sh.Index, Text = sh.PrintArea });
        if (!dns.HasChildren) { if (dns.Parent != null) dns.Remove(); return; }
        if (dns.Parent == null) wbRoot.DefinedNames = dns;
    }

    // ---------------- riscrittura di una parte foglio ----------------
    static byte[] RewriteSheet(Sheet sh, byte[] orig)
    {
        string text = Encoding.UTF8.GetString(orig);
        if (text.Length > 0 && text[0] == '﻿') text = text[1..];
        // individua <sheetData ...> ... </sheetData>
        var open = Regex.Match(text, @"<((?<p>[\w.-]+):)?sheetData(\s[^>]*)?(?<sc>/)?>");
        if (!open.Success) throw new XlsxException("Foglio senza sheetData");
        string prefix = open.Groups["p"].Success ? open.Groups["p"].Value + ":" : "";
        int dataStart = open.Index, dataEnd;
        string inner = "";
        if (open.Groups["sc"].Success) dataEnd = open.Index + open.Length;
        else
        {
            int close = text.IndexOf("</" + prefix + "sheetData>", open.Index + open.Length, StringComparison.Ordinal);
            if (close < 0) throw new XlsxException("sheetData non chiuso");
            inner = text.Substring(open.Index + open.Length, close - open.Index - open.Length);
            dataEnd = close + ("</" + prefix + "sheetData>").Length;
        }
        string head = text[..dataStart], tail = text[dataEnd..];
        const string Mark = "@@PT_SHEETDATA@@";
        var doc = XDocument.Parse(head + "<" + prefix + "sheetData>" + Mark + "</" + prefix + "sheetData>" + tail, LoadOptions.PreserveWhitespace);
        var root = doc.Root!;
        var ns = root.Name.Namespace;

        // dimension
        int maxR = Math.Max(1, sh.MaxRow), maxC = Math.Max(1, sh.MaxCol);
        SetChild(root, ns, "dimension", new XAttribute("ref", sh.MaxRow == 0 ? "A1" : "A1:" + Addr.A1(maxR, maxC)), replace: true);
        if (sh.LayoutDirty) ApplyLayout(sh, root, ns);
        foreach (var op in sh.LineOps) ApplyLineOp(root, ns, op);

        // righe
        var sb = new StringBuilder(inner.Length + 4096);
        var slices = SliceRows(inner);
        var rowsInStore = new SortedSet<int>(slices.Keys);
        foreach (var r in sh.DirtyRows) rowsInStore.Add(r);
        // righe che hanno celle nello store ma non erano nell'originale e non sono sporche non esistono (tutte le scritture marcano sporco)
        foreach (var r in rowsInStore)
        {
            if (!sh.DirtyRows.Contains(r)) { var (a, b) = slices[r]; sb.Append(inner, a, b - a); continue; }
            WriteRow(sb, sh, r, prefix);
        }
        string xml = root.ToString(SaveOptions.DisableFormatting);
        var m = Regex.Match(xml, @"<" + Regex.Escape(prefix) + @"sheetData>" + Mark + "</" + Regex.Escape(prefix) + "sheetData>");
        if (!m.Success) throw new XlsxException("marcatore perso");
        string body = sb.Length == 0 ? "<" + prefix + "sheetData/>" : "<" + prefix + "sheetData>" + sb + "</" + prefix + "sheetData>";
        string full = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" + xml[..m.Index] + body + xml[(m.Index + m.Length)..];
        return new UTF8Encoding(false).GetBytes(full);
    }

    static Dictionary<int, (int Start, int End)> SliceRows(string inner)
    {
        var res = new Dictionary<int, (int, int)>();
        int pos = 0; int lastRow = 0;
        while (true)
        {
            int s = FindRowStart(inner, pos);
            if (s < 0) break;
            int gt = inner.IndexOf('>', s);
            bool selfClose = inner[gt - 1] == '/';
            int e;
            if (selfClose) e = gt + 1;
            else
            {
                int c = inner.IndexOf("</", gt, StringComparison.Ordinal);
                // chiusura </row> o </x:row>: la prima chiusura dopo la riga è sempre quella di riga perché le celle non contengono <row
                while (c >= 0 && !IsRowClose(inner, c)) c = inner.IndexOf("</", c + 2, StringComparison.Ordinal);
                if (c < 0) throw new XlsxException("riga non chiusa");
                e = inner.IndexOf('>', c) + 1;
            }
            string tag = inner.Substring(s, gt - s);
            var rm = Regex.Match(tag, @"\sr=""(\d+)""");
            int rn = rm.Success ? int.Parse(rm.Groups[1].Value) : lastRow + 1;
            lastRow = rn;
            res[rn] = (s, e);
            pos = e;
        }
        return res;
    }

    static int FindRowStart(string t, int pos)
    {
        while (true)
        {
            int i = t.IndexOf('<', pos);
            if (i < 0) return -1;
            int j = i + 1;
            int colon = t.IndexOf(':', j, Math.Min(12, t.Length - j));
            int nameStart = colon > 0 && colon < j + 10 && t[j..colon].All(char.IsLetterOrDigit) ? colon + 1 : j;
            if (string.CompareOrdinal(t, nameStart, "row", 0, 3) == 0 && nameStart + 3 < t.Length && (t[nameStart + 3] is ' ' or '>' or '/' or '\r' or '\n' or '\t')) return i;
            pos = j;
        }
    }

    static bool IsRowClose(string t, int c)
    {
        int j = c + 2; int colon = t.IndexOf(':', j, Math.Min(12, t.Length - j));
        int nameStart = colon > 0 && t[j..colon].All(char.IsLetterOrDigit) ? colon + 1 : j;
        return string.CompareOrdinal(t, nameStart, "row", 0, 3) == 0 && t[nameStart + 3] == '>';
    }

    static void WriteRow(StringBuilder sb, Sheet sh, int row, string p)
    {
        sh.Rows.TryGetValue(row, out var ri);
        var cells = new List<(int Col, Cell C)>();
        // celle della riga: scansione colonne fino a MaxCol (le righe larghe sono rare); per fogli molto larghi si limita alle colonne con celle
        int maxC = sh.MaxCol;
        for (int c = 1; c <= maxC; c++)
            if (sh.Cells.TryGetValue(Addr.Key(row, c), out var cell)) cells.Add((c, cell));
        if (cells.Count == 0 && ri == null) return;
        sb.Append('<').Append(p).Append("row r=\"").Append(row).Append('"');
        if (ri != null)
        {
            if (ri.CustomFormat || ri.Style != 0) sb.Append(" s=\"").Append(ri.Style).Append("\" customFormat=\"1\"");
            if (ri.Height is { } h) sb.Append(" ht=\"").Append(h.ToString("R", CultureInfo.InvariantCulture)).Append('"');
            if (ri.Hidden) sb.Append(" hidden=\"1\"");
            if (ri.CustomHeight || (ri.Height != null && ri.CustomHeight)) sb.Append(" customHeight=\"1\"");
            foreach (var (n, px, nsu, val) in ri.Extra)
            {
                sb.Append(' ');
                if (px.Length > 0) sb.Append(px).Append(':');
                sb.Append(n).Append("=\"").Append(Esc(val)).Append('"');
            }
        }
        if (cells.Count == 0) { sb.Append("/>"); return; }
        sb.Append('>');
        foreach (var (col, c) in cells) WriteCell(sb, row, col, c, p, sh);
        sb.Append('<').Append('/').Append(p).Append("row>");
    }

    static void WriteCell(StringBuilder sb, int row, int col, in Cell c, string p, Sheet sh)
    {
        sb.Append('<').Append(p).Append("c r=\"").Append(Addr.ColName(col)).Append(row).Append('"');
        if (c.Style != 0) sb.Append(" s=\"").Append(c.Style).Append('"');
        string? t = c.Kind switch { CellKind.Str => c.F != null ? "str" : "inlineStr", CellKind.Bool => "b", CellKind.Err => "e", _ => null };
        if (t != null) sb.Append(" t=\"").Append(t).Append('"');
        if (c.F?.CellAttrs != null) sb.Append(c.F.CellAttrs);
        if (c.Kind == CellKind.Empty && c.F == null) { sb.Append("/>"); return; }
        sb.Append('>');
        if (c.F != null)
        {
            if (c.F.RawF != null) sb.Append(c.F.RawF);
            else sb.Append('<').Append(p).Append("f>").Append(Esc(c.F.Text)).Append("</").Append(p).Append("f>");
        }
        switch (c.Kind)
        {
            case CellKind.Num: sb.Append('<').Append(p).Append("v>").Append(c.Num.ToString("R", CultureInfo.InvariantCulture)).Append("</").Append(p).Append("v>"); break;
            case CellKind.Bool: sb.Append('<').Append(p).Append("v>").Append(c.Num != 0 ? '1' : '0').Append("</").Append(p).Append("v>"); break;
            case CellKind.Err: sb.Append('<').Append(p).Append("v>").Append(Errors.Text((CellError)(int)c.Num)).Append("</").Append(p).Append("v>"); break;
            case CellKind.Str:
                if (c.F != null) sb.Append('<').Append(p).Append("v>").Append(Esc((string?)c.Obj ?? "")).Append("</").Append(p).Append("v>");
                else if (sh.Rich.TryGetValue(Addr.Key(row, col), out var rt) && rt.Plain == (string?)c.Obj) sb.Append('<').Append(p).Append("is>").Append(rt.Xml).Append("</").Append(p).Append("is>");
                else sb.Append('<').Append(p).Append("is><").Append(p).Append("t xml:space=\"preserve\">").Append(Esc((string?)c.Obj ?? "")).Append("</").Append(p).Append("t></").Append(p).Append("is>"); // ponytail: stringhe scritte inline, niente rich text né sst per le righe riscritte
                break;
        }
        sb.Append("</").Append(p).Append("c>");
    }

    static string Esc(string s)
    {
        StringBuilder? sb = null;
        for (int i = 0; i < s.Length; i++)
        {
            char ch = s[i]; string? rep = ch switch { '&' => "&amp;", '<' => "&lt;", '>' => "&gt;", '"' => "&quot;", _ => null };
            if (rep == null && ch < 32 && ch is not ('\t' or '\n' or '\r')) rep = "";
            if (rep != null) { sb ??= new StringBuilder(s, 0, i, s.Length + 16); sb.Append(rep); }
            else sb?.Append(ch);
        }
        return sb?.ToString() ?? s;
    }

    /// <summary>Riporta un inserimento/eliminazione su formattazione condizionale, convalida dati e collegamenti.</summary>
    static void ApplyLineOp(XElement root, XNamespace ns, LineOp op)
    {
        string Adj(string f) => Formulas.FormulaText.AdjustLines(f, op.Sheet, op.Sheet, op.Rows, op.At, op.Count);
        foreach (var cf in root.Elements(ns + "conditionalFormatting").ToList())
        {
            string sq = Formulas.FormulaText.AdjustSqref((string?)cf.Attribute("sqref") ?? "", op.Rows, op.At, op.Count);
            if (sq.Length == 0) { cf.Remove(); continue; }
            cf.SetAttributeValue("sqref", sq);
            foreach (var f in cf.Descendants(ns + "formula")) f.Value = Adj(f.Value);
        }
        var dvs = root.Element(ns + "dataValidations");
        if (dvs != null)
        {
            foreach (var dv in dvs.Elements(ns + "dataValidation").ToList())
            {
                string sq = Formulas.FormulaText.AdjustSqref((string?)dv.Attribute("sqref") ?? "", op.Rows, op.At, op.Count);
                if (sq.Length == 0) { dv.Remove(); continue; }
                dv.SetAttributeValue("sqref", sq);
                foreach (var f in dv.Elements().Where(e => e.Name.LocalName is "formula1" or "formula2")) f.Value = Adj(f.Value);
            }
            if (!dvs.HasElements) dvs.Remove(); else dvs.SetAttributeValue("count", dvs.Elements().Count());
        }
        var hl = root.Element(ns + "hyperlinks");
        if (hl != null)
        {
            foreach (var h in hl.Elements(ns + "hyperlink").ToList())
            {
                string sq = Formulas.FormulaText.AdjustSqref((string?)h.Attribute("ref") ?? "", op.Rows, op.At, op.Count);
                if (sq.Length == 0) h.Remove(); else h.SetAttributeValue("ref", sq);
            }
            if (!hl.HasElements) hl.Remove();
        }
    }

    // ---------------- layout (cols, merge, pane, filtro) ----------------
    static void InsertOrdered(XElement root, XNamespace ns, XElement e)
    {
        int idx = Array.IndexOf(SheetOrder, e.Name.LocalName);
        var before = root.Elements().FirstOrDefault(x => Array.IndexOf(SheetOrder, x.Name.LocalName) > idx);
        if (before != null) before.AddBeforeSelf(e); else root.Add(e);
    }

    static void SetChild(XElement root, XNamespace ns, string name, XAttribute attr, bool replace)
    {
        var ex = root.Element(ns + name);
        if (ex != null) { ex.SetAttributeValue(attr.Name, attr.Value); return; }
        InsertOrdered(root, ns, new XElement(ns + name, attr));
    }

    static void ApplyLayout(Sheet sh, XElement root, XNamespace ns)
    {
        // cols
        root.Element(ns + "cols")?.Remove();
        if (sh.Cols.Count > 0)
        {
            var cols = new XElement(ns + "cols");
            var keys = sh.Cols.Keys.OrderBy(k => k).ToList();
            for (int i = 0; i < keys.Count;)
            {
                int a = keys[i], b = a; var ci = sh.Cols[a];
                int j = i + 1;
                while (j < keys.Count && keys[j] == b + 1 && Same(sh.Cols[keys[j]], ci)) { b = keys[j]; j++; }
                var e = new XElement(ns + "col", new XAttribute("min", a), new XAttribute("max", b),
                    new XAttribute("width", ci.Width.ToString("R", CultureInfo.InvariantCulture)));
                if (ci.Style != 0) e.SetAttributeValue("style", ci.Style);
                if (ci.Hidden) e.SetAttributeValue("hidden", 1);
                if (ci.CustomWidth || Math.Abs(ci.Width - sh.DefaultColWidth) > 1e-9) e.SetAttributeValue("customWidth", 1);
                cols.Add(e);
                i = j;
            }
            InsertOrdered(root, ns, cols);
        }
        // merge
        root.Element(ns + "mergeCells")?.Remove();
        if (sh.Merges.Count > 0)
        {
            var mc = new XElement(ns + "mergeCells", new XAttribute("count", sh.Merges.Count));
            foreach (var m in sh.Merges) mc.Add(new XElement(ns + "mergeCell", new XAttribute("ref", Addr.A1(m.R1, m.C1) + ":" + Addr.A1(m.R2, m.C2))));
            InsertOrdered(root, ns, mc);
        }
        // riquadri bloccati
        var svs = root.Element(ns + "sheetViews");
        if (svs == null) { svs = new XElement(ns + "sheetViews", new XElement(ns + "sheetView", new XAttribute("workbookViewId", 0))); InsertOrdered(root, ns, svs); }
        var sv = svs.Elements(ns + "sheetView").First();
        sv.Element(ns + "pane")?.Remove();
        foreach (var s in sv.Elements(ns + "selection").ToList()) s.Remove();
        if (sh.FrozenRows > 0 || sh.FrozenCols > 0)
        {
            string active = sh.FrozenRows > 0 && sh.FrozenCols > 0 ? "bottomRight" : sh.FrozenRows > 0 ? "bottomLeft" : "topRight";
            var pane = new XElement(ns + "pane");
            if (sh.FrozenCols > 0) pane.SetAttributeValue("xSplit", sh.FrozenCols);
            if (sh.FrozenRows > 0) pane.SetAttributeValue("ySplit", sh.FrozenRows);
            pane.SetAttributeValue("topLeftCell", Addr.A1(sh.FrozenRows + 1, sh.FrozenCols + 1));
            pane.SetAttributeValue("activePane", active);
            pane.SetAttributeValue("state", "frozen");
            sv.AddFirst(pane);
            sv.Add(new XElement(ns + "selection", new XAttribute("pane", active), new XAttribute("activeCell", Addr.A1(sh.FrozenRows + 1, sh.FrozenCols + 1)), new XAttribute("sqref", Addr.A1(sh.FrozenRows + 1, sh.FrozenCols + 1))));
        }
        else sv.Add(new XElement(ns + "selection", new XAttribute("activeCell", "A1"), new XAttribute("sqref", "A1")));
        // filtro
        if (sh.Filter is { Dirty: true } f)
        {
            root.Element(ns + "autoFilter")?.Remove();
            var af = new XElement(ns + "autoFilter", new XAttribute("ref", Addr.A1(f.R1, f.C1) + ":" + Addr.A1(f.R2, f.C2)));
            foreach (var (col, vals) in f.Allowed.OrderBy(k => k.Key))
            {
                var fc = new XElement(ns + "filterColumn", new XAttribute("colId", col - f.C1));
                var fl = new XElement(ns + "filters");
                foreach (var v in vals.OrderBy(x => x, StringComparer.Ordinal)) fl.Add(new XElement(ns + "filter", new XAttribute("val", v)));
                fc.Add(fl); af.Add(fc);
            }
            InsertOrdered(root, ns, af);
            f.Dirty = false;
        }
        else if (sh.Filter == null) root.Element(ns + "autoFilter")?.Remove();
    }

    static bool Same(ColInfo a, ColInfo b) => a.Width.Equals(b.Width) && a.Hidden == b.Hidden && a.Style == b.Style && a.CustomWidth == b.CustomWidth;
}
