using System.Globalization;
using System.IO.Packaging;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using PowerTables.Core.Formulas;

namespace PowerTables.Core.IO;

/// <summary>Stato del package originale mantenuto aperto per il round-trip.</summary>
internal sealed class PackageState : IDisposable
{
    public MemoryStream Ms = null!;
    public SpreadsheetDocument Doc = null!;
    public bool FormulasTouched;
    public void Dispose() { Doc?.Dispose(); Ms?.Dispose(); }
}

public sealed class XlsxException(string msg, Exception? inner = null) : Exception(msg, inner);

public static class XlsxReader
{
    public const string MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    public const string RelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    static XmlReaderSettings Settings() => new() { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, IgnoreWhitespace = false, IgnoreComments = true, CheckCharacters = false };

    /// <summary>Apre un .xlsx. I file sono considerati non fidati: ogni errore di struttura diventa XlsxException.</summary>
    public static Workbook Open(string path)
    {
        byte[] bytes;
        try { bytes = File.ReadAllBytes(path); }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { throw new XlsxException("Impossibile leggere il file: " + e.Message, e); }
        return Open(bytes, path);
    }

    public static Workbook Open(byte[] bytes, string? path = null)
    {
        var st = new PackageState { Ms = new MemoryStream() };
        st.Ms.Write(bytes); st.Ms.Position = 0;
        try
        {
            st.Doc = SpreadsheetDocument.Open(st.Ms, true, new OpenSettings { AutoSave = false });
            var wbp = st.Doc.WorkbookPart ?? throw new XlsxException("Il file non contiene una cartella di lavoro.");
            var wb = new Workbook { FilePath = path, Pkg = st };
            LoadWorkbook(wb, wbp, st);
            wb.Modified = false;
            return wb;
        }
        catch (XlsxException) { st.Dispose(); throw; }
        catch (Exception e) when (e is InvalidDataException or IOException or XmlException or OpenXmlPackageException or FileFormatException or InvalidOperationException or ArgumentException or FormatException or NullReferenceException or KeyNotFoundException)
        {
            st.Dispose();
            throw new XlsxException("File .xlsx non valido o danneggiato: " + e.Message, e);
        }
    }

    static XDocument LoadXml(Stream s)
    {
        using var r = XmlReader.Create(s, Settings());
        return XDocument.Load(r, LoadOptions.PreserveWhitespace);
    }

    static void LoadWorkbook(Workbook wb, WorkbookPart wbp, PackageState st)
    {
        XNamespace m = MainNs, r = RelNs;
        XDocument wdoc;
        using (var s = wbp.GetStream(FileMode.Open, FileAccess.Read)) wdoc = LoadXml(s);
        wb.Date1904 = (bool?)wdoc.Root!.Element(m + "workbookPr")?.Attribute("date1904") ?? false;

        // stili
        var sp = wbp.WorkbookStylesPart;
        if (sp != null)
        {
            using var s = sp.GetStream(FileMode.Open, FileAccess.Read);
            wb.Styles = new StyleTable(LoadXml(s));
        }

        // stringhe condivise
        string[] sst = []; Dictionary<int, string> richSst = [];
        var ssp = wbp.SharedStringTablePart;
        if (ssp != null) { using var s = ssp.GetStream(FileMode.Open, FileAccess.Read); sst = ReadSharedStrings(s, out richSst); }

        int idx = 0;
        foreach (var se in wdoc.Root!.Element(m + "sheets")?.Elements(m + "sheet") ?? [])
        {
            string name = (string?)se.Attribute("name") ?? "Foglio" + (idx + 1);
            string? rid = (string?)se.Attribute(r + "id");
            var sh = new Sheet(name) { Book = wb, Index = idx++, PartId = rid, OrigName = name };
            wb.Sheets.Add(sh);
            if (rid == null || wbp.GetPartById(rid) is not WorksheetPart wsp) { sh.Loaded = false; continue; } // chartsheet ecc.: intatto
            byte[] sb;
            using (var s = wsp.GetStream(FileMode.Open, FileAccess.Read)) { var ms = new MemoryStream(); s.CopyTo(ms); sb = ms.ToArray(); }
            LoadSheet(sh, sb, sst, richSst);
        }
        foreach (var dn in wdoc.Root.Element(m + "definedNames")?.Elements(m + "definedName") ?? [])
        {
            string n = (string?)dn.Attribute("name") ?? "";
            if (n == "_xlnm.Print_Area" && (int?)dn.Attribute("localSheetId") is { } ps && ps >= 0 && ps < wb.Sheets.Count) wb.Sheets[ps].PrintArea = dn.Value;
            if (n.StartsWith("_xlnm.") || (bool?)dn.Attribute("hidden") == true && n.StartsWith("_")) continue;
            if ((int?)dn.Attribute("localSheetId") is { } ls) { wb.LocalNames[(ls, n.ToUpperInvariant())] = dn.Value; wb.LocalNameCase[(ls, n.ToUpperInvariant())] = n; }
            else wb.Names[n] = dn.Value;
        }
        wb.Styles.MarkClean();
    }

    /// <summary>Testo semplice di ogni stringa condivisa; per quelle con formattazione (runs) anche il contenuto XML originale di &lt;si&gt;.</summary>
    static string[] ReadSharedStrings(Stream s, out Dictionary<int, string> rich)
    {
        var list = new List<string>(); rich = [];
        using var r = XmlReader.Create(s, Settings());
        while (r.Read())
        {
            if (r.NodeType != XmlNodeType.Element || r.LocalName != "si") continue;
            var (plain, xml) = ParseSi(r.ReadOuterXml());
            if (xml != null) rich[list.Count] = xml;
            list.Add(plain);
            // ReadOuterXml avanza oltre </si>: se il nodo corrente è un altro <si> va rielaborato
            while (r.NodeType == XmlNodeType.Element && r.LocalName == "si")
            {
                (plain, xml) = ParseSi(r.ReadOuterXml());
                if (xml != null) rich[list.Count] = xml;
                list.Add(plain);
            }
        }
        return [.. list];
    }

    /// <summary>Estrae il testo (senza fonetica) da un &lt;si&gt;/&lt;is&gt;; restituisce il contenuto XML originale solo se contiene runs formattati.</summary>
    static (string Plain, string? RichXml) ParseSi(string xml)
    {
        var el = XElement.Parse(xml, LoadOptions.PreserveWhitespace);
        var sb = new StringBuilder();
        foreach (var t in el.Descendants().Where(d => d.Name.LocalName == "t" && !d.Ancestors().Any(a => a.Name.LocalName == "rPh"))) sb.Append(t.Value);
        bool hasRuns = el.Elements().Any(e => e.Name.LocalName == "r");
        return (sb.ToString(), hasRuns ? string.Concat(el.Nodes().Select(n => n.ToString(SaveOptions.DisableFormatting))) : null);
    }

    internal static void LoadSheet(Sheet sh, byte[] bytes, string[] sst, Dictionary<int, string> richSst)
    {
        using var ms = new MemoryStream(bytes);
        using var r = XmlReader.Create(ms, Settings());
        var shared = new Dictionary<int, (string Text, int Row, int Col)>();
        while (r.Read())
        {
            if (r.NodeType != XmlNodeType.Element) continue;
            switch (r.LocalName)
            {
                case "sheetFormatPr":
                    if (double.TryParse(r.GetAttribute("defaultRowHeight"), NumberStyles.Float, CultureInfo.InvariantCulture, out var drh)) sh.DefaultRowHeight = drh;
                    if (double.TryParse(r.GetAttribute("defaultColWidth"), NumberStyles.Float, CultureInfo.InvariantCulture, out var dcw)) sh.DefaultColWidth = dcw;
                    break;
                case "sheetView":
                    if (r.GetAttribute("showGridLines") is "0" or "false") sh.ShowGridLines = false;
                    break;
                case "pane":
                    if (r.GetAttribute("state") is "frozen" or "frozenSplit")
                    {
                        sh.FrozenCols = (int)(double.TryParse(r.GetAttribute("xSplit"), NumberStyles.Float, CultureInfo.InvariantCulture, out var xs) ? xs : 0);
                        sh.FrozenRows = (int)(double.TryParse(r.GetAttribute("ySplit"), NumberStyles.Float, CultureInfo.InvariantCulture, out var ys) ? ys : 0);
                    }
                    break;
                case "col": ReadCol(sh, r); break;
                case "sheetData": if (!r.IsEmptyElement) ReadSheetData(sh, r, sst, richSst, shared); break;
                case "mergeCell":
                    if (r.GetAttribute("ref") is { } rf && Addr.TryParseRange(rf, out var r1, out var c1, out var r2, out var c2) && (r1 != r2 || c1 != c2))
                        sh.Merges.Add(new MergeRange(r1, c1, r2, c2));
                    break;
                case "autoFilter":
                    if (r.GetAttribute("ref") is { } af && Addr.TryParseRange(af, out var a1, out var b1, out var a2, out var b2))
                        sh.Filter = new AutoFilterState { R1 = a1, C1 = b1, R2 = a2, C2 = b2 };
                    break;
            }
        }
        sh.DirtyRows.Clear();
        sh.LayoutDirty = false;
    }

    static void ReadCol(Sheet sh, XmlReader r)
    {
        if (!int.TryParse(r.GetAttribute("min"), out var min) || !int.TryParse(r.GetAttribute("max"), out var max)) return;
        double.TryParse(r.GetAttribute("width"), NumberStyles.Float, CultureInfo.InvariantCulture, out var w);
        bool hidden = r.GetAttribute("hidden") is "1" or "true";
        bool custom = r.GetAttribute("customWidth") is "1" or "true";
        uint.TryParse(r.GetAttribute("style"), out var style);
        if (w <= 0) w = sh.DefaultColWidth;
        max = Math.Min(max, Addr.MaxCol);
        for (int c = min; c <= max && c >= 1; c++)
            sh.Cols[c] = new ColInfo { Width = w, Hidden = hidden, CustomWidth = custom, Style = style };
    }

    static void ReadSheetData(Sheet sh, XmlReader r, string[] sst, Dictionary<int, string> richSst, Dictionary<int, (string, int, int)> shared)
    {
        int depth = r.Depth;
        int row = 0, nextCol = 1;
        while (r.Read() && !(r.NodeType == XmlNodeType.EndElement && r.Depth == depth))
        {
            if (r.NodeType != XmlNodeType.Element) continue;
            if (r.LocalName == "row")
            {
                row = int.TryParse(r.GetAttribute("r"), out var rr) ? rr : row + 1;
                nextCol = 1;
                ReadRowAttrs(sh, r, row);
                if (r.IsEmptyElement) continue;
            }
            else if (r.LocalName == "c") ReadCell(sh, r, row, ref nextCol, sst, richSst, shared);
        }
    }

    static void ReadRowAttrs(Sheet sh, XmlReader r, int row)
    {
        RowInfo? ri = null;
        if (r.HasAttributes)
        {
            r.MoveToFirstAttribute();
            do
            {
                string n = r.LocalName;
                if (n is "r" or "spans" or "xmlns") continue;
                if (r.Prefix == "xmlns") continue;
                ri ??= new RowInfo();
                switch (n)
                {
                    case "ht": if (r.NamespaceURI == "" && double.TryParse(r.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var h)) ri.Height = h; break;
                    case "hidden": ri.Hidden = r.Value is "1" or "true"; break;
                    case "customHeight": ri.CustomHeight = r.Value is "1" or "true"; break;
                    case "s": uint.TryParse(r.Value, out ri.Style); break;
                    case "customFormat": ri.CustomFormat = r.Value is "1" or "true"; break;
                    default: ri.Extra.Add((n, r.Prefix, r.NamespaceURI, r.Value)); break;
                }
            } while (r.MoveToNextAttribute());
            r.MoveToElement();
        }
        if (ri != null) sh.Rows[row] = ri;
    }

    static void ReadCell(Sheet sh, XmlReader r, int row, ref int nextCol, string[] sst, Dictionary<int, string> richSst, Dictionary<int, (string Text, int Row, int Col)> shared)
    {
        int col = nextCol;
        string? t = null; uint style = 0; StringBuilder? extra = null;
        string? refAttr = null;
        if (r.HasAttributes)
        {
            r.MoveToFirstAttribute();
            do
            {
                switch (r.LocalName)
                {
                    case "r": refAttr = r.Value; break;
                    case "s": uint.TryParse(r.Value, out style); break;
                    case "t": t = r.Value; break;
                    case "cm" or "vm": (extra ??= new()).Append(' ').Append(r.LocalName).Append("=\"").Append(System.Security.SecurityElement.Escape(r.Value)).Append('"'); break;
                }
            } while (r.MoveToNextAttribute());
            r.MoveToElement();
        }
        if (refAttr != null && Addr.TryParse(refAttr, out var rr, out var cc)) { col = cc; row = rr; }
        nextCol = col + 1;
        string? v = null, fText = null, fRaw = null, fType = null; int fSi = -1; string? inline = null, inlineRich = null; bool hasF = false;
        if (!r.IsEmptyElement)
        {
            int depth = r.Depth;
            bool advance = true;
            while (true)
            {
                if (advance && !r.Read()) break;
                advance = true;
                if (r.NodeType == XmlNodeType.EndElement && r.Depth == depth) break;
                if (r.NodeType != XmlNodeType.Element) continue;
                switch (r.LocalName)
                {
                    case "v": v = r.ReadElementContentAsString(); advance = false; break;
                    case "f":
                        hasF = true;
                        fType = r.GetAttribute("t");
                        if (int.TryParse(r.GetAttribute("si"), out var si)) fSi = si;
                        string? refA = r.GetAttribute("ref");
                        if (fType is "array" or "dataTable") { fRaw = r.ReadOuterXml(); fText = XElement.Parse(fRaw).Value; }
                        else
                        {
                            fText = r.ReadElementContentAsString();
                            if (fType == "shared" && refA != null && fSi >= 0) shared[fSi] = (fText, row, col);
                        }
                        advance = false;
                        break;
                    case "is":
                        {
                            var (pl, rx) = ParseSi(r.ReadOuterXml());
                            inline = pl; inlineRich = rx; advance = false; break;
                        }
                }
            }
        }
        var cell = new Cell { Style = style };
        // valore
        switch (t)
        {
            case "s":
                if (int.TryParse(v, out var si2) && si2 >= 0 && si2 < sst.Length) { cell.Kind = CellKind.Str; cell.Obj = sst[si2]; if (richSst.TryGetValue(si2, out var rx)) sh.Rich[Addr.Key(row, col)] = (sst[si2], rx); }
                break;
            case "str": cell.Kind = CellKind.Str; cell.Obj = v ?? ""; break;
            case "inlineStr": cell.Kind = CellKind.Str; cell.Obj = inline ?? ""; if (inlineRich != null) sh.Rich[Addr.Key(row, col)] = (inline ?? "", inlineRich); break;
            case "b": if (v != null) { cell.Kind = CellKind.Bool; cell.Num = v is "1" or "true" ? 1 : 0; } break;
            case "e": if (v != null && Errors.TryParse(v, out var ce)) { cell.Kind = CellKind.Err; cell.Num = (int)ce; } break;
            case "d": cell.Kind = CellKind.Str; cell.Obj = v ?? ""; break; // ponytail: date ISO trattate come testo
            default:
                if (v != null && double.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out var d)) { cell.Kind = CellKind.Num; cell.Num = d; }
                else if (v != null && v.Length > 0) { cell.Kind = CellKind.Str; cell.Obj = v; }
                break;
        }
        if (hasF)
        {
            var fi = new FormulaInfo { CellAttrs = extra?.ToString() };
            if (fRaw != null) { fi.RawF = fRaw; fi.Text = fText ?? ""; }
            else if (fType == "shared" && string.IsNullOrEmpty(fText) && shared.TryGetValue(fSi, out var m))
                fi.Text = FormulaText.Shift(m.Text, row - m.Row, col - m.Col);
            else fi.Text = fText ?? "";
            if (fi.RawF == null && fi.Text.Length == 0) { sh.Put(row, col, cell); return; }
            cell.F = fi;
        }
        if (!cell.IsEmpty) sh.Put(row, col, cell);
    }
}
