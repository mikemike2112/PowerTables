using System.Xml.Linq;

namespace PowerTables.Core;

public sealed partial class Workbook
{
    static readonly string[] LinkSchemes = ["http", "https", "mailto", "ftp"];

    /// <summary>Verifica l'indirizzo di un collegamento esterno (solo http, https, ftp, mailto). Restituisce l'URI normalizzato o lancia ArgumentException.</summary>
    public static string ValidateUrl(string url)
    {
        url = url.Trim();
        if (url.Length == 0) throw new ArgumentException("Indicare l'indirizzo.");
        if (!url.Contains(':') && !url.Contains('@')) url = "https://" + url;
        else if (!url.Contains(':')) url = "mailto:" + url;
        if (!Uri.TryCreate(url, UriKind.Absolute, out var u) || !LinkSchemes.Contains(u.Scheme)) throw new ArgumentException("Indirizzo non valido: sono ammessi http, https, ftp e mailto.");
        return u.OriginalString;
    }

    public (string? Url, string? Location, string? Tooltip)? GetHyperlink(Sheet s, int r, int c)
    {
        var h = s.Links.FirstOrDefault(l => ParseSqref((string?)l.Attribute("ref") ?? "").Any(x => r >= x.R1 && r <= x.R2 && c >= x.C1 && c <= x.C2));
        if (h == null) return null;
        string? url = (string?)h.Attribute(IO.XlsxWriter.PtUrl);
        if (url == null && (string?)h.Attribute(XName.Get("id", IO.XlsxReader.RelNs)) is { } rid) url = LinkTarget(s, rid);
        return (url, (string?)h.Attribute("location"), (string?)h.Attribute("tooltip"));
    }

    string? LinkTarget(Sheet s, string rid)
    {
        try
        {
            if (Pkg?.Doc.WorkbookPart is { } wbp && s.PartId != null && wbp.GetPartById(s.PartId) is DocumentFormat.OpenXml.Packaging.WorksheetPart wsp)
                return wsp.HyperlinkRelationships.FirstOrDefault(x => x.Id == rid)?.Uri.OriginalString;
        }
        catch (Exception ex) when (ex is KeyNotFoundException or ArgumentException or InvalidOperationException) { }
        return null;
    }

    /// <summary>Inserisce un collegamento (esterno con <paramref name="url"/> oppure interno con <paramref name="location"/>, es. Foglio2!A1).
    /// Se <paramref name="text"/> non è null (o la cella è vuota) imposta il testo della cella; applica lo stile blu sottolineato.</summary>
    public void SetHyperlink(Sheet s, int r, int c, string? url, string? location, string? text, string? tooltip = null)
    {
        string? u = string.IsNullOrWhiteSpace(url) ? null : ValidateUrl(url);
        string? loc = string.IsNullOrWhiteSpace(location) ? null : location.Trim();
        if (u == null && loc == null) throw new ArgumentException("Indicare un indirizzo o una posizione nel documento.");
        if (loc != null)
        {
            string target = loc.Contains('!') ? loc[(loc.LastIndexOf('!') + 1)..] : loc;
            string sh = loc.Contains('!') ? loc[..loc.LastIndexOf('!')].Trim('\'') : s.Name;
            if (FindSheet(sh.Replace("''", "'")) == null || !(Addr.TryParseRange(target.Replace("$", ""), out _, out _, out _, out _) || Addr.TryParse(target.Replace("$", ""), out _, out _)))
                throw new ArgumentException("La posizione nel documento non è valida.");
        }
        ChangeSections(s, "Collegamento", () =>
        {
            RemoveLinkAt(s, r, c);
            var e = new XElement(MainX + "hyperlink", new XAttribute("ref", Addr.A1(r, c)));
            if (u != null) e.SetAttributeValue(IO.XlsxWriter.PtUrl, u);
            if (loc != null) e.SetAttributeValue("location", loc);
            if (!string.IsNullOrEmpty(tooltip)) e.SetAttributeValue("tooltip", tooltip);
            if (!string.IsNullOrEmpty(text)) e.SetAttributeValue("display", text);
            s.Links.Add(e);
            var cell = s.Get(r, c);
            if (!string.IsNullOrEmpty(text)) { cell.SetValue(text); cell.F = null; }
            else if (cell.Kind == CellKind.Empty) cell.SetValue(u ?? loc);
            cell.Style = Styles.Derive(cell.Style, x => x with { Font = x.Font with { Color = "FF0563C1", Underline = true } });
            Write(s, r, c, cell);
        });
    }

    void RemoveLinkAt(Sheet s, int r, int c)
    {
        foreach (var l in s.Links.ToList())
        {
            var left = ParseSqref((string?)l.Attribute("ref") ?? "").SelectMany(x => Subtract(x, (r, c, r, c))).ToList();
            if (left.Count == 0) s.Links.Remove(l); else l.SetAttributeValue("ref", FormatSqref(left));   // ponytail: un collegamento su più celle resta sulle altre
        }
    }

    public void RemoveHyperlink(Sheet s, int r, int c)
    {
        if (GetHyperlink(s, r, c) == null) return;
        ChangeSections(s, "Rimuovi collegamento", () =>
        {
            RemoveLinkAt(s, r, c);
            var cell = s.Get(r, c);
            cell.Style = Styles.Derive(cell.Style, x => x with { Font = x.Font with { Color = null, Underline = false } });
            Write(s, r, c, cell);
        });
    }
}
