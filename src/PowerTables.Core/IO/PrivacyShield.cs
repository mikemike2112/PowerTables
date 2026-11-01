using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace PowerTables.Core.IO;

/// <summary>
/// Scudo privacy: analizza un .xlsx e crea una COPIA senza i dati personali scelti. Non modifica mai il file originale. Tutto in locale.
/// </summary>
public static class PrivacyShield
{
    public enum Kind { Author, LastAuthor, Company, Manager, CustomProperty, CommentAuthor, PersonName, RevisionUser, LocalPath, Thumbnail }

    public sealed class Finding
    {
        public Kind Kind; public string Label = ""; public string Value = ""; public string Part = "";
        public bool Selected = true;
        public override string ToString() => $"{Label}: {Value}";
    }

    public sealed class Report
    {
        public List<Finding> Findings = [];
        public List<string> Warnings = [];
        public bool HasComments;
    }

    public enum CommentsMode { Anonymize, Remove }

    public const string Anonymous = "Anonimo";

    static readonly Regex UserPath = new(@"(?i)(file:///)?[A-Z]:[\\/]+Users[\\/]+([^\\/""<>|\r\n]+)", RegexOptions.Compiled);
    static readonly Regex AnyLocalPath = new(@"(?i)(file:///)?[A-Z]:[\\/][^""<>|\r\n]*", RegexOptions.Compiled);

    static string? Read(ZipArchive z, string name)
    {
        var e = z.GetEntry(name); if (e == null) return null;
        using var r = new StreamReader(e.Open()); return r.ReadToEnd();
    }

    static void Write(ZipArchive z, string name, string text)
    {
        z.GetEntry(name)?.Delete();
        var e = z.CreateEntry(name, CompressionLevel.Optimal);
        using var w = new StreamWriter(e.Open(), new UTF8Encoding(false)); w.Write(text);
    }

    static string Unesc(string s) => System.Net.WebUtility.HtmlDecode(s);

    /// <summary>Analizza un file (in sola lettura).</summary>
    public static Report Analyze(string path)
    {
        var rep = new Report();
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var z = new ZipArchive(fs, ZipArchiveMode.Read);
        string? core = Read(z, "docProps/core.xml");
        if (core != null)
        {
            Add(rep, Kind.Author, "Autore", Tag(core, "dc:creator"), "docProps/core.xml");
            Add(rep, Kind.LastAuthor, "Ultimo autore", Tag(core, "cp:lastModifiedBy"), "docProps/core.xml");
        }
        string? app = Read(z, "docProps/app.xml");
        if (app != null)
        {
            Add(rep, Kind.Company, "Società", Tag(app, "Company"), "docProps/app.xml");
            Add(rep, Kind.Manager, "Responsabile", Tag(app, "Manager"), "docProps/app.xml");
        }
        string? custom = Read(z, "docProps/custom.xml");
        if (custom != null)
            foreach (var p in XDocument.Parse(custom).Descendants().Where(d => d.Name.LocalName == "property"))
                Add(rep, Kind.CustomProperty, "Proprietà personalizzata «" + (string?)p.Attribute("name") + "»", p.Value, "docProps/custom.xml");
        foreach (var e in z.Entries)
        {
            string n = e.FullName;
            if (Regex.IsMatch(n, @"^xl/comments\d*\.xml$"))
            {
                rep.HasComments = true;
                foreach (Match m in Regex.Matches(Read(z, n)!, "<author>(.*?)</author>")) Add(rep, Kind.CommentAuthor, "Autore di commenti", Unesc(m.Groups[1].Value), n);
            }
            else if (Regex.IsMatch(n, @"^xl/persons/.*\.xml$"))
                foreach (Match m in Regex.Matches(Read(z, n)!, "displayName=\"([^\"]*)\"")) Add(rep, Kind.PersonName, "Nome in persone/thread", Unesc(m.Groups[1].Value), n);
            else if (Regex.IsMatch(n, @"^xl/threadedComments/.*\.xml$")) rep.HasComments = true;
            else if (Regex.IsMatch(n, @"^xl/revisions/.*\.xml$"))
                foreach (var u in Regex.Matches(Read(z, n)!, "userName=\"([^\"]*)\"").Select(m => Unesc(m.Groups[1].Value)).Distinct()) Add(rep, Kind.RevisionUser, "Utente delle revisioni", u, n);
            else if (n == "docProps/thumbnail.jpeg" || n == "docProps/thumbnail.png" || n == "docProps/thumbnail.wmf") rep.Findings.Add(new Finding { Kind = Kind.Thumbnail, Label = "Anteprima miniatura", Value = n + " (" + e.Length + " byte)", Part = n });
        }
        // percorsi locali e nomi utente nei metadati e nella cartella di lavoro
        foreach (var n in new[] { "xl/workbook.xml", "docProps/app.xml", "docProps/core.xml" })
        {
            string? t = Read(z, n); if (t == null) continue;
            foreach (var u in UserPath.Matches(t).Where(m => m.Groups[2].Value != Anonymous).Select(m => m.Value).Distinct()) rep.Findings.Add(new Finding { Kind = Kind.LocalPath, Label = "Percorso locale con nome utente", Value = Unesc(u), Part = n });
        }
        // avvisi: ciò che non si sa pulire
        foreach (var e in z.Entries)
        {
            string n = e.FullName;
            if (Regex.IsMatch(n, @"(?i)^xl/media/.*\.(jpe?g|tiff?)$") && ContainsExif(e)) rep.Warnings.Add($"L'immagine {n} contiene metadati EXIF (fotocamera, data, posizione): non vengono rimossi.");
            else if (n.StartsWith("xl/embeddings/", StringComparison.OrdinalIgnoreCase)) rep.Warnings.Add($"Oggetto incorporato {n}: può contenere dati personali, non analizzato.");
            else if (n.Equals("xl/vbaProject.bin", StringComparison.OrdinalIgnoreCase)) rep.Warnings.Add("Il file contiene macro (vbaProject.bin): non analizzate né rimosse.");
            else if (Regex.IsMatch(n, @"^xl/externalLinks/_rels/.*\.rels$") && AnyLocalPath.IsMatch(Read(z, n)!)) rep.Warnings.Add($"Collegamenti esterni a percorsi locali ({n}): non modificati per non romperli.");
            else if (Regex.IsMatch(n, @"^xl/worksheets/sheet\d+\.xml$"))
            {
                string x = Read(z, n)!;
                if (Regex.IsMatch(x, "<row [^>]*hidden=\"1\"") || Regex.IsMatch(x, "<col [^>]*hidden=\"1\"")) rep.Warnings.Add($"{n}: righe o colonne nascoste (il contenuto resta nel file).");
            }
        }
        string? wb = Read(z, "xl/workbook.xml");
        if (wb != null && Regex.IsMatch(wb, "<sheet [^>]*state=\"(hidden|veryHidden)\"")) rep.Warnings.Add("Il file contiene fogli nascosti: il contenuto non viene analizzato né rimosso.");
        if (rep.Findings.Count > 0 || rep.HasComments) rep.Warnings.Add("Il testo delle celle, i nomi dei fogli e il contenuto dei commenti non sono analizzati: se contengono dati personali restano.");
        return rep;
    }

    static bool ContainsExif(ZipArchiveEntry e)
    {
        using var s = e.Open(); var buf = new byte[(int)Math.Min(e.Length, 65536)]; int n = 0;
        while (n < buf.Length) { int r = s.Read(buf, n, buf.Length - n); if (r <= 0) break; n += r; }
        return Encoding.Latin1.GetString(buf, 0, n).Contains("Exif");
    }

    static void Add(Report rep, Kind k, string label, string? value, string part)
    {
        if (string.IsNullOrWhiteSpace(value) || value == Anonymous) return;
        if (rep.Findings.Any(f => f.Kind == k && f.Value == value && f.Part == part)) return;
        rep.Findings.Add(new Finding { Kind = k, Label = label, Value = value, Part = part });
    }

    static string? Tag(string xml, string tag)
    {
        var m = Regex.Match(xml, $"<{Regex.Escape(tag)}(?: [^>]*)?>(.*?)</{Regex.Escape(tag)}>", RegexOptions.Singleline);
        return m.Success ? Unesc(m.Groups[1].Value) : null;
    }

    /// <summary>Crea una copia pulita. L'originale non viene mai aperto in scrittura. Restituisce la nuova analisi della copia.</summary>
    public static Report CreateCleanCopy(string source, string dest, IEnumerable<Finding> selected, CommentsMode comments)
    {
        if (Path.GetFullPath(source).Equals(Path.GetFullPath(dest), StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("La copia non può avere lo stesso percorso dell'originale.");
        string tmp = dest + ".tmp~";
        File.Copy(source, tmp, true);
        try
        {
            var sel = selected.ToList();
            var values = sel.Select(f => f.Value).ToHashSet();
            bool removeThumb = sel.Any(f => f.Kind == Kind.Thumbnail);
            using (var z = ZipFile.Open(tmp, ZipArchiveMode.Update))
            {
                foreach (var e in z.Entries.ToList())
                {
                    string n = e.FullName;
                    bool isXml = n.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) || n.EndsWith(".rels", StringComparison.OrdinalIgnoreCase);
                    if (!isXml) continue;
                    string x = Read(z, n)!, y = x;
                    if (n == "docProps/core.xml")
                    {
                        if (sel.Any(f => f.Kind == Kind.Author)) y = Blank(y, "dc:creator");
                        if (sel.Any(f => f.Kind == Kind.LastAuthor)) y = Blank(y, "cp:lastModifiedBy");
                    }
                    else if (n == "docProps/app.xml")
                    {
                        if (sel.Any(f => f.Kind == Kind.Company)) y = Blank(y, "Company");
                        if (sel.Any(f => f.Kind == Kind.Manager)) y = Blank(y, "Manager");
                    }
                    else if (n == "docProps/custom.xml" && sel.Any(f => f.Kind == Kind.CustomProperty)) y = CleanCustom(y, sel.Where(f => f.Kind == Kind.CustomProperty).Select(f => f.Value).ToHashSet());
                    else if (Regex.IsMatch(n, @"^xl/comments\d*\.xml$") && comments == CommentsMode.Anonymize && sel.Any(f => f.Kind == Kind.CommentAuthor))
                        y = Regex.Replace(y, "<author>(.*?)</author>", m => values.Contains(Unesc(m.Groups[1].Value)) ? "<author>" + Anonymous + "</author>" : m.Value);
                    else if (Regex.IsMatch(n, @"^xl/persons/.*\.xml$") && sel.Any(f => f.Kind == Kind.PersonName))
                        y = Regex.Replace(y, "(displayName|userId)=\"([^\"]*)\"", m => values.Contains(Unesc(m.Groups[2].Value)) || m.Groups[1].Value == "userId" ? m.Groups[1].Value + "=\"" + Anonymous + "\"" : m.Value);
                    else if (Regex.IsMatch(n, @"^xl/revisions/.*\.xml$") && sel.Any(f => f.Kind == Kind.RevisionUser))
                        y = Regex.Replace(y, "userName=\"([^\"]*)\"", m => values.Contains(Unesc(m.Groups[1].Value)) ? "userName=\"" + Anonymous + "\"" : m.Value);
                    if (sel.Any(f => f.Kind == Kind.LocalPath) && (n == "xl/workbook.xml" || n == "docProps/app.xml" || n == "docProps/core.xml"))
                        y = UserPath.Replace(y, m => m.Groups[1].Value + "C:\\Users\\" + Anonymous);
                    if (y != x) Write(z, n, y);
                }
                if (removeThumb)
                {
                    foreach (var n in new[] { "docProps/thumbnail.jpeg", "docProps/thumbnail.png", "docProps/thumbnail.wmf" }) z.GetEntry(n)?.Delete();
                    string? rels = Read(z, "_rels/.rels");
                    if (rels != null) Write(z, "_rels/.rels", Regex.Replace(rels, "<Relationship [^>]*thumbnail[^>]*/>", ""));
                }
            }
            if (comments == CommentsMode.Remove && sel.Any(f => f.Kind is Kind.CommentAuthor or Kind.PersonName)) RemoveComments(tmp);
            File.Move(tmp, dest, true);
        }
        catch { try { File.Delete(tmp); } catch (IOException) { } throw; }
        return Analyze(dest);
    }

    static string Blank(string xml, string tag) => Regex.Replace(xml, $"(<{Regex.Escape(tag)}(?: [^>]*)?>)(.*?)(</{Regex.Escape(tag)}>)", "$1$3", RegexOptions.Singleline);

    static string CleanCustom(string xml, HashSet<string> values)
    {
        var d = XDocument.Parse(xml);
        foreach (var p in d.Descendants().Where(e => e.Name.LocalName == "property" && values.Contains(e.Value)).ToList()) p.Remove();
        return d.Declaration + "\r\n" + d.Root!.ToString(SaveOptions.DisableFormatting);
    }

    /// <summary>Rimuove commenti e thread con l'SDK (relazioni, tipi di contenuto e riferimenti nel foglio restano coerenti).</summary>
    static void RemoveComments(string path)
    {
        using var doc = SpreadsheetDocument.Open(path, true);
        foreach (var wsp in doc.WorkbookPart!.WorksheetParts)
        {
            var ws = wsp.Worksheet ?? throw new InvalidDataException("foglio illeggibile");
            if (wsp.WorksheetCommentsPart != null)
            {
                var legacyIds = ws.Elements<DocumentFormat.OpenXml.Spreadsheet.LegacyDrawing>().Select(l => l.Id?.Value).ToList();
                wsp.DeletePart(wsp.WorksheetCommentsPart);
                foreach (var lid in legacyIds) if (lid != null && wsp.TryGetPartById(lid, out var vp) && vp is VmlDrawingPart) wsp.DeletePart(vp);
                foreach (var l in ws.Elements<DocumentFormat.OpenXml.Spreadsheet.LegacyDrawing>().ToList()) l.Remove();
                ws.Save();
            }
            foreach (var t in wsp.GetPartsOfType<WorksheetThreadedCommentsPart>().ToList()) wsp.DeletePart(t);
        }
        foreach (var p in doc.WorkbookPart.GetPartsOfType<WorkbookPersonPart>().ToList()) doc.WorkbookPart.DeletePart(p);
    }
}
