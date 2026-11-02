using System.Globalization;
using System.Xml.Linq;

namespace PowerTables.Core;

/// <summary>Margini in pollici (come nel file).</summary>
public readonly record struct Margins(double Left, double Right, double Top, double Bottom, double Header, double Footer)
{
    public static readonly Margins Normal = new(0.7, 0.7, 0.75, 0.75, 0.3, 0.3), Wide = new(1, 1, 1, 1, 0.5, 0.5), Narrow = new(0.25, 0.25, 0.75, 0.75, 0.3, 0.3);
}

public sealed partial class Workbook
{
    /// <summary>Formati carta supportati: (codice OOXML, nome, larghezza mm, altezza mm in verticale).</summary>
    public static readonly (int Code, string Name, double W, double H)[] PaperSizes =
    [
        (9, "A4", 210, 297), (8, "A3", 297, 420), (11, "A5", 148, 210), (1, "Lettera", 215.9, 279.4), (5, "Legale", 215.9, 355.6),
    ];

    public Margins GetMargins(Sheet s)
    {
        var m = s.PageMargins;
        double A(string n, double d) => m != null && double.TryParse((string?)m.Attribute(n), NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v : d;
        return new(A("left", 0.7), A("right", 0.7), A("top", 0.75), A("bottom", 0.75), A("header", 0.3), A("footer", 0.3));
    }

    public bool IsLandscape(Sheet s) => (string?)s.PageSetup?.Attribute("orientation") == "landscape";

    public int PaperSize(Sheet s) => (int?)s.PageSetup?.Attribute("paperSize") ?? 9;

    public bool PrintGridLines(Sheet s) => (bool?)s.PrintOptions?.Attribute("gridLines") ?? false;

    public void SetMargins(Sheet s, Margins m)
    {
        if (m.Left < 0 || m.Right < 0 || m.Top < 0 || m.Bottom < 0 || m.Header < 0 || m.Footer < 0 || m.Left + m.Right >= 20 || m.Top + m.Bottom >= 20) throw new ArgumentException("Margini non validi.");
        string F(double d) => d.ToString("R", CultureInfo.InvariantCulture);
        ChangeSections(s, "Margini", () => s.PageMargins = new XElement(MainX + "pageMargins", new XAttribute("left", F(m.Left)), new XAttribute("right", F(m.Right)), new XAttribute("top", F(m.Top)), new XAttribute("bottom", F(m.Bottom)), new XAttribute("header", F(m.Header)), new XAttribute("footer", F(m.Footer))));
    }

    public void SetOrientation(Sheet s, bool landscape) =>
        ChangeSections(s, "Orientamento", () => (s.PageSetup ??= new XElement(MainX + "pageSetup")).SetAttributeValue("orientation", landscape ? "landscape" : "portrait"));

    public void SetPaperSize(Sheet s, int code)
    {
        if (!PaperSizes.Any(p => p.Code == code)) throw new ArgumentException("Formato carta non supportato.");
        ChangeSections(s, "Dimensioni", () => (s.PageSetup ??= new XElement(MainX + "pageSetup")).SetAttributeValue("paperSize", code));
    }

    public void SetPrintGridLines(Sheet s, bool on) =>
        ChangeSections(s, "Griglia: stampa", () =>
        {
            s.PrintOptions ??= new XElement(MainX + "printOptions");
            s.PrintOptions.SetAttributeValue("gridLines", on ? 1 : null);
            if (!s.PrintOptions.HasAttributes) s.PrintOptions = null;
        });
}
