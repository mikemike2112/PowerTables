using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Media;
using PowerTables.Core;

namespace PowerTables.App;

/// <summary>Finestra «Formato celle»: numero, allineamento, carattere, riempimento. Un solo Ctrl+Z annulla tutto.</summary>
public partial class MainWindow
{
    static readonly (string Label, string Code)[] NumberPresets =
    [
        ("Generale", "General"), ("Numero (2 decimali)", "0.00"), ("Migliaia", "#,##0"), ("Valuta", "#,##0.00\\ \"€\""), ("Percentuale", "0.00%"),
        ("Data breve", "dd/mm/yyyy"), ("Data estesa", "dddd d mmmm yyyy"), ("Ora", "hh:mm:ss"), ("Scientifico", "0.00E+00"), ("Testo", "@"),
    ];
    static readonly string[] HAligns = ["Generale", "Sinistra", "Centro", "Destra"], HAlignVals = ["", "left", "center", "right"];
    static readonly string[] VAligns = ["In alto", "Al centro", "In basso"], VAlignVals = ["top", "center", "bottom"];
    static readonly Regex HexRx = new("^[0-9A-Fa-f]{6}$");

    static string? HexOrNull(string t, string what)
    {
        t = t.Trim().TrimStart('#');
        if (t.Length == 0) return null;
        if (!HexRx.IsMatch(t)) throw new ArgumentException($"{what}: usare sei cifre esadecimali (es. FF0000).");
        return "FF" + t.ToUpperInvariant();
    }

    void FormatCellsDialog()
    {
        var xf = ActiveXf; string code = Wb.Styles.FormatCode(Cur.Get(G.ActiveRow, G.ActiveCol).Style);
        var d = Form("Formato celle");
        // numero
        var num = new ComboBox { IsEditable = true, MinWidth = 360, Padding = new Thickness(6, 4, 6, 4), ItemsSource = NumberPresets.Select(p => p.Code).ToArray(), Text = code };
        StyleInput(num); AutomationProperties.SetAutomationId(num, "FcNumber"); AutomationProperties.SetName(num, "Codice formato numero");
        d.Body.Children.Add(FormLabel("Numero: codice formato (es. 0.00, #,##0, 0%, dd/mm/yyyy, @)")); d.Body.Children.Add(num);
        // allineamento
        var hal = FormCombo(d, "Allineamento orizzontale", HAligns, Math.Max(0, Array.IndexOf(HAlignVals, xf.HAlign ?? "")), "FcHAlign");
        var val = FormCombo(d, "Allineamento verticale", VAligns, Math.Max(0, Array.IndexOf(VAlignVals, xf.VAlign ?? "bottom")), "FcVAlign");
        var wrap = FormCheck(d, "Testo a capo", xf.Wrap, "FcWrap");
        var indent = FormText(d, "Rientro (0-15)", xf.Indent.ToString(), "FcIndent", 80);
        // carattere
        var font = FormCombo(d, "Tipo di carattere", ["Calibri", "Arial", "Segoe UI", "Times New Roman", "Courier New", "Verdana", "Georgia", "Consolas"], 0, "FcFont");
        font.IsEditable = true; font.Text = xf.Font.Name;
        var size = FormText(d, "Dimensione", xf.Font.Size.ToString("0.#", Ci), "FcSize", 80);
        var bold = FormCheck(d, "Grassetto", xf.Font.Bold, "FcBold"); var ital = FormCheck(d, "Corsivo", xf.Font.Italic, "FcItalic");
        var und = FormCheck(d, "Sottolineato", xf.Font.Underline, "FcUnder"); var strike = FormCheck(d, "Barrato", xf.Font.Strike, "FcStrike");
        var row = new WrapPanel();
        foreach (var cb in new[] { bold, ital, und, strike }) { d.Body.Children.Remove(cb); cb.Margin = new Thickness(0, 6, 16, 2); row.Children.Add(cb); }
        d.Body.Children.Add(row);
        var fcol = FormText(d, "Colore carattere (RRGGBB, vuoto = automatico)", xf.Font.Color is { Length: 8 } fc ? fc[2..] : "", "FcColor", 120);
        var fill = FormText(d, "Colore di riempimento (RRGGBB, vuoto = nessuno)", xf.FillColor is { Length: 8 } fl ? fl[2..] : "", "FcFill", 120);
        var err = new TextBlock { Foreground = Theme.Brush(Color.FromRgb(0xe5, 0x5b, 0x5b)), FontFamily = new FontFamily("Segoe UI"), FontSize = 12, Margin = new Thickness(0, 8, 0, 0), TextWrapping = TextWrapping.Wrap };
        AutomationProperties.SetAutomationId(err, "FcError");
        d.Body.Children.Add(err);
        // la finestra è alta: scorrimento
        var body = d.Body;
        d.AddButton("Annulla", "Cancel", cancel: true);
        var ok = d.AddButton("OK", "Ok", primary: true, close: false);
        ok.Click += (_, _) =>
        {
            try
            {
                if (!double.TryParse(size.Text, System.Globalization.NumberStyles.Float, Ci, out var sz) || sz < 1 || sz > 409) throw new ArgumentException("Dimensione carattere non valida (1-409).");
                if (!int.TryParse(indent.Text, out var ind) || ind < 0 || ind > 15) throw new ArgumentException("Rientro non valido (0-15).");
                string nf = num.Text.Trim(); if (nf.Length == 0) nf = "General";
                if (nf.Length > 255) throw new ArgumentException("Codice formato troppo lungo.");
                string? fcolor = HexOrNull(fcol.Text, "Colore carattere"), fillc = HexOrNull(fill.Text, "Riempimento");
                string fname = string.IsNullOrWhiteSpace(font.Text) ? xf.Font.Name : font.Text.Trim();
                string? h = HAlignVals[hal.SelectedIndex] is { Length: > 0 } hv ? hv : null;
                string? v = VAlignVals[val.SelectedIndex] == "bottom" ? null : VAlignVals[val.SelectedIndex];
                Fmt(x => x with
                {
                    NumFmtId = nf == "General" ? 0 : 164, NumFmtCode = nf == "General" ? null : nf,
                    HAlign = h, VAlign = v, Wrap = wrap.IsChecked == true, Indent = ind,
                    FillColor = fillc,
                    Font = x.Font with { Name = fname, Size = sz, Bold = bold.IsChecked == true, Italic = ital.IsChecked == true, Underline = und.IsChecked == true, Strike = strike.IsChecked == true, Color = fcolor },
                });
                d.DialogResult = true;
            }
            catch (ArgumentException ex) { err.Text = ex.Message; }
        };
        d.ShowDialog();
    }
}
