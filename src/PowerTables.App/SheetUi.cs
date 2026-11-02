using System.Diagnostics;
using System.IO;
using System.Windows.Input;
using System.Globalization;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using PowerTables.Core;
using PowerTables.Core.Formulas;
using PowerTables.Core.IO;

namespace PowerTables.App;

/// <summary>Convalida dati, formattazione condizionale, collegamenti, imposta pagina.</summary>
public partial class MainWindow
{
    void InitSheetUi()
    {
        G.ValidationListRequested += (r, c, pt) => ShowValidationList(r, c, pt);
        G.LinkClicked += OpenLink;
    }

    /// <summary>Testo digitato → formula in forma di file: «=...» è una formula, un numero resta numero, il resto diventa testo tra virgolette.</summary>
    string ToFormulaValue(string input)
    {
        input = input.Trim();
        if (input.StartsWith('=')) return FormulaText.ToFile(FormulaText.FromUser(input[1..]));
        if (double.TryParse(input, NumberStyles.Float, Ci, out var d) || double.TryParse(input, NumberStyles.Float, CultureInfo.InvariantCulture, out d)) return d.ToString("R", CultureInfo.InvariantCulture);
        return "\"" + input.Replace("\"", "\"\"") + "\"";
    }

    // ---------- Convalida dati ----------
    static readonly (string Label, string Op)[] DvOps =
    [
        ("tra", "between"), ("non tra", "notBetween"), ("uguale a", "equal"), ("diverso da", "notEqual"),
        ("maggiore di", "greaterThan"), ("minore di", "lessThan"), ("maggiore o uguale a", "greaterThanOrEqual"), ("minore o uguale a", "lessThanOrEqual"),
    ];
    static readonly (string Label, string Type)[] DvTypes = [("Qualsiasi valore", "none"), ("Numero intero", "whole"), ("Decimale", "decimal"), ("Elenco", "list"), ("Lunghezza testo", "textLength")];

    void DataValidationDialog()
    {
        var (r1, c1, r2, c2) = G.Selection;
        r2 = Math.Min(r2, Math.Max(r1, Cur.MaxRow + 1000)); c2 = Math.Min(c2, Math.Max(c1, Cur.MaxCol + 50));   // ponytail: intere righe/colonne limitate all'area usata + margine
        var cur = Wb.GetValidation(Cur, r1, c1);
        var d = Form("Convalida dati", "Limita i valori che si possono immettere nelle celle selezionate.");
        var type = FormCombo(d, "Consenti", DvTypes.Select(t => t.Label).ToArray(), Math.Max(0, Array.FindIndex(DvTypes, t => t.Type == cur?.Type)), "DvType");
        var op = FormCombo(d, "Dati", DvOps.Select(t => t.Label).ToArray(), Math.Max(0, Array.FindIndex(DvOps, t => t.Op == cur?.Op)), "DvOp");
        string Show(string? f) => f == null ? "" : cur is { Type: "list" } && f.StartsWith('"') && f.EndsWith('"') ? f[1..^1].Replace(',', ';') : FormulaText.ToDisplay(f) is var t && cur is { Type: "list" } ? "=" + t : t;
        var v1 = FormText(d, "Minimo / valore / origine (elenco: voci separate da ;  oppure =$A$1:$A$5)", Show(cur?.F1), "DvF1");
        var v2 = FormText(d, "Massimo", Show(cur?.F2), "DvF2");
        var blank = FormCheck(d, "Ignora celle vuote", cur?.AllowBlank ?? true, "DvBlank");
        var arrow = FormCheck(d, "Mostra freccia nell'elenco", cur?.Dropdown ?? true, "DvArrow");
        var msg = FormText(d, "Messaggio di errore (facoltativo)", cur?.Err ?? "", "DvMsg");
        var err = new TextBlock { Foreground = Theme.Brush(Color.FromRgb(0xe5, 0x5b, 0x5b)), FontFamily = new FontFamily("Segoe UI"), FontSize = 12, Margin = new Thickness(0, 8, 0, 0), TextWrapping = TextWrapping.Wrap };
        AutomationProperties.SetAutomationId(err, "DvError");
        d.Body.Children.Add(err);
        void Sync()
        {
            string t = DvTypes[type.SelectedIndex].Type;
            op.IsEnabled = t is "whole" or "decimal" or "textLength";
            bool two = op.IsEnabled && DvOps[op.SelectedIndex].Op is "between" or "notBetween";
            v2.IsEnabled = two; arrow.IsEnabled = t == "list"; v1.IsEnabled = t != "none";
        }
        type.SelectionChanged += (_, _) => Sync(); op.SelectionChanged += (_, _) => Sync(); Sync();
        d.AddButton("Cancella tutto", "Clear", close: false).Click += (_, _) => { Wb.SetValidation(Cur, r1, c1, r2, c2, null); d.DialogResult = true; G.InvalidateVisual(); };
        d.AddButton("Annulla", "Cancel", cancel: true);
        var ok = d.AddButton("OK", "Ok", primary: true, close: false);
        ok.Click += (_, _) =>
        {
            try
            {
                string t = DvTypes[type.SelectedIndex].Type;
                if (t == "none") Wb.SetValidation(Cur, r1, c1, r2, c2, null);
                else
                {
                    string f1 = t == "list"
                        ? (v1.Text.TrimStart().StartsWith('=') ? ToFormulaValue(v1.Text) : "\"" + v1.Text.Replace("\"", "\"\"").Replace(';', ',') + "\"")
                        : ToFormulaValue(v1.Text);
                    string? f2 = v2.IsEnabled ? ToFormulaValue(v2.Text) : null;
                    if (t == "list" && f1.Length > 257) throw new ArgumentException("L'elenco è troppo lungo (massimo 255 caratteri).");
                    Wb.SetValidation(Cur, r1, c1, r2, c2, new DvSpec(t, t == "list" ? "between" : DvOps[op.SelectedIndex].Op, f1, f2, blank.IsChecked == true, true, null, msg.Text, arrow.IsChecked == true));
                }
                d.DialogResult = true; G.InvalidateVisual();
            }
            catch (Exception ex) when (ex is ArgumentException or FormatException) { err.Text = ex.Message; }
        };
        d.ShowDialog();
    }

    /// <summary>Menu a discesa con le voci dell'elenco di convalida: la scelta scrive il valore nella cella.</summary>
    void ShowValidationList(int r, int c, Point screen)
    {
        var items = Wb.ValidationList(Cur, r, c, Ci);
        if (items == null || items.Count == 0) return;
        var pop = new Popup { Placement = PlacementMode.AbsolutePoint, HorizontalOffset = screen.X, VerticalOffset = screen.Y, StaysOpen = false, AllowsTransparency = true };
        var sp = new StackPanel { Margin = new Thickness(2) };
        foreach (var it in items.Take(200))
        {
            var text = it;
            var b = new Button { Content = text, HorizontalContentAlignment = HorizontalAlignment.Left, MinWidth = 140 };
            AutomationProperties.SetName(b, text);
            b.Click += (_, _) => { pop.IsOpen = false; Wb.Do("Modifica cella", () => Wb.SetInput(Cur, r, c, text, Ci)); G.Focus(); };
            sp.Children.Add(b);
        }
        pop.Child = new Border { Background = R("PanelBrush"), BorderBrush = R("AccentBrush"), BorderThickness = new Thickness(1), Child = new ScrollViewer { Content = sp, MaxHeight = 260, VerticalScrollBarVisibility = ScrollBarVisibility.Auto } };
        pop.IsOpen = true;
    }

    // ---------- Formattazione condizionale ----------
    static readonly (string Label, string? Font, string? Fill, bool Bold)[] CfFormats =
    [
        ("Riempimento rosso chiaro con testo rosso scuro", "FF9C0006", "FFFFC7CE", false),
        ("Riempimento giallo con testo giallo scuro", "FF9C5700", "FFFFEB9C", false),
        ("Riempimento verde con testo verde scuro", "FF006100", "FFC6EFCE", false),
        ("Testo rosso", "FFC00000", null, false),
        ("Testo in grassetto", null, null, true),
    ];

    void CondFormatMenu(Button anchor) => DropMenu(anchor,
        ("Maggiore di...", () => CondFormatDialog("greaterThan", "Maggiore di")),
        ("Minore di...", () => CondFormatDialog("lessThan", "Minore di")),
        ("Tra...", () => CondFormatDialog("between", "Tra")),
        ("Uguale a...", () => CondFormatDialog("equal", "Uguale a")),
        ("Il testo contiene...", () => CondFormatDialog("containsText", "Il testo contiene")),
        ("Cancella regole dalle celle selezionate", () => { var (r1, c1, r2, c2) = G.Selection; Wb.ClearCondFormat(Cur, r1, c1, Math.Min(r2, Math.Max(r1, Cur.MaxRow)), Math.Min(c2, Math.Max(c1, Cur.MaxCol))); G.InvalidateVisual(); }),
        ("Cancella regole da tutto il foglio", () => { Wb.ClearCondFormat(Cur, 1, 1, Addr.MaxRow, Addr.MaxCol); G.InvalidateVisual(); }));

    void CondFormatDialog(string op, string title)
    {
        var (r1, c1, r2, c2) = G.Selection;
        var d = Form(title, "Formatta le celle selezionate che soddisfano la condizione.");
        bool text = op == "containsText", two = op == "between";
        var v1 = FormText(d, text ? "Testo" : two ? "Da" : "Valore (o =riferimento)", "", "CfV1");
        var v2 = two ? FormText(d, "A", "", "CfV2") : null;
        var fmt = FormCombo(d, "Formato", CfFormats.Select(f => f.Label).ToArray(), 0, "CfFormat");
        var err = new TextBlock { Foreground = Theme.Brush(Color.FromRgb(0xe5, 0x5b, 0x5b)), FontFamily = new FontFamily("Segoe UI"), FontSize = 12, Margin = new Thickness(0, 8, 0, 0), TextWrapping = TextWrapping.Wrap };
        AutomationProperties.SetAutomationId(err, "CfError");
        d.Body.Children.Add(err);
        d.FocusFirst(v1);
        d.AddButton("Annulla", "Cancel", cancel: true);
        var ok = d.AddButton("OK", "Ok", primary: true, close: false);
        ok.Click += (_, _) =>
        {
            try
            {
                var f = CfFormats[fmt.SelectedIndex];
                var spec = text ? new CfSpec("containsText", "", "", null, v1.Text, f.Font, f.Fill, f.Bold)
                                : new CfSpec("cellIs", op, v1.Text.Trim().Length == 0 ? "" : ToFormulaValue(v1.Text), v2 == null ? null : ToFormulaValue(v2.Text), null, f.Font, f.Fill, f.Bold);
                Wb.AddCondFormat(Cur, r1, c1, Math.Min(r2, Math.Max(r1, Cur.MaxRow + 1000)), Math.Min(c2, Math.Max(c1, Cur.MaxCol + 50)), spec);
                d.DialogResult = true; G.InvalidateVisual();
            }
            catch (Exception ex) when (ex is ArgumentException or FormatException) { err.Text = ex.Message; }
        };
        d.ShowDialog();
    }

    // ---------- Collegamento ----------
    void LinkDialog()
    {
        int r = G.ActiveRow, c = G.ActiveCol;
        var cur = Wb.GetHyperlink(Cur, r, c);
        var d = Form("Inserisci collegamento");
        var kind = FormCombo(d, "Collega a", ["Pagina web o indirizzo e-mail", "Posizione in questo documento"], cur?.Location != null && cur.Value.Url == null ? 1 : 0, "LinkKind");
        var text = FormText(d, "Testo da visualizzare", Cur.Get(r, c).Kind == CellKind.Empty ? "" : Wb.DisplayText(Cur, r, c, Ci), "LinkText");
        var addr = FormText(d, "Indirizzo (es. https://esempio.it, mailto:a@b.it) oppure riferimento (es. Foglio2!A1)", cur?.Url ?? cur?.Location ?? "", "LinkAddr");
        var tip = FormText(d, "Descrizione (facoltativa)", cur?.Tooltip ?? "", "LinkTip");
        var err = new TextBlock { Foreground = Theme.Brush(Color.FromRgb(0xe5, 0x5b, 0x5b)), FontFamily = new FontFamily("Segoe UI"), FontSize = 12, Margin = new Thickness(0, 8, 0, 0), TextWrapping = TextWrapping.Wrap };
        AutomationProperties.SetAutomationId(err, "LinkError");
        d.Body.Children.Add(err);
        d.FocusFirst(addr);
        if (cur != null) d.AddButton("Rimuovi collegamento", "Remove", close: false).Click += (_, _) => { Wb.RemoveHyperlink(Cur, r, c); d.DialogResult = true; };
        d.AddButton("Annulla", "Cancel", cancel: true);
        var ok = d.AddButton("OK", "Ok", primary: true, close: false);
        ok.Click += (_, _) =>
        {
            try
            {
                bool internalLink = kind.SelectedIndex == 1;
                Wb.SetHyperlink(Cur, r, c, internalLink ? null : addr.Text, internalLink ? addr.Text : null, text.Text.Length == 0 ? null : text.Text, tip.Text);
                d.DialogResult = true;
            }
            catch (ArgumentException ex) { err.Text = ex.Message; }
        };
        d.ShowDialog();
    }

    /// <summary>Ctrl+clic: salta alla posizione interna oppure apre l'indirizzo esterno (solo http, https, ftp, mailto: già validati in Core).</summary>
    void OpenLink(int r, int c)
    {
        if (Wb.GetHyperlink(Cur, r, c) is not { } h) return;
        if (h.Url != null)
        {
            try { Workbook.ValidateUrl(h.Url); Process.Start(new ProcessStartInfo(h.Url) { UseShellExecute = true }); }
            catch (Exception ex) when (ex is ArgumentException or System.ComponentModel.Win32Exception or InvalidOperationException) { ThemedDialog.Notify(this, "Impossibile aprire il collegamento: " + ex.Message, DialogIcon.Warning); }
        }
        else if (h.Location != null)
        {
            string loc = h.Location; int bang = loc.LastIndexOf('!');
            var sh = bang > 0 ? Wb.FindSheet(loc[..bang].Trim('\'').Replace("''", "'")) : Cur;
            string target = (bang > 0 ? loc[(bang + 1)..] : loc).Replace("$", "");
            if (sh == null) return;
            if (!ReferenceEquals(sh, Cur)) { G.SwitchSheet(sh); RebuildSheetTabs(); }
            if (Addr.TryParseRange(target, out var a, out var b, out var c2, out var d2)) { G.SelectRange(a, b, c2, d2); G.EnsureVisible(a, b); }
            else if (Addr.TryParse(target, out var rr, out var cc)) { G.Select(rr, cc); G.EnsureVisible(rr, cc); }
        }
    }

    // ---------- Imposta pagina ----------
    void MarginsMenu(Button anchor) => DropMenu(anchor,
        ("Normali", () => Wb.SetMargins(Cur, Margins.Normal)), ("Larghi", () => Wb.SetMargins(Cur, Margins.Wide)), ("Stretti", () => Wb.SetMargins(Cur, Margins.Narrow)));

    void OrientationMenu(Button anchor) => DropMenu(anchor, ("Verticale", () => Wb.SetOrientation(Cur, false)), ("Orizzontale", () => Wb.SetOrientation(Cur, true)));

    void PaperMenu(Button anchor) => DropMenu(anchor, Workbook.PaperSizes.Select(p => ($"{p.Name}  ({p.W:0.#} x {p.H:0.#} mm)", (Action)(() => Wb.SetPaperSize(Cur, p.Code)))).ToArray());

    void PrintAreaMenu(Button anchor) => DropMenu(anchor,
        ("Imposta area di stampa", () => { var (r1, c1, r2, c2) = G.Selection; Wb.SetPrintArea(Cur, (r1, c1, r2, c2)); }),
        ("Cancella area di stampa", () => Wb.SetPrintArea(Cur, null)));
}

public partial class MainWindow
{
    // ---------- Esporta PDF e Stampa ----------
    void ExportPdf()
    {
        var res = ThemedDialog.AskSave(this, ThemedDialog.SaveMode.SaveAs, null, "", Cur.Name, DefaultFolder(), ".pdf", "PDF (*.pdf)|*.pdf");
        if (res.Choice != SaveLogic.Choice.Save || res.Path == null) return;
        try
        {
            Cursor = Cursors.Wait;
            int n = PowerTables.Core.IO.PdfExport.Export(Wb, Cur, res.Path, Ci);
            StatusRight.Text = $"PDF esportato ({n} pagine)";
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException or InvalidOperationException) { ThemedDialog.Notify(this, ex.Message, DialogIcon.Error, "Esportazione non riuscita"); }
        finally { Cursor = null; }
    }

    /// <summary>Stampa: crea un PDF temporaneo del foglio e lo apre nel programma predefinito, da cui si stampa.</summary>
    void PrintSheet()
    {
        try
        {
            string dir = Path.Combine(Path.GetTempPath(), "PowerTables"); Directory.CreateDirectory(dir);
            string path = Path.Combine(dir, "Stampa-" + Guid.NewGuid().ToString("N")[..8] + ".pdf");
            int n = PowerTables.Core.IO.PdfExport.Export(Wb, Cur, path, Ci);
            StatusRight.Text = $"Anteprima di stampa aperta ({n} pagine)";
            Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException or InvalidOperationException or System.ComponentModel.Win32Exception) { ThemedDialog.Notify(this, ex.Message, DialogIcon.Error, "Stampa non riuscita"); }
    }
}
