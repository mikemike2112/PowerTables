using System.IO;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Microsoft.Win32;
using PowerTables.Core;

namespace PowerTables.App;

/// <summary>Inserisci immagini, forme, caselle di testo, grafici e commenti.</summary>
public partial class MainWindow
{
    void Warn(Exception ex, string title) => ThemedDialog.Notify(this, ex.Message, DialogIcon.Warning, title);

    /// <summary>Inserisce l'immagine dai byte (la finestra di scelta file chiama questo): la dimensione in pixel viene letta dall'immagine stessa.</summary>
    void InsertPictureBytes(byte[] data)
    {
        try
        {
            int w = 64, h = 40;
            try { var bi = new BitmapImage(); bi.BeginInit(); bi.CacheOption = BitmapCacheOption.OnLoad; bi.StreamSource = new MemoryStream(data); bi.EndInit(); w = bi.PixelWidth; h = bi.PixelHeight; }
            catch (Exception e) when (e is NotSupportedException or InvalidOperationException or FileFormatException) { throw new ArgumentException("Il file non è un'immagine leggibile."); }
            Wb.AddPicture(Cur, G.ActiveRow, G.ActiveCol, data, w, h);
            G.InvalidateVisual();
        }
        catch (ArgumentException ex) { Warn(ex, "Inserisci immagine"); }
    }

    void InsertPicture()
    {
        var dlg = new OpenFileDialog { Filter = "Immagini (*.png;*.jpg;*.jpeg;*.gif)|*.png;*.jpg;*.jpeg;*.gif" };
        if (dlg.ShowDialog(this) != true) return;
        try
        {
            if (new FileInfo(dlg.FileName).Length > 10_000_000) throw new ArgumentException("Immagine più grande di 10 MB.");
            InsertPictureBytes(File.ReadAllBytes(dlg.FileName));
        }
        catch (ArgumentException ex) { Warn(ex, "Inserisci immagine"); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { ThemedDialog.Notify(this, "Impossibile leggere il file: " + ex.Message, DialogIcon.Error, "Inserisci immagine"); }
    }

    void ShapesMenu(Button anchor) => DropMenu(anchor,
        ("Rettangolo", () => ShapeDialog("rect", false)), ("Rettangolo arrotondato", () => ShapeDialog("roundRect", false)), ("Ellisse", () => ShapeDialog("ellipse", false)),
        ("Triangolo", () => ShapeDialog("triangle", false)), ("Freccia a destra", () => ShapeDialog("rightArrow", false)));

    /// <summary>Chiede il testo e inserisce la forma o la casella di testo sulla cella attiva.</summary>
    void ShapeDialog(string geom, bool textBox)
    {
        var d = Form(textBox ? "Casella di testo" : "Forma");
        var t = FormText(d, "Testo (vuoto = nessun testo)", "", "ShapeText");
        var err = new TextBlock { Foreground = Theme.Brush(Color.FromRgb(0xe5, 0x5b, 0x5b)), FontFamily = new FontFamily("Segoe UI"), FontSize = 12, Margin = new Thickness(0, 8, 0, 0) };
        AutomationProperties.SetAutomationId(err, "ShapeError"); d.Body.Children.Add(err);
        d.FocusFirst(t);
        d.AddButton("Annulla", "Cancel", cancel: true);
        var ok = d.AddButton("OK", "Ok", primary: true, close: false);
        ok.Click += (_, _) =>
        {
            try { Wb.AddShape(Cur, G.ActiveRow, G.ActiveCol, textBox, geom, t.Text); G.InvalidateVisual(); d.DialogResult = true; }
            catch (ArgumentException ex) { err.Text = ex.Message; }
        };
        d.ShowDialog();
    }

    void InsertChart(string type)
    {
        try { var (r1, c1, r2, c2) = UsedSelection(); Wb.AddChart(Cur, r1, c1, r2, c2, type); G.InvalidateVisual(); }
        catch (Exception ex) when (ex is ArgumentException or InvalidOperationException) { Warn(ex, "Inserisci grafico"); }
    }

    /// <summary>Grafici consigliati: propone istogramma o linee in base alla prima colonna (date o testo con tante righe → linee).</summary>
    void RecommendedCharts(Button anchor) => DropMenu(anchor, ("Istogramma a colonne", () => InsertChart("col")), ("Linee", () => InsertChart("line")));

    void CommentDialog()
    {
        Cur.Comments.TryGetValue(Addr.Key(G.ActiveRow, G.ActiveCol), out var old);
        var d = Form("Commento", old == null ? "Cella " + Addr.A1(G.ActiveRow, G.ActiveCol) : "Cella " + Addr.A1(G.ActiveRow, G.ActiveCol) + " (testo vuoto = elimina il commento)");
        var t = FormText(d, "Testo del commento", old?.Text ?? "", "CommentText");
        var err = new TextBlock { Foreground = Theme.Brush(Color.FromRgb(0xe5, 0x5b, 0x5b)), FontFamily = new FontFamily("Segoe UI"), FontSize = 12, Margin = new Thickness(0, 8, 0, 0) };
        AutomationProperties.SetAutomationId(err, "CommentError"); d.Body.Children.Add(err);
        d.FocusFirst(t);
        d.AddButton("Annulla", "Cancel", cancel: true);
        var ok = d.AddButton("OK", "Ok", primary: true, close: false);
        ok.Click += (_, _) =>
        {
            try { Wb.SetComment(Cur, G.ActiveRow, G.ActiveCol, t.Text, Environment.UserName.Length > 0 ? "PowerTables" : ""); G.InvalidateVisual(); d.DialogResult = true; }
            catch (ArgumentException ex) { err.Text = ex.Message; }
        };
        d.ShowDialog();
    }
}

public partial class MainWindow
{
    static bool ParseArea(string text, out (int R1, int C1, int R2, int C2) a)
    {
        string t = text.Trim(); t = t[(t.LastIndexOf('!') + 1)..].Replace("$", "");
        a = default;
        if (Addr.TryParseRange(t, out var r1, out var c1, out var r2, out var c2)) { a = (r1, c1, r2, c2); return true; }
        if (Addr.TryParse(t, out var r, out var c)) { a = (r, c, r, c); return true; }
        return false;
    }

    /// <summary>Dati &gt; Avanzate: filtro con intervallo dei criteri, sul posto o con copia, anche solo record univoci.</summary>
    void AdvancedFilterDialog()
    {
        var (r1, c1, r2, c2) = CurrentRegion();
        var d = Form("Filtro avanzato", "Le righe sotto l'intestazione dei criteri sono alternative (O); più colonne nella stessa riga si combinano (E).");
        var list = FormText(d, "Intervallo elenco", Addr.A1(r1, c1) + ":" + Addr.A1(r2, c2), "AdvList");
        var crit = FormText(d, "Intervallo criteri (intestazione + almeno una riga)", "", "AdvCrit");
        var copy = FormCheck(d, "Copia in un'altra posizione", false, "AdvCopy");
        var dest = FormText(d, "Copia in (cella iniziale)", "", "AdvDest");
        var uniq = FormCheck(d, "Solo record univoci", false, "AdvUnique");
        var err = new TextBlock { Foreground = Theme.Brush(Color.FromRgb(0xe5, 0x5b, 0x5b)), FontFamily = new FontFamily("Segoe UI"), FontSize = 12, Margin = new Thickness(0, 8, 0, 0), TextWrapping = TextWrapping.Wrap };
        AutomationProperties.SetAutomationId(err, "AdvError"); d.Body.Children.Add(err);
        d.FocusFirst(crit);
        d.AddButton("Annulla", "Cancel", cancel: true);
        var ok = d.AddButton("OK", "Ok", primary: true, close: false);
        ok.Click += (_, _) =>
        {
            try
            {
                if (!ParseArea(list.Text, out var la)) throw new ArgumentException("Intervallo elenco non valido.");
                if (!ParseArea(crit.Text, out var ca)) throw new ArgumentException("Intervallo criteri non valido.");
                int n;
                if (copy.IsChecked == true)
                {
                    if (!ParseArea(dest.Text, out var da)) throw new ArgumentException("Cella di destinazione non valida.");
                    n = Wb.AdvancedFilterCopy(Cur, la, ca, uniq.IsChecked == true, (da.R1, da.C1), Ci);
                }
                else n = Wb.AdvancedFilterInPlace(Cur, la, ca, uniq.IsChecked == true, Ci);
                G.RelayoutAll(); StatusRight.Text = $"Filtro avanzato: {n} record";
                d.DialogResult = true;
            }
            catch (Exception ex) when (ex is ArgumentException or InvalidOperationException) { err.Text = ex.Message; }
        };
        d.ShowDialog();
    }

    /// <summary>Dati &gt; Avanzate &gt; Mostra tutto: toglie il filtro avanzato sul posto.</summary>
    void ClearAdvancedFilter()
    {
        var (r1, _, r2, _) = CurrentRegion();
        Wb.ShowAllRows(Cur, r1, r2); G.RelayoutAll();
    }
}

public partial class MainWindow
{
    TextBlock ErrBlock(ThemedDialog d, string id)
    {
        var err = new TextBlock { Foreground = Theme.Brush(Color.FromRgb(0xe5, 0x5b, 0x5b)), FontFamily = new FontFamily("Segoe UI"), FontSize = 12, Margin = new Thickness(0, 8, 0, 0), TextWrapping = TextWrapping.Wrap };
        AutomationProperties.SetAutomationId(err, id); d.Body.Children.Add(err);
        return err;
    }

    /// <summary>Dati &gt; Analisi di simulazione &gt; Ricerca obiettivo.</summary>
    void GoalSeekDialog()
    {
        var d = Form("Ricerca obiettivo");
        var set = FormText(d, "Imposta la cella (con formula)", Addr.A1(G.ActiveRow, G.ActiveCol), "GsSet");
        var val = FormText(d, "Al valore", "0", "GsValue");
        var chg = FormText(d, "Cambiando la cella (numero)", "", "GsChange");
        var err = ErrBlock(d, "GsError");
        d.FocusFirst(val);
        d.AddButton("Annulla", "Cancel", cancel: true);
        var ok = d.AddButton("OK", "Ok", primary: true, close: false);
        ok.Click += (_, _) =>
        {
            try
            {
                if (!ParseArea(set.Text, out var sa) || !ParseArea(chg.Text, out var ca)) throw new ArgumentException("Riferimento di cella non valido.");
                if (!double.TryParse(val.Text, System.Globalization.NumberStyles.Float, Ci, out var target)) throw new ArgumentException("Valore non numerico.");
                var res = Wb.GoalSeek(Cur, sa.R1, sa.C1, target, ca.R1, ca.C1);
                if (res == null) { err.Text = "Nessuna soluzione trovata: la cella da cambiare è rimasta invariata."; return; }
                StatusRight.Text = $"Ricerca obiettivo: {Addr.A1(ca.R1, ca.C1)} = {res.Value.ToString("G6", Ci)}";
                d.DialogResult = true;
            }
            catch (Exception ex) when (ex is ArgumentException or InvalidOperationException) { err.Text = ex.Message; }
        };
        d.ShowDialog();
    }

    void ForecastDialog()
    {
        var (r1, c1, r2, c2) = CurrentRegion();
        var d = Form("Crea foglio previsione", "Previsione lineare (regressione) con limiti di confidenza al 95%.");
        var tl = FormText(d, "Cronologia (date o numeri crescenti)", c2 > c1 ? Addr.A1(r1 + 1, c1) + ":" + Addr.A1(r2, c1) : "", "FcTimeline");
        var vals = FormText(d, "Valori", c2 > c1 ? Addr.A1(r1 + 1, c1 + 1) + ":" + Addr.A1(r2, c1 + 1) : "", "FcValues");
        var per = FormText(d, "Periodi da prevedere", "5", "FcPeriods", 80);
        var err = ErrBlock(d, "FcErrorF");
        d.FocusFirst(tl);
        d.AddButton("Annulla", "Cancel", cancel: true);
        var ok = d.AddButton("OK", "Ok", primary: true, close: false);
        ok.Click += (_, _) =>
        {
            try
            {
                if (!ParseArea(tl.Text, out var ta) || !ParseArea(vals.Text, out var va)) throw new ArgumentException("Intervallo non valido.");
                if (!int.TryParse(per.Text, out var p)) throw new ArgumentException("Numero di periodi non valido.");
                var ns = Wb.CreateForecastSheet(Cur, va, ta, p);
                G.SwitchSheet(ns); RebuildSheetTabs(); DumpState();
                d.DialogResult = true;
            }
            catch (Exception ex) when (ex is ArgumentException or InvalidOperationException) { err.Text = ex.Message; }
        };
        d.ShowDialog();
    }
}

public partial class MainWindow
{
    /// <summary>Inserisci &gt; Tabella pivot: un campo riga, un campo valori (somma o conteggio), in un nuovo foglio o in una cella.</summary>
    void PivotDialog()
    {
        var (r1, c1, r2, c2) = CurrentRegion();
        string Head(int c) => Cur.Get(r1, c).Obj as string ?? "";
        var d = Form("Tabella pivot", "L'origine deve avere intestazioni di testo nella prima riga.");
        var src = FormText(d, "Origine (con intestazioni)", Addr.A1(r1, c1) + ":" + Addr.A1(r2, c2), "PvSource");
        var rowF = FormText(d, "Campo riga (intestazione)", c2 > c1 ? Head(c1) : "", "PvRow");
        var valF = FormText(d, "Campo valori (intestazione)", c2 > c1 ? Head(c1 + 1) : "", "PvValue");
        var fn = FormCombo(d, "Funzione", ["Somma", "Conteggio"], 0, "PvFunc");
        var newSheet = FormCheck(d, "Nuovo foglio", true, "PvNewSheet");
        var dest = FormText(d, "Altrimenti: cella di destinazione nel foglio corrente", "", "PvDest");
        var err = ErrBlock(d, "PvError");
        d.FocusFirst(rowF);
        d.AddButton("Annulla", "Cancel", cancel: true);
        var ok = d.AddButton("OK", "Ok", primary: true, close: false);
        ok.Click += (_, _) =>
        {
            try
            {
                if (!ParseArea(src.Text, out var a)) throw new ArgumentException("Origine non valida.");
                int Find(string name)
                {
                    for (int c = a.C1; c <= a.C2; c++) if (string.Equals(Cur.Get(a.R1, c).Obj as string, name.Trim(), StringComparison.OrdinalIgnoreCase)) return c;
                    throw new ArgumentException($"Intestazione «{name.Trim()}» non trovata nella prima riga dell'origine.");
                }
                int rc = Find(rowF.Text), dc = Find(valF.Text); bool count = fn.SelectedIndex == 1;
                var srcSheet = Cur; Sheet? created = null;
                if (newSheet.IsChecked == true)
                {
                    try { Wb.Do("Tabella pivot", () => { created = Wb.InsertSheet(); Wb.CreatePivot(srcSheet, a, rc, dc, count, created, 1, 1, Ci); }); }
                    catch (ArgumentException) { if (created != null) { Wb.Undo(); RebuildSheetTabs(); } throw; }   // niente foglio vuoto rimasto dopo un errore
                }
                else
                {
                    if (!ParseArea(dest.Text, out var da)) throw new ArgumentException("Cella di destinazione non valida.");
                    Wb.CreatePivot(srcSheet, a, rc, dc, count, srcSheet, da.R1, da.C1, Ci);
                }
                if (created != null) { G.SwitchSheet(created); RebuildSheetTabs(); DumpState(); }
                G.RelayoutAll();
                d.DialogResult = true;
            }
            catch (Exception ex) when (ex is ArgumentException or InvalidOperationException) { err.Text = ex.Message; }
        };
        d.ShowDialog();
    }
}
