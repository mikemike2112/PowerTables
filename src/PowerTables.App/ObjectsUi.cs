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
