using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using PowerTables.Core;

namespace PowerTables.App;

/// <summary>Strumenti Dati e Appunti: copia formato, rimuovi duplicati, testo in colonne, e piccoli elementi per i dialoghi.</summary>
public partial class MainWindow
{
    // ---------- elementi di modulo per i dialoghi ----------
    static void StyleInput(Control c)
    {
        c.FontFamily = new FontFamily("Segoe UI"); c.FontSize = 13; c.Margin = new Thickness(0, 4, 0, 4);
        c.Background = Theme.Brush(Theme.PanelAlt); c.Foreground = Theme.Brush(Theme.Text); c.BorderBrush = Theme.Brush(Theme.Border);
    }

    static TextBlock FormLabel(string t) => new() { Text = t, FontFamily = new FontFamily("Segoe UI"), FontSize = 12, Foreground = Theme.Brush(Theme.TextDim), Margin = new Thickness(0, 10, 0, 0) };

    static TextBox FormText(ThemedDialog d, string label, string def, string id, double minW = 360)
    {
        var tb = new TextBox { Text = def, MinWidth = minW, Padding = new Thickness(6, 4, 6, 4) }; StyleInput(tb);
        AutomationProperties.SetAutomationId(tb, id); AutomationProperties.SetName(tb, label);
        d.Body.Children.Add(FormLabel(label)); d.Body.Children.Add(tb);
        return tb;
    }

    static CheckBox FormCheck(ThemedDialog d, string label, bool on, string id)
    {
        var cb = new CheckBox { Content = label, IsChecked = on, Margin = new Thickness(0, 6, 0, 2), Foreground = Theme.Brush(Theme.Text), FontFamily = new FontFamily("Segoe UI"), FontSize = 13 };
        AutomationProperties.SetAutomationId(cb, id); AutomationProperties.SetName(cb, label);
        d.Body.Children.Add(cb);
        return cb;
    }

    static ComboBox FormCombo(ThemedDialog d, string label, string[] items, int sel, string id)
    {
        var cb = new ComboBox { ItemsSource = items, SelectedIndex = sel, MinWidth = 360, Padding = new Thickness(6, 4, 6, 4) }; StyleInput(cb);
        AutomationProperties.SetAutomationId(cb, id); AutomationProperties.SetName(cb, label);
        d.Body.Children.Add(FormLabel(label)); d.Body.Children.Add(cb);
        return cb;
    }

    ThemedDialog Form(string title, string msg = "") => new(this, title, msg, DialogIcon.Info);

    /// <summary>Selezione ridotta all'area usata del foglio (le selezioni di intere righe/colonne diventano finite).</summary>
    (int R1, int C1, int R2, int C2) UsedSelection()
    {
        var (r1, c1, r2, c2) = G.Selection;
        return (r1, c1, Math.Min(r2, Math.Max(r1, Cur.MaxRow)), Math.Min(c2, Math.Max(c1, Cur.MaxCol)));
    }

    // ---------- Copia formato ----------
    (int R1, int C1, int R2, int C2)? _fmtSrc; Sheet? _fmtSheet; Button? _fmtBtn;

    void InitTools()
    {
        PreviewMouseLeftButtonUp += (_, e) =>
        {
            if (_fmtSrc == null || !(G.IsMouseOver || G.IsMouseCaptured)) return;
            Dispatcher.BeginInvoke(new Action(ApplyCopyFormat));
        };
    }

    void CopyFormat()
    {
        if (_fmtSrc != null) { _fmtSrc = null; _fmtSheet = null; Cursor = null; StatusRight.Text = "Pronto"; return; }
        _fmtSrc = G.Selection; _fmtSheet = Cur;
        StatusRight.Text = "Copia formato: selezionare l'area di destinazione";
    }

    /// <summary>Applica il formato memorizzato alla selezione corrente (una cella sola: si estende alle dimensioni dell'origine).</summary>
    void ApplyCopyFormat()
    {
        if (_fmtSrc is not { } s || _fmtSheet == null) return;
        var (r1, c1, r2, c2) = G.Selection;
        if (r1 == r2 && c1 == c2) { r2 = Math.Min(Addr.MaxRow, r1 + s.R2 - s.R1); c2 = Math.Min(Addr.MaxCol, c1 + s.C2 - s.C1); }
        _fmtSrc = null; StatusRight.Text = "Pronto";
        if ((long)(r2 - r1 + 1) * (c2 - c1 + 1) > 500_000) { r2 = Math.Min(r2, Math.Max(Cur.MaxRow, r1)); c2 = Math.Min(c2, Math.Max(Cur.MaxCol, c1)); }
        Wb.PasteFormats(_fmtSheet, s.R1, s.C1, s.R2, s.C2, Cur, r1, c1, r2, c2);
        _fmtSheet = null;
        G.InvalidateVisual(); UpdateToggles();
    }

    // ---------- Rimuovi duplicati ----------
    void RemoveDuplicatesDialog()
    {
        var (r1, c1, r2, c2) = CurrentRegion();
        if (r1 == r2) { ThemedDialog.Notify(this, "Selezionare un intervallo con almeno due righe.", DialogIcon.Warning); return; }
        var d = Form("Rimuovi duplicati", "Scegliere le colonne in cui cercare i duplicati.");
        var header = FormCheck(d, "I dati hanno intestazioni", true, "DupHeader");
        var boxes = new List<CheckBox>();
        void Labels()
        {
            for (int i = 0; i < boxes.Count; i++)
                boxes[i].Content = header.IsChecked == true ? Wb.DisplayText(Cur, r1, c1 + i, Ci) is { Length: > 0 } t ? t : "Colonna " + Addr.ColName(c1 + i) : "Colonna " + Addr.ColName(c1 + i);
        }
        for (int c = c1; c <= c2; c++) { var cb = FormCheck(d, "", true, "DupCol" + (c - c1)); boxes.Add(cb); }
        header.Click += (_, _) => Labels(); Labels();
        d.AddButton("Annulla", "Cancel", cancel: true);
        d.AddButton("OK", "Ok", primary: true);
        if (d.ShowAndGet() != "Ok") return;
        var cols = boxes.Select((b, i) => (b, i)).Where(x => x.b.IsChecked == true).Select(x => c1 + x.i).ToList();
        if (cols.Count == 0) { ThemedDialog.Notify(this, "Selezionare almeno una colonna.", DialogIcon.Warning); return; }
        int n = Wb.RemoveDuplicates(Cur, r1, c1, r2, c2, cols, header.IsChecked == true);
        int left = (r2 - r1 + 1) - n - (header.IsChecked == true ? 1 : 0);
        ThemedDialog.Notify(this, n == 0 ? "Nessun valore duplicato trovato." : $"{n} valori duplicati trovati e rimossi; {left} valori univoci rimasti.", DialogIcon.Info, "Rimuovi duplicati");
        G.SelectRange(r1, c1, r2, c2);
    }

    // ---------- Testo in colonne ----------
    void TextToColumnsDialog()
    {
        var (r1, c1, r2, c2) = G.Selection;
        r2 = Math.Min(r2, Math.Max(r1, Cur.MaxRow));
        if (c1 != c2) { ThemedDialog.Notify(this, "Selezionare una sola colonna.", DialogIcon.Warning); return; }
        var d = Form("Testo in colonne", "Divide il testo della colonna selezionata nelle colonne a destra. I valori già presenti a destra vengono sovrascritti.");
        var tab = FormCheck(d, "Tabulazione", false, "TtcTab"); var semi = FormCheck(d, "Punto e virgola", true, "TtcSemi");
        var comma = FormCheck(d, "Virgola", false, "TtcComma"); var space = FormCheck(d, "Spazio", false, "TtcSpace");
        var other = FormText(d, "Altro delimitatore", "", "TtcOther", 80);
        var merge = FormCheck(d, "Considera consecutivi i delimitatori come uno", false, "TtcMerge");
        d.AddButton("Annulla", "Cancel", cancel: true);
        d.AddButton("Fine", "Ok", primary: true);
        if (d.ShowAndGet() != "Ok") return;
        string delims = (tab.IsChecked == true ? "\t" : "") + (semi.IsChecked == true ? ";" : "") + (comma.IsChecked == true ? "," : "") + (space.IsChecked == true ? " " : "") + other.Text;
        if (delims.Length == 0) { ThemedDialog.Notify(this, "Indicare almeno un delimitatore.", DialogIcon.Warning); return; }
        int n = Wb.TextToColumns(Cur, r1, c1, r2, delims, merge.IsChecked == true, '"', Ci);
        G.SelectRange(r1, c1, r2, c1 + n - 1);
    }
}
