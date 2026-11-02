using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Media;
using PowerTables.Core;

namespace PowerTables.App;

public partial class MainWindow
{
    ThemedDialog? _findDlg;

    void EditLines(bool rows, bool insert)
    {
        var (r1, c1, r2, c2) = G.Selection;
        int at = rows ? r1 : c1, n = rows ? r2 - r1 + 1 : c2 - c1 + 1;
        try { if (insert) Wb.InsertLines(Cur, rows, at, n); else Wb.DeleteLines(Cur, rows, at, n); }
        catch (InvalidOperationException ex) { ThemedDialog.Notify(this, ex.Message, DialogIcon.Warning); return; }
        G.RelayoutAll(); G.Focus();
    }

    void ShowFind(bool replace)
    {
        if (_findDlg != null) { _findDlg.Close(); }
        var d = _findDlg = new ThemedDialog(this, replace ? "Sostituisci" : "Trova", "", DialogIcon.Info);
        TextBox Box(string id, string name)
        {
            var t = new TextBox { Margin = new Thickness(0, 6, 0, 0), Padding = new Thickness(8, 6, 8, 6), FontSize = 13, FontFamily = new FontFamily("Segoe UI"), MinWidth = 320 };
            t.SetResourceReference(Control.BackgroundProperty, "PanelAltBrush"); t.SetResourceReference(Control.ForegroundProperty, "TextBrush"); t.SetResourceReference(Control.BorderBrushProperty, "BorderBrush2");
            AutomationProperties.SetAutomationId(t, id); AutomationProperties.SetName(t, name);
            return t;
        }
        CheckBox Chk(string text, string id)
        {
            var c = new CheckBox { Content = text, Margin = new Thickness(0, 8, 12, 0) };
            c.SetResourceReference(Control.ForegroundProperty, "TextBrush"); AutomationProperties.SetAutomationId(c, id);
            return c;
        }
        var find = Box("FindText", "Trova"); var rep = Box("ReplaceText", "Sostituisci con");
        var mc = Chk("Maiuscole/minuscole", "FindCase"); var wc = Chk("Intero contenuto della cella", "FindWhole");
        var status = new TextBlock { Margin = new Thickness(0, 8, 0, 0), FontSize = 12 }; status.SetResourceReference(TextBlock.ForegroundProperty, "TextDimBrush");
        AutomationProperties.SetAutomationId(status, "FindStatus");
        d.MaxWidth = 760;   // 5 pulsanti: col limite standard (620) "Trova precedente" veniva tagliato
        TextBlock Label(string t) { var l = new TextBlock { Text = t, Margin = new Thickness(0, 8, 0, 0), FontSize = 12 }; l.SetResourceReference(TextBlock.ForegroundProperty, "TextDimBrush"); return l; }
        d.Body.Children.Add(Label("Trova:")); d.Body.Children.Add(find);
        if (replace) { d.Body.Children.Add(Label("Sostituisci con:")); d.Body.Children.Add(rep); }
        d.Body.Children.Add(new StackPanel { Orientation = Orientation.Horizontal, Children = { mc, wc } });
        d.Body.Children.Add(status);
        void Next(bool fwd)
        {
            var hit = Wb.Find(Cur, find.Text, G.ActiveRow, G.ActiveCol, fwd, mc.IsChecked == true, wc.IsChecked == true, Ci);
            if (hit is { } h) { G.Select(h.Row, h.Col); status.Text = ""; } else status.Text = "Nessuna corrispondenza.";
        }
        var prev = d.AddButton("Trova precedente", "FindPrev", close: false); prev.Click += (_, _) => Next(false);
        var next = d.AddButton("Trova successivo", "FindNext", primary: !replace, close: false); next.Click += (_, _) => Next(true);
        if (replace)
        {
            var one = d.AddButton("Sostituisci", "ReplaceOne", primary: true, close: false);
            one.Click += (_, _) =>
            {
                // sostituisce nella cella attiva se corrisponde, poi passa alla successiva
                var (ar, ac) = (G.ActiveRow, G.ActiveCol);
                int n = Wb.ReplaceAll(Cur, find.Text, rep.Text, mc.IsChecked == true, wc.IsChecked == true, (ar, ac, ar, ac), Ci);
                if (n > 0) G.RelayoutAll();
                Next(true);
            };
            var all = d.AddButton("Sostituisci tutto", "ReplaceAll", close: false);
            all.Click += (_, _) =>
            {
                var (r1, c1, r2, c2) = G.Selection;
                var area = r1 == r2 && c1 == c2 ? ((int, int, int, int)?)null : (r1, c1, r2, c2);
                int n = Wb.ReplaceAll(Cur, find.Text, rep.Text, mc.IsChecked == true, wc.IsChecked == true, area, Ci);
                status.Text = n == 0 ? "Nessuna corrispondenza." : $"{n} sostituzioni.";
                G.RelayoutAll();
            };
        }
        var close = d.AddButton("Chiudi", "FindClose", cancel: true, close: false); close.Click += (_, _) => d.Close();
        d.Closed += (_, _) => { if (ReferenceEquals(_findDlg, d)) _findDlg = null; };
        d.FocusFirst(find);
        d.Show();
    }
}
