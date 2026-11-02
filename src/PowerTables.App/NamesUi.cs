using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Media;
using PowerTables.Core;

namespace PowerTables.App;

/// <summary>Formule &gt; Definisci nome e Gestione nomi.</summary>
public partial class MainWindow
{
    string SelectionRef()
    {
        var (r1, c1, r2, c2) = G.Selection;
        string a(int r, int c) => "$" + Addr.ColName(c) + "$" + r;
        return Workbook.QuoteSheet(Cur.Name) + "!" + (r1 == r2 && c1 == c2 ? a(r1, c1) : a(r1, c1) + ":" + a(r2, c2));
    }

    /// <summary>Finestra di definizione; restituisce true se il nome è stato salvato. <paramref name="edit"/> = nome esistente da modificare.</summary>
    bool NameEditor(string title, (string Name, Sheet? Scope, string RefersTo)? edit)
    {
        var d = Form(title);
        var name = FormText(d, "Nome", edit?.Name ?? "", "NameName");
        var scopes = new[] { "Cartella di lavoro" }.Concat(Wb.Sheets.Select(s => s.Name)).ToArray();
        var scope = FormCombo(d, "Ambito", scopes, edit?.Scope == null ? 0 : edit.Value.Scope!.Index + 1, "NameScope");
        var refs = FormText(d, "Riferito a", edit?.RefersTo is { } rt ? "=" + rt : "=" + SelectionRef(), "NameRefers");
        if (edit != null) { name.IsEnabled = false; scope.IsEnabled = false; }
        d.FocusFirst(edit != null ? refs : name);
        var err = new TextBlock { Foreground = Theme.Brush(Color.FromRgb(0xe5, 0x5b, 0x5b)), FontFamily = new FontFamily("Segoe UI"), FontSize = 12, Margin = new Thickness(0, 8, 0, 0), TextWrapping = TextWrapping.Wrap };
        AutomationProperties.SetAutomationId(err, "NameError");
        d.Body.Children.Add(err);
        d.AddButton("Annulla", "Cancel", cancel: true);
        var ok = d.AddButton("OK", "Ok", primary: true, close: false);
        bool saved = false;
        ok.Click += (_, _) =>
        {
            try
            {
                var sc = scope.SelectedIndex <= 0 ? null : Wb.Sheets[scope.SelectedIndex - 1];
                if (edit == null && Wb.HasName(name.Text.Trim(), sc)) throw new ArgumentException("Esiste già un nome uguale nello stesso ambito.");
                Wb.DefineName(name.Text, refs.Text, edit != null ? edit.Value.Scope : sc);
                saved = true; d.DialogResult = true;
            }
            catch (ArgumentException ex) { err.Text = ex.Message; }
        };
        d.ShowDialog();
        return saved;
    }

    void DefineNameDialog() => NameEditor("Nuovo nome", null);

    void NameManagerDialog()
    {
        var d = Form("Gestione nomi");
        var list = new ListBox { Height = 240, MinWidth = 560, FontFamily = new FontFamily("Consolas"), FontSize = 12, Margin = new Thickness(0, 10, 0, 0) };
        list.Background = Theme.Brush(Theme.PanelAlt); list.Foreground = Theme.Brush(Theme.Text); list.BorderBrush = Theme.Brush(Theme.Border);
        AutomationProperties.SetAutomationId(list, "NameList");
        List<(string Name, Sheet? Scope, string RefersTo)> items = [];
        void Fill()
        {
            items = Wb.ListNames();
            list.ItemsSource = items.Select(n => $"{n.Name,-22} {(n.Scope?.Name ?? "Cartella"),-14} ={n.RefersTo}").ToList();
            if (items.Count > 0) list.SelectedIndex = 0;
        }
        d.Body.Children.Add(list);
        d.AddButton("Chiudi", "Close", cancel: true);
        var del = d.AddButton("Elimina", "Del", close: false);
        var edit = d.AddButton("Modifica...", "Edit", close: false);
        var add = d.AddButton("Nuovo...", "New", close: false);
        add.Click += (_, _) => { if (NameEditor("Nuovo nome", null)) Fill(); };
        edit.Click += (_, _) => { if (list.SelectedIndex >= 0 && NameEditor("Modifica nome", items[list.SelectedIndex])) Fill(); };
        del.Click += (_, _) => { if (list.SelectedIndex >= 0) { var n = items[list.SelectedIndex]; Wb.DeleteName(n.Name, n.Scope); Fill(); } };
        Fill();
        d.ShowDialog();
    }
}
