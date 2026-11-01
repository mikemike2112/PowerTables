using System.IO;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Media;
using PowerTables.Core.IO;

namespace PowerTables.App;

/// <summary>Scudo privacy: mostra cosa contiene il file salvato e crea una COPIA senza i dati scelti. Il file aperto non viene mai toccato.</summary>
public partial class MainWindow
{
    TextBlock Note(string t, double size = 12, bool dim = true, double top = 8)
    {
        var tb = new TextBlock { Text = t, FontFamily = new FontFamily("Segoe UI"), FontSize = size, TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, top, 0, 0), MaxWidth = 540 };
        tb.SetResourceReference(TextBlock.ForegroundProperty, dim ? "TextDimBrush" : "TextBrush");
        return tb;
    }

    void PrivacyShieldDialog()
    {
        string? path = Wb.FilePath;
        if (path == null || Wb.Modified || !path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            ThemedDialog.Notify(this, "Lo Scudo privacy analizza il file salvato su disco. Salva il documento (.xlsx) e riprova: il file aperto non verrà modificato.", DialogIcon.Info, "Scudo privacy");
            return;
        }
        PrivacyShield.Report rep;
        try { rep = PrivacyShield.Analyze(path); }
        catch (Exception ex) when (ex is IOException or InvalidDataException or System.Xml.XmlException or UnauthorizedAccessException)
        {
            ThemedDialog.Notify(this, "Impossibile analizzare il file: " + ex.Message, DialogIcon.Error, "Scudo privacy"); return;
        }
        var d = new ThemedDialog(this, "Scudo privacy", rep.Findings.Count == 0 ? "Tra i dati controllati non è stato trovato nulla di personale." : $"Trovate {rep.Findings.Count} voci nei metadati. Scegli cosa togliere dalla copia.", rep.Findings.Count == 0 ? DialogIcon.Info : DialogIcon.Warning);
        var list = new StackPanel(); var boxes = new List<(CheckBox Box, PrivacyShield.Finding F)>();
        int n = 0;
        foreach (var f in rep.Findings)
        {
            var cb = new CheckBox { Content = $"{f.Label}: {f.Value}", IsChecked = true, Margin = new Thickness(0, 3, 0, 3), FontFamily = new FontFamily("Segoe UI"), FontSize = 13 };
            cb.SetResourceReference(Control.ForegroundProperty, "TextBrush");
            AutomationProperties.SetAutomationId(cb, "Priv_" + n++); AutomationProperties.SetName(cb, cb.Content.ToString());
            list.Children.Add(cb); boxes.Add((cb, f));
        }
        RadioButton? anon = null, remove = null;
        if (rep.HasComments)
        {
            list.Children.Add(Note("Commenti e revisioni", 12, true, 12));
            anon = new RadioButton { Content = "Anonimizza gli autori (il contenuto resta invariato)", IsChecked = true, GroupName = "c", Margin = new Thickness(0, 4, 0, 2), FontSize = 13 };
            remove = new RadioButton { Content = "Rimuovi commenti e thread", GroupName = "c", Margin = new Thickness(0, 2, 0, 2), FontSize = 13 };
            foreach (var r in new[] { anon, remove }) r.SetResourceReference(Control.ForegroundProperty, "TextBrush");
            AutomationProperties.SetAutomationId(anon, "Priv_Anonymize"); AutomationProperties.SetAutomationId(remove, "Priv_Remove");
            list.Children.Add(anon); list.Children.Add(remove);
        }
        if (rep.Warnings.Count > 0)
        {
            list.Children.Add(Note("Avvisi: cosa lo Scudo non sa pulire", 12, true, 12));
            foreach (var w in rep.Warnings) list.Children.Add(Note("• " + w, 12, true, 3));
        }
        list.Children.Add(Note("Il file aperto non viene modificato: viene creata una copia. Tutto avviene in locale.", 12, true, 12));
        d.Body.Children.Add(new ScrollViewer { Content = list, MaxHeight = 330, VerticalScrollBarVisibility = ScrollBarVisibility.Auto, Margin = new Thickness(0, 10, 0, 0) });
        d.AddButton("Chiudi", "Cancel", cancel: true);
        var make = d.AddButton("Crea copia pulita…", "Make", primary: true, close: false);
        AutomationProperties.SetAutomationId(make, "Priv_Make");
        if (rep.Findings.Count == 0 && !rep.HasComments) make.IsEnabled = false;
        make.Click += (_, _) =>
        {
            var sel = boxes.Where(b => b.Box.IsChecked == true).Select(b => b.F).ToList();
            if (sel.Count == 0) { ThemedDialog.Notify(d, "Seleziona almeno una voce da togliere.", DialogIcon.Info, "Scudo privacy"); return; }
            string stem = Path.GetFileNameWithoutExtension(path);
            var save = ThemedDialog.AskSave(d, ThemedDialog.SaveMode.SaveAs, null, "", stem + " (pulito)", Path.GetDirectoryName(path)!, ".xlsx", XlsxFilter);
            if (save.Choice != SaveLogic.Choice.Save || save.Path == null) return;
            if (Path.GetFullPath(save.Path).Equals(Path.GetFullPath(path), StringComparison.OrdinalIgnoreCase)) { ThemedDialog.Notify(d, "La copia non può sostituire il file originale: scegli un altro nome.", DialogIcon.Warning, "Scudo privacy"); return; }
            try
            {
                var after = PrivacyShield.CreateCleanCopy(path, save.Path, sel, remove?.IsChecked == true ? PrivacyShield.CommentsMode.Remove : PrivacyShield.CommentsMode.Anonymize);
                d.Close();
                ShowPrivacyResult(save.Path, sel, after);
            }
            catch (Exception ex) when (ex is IOException or InvalidDataException or UnauthorizedAccessException or InvalidOperationException or System.Xml.XmlException)
            {
                ThemedDialog.Notify(d, "Copia non riuscita, il file originale non è stato toccato: " + ex.Message, DialogIcon.Error, "Scudo privacy");
            }
        };
        d.ShowDialog();
    }

    void ShowPrivacyResult(string dest, List<PrivacyShield.Finding> removed, PrivacyShield.Report after)
    {
        var sb = new System.Text.StringBuilder();
        sb.AppendLine($"Copia creata: {Path.GetFileName(dest)}");
        sb.AppendLine($"Voci tolte o anonimizzate: {removed.Count}.");
        sb.AppendLine(after.Findings.Count == 0 ? "Ricontrollo della copia: tra i dati controllati non è rimasto nulla." : "Ricontrollo della copia: restano " + after.Findings.Count + " voci (" + string.Join("; ", after.Findings.Take(3).Select(f => f.Label + ": " + f.Value)) + ").");
        if (after.Warnings.Count > 0) { sb.AppendLine(); sb.AppendLine("Riserve:"); foreach (var w in after.Warnings) sb.AppendLine("• " + w); }
        sb.AppendLine(); sb.Append("Questo controllo copre solo i metadati elencati: non garantisce l'assenza di altri dati personali nel contenuto.");
        var d = new ThemedDialog(this, "Scudo privacy", "", after.Findings.Count == 0 ? DialogIcon.Info : DialogIcon.Warning);
        d.Body.Children.Add(Note(sb.ToString(), 13, false, 8));
        AutomationProperties.SetAutomationId(d.Body, "Priv_Result");
        d.AddButton("OK", "Ok", primary: true, cancel: true);
        d.ShowDialog();
    }
}
