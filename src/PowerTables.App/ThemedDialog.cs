using System.IO;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Effects;
using PowerTables.Core.IO;

namespace PowerTables.App;

public enum DialogIcon { Info, Warning, Error, Question }

/// <summary>Finestra di dialogo propria: tema attivo, angoli arrotondati, ombra, nessun bordo di sistema. Sostituisce MessageBox.</summary>
public sealed class ThemedDialog : Window
{
    public string ResultId { get; private set; } = "";
    /// <summary>Valore scelto da dialoghi con elenco (es. nome funzione).</summary>
    public string? ResultId2;
    readonly StackPanel _body = new();
    readonly StackPanel _buttons = new() { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 18, 0, 0) };
    public readonly List<Button> Buttons = [];
    Button? _primary;
    Control? _focusFirst;

    public ThemedDialog(Window? owner, string title, string message, DialogIcon icon)
    {
        Owner = owner is { IsVisible: true } ? owner : null;
        WindowStartupLocation = Owner != null ? WindowStartupLocation.CenterOwner : WindowStartupLocation.CenterScreen;
        WindowStyle = WindowStyle.None; AllowsTransparency = true; Background = Brushes.Transparent; ShowInTaskbar = false;
        SizeToContent = SizeToContent.WidthAndHeight; ResizeMode = ResizeMode.NoResize; MinWidth = 460; MaxWidth = 620;
        Title = title;
        AutomationProperties.SetName(this, title);

        var head = new TextBlock { Text = title, FontFamily = new FontFamily("Segoe UI Semibold"), FontSize = 16, Foreground = Theme.Brush(Theme.Text), TextWrapping = TextWrapping.Wrap };
        var text = new StackPanel(); text.Children.Add(head);
        if (message.Length > 0) text.Children.Add(new TextBlock { Text = message, FontFamily = new FontFamily("Segoe UI"), FontSize = 13, Foreground = Theme.Brush(Theme.TextDim), TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 8, 0, 0) });
        text.Children.Add(_body);
        var ic = IconFor(icon);
        var row = new DockPanel();
        DockPanel.SetDock(ic, Dock.Left); ic.Margin = new Thickness(0, 2, 16, 0); ic.VerticalAlignment = VerticalAlignment.Top;
        row.Children.Add(ic); row.Children.Add(text);
        var stack = new StackPanel { Margin = new Thickness(24, 22, 24, 20) };
        stack.Children.Add(row); stack.Children.Add(_buttons);
        var card = new Border { Background = Theme.Brush(Theme.Panel), BorderBrush = Theme.Brush(Theme.Border), BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(10), Child = stack, Margin = new Thickness(16), Effect = new DropShadowEffect { BlurRadius = 18, ShadowDepth = 3, Opacity = 0.35, Color = Colors.Black } };
        Content = card;
        card.MouseLeftButtonDown += (_, e) => { if (e.OriginalSource is not (TextBox or Button or System.Windows.Controls.Primitives.TextBoxBase)) { try { DragMove(); } catch (InvalidOperationException) { } } };
        Loaded += (_, _) => { (_focusFirst ?? _primary)?.Focus(); if (_focusFirst is TextBox tb) tb.SelectAll(); };
    }

    public StackPanel Body => _body;
    public void FocusFirst(Control c) => _focusFirst = c;

    static FrameworkElement IconFor(DialogIcon icon)
    {
        // geometrie proprie su griglia 24x24
        (string path, Color color) = icon switch
        {
            DialogIcon.Error => ("M12 3a9 9 0 1 0 .01 0 M8.5 8.5l7 7 M15.5 8.5l-7 7", Color.FromRgb(0xe5, 0x5b, 0x5b)),
            DialogIcon.Warning => ("M12 3L2 20h20z M12 9v5 M12 17v.5", Color.FromRgb(0xf0, 0xa8, 0x30)),
            DialogIcon.Question => ("M12 3a9 9 0 1 0 .01 0 M9.5 9.5a2.5 2.5 0 1 1 3.5 2.3c-.7.4-1 1-1 1.7 M12 17v.5", Theme.Accent),
            _ => ("M12 3a9 9 0 1 0 .01 0 M12 11v6 M12 7.5v.5", Theme.Accent),
        };
        var p = new System.Windows.Shapes.Path { Data = Geometry.Parse(path), Stroke = Theme.Brush(color), StrokeThickness = 1.7, StrokeLineJoin = PenLineJoin.Round, StrokeStartLineCap = PenLineCap.Round, StrokeEndLineCap = PenLineCap.Round, Fill = Brushes.Transparent };
        return new Viewbox { Width = 34, Height = 34, Child = new Canvas { Width = 24, Height = 24, Children = { p } } };
    }

    static ControlTemplate ButtonTemplate(bool primary)
    {
        var bd = new FrameworkElementFactory(typeof(Border));
        bd.SetValue(Border.CornerRadiusProperty, new CornerRadius(6));
        bd.SetValue(Border.BackgroundProperty, primary ? Theme.Brush(Theme.Accent) : Theme.Brush(Theme.PanelAlt));
        bd.SetValue(Border.BorderBrushProperty, primary ? Theme.Brush(Theme.Accent) : Theme.Brush(Theme.Border));
        bd.SetValue(Border.BorderThicknessProperty, new Thickness(1));
        bd.SetValue(Border.PaddingProperty, new Thickness(14, 6, 14, 6));
        var cp = new FrameworkElementFactory(typeof(ContentPresenter));
        cp.SetValue(HorizontalAlignmentProperty, HorizontalAlignment.Center); cp.SetValue(VerticalAlignmentProperty, VerticalAlignment.Center);
        bd.AppendChild(cp);
        return new ControlTemplate(typeof(Button)) { VisualTree = bd };
    }

    /// <summary>Aggiunge un pulsante (a destra, nell'ordine di chiamata). Primario = accento + Invio; annulla = Esc.</summary>
    public Button AddButton(string label, string id, bool primary = false, bool cancel = false, bool close = true)
    {
        var b = new Button { Content = label, MinWidth = 92, Margin = new Thickness(8, 0, 0, 0), FontFamily = new FontFamily("Segoe UI"), FontSize = 13, IsDefault = primary, IsCancel = cancel, Template = ButtonTemplate(primary), Foreground = primary ? Brushes.White : Theme.Brush(Theme.Text) };
        AutomationProperties.SetName(b, label); AutomationProperties.SetAutomationId(b, "Dlg_" + id);
        if (close) b.Click += (_, _) => { ResultId = id; DialogResult = true; };
        _buttons.Children.Add(b); Buttons.Add(b);
        if (primary) _primary = b;
        return b;
    }

    /// <summary>Mostra il dialogo e restituisce l'id del pulsante premuto ("" se chiuso con Esc senza pulsante annulla).</summary>
    public string ShowAndGet() { ShowDialog(); return ResultId; }

    /// <summary>Messaggio con un solo OK (errore, avviso, informazione).</summary>
    public static void Notify(Window? owner, string message, DialogIcon icon = DialogIcon.Error, string title = "PowerTables")
    {
        var d = new ThemedDialog(owner, title, message, icon);
        d.AddButton("OK", "Ok", primary: true, cancel: true);
        d.ShowDialog();
    }

    /// <summary>Domanda con due pulsanti (primario = conferma).</summary>
    public static bool Confirm(Window? owner, string title, string message, string yes = "Sì", string no = "No", DialogIcon icon = DialogIcon.Question)
    {
        var d = new ThemedDialog(owner, title, message, icon);
        d.AddButton(no, "No", cancel: true); d.AddButton(yes, "Yes", primary: true);
        return d.ShowAndGet() == "Yes";
    }

    // ---------------- richiesta di salvataggio / Salva con nome ----------------
    public enum SaveMode { PromptDirty, SaveAs }

    /// <summary>Esito: scelta e, se Save con documento senza file o Salva con nome, il percorso completo scelto.</summary>
    public readonly record struct SaveResult(SaveLogic.Choice Choice, string? Path);

    /// <param name="existingPath">file già associato al documento (null = documento nuovo)</param>
    /// <param name="ext">estensione con punto (".xlsx", ".csv")</param>
    /// <param name="filter">filtro della finestra di sistema, senza nomi di terzi</param>
    public static SaveResult AskSave(Window? owner, SaveMode mode, string? existingPath, string docName, string suggestedName, string folder, string ext, string filter)
    {
        bool needPath = mode == SaveMode.SaveAs || existingPath == null;
        string title = mode == SaveMode.SaveAs ? "Salva con nome" : $"Vuoi salvare le modifiche a «{docName}»?";
        string msg = mode == SaveMode.SaveAs ? "" : "Se non salvi, le modifiche andranno perse.";
        var d = new ThemedDialog(owner, title, msg, mode == SaveMode.SaveAs ? DialogIcon.Info : DialogIcon.Question);
        TextBox? nameBox = null, folderBox = null; TextBlock? err = null;
        if (needPath)
        {
            nameBox = Box(mode == SaveMode.PromptDirty ? Path.GetFileNameWithoutExtension(docName) : suggestedName, "FileName"); // stesso nome mostrato nella domanda
            d.Body.Children.Add(Label("Nome file")); d.Body.Children.Add(nameBox);
            var row = new DockPanel();
            var browse = new Button { Content = "Sfoglia…", Margin = new Thickness(8, 0, 0, 0), Template = ButtonTemplate(false), Foreground = Theme.Brush(Theme.Text), FontFamily = new FontFamily("Segoe UI"), FontSize = 13 };
            AutomationProperties.SetName(browse, "Sfoglia…"); AutomationProperties.SetAutomationId(browse, "Dlg_Browse");
            DockPanel.SetDock(browse, Dock.Right);
            folderBox = Box(folder, "Folder");
            row.Children.Add(browse); row.Children.Add(folderBox);
            d.Body.Children.Add(Label("Percorso")); d.Body.Children.Add(row);
            var nb = nameBox; var fb = folderBox;
            browse.Click += (_, _) =>
            {
                var dlg = new Microsoft.Win32.SaveFileDialog { Filter = filter + "|Tutti i file (*.*)|*.*", FileName = nb.Text, InitialDirectory = Directory.Exists(fb.Text) ? fb.Text : null, OverwritePrompt = false };
                if (dlg.ShowDialog(d) == true) { fb.Text = Path.GetDirectoryName(dlg.FileName) ?? fb.Text; nb.Text = Path.GetFileName(dlg.FileName); }
            };
            err = new TextBlock { Foreground = Theme.Brush(Color.FromRgb(0xe5, 0x5b, 0x5b)), FontFamily = new FontFamily("Segoe UI"), FontSize = 12, Margin = new Thickness(0, 8, 0, 0), TextWrapping = TextWrapping.Wrap, Visibility = Visibility.Collapsed };
            AutomationProperties.SetAutomationId(err, "Dlg_Error"); AutomationProperties.SetName(err, "Errore");
            d.Body.Children.Add(err);
            d.FocusFirst(nameBox);
        }
        SaveResult result = new(SaveLogic.Choice.Cancel, null);
        d.AddButton("Annulla", "Cancel", cancel: true);
        if (mode == SaveMode.PromptDirty) d.AddButton("Non salvare", "DontSave").Click += (_, _) => result = new(SaveLogic.Choice.DontSave, null);
        var save = d.AddButton("Salva", "Save", primary: true, close: false);
        save.Click += (_, _) =>
        {
            string? path = null;
            if (needPath)
            {
                string? e = SaveLogic.ValidateName(nameBox!.Text) ?? SaveLogic.ValidateFolder(folderBox!.Text.Trim());
                if (e != null) { err!.Text = e; err.Visibility = Visibility.Visible; return; }
                path = SaveLogic.BuildPath(folderBox!.Text.Trim(), nameBox!.Text, ext);
                if (SaveLogic.NeedsOverwriteConfirmation(path) && !Confirm(d, "Sostituire il file?", $"Esiste già un file «{Path.GetFileName(path)}» in questa cartella. Vuoi sostituirlo?", "Sostituisci", "Annulla", DialogIcon.Warning)) return;
            }
            result = new(SaveLogic.Choice.Save, path);
            d.ResultId = "Save"; d.DialogResult = true;
        };
        d.ShowDialog();
        return result;
    }

    static TextBlock Label(string t) => new() { Text = t, Margin = new Thickness(0, 14, 0, 4), FontFamily = new FontFamily("Segoe UI"), FontSize = 12, Foreground = Theme.Brush(Theme.TextDim) };

    static TextBox Box(string text, string id)
    {
        var tb = new TextBox { Text = text, FontFamily = new FontFamily("Segoe UI"), FontSize = 13, Padding = new Thickness(8, 6, 8, 6), Background = Theme.Brush(Theme.PanelAlt), Foreground = Theme.Brush(Theme.Text), CaretBrush = Theme.Brush(Theme.Text), BorderBrush = Theme.Brush(Theme.Border), MinWidth = 300 };
        AutomationProperties.SetAutomationId(tb, "Dlg_" + id); AutomationProperties.SetName(tb, id == "FileName" ? "Nome file" : "Percorso");
        return tb;
    }
}
