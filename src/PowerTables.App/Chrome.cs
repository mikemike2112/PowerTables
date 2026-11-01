using System.Globalization;
using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Shell;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using Path = System.Windows.Shapes.Path;
using PowerTables.Core;

namespace PowerTables.App;

/// <summary>Barra del titolo, scheda File, ricerca, barra di stato, zoom e sonde per i test UI.</summary>
public partial class MainWindow
{
    Button? _undoBtn, _redoBtn, _maxBtn;
    TextBlock? _zoomText;
    readonly bool _hooks = Environment.GetEnvironmentVariable("PT_TEST_HOOKS") == "1";

    Button Glyph(string icon, string tip, Action act, double size = 18)
    {
        var b = new Button { Content = Icons.Make(icon, size), ToolTip = tip, Padding = new Thickness(6, 4, 6, 4), Margin = new Thickness(1) };
        b.SetValue(AutomationProperties.NameProperty, tip);
        b.Click += (_, _) => act();
        b.SetValue(WindowChrome.IsHitTestVisibleInChromeProperty, true);
        return b;
    }

    void BuildChrome()
    {
        var logo = new Image { Source = new System.Windows.Media.Imaging.BitmapImage(new Uri("pack://application:,,,/Assets/powertables.png")), Width = 24, Height = 24, Margin = new Thickness(0, 0, 8, 0), VerticalAlignment = VerticalAlignment.Center };
        RenderOptions.SetBitmapScalingMode(logo, BitmapScalingMode.HighQuality);
        QuickAccess.Children.Add(logo);
        QuickAccess.Children.Add(Glyph("save", "Salva (Ctrl+S)", () => Save(false)));
        _undoBtn = Glyph("undo", "Annulla (Ctrl+Z)", () => Wb.Undo()); _redoBtn = Glyph("redo", "Ripeti (Ctrl+Y)", () => Wb.Redo());
        QuickAccess.Children.Add(_undoBtn); QuickAccess.Children.Add(_redoBtn);
        TitleBar.MouseLeftButtonDown += (_, e) => { if (e.ClickCount == 2) ToggleMax(); else if (e.OriginalSource is not (Button or TextBox)) { try { DragMove(); } catch (InvalidOperationException) { } } };
        SearchIcon.Content = Icons.Make("search", 16);
        SearchBox.TextChanged += (_, _) => RunSearch();
        SearchBox.PreviewKeyDown += (_, e) =>
        {
            if (e.Key == Key.Enter) { var first = SearchResults.Children.OfType<Button>().FirstOrDefault(b => b.IsEnabled); first?.RaiseEvent(new RoutedEventArgs(Button.ClickEvent)); e.Handled = true; }
            else if (e.Key == Key.Escape) { SearchBox.Clear(); SearchPopup.IsOpen = false; G.Focus(); e.Handled = true; }
        };
        // pulsanti finestra
        WinButtons.Children.Add(CaptionBtn("M5 12h14", "Riduci a icona", () => WindowState = WindowState.Minimized));
        _maxBtn = CaptionBtn("M6 6h12v12H6z", "Ingrandisci", ToggleMax); WinButtons.Children.Add(_maxBtn);
        WinButtons.Children.Add(CaptionBtn("M6 6l12 12 M18 6L6 18", "Chiudi", Close));
        // barra della formula: annulla/conferma
        FormulaButtons.Children.Add(Glyph("close", "Annulla modifica", () => { Host.Cancel(); RefreshBar(); }, 14));
        FormulaButtons.Children.Add(Glyph("func", "Inserisci funzione", () => { }, 14));
        FormulaButtons.Children[1].IsEnabled = false; ((Button)FormulaButtons.Children[1]).ToolTip = "Non ancora disponibile";
        // barra di stato: viste e zoom
        StatusRightPanel.Children.Add(StatusView("gridview", "Normale", true));
        StatusRightPanel.Children.Add(StatusView("page", "Layout di pagina", false));
        StatusRightPanel.Children.Add(StatusView("table", "Anteprima interruzioni di pagina", false));
        StatusRightPanel.Children.Add(Glyph("minus", "Riduci zoom", () => SetZoom(Host.ZoomPct - 10), 12));
        _zoomText = new TextBlock { Text = "100%", MinWidth = 40, TextAlignment = TextAlignment.Center, VerticalAlignment = VerticalAlignment.Center, FontSize = 12, FontFamily = new FontFamily("Segoe UI"), Cursor = Cursors.Hand };
        _zoomText.SetResourceReference(TextBlock.ForegroundProperty, "TextBrush");
        _zoomText.SetValue(AutomationProperties.AutomationIdProperty, "ZoomText");
        _zoomText.MouseLeftButtonUp += (_, _) => AskZoom();
        StatusRightPanel.Children.Add(_zoomText);
        StatusRightPanel.Children.Add(Glyph("plus", "Aumenta zoom", () => SetZoom(Host.ZoomPct + 10), 12));
        BuildBackstage();
        StateChanged += (_, _) => { Root.Margin = WindowState == WindowState.Maximized ? new Thickness(7) : new Thickness(0); };
        Host.Canvas.SelectionChanged += () => UpdateProbe(); Host.Canvas.ViewChanged += UpdateProbe; Loaded += (_, _) => UpdateProbe();
        Theme.Changed += UpdateProbe;
        if (_hooks) { Topmost = true; Probe.Visibility = Visibility.Visible; Probe.Opacity = 0; Probe.Width = 1; Probe.Height = 1; Probe.HorizontalAlignment = HorizontalAlignment.Left; }
    }

    Button CaptionBtn(string path, string tip, Action act)
    {
        var p = new Path { Data = Geometry.Parse(path), StrokeThickness = 1.4, Fill = Brushes.Transparent };
        p.SetResourceReference(System.Windows.Shapes.Shape.StrokeProperty, "TextBrush");
        var b = new Button { Content = new Viewbox { Width = 14, Height = 14, Child = new Canvas { Width = 24, Height = 24, Children = { p } } }, Width = 46, Height = 34, ToolTip = tip, Margin = new Thickness(0) };
        b.SetValue(AutomationProperties.NameProperty, tip);
        b.Click += (_, _) => act();
        return b;
    }

    FrameworkElement StatusView(string icon, string tip, bool on)
    {
        var b = new Button { Content = Icons.Make(icon, 14), ToolTip = on ? tip : "Non ancora disponibile", IsEnabled = on, Padding = new Thickness(8, 3, 8, 3), BorderBrush = on ? R("AccentBrush") : Brushes.Transparent };
        b.SetValue(AutomationProperties.NameProperty, tip);
        return b;
    }

    void ToggleMax() => WindowState = WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized;

    // ---------- zoom ----------
    void SetZoom(int pct)
    {
        Host.SetZoom(pct);
        if (_zoomText != null) _zoomText.Text = Host.ZoomPct + "%";
        G.RelayoutAll(); UpdateProbe();
    }

    void AskZoom()
    {
        var s = Ask("Zoom", "Percentuale (10-400):", Host.ZoomPct.ToString());
        if (s != null && int.TryParse(s.Trim().TrimEnd('%'), out var z)) SetZoom(z);
    }

    void RecalcNow()
    {
        var sw = System.Diagnostics.Stopwatch.StartNew(); Wb.RecalcAll(); G.InvalidateVisual();
        StatusRight.Text = $"Ricalcolo: {sw.ElapsedMilliseconds} ms ({Wb.LastRecalcCount:N0} formule)"; UpdateStatus();
    }

    // ---------- ricerca: comandi e testo nelle celle ----------
    void RunSearch()
    {
        string q = SearchBox.Text.Trim();
        SearchHint.Visibility = q.Length == 0 ? Visibility.Visible : Visibility.Collapsed;
        SearchResults.Children.Clear();
        if (q.Length == 0) { SearchPopup.IsOpen = false; return; }
        foreach (var (label, tab, run) in _commands.Where(c => c.Label.Contains(q, StringComparison.CurrentCultureIgnoreCase)).DistinctBy(c => c.Label).Take(8))
        {
            var r = run;
            var b = new Button { Content = $"{label}   ·   {tab}", HorizontalContentAlignment = HorizontalAlignment.Left, IsEnabled = r != null, ToolTip = r == null ? "Non ancora disponibile" : null, Padding = new Thickness(10, 5, 10, 5) };
            b.Click += (_, _) => { SearchPopup.IsOpen = false; SearchBox.Clear(); r?.Invoke(); };
            SearchResults.Children.Add(b);
        }
        var find = new Button { Content = $"Cerca \"{q}\" nel foglio", HorizontalContentAlignment = HorizontalAlignment.Left, Padding = new Thickness(10, 5, 10, 5) };
        find.Click += (_, _) => { SearchPopup.IsOpen = false; FindInSheet(q); };
        SearchResults.Children.Add(find);
        SearchPopup.IsOpen = true;
    }

    /// <summary>Seleziona la prossima cella (dopo l'attiva, con ritorno a capo) il cui testo visualizzato contiene q.</summary>
    public bool FindInSheet(string q)
    {
        var keys = Cur.CellMap.Keys.OrderBy(k => k).ToList();
        long cur = Addr.Key(G.ActiveRow, G.ActiveCol);
        foreach (var k in keys.Where(k => k > cur).Concat(keys.Where(k => k <= cur)))
            if (Wb.DisplayText(Cur, Addr.Row(k), Addr.Col(k), Ci).Contains(q, StringComparison.CurrentCultureIgnoreCase)) { G.Select(Addr.Row(k), Addr.Col(k)); return true; }
        StatusRight.Text = $"Nessun risultato per \"{q}\"";
        return false;
    }

    // ---------- scheda File ----------
    StackPanel? _bsPane;

    void BuildBackstage()
    {
        var left = new StackPanel { Width = 240, Margin = new Thickness(0, 10, 0, 0) };
        var right = new StackPanel { Margin = new Thickness(40, 30, 0, 0) };
        _bsPane = right;
        void Item(string text, string icon, Action? act, string desc)
        {
            var b = new Button { Content = new StackPanel { Orientation = Orientation.Horizontal, Children = { Icons.Make(icon, 20), new TextBlock { Text = text, Margin = new Thickness(12, 0, 0, 0), FontSize = 14, FontFamily = new FontFamily("Segoe UI"), VerticalAlignment = VerticalAlignment.Center, Foreground = R("TextBrush") } } }, HorizontalContentAlignment = HorizontalAlignment.Left, Padding = new Thickness(18, 9, 10, 9), Margin = new Thickness(6, 1, 6, 1) };
            b.SetValue(AutomationProperties.NameProperty, text);
            if (act == null) { b.IsEnabled = false; b.ToolTip = "Non ancora disponibile"; }
            else b.Click += (_, _) => { ShowBackstage(false); act(); };
            b.MouseEnter += (_, _) => { right.Children.Clear(); right.Children.Add(new TextBlock { Text = text, FontSize = 24, FontFamily = new FontFamily("Segoe UI Semibold"), Foreground = R("TextBrush") }); right.Children.Add(new TextBlock { Text = desc, Margin = new Thickness(0, 12, 0, 0), FontSize = 13, FontFamily = new FontFamily("Segoe UI"), Foreground = R("TextDimBrush"), TextWrapping = TextWrapping.Wrap, MaxWidth = 520 }); };
            left.Children.Add(b);
            _commands.Add((text, "File", act == null ? null : () => act()));
        }
        var back = new Button { Content = Icons.Make("undo", 20), Margin = new Thickness(12, 4, 0, 10), HorizontalAlignment = HorizontalAlignment.Left, ToolTip = "Indietro" };
        back.SetValue(AutomationProperties.NameProperty, "Indietro");
        back.Click += (_, _) => ShowBackstage(false);
        left.Children.Add(back);
        Item("Nuovo", "new", NewDoc, "Crea una nuova cartella di lavoro vuota.");
        Item("Apri", "open", Open, "Apre un file .xlsx o .csv.");
        Item("Salva", "save", () => Save(false), "Salva la cartella di lavoro in formato .xlsx; il file originale resta intatto se il salvataggio fallisce.");
        Item("Salva con nome", "save", () => Save(true), "Salva con un altro nome.");
        Item("Importa CSV", "import", ImportCsv, "Importa un file CSV nel foglio corrente.");
        Item("Esporta CSV", "export", ExportCsv, "Esporta il foglio corrente in CSV.");
        Item("Esporta PDF", "pdf", null, "");
        Item("Stampa", "print", null, "");
        Item("Chiudi", "close", Close, "Chiude PowerTables.");
        var dock = new DockPanel();
        DockPanel.SetDock(left, Dock.Left);
        dock.Children.Add(left); dock.Children.Add(right);
        var host = new Border { Background = R("PanelBrush"), Child = dock };
        host.SetResourceReference(Border.BackgroundProperty, "PanelBrush");
        Backstage.Children.Add(host);
    }

    void ShowBackstage(bool on)
    {
        Backstage.Visibility = on ? Visibility.Visible : Visibility.Collapsed;
        if (!on) { _tabButtons[_curTab].IsChecked = true; G.Focus(); }
    }

    /// <summary>Solo per test e screenshot (--demo-dialog tipo[:cartella]): mostra una finestra di dialogo e scrive l'esito in PT_DEMO_RESULT.</summary>
    public void DemoDialog(string spec)
    {
        string kind = spec.Split(':')[0]; string folder = spec.Contains(':') ? spec[(spec.IndexOf(':') + 1)..] : DefaultFolder();
        string result;
        switch (kind)
        {
            case "save": { var r = ThemedDialog.AskSave(this, ThemedDialog.SaveMode.PromptDirty, null, "Senza titolo.xlsx", "Senza titolo", folder, ".xlsx", XlsxFilter); result = r.Choice + "|" + r.Path; break; }
            case "savefile": { var r = ThemedDialog.AskSave(this, ThemedDialog.SaveMode.PromptDirty, System.IO.Path.Combine(folder, "bilancio.xlsx"), "bilancio.xlsx", "x", folder, ".xlsx", XlsxFilter); result = r.Choice + "|" + r.Path; break; }
            case "saveas": { var r = ThemedDialog.AskSave(this, ThemedDialog.SaveMode.SaveAs, null, "", "Senza titolo", folder, ".xlsx", XlsxFilter); result = r.Choice + "|" + r.Path; break; }
            default: ThemedDialog.Notify(this, "Il file non è un documento .xlsx valido o è danneggiato.", DialogIcon.Error, "Apertura non riuscita"); result = "ok"; break;
        }
        var path = Environment.GetEnvironmentVariable("PT_DEMO_RESULT");
        if (path != null) File.WriteAllText(path, result);
    }

    // ---------- sonda e dump per i test UI (solo con PT_TEST_HOOKS=1) ----------
    void UpdateProbe()
    {
        if (!_hooks || Cur == null) return;
        var (r1, c1, r2, c2) = G.Selection;
        Probe.SetValue(AutomationProperties.NameProperty,
            $"sel={r1},{c1},{r2},{c2};active={G.ActiveRow},{G.ActiveCol};edit={(Host.Editing ? 1 : 0)};dark={(Theme.Dark ? 1 : 0)};zoom={Host.ZoomPct};sheet={Cur.Name};bar={FormulaBar.Text};canvas={(int)G.ActualWidth}x{(int)G.ActualHeight};filter={(Cur.Filter != null ? 1 : 0)};{Origin()}filter2=1;frozen={Cur.FrozenRows},{Cur.FrozenCols};modified={(Wb.Modified ? 1 : 0)}");
        _undoBtn?.SetCurrentValue(IsEnabledProperty, Wb.CanUndo); _redoBtn?.SetCurrentValue(IsEnabledProperty, Wb.CanRedo);
    }

    string Origin()
    {
        try
        {
            if (!G.IsLoaded) return "";
            var o = G.PointToScreen(new Point(0, 0)); var dpi = VisualTreeHelper.GetDpi(G);
            return FormattableString.Invariant($"origin={o.X:0.##},{o.Y:0.##};scale={dpi.DpiScaleX:0.####};zoomscale={Host.ZoomPct / 100.0:0.###};");
        }
        catch (InvalidOperationException) { return ""; }
    }

    void DumpState()
    {
        var path = Environment.GetEnvironmentVariable("PT_TEST_DUMP");
        if (path == null || Cur == null) return;
        try
        {
            var cells = new Dictionary<string, object>();
            foreach (var kv in Cur.CellMap)
            {
                int r = Addr.Row(kv.Key), c = Addr.Col(kv.Key); if (r > 200 || c > 60) continue;
                cells[Addr.A1(r, c)] = new { text = Wb.DisplayText(Cur, r, c, Ci), f = kv.Value.F != null ? "=" + Core.Formulas.FormulaText.ToDisplay(kv.Value.F.Text) : null, kind = kv.Value.Kind.ToString() };
            }
            var hidden = Cur.Rows.Where(kv => kv.Value.Hidden).Select(kv => kv.Key).Order().ToList();
            File.WriteAllText(path, JsonSerializer.Serialize(new { sheet = Cur.Name, cells, merges = Cur.Merges.Select(m => Addr.A1(m.R1, m.C1) + ":" + Addr.A1(m.R2, m.C2)), frozen = new[] { Cur.FrozenRows, Cur.FrozenCols }, hiddenRows = hidden, dark = Theme.Dark, modified = Wb.Modified }));
        }
        catch (IOException) { }
    }
}
