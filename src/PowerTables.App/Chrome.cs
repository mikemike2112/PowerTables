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
        StateChanged += (_, _) => UpdateMaxButton();
        Loaded += (_, _) => { UpdateMaxButton(); ApplyResponsive(); };
        SizeChanged += (_, _) => ApplyResponsive();
        RibbonOuter.SizeChanged += (_, _) => LayoutRibbon();
        WinButtons.Children.Add(CaptionBtn("M6 6l12 12 M18 6L6 18", "Chiudi", Close));
        // barra della formula: annulla/conferma
        FormulaButtons.Children.Add(Glyph("close", "Annulla modifica", () => { Host.Cancel(); RefreshBar(); }, 14));
        FormulaButtons.Children.Add(Glyph("func", "Inserisci funzione", () => FunctionDialog(), 14));
        // barra di stato: viste e zoom
        StatusRightPanel.Children.Add(StatusView("gridview", "Normale", true, () => SetView(0)));
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

    static FrameworkElement GlyphOf(string path)
    {
        var p = new Path { Data = Geometry.Parse(path), StrokeThickness = 1.4, Fill = Brushes.Transparent };
        p.SetResourceReference(System.Windows.Shapes.Shape.StrokeProperty, "TextBrush");
        return new Viewbox { Width = 14, Height = 14, Child = new Canvas { Width = 24, Height = 24, Children = { p } } };
    }

    /// <summary>Un quadrato quando si può ingrandire, due quadrati sovrapposti quando si può ripristinare la finestra.</summary>
    void UpdateMaxButton()
    {
        if (_maxBtn == null) return;
        bool max = WindowState == WindowState.Maximized;
        _maxBtn.Content = GlyphOf(max ? "M8 4h12v12 M4 8h12v12H4z" : "M6 6h12v12H6z");
        string tip = max ? "Ripristina giù" : "Ingrandisci";
        _maxBtn.ToolTip = tip; AutomationProperties.SetName(_maxBtn, tip);
    }

    // ---------- adattamento alla larghezza ----------
    Button? _tabOverflow; readonly Popup _tabPop = new() { StaysOpen = false, AllowsTransparency = true, Placement = PlacementMode.Bottom };
    Button? _ribbonOverflow; readonly Popup _ribbonPop = new() { StaysOpen = false, AllowsTransparency = true, Placement = PlacementMode.Bottom };
    readonly StackPanel _ribbonPopPanel = new();

    void ApplyResponsive()
    {
        if (!IsLoaded || TabStrip.Children.Count == 0) return;
        double w = ActualWidth;
        SearchBorder.Width = Math.Clamp(w - 640, 0, 460);
        SearchBorder.Visibility = SearchBorder.Width < 110 ? Visibility.Collapsed : Visibility.Visible;
        DocTitle.MaxWidth = Math.Clamp(w - 560, 0, 420); DocTitle.TextTrimming = TextTrimming.CharacterEllipsis;
        DocTitle.Visibility = DocTitle.MaxWidth < 60 ? Visibility.Collapsed : Visibility.Visible;
        // schede: quelle che non entrano finiscono nel menu "⋯"
        if (_tabOverflow == null)
        {
            _tabOverflow = new Button { Content = "⋯", ToolTip = "Altre schede", Padding = new Thickness(10, 4, 10, 4), VerticalAlignment = VerticalAlignment.Center };
            AutomationProperties.SetName(_tabOverflow, "Altre schede"); AutomationProperties.SetAutomationId(_tabOverflow, "TabOverflow");
            _tabOverflow.Click += (_, _) => ShowTabOverflow();
            TabStrip.Children.Add(_tabOverflow);
        }
        var tabs = TabStrip.Children.OfType<ToggleButton>().ToList();
        foreach (var t in tabs) t.Visibility = Visibility.Visible;
        _tabOverflow.Visibility = Visibility.Collapsed;
        double avail = w - 30 - 100;
        double acc = 0; int firstHidden = -1;
        for (int i = 0; i < tabs.Count; i++) { tabs[i].Measure(new Size(double.PositiveInfinity, 40)); acc += tabs[i].DesiredSize.Width; if (acc > avail - 46 && firstHidden < 0) firstHidden = i; }
        if (acc > avail && firstHidden >= 0)
        {
            for (int i = firstHidden; i < tabs.Count; i++) tabs[i].Visibility = Visibility.Collapsed;
            _tabOverflow.Visibility = Visibility.Visible;
        }
        LayoutRibbon();
    }

    void ShowTabOverflow()
    {
        var sp = new StackPanel { Margin = new Thickness(2) };
        foreach (var t in TabStrip.Children.OfType<ToggleButton>().Where(t => t.Visibility == Visibility.Collapsed))
        {
            string n = (string)t.Content;
            var b = new Button { Content = n, MinWidth = 180, HorizontalContentAlignment = HorizontalAlignment.Left };
            AutomationProperties.SetName(b, "Scheda " + n);
            b.Click += (_, _) => { _tabPop.IsOpen = false; if (n == "File") ShowBackstage(true); else ShowTab(n); };
            sp.Children.Add(b);
        }
        _tabPop.Child = new Border { Background = R("PanelBrush"), BorderBrush = R("AccentBrush"), BorderThickness = new Thickness(1), Child = sp };
        _tabPop.PlacementTarget = _tabOverflow; _tabPop.IsOpen = true;
    }

    /// <summary>Ribbon: prima i gruppi si compattano (solo icone), poi i gruppi in più vanno in un menu a discesa.</summary>
    bool _ribbonCompact, _ribbonOverflowOn; double _ribbonFull, _ribbonAvail;   // stato dell'ultimo adattamento (letto dalla sonda dei test)
    void LayoutRibbon()
    {
        if (!IsLoaded || !_tabGroups.TryGetValue(_curTab, out var groups) || RibbonHost.Content is not ScrollViewer sv || sv.Content is not StackPanel sp) return;
        _ribbonPop.IsOpen = false;
        _ribbonPopPanel.Children.Clear();
        sp.Children.Clear();
        foreach (var g in groups) { sp.Children.Add(g); SetLabels(g, true); }
        double avail = Math.Max(120, RibbonOuter.ActualWidth - 16);
        _ribbonCompact = false; _ribbonOverflowOn = false;
        // le misure vanno ricalcolate davvero: WPF non invalida gli antenati quando cambia la visibilità di un discendente
        double Width() { sp.InvalidateMeasure(); sp.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity)); return sp.DesiredSize.Width; }
        _ribbonAvail = avail; _ribbonFull = Width();
        if (_ribbonFull <= avail) return;
        _ribbonCompact = true;
        for (int i = groups.Count - 1; i >= 0 && Width() > avail; i--) SetLabels(groups[i], false);
        if (Width() <= avail) return;
        _ribbonOverflowOn = true;
        _ribbonOverflow ??= MakeRibbonOverflow();
        sp.Children.Add(_ribbonOverflow);
        for (int i = groups.Count - 1; i > 0 && Width() > avail; i--)
        {
            sp.Children.Remove(groups[i]); SetLabels(groups[i], true); _ribbonPopPanel.Children.Insert(0, groups[i]);
        }
    }

    Button MakeRibbonOverflow()
    {
        var b = new Button { Content = "»", ToolTip = "Altri comandi", Padding = new Thickness(12, 4, 12, 4), VerticalAlignment = VerticalAlignment.Center, FontSize = 16 };
        AutomationProperties.SetName(b, "Altri comandi"); AutomationProperties.SetAutomationId(b, "RibbonOverflow");
        _ribbonPop.Child = new Border { Background = R("PanelBrush"), BorderBrush = R("AccentBrush"), BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(6), Child = new ScrollViewer { Content = _ribbonPopPanel, MaxHeight = 420, VerticalScrollBarVisibility = ScrollBarVisibility.Auto } };
        b.Click += (_, _) => { _ribbonPop.PlacementTarget = b; _ribbonPop.IsOpen = true; };
        return b;
    }

    void SetLabels(Border g, bool visible)
    {
        if (!_groupLabels.TryGetValue(g, out var l)) return;
        foreach (var t in l)
        {
            t.Visibility = visible ? Visibility.Visible : Visibility.Collapsed;
            DependencyObject? p = t;
            while (p != null && p is not ScrollViewer) { (p as UIElement)?.InvalidateMeasure(); p = VisualTreeHelper.GetParent(p); }
        }
    }

    Button CaptionBtn(string path, string tip, Action act)
    {
        var b = new Button { Content = GlyphOf(path), Width = 46, Height = 34, ToolTip = tip, Margin = new Thickness(0) };
        b.SetValue(AutomationProperties.NameProperty, tip);
        b.Click += (_, _) => act();
        return b;
    }

    FrameworkElement StatusView(string icon, string tip, bool on, Action? click = null)
    {
        var b = new Button { Content = Icons.Make(icon, 14), ToolTip = on ? tip : "Non ancora disponibile", IsEnabled = on, Padding = new Thickness(8, 3, 8, 3), BorderBrush = on ? R("AccentBrush") : Brushes.Transparent };
        b.SetValue(AutomationProperties.NameProperty, tip);
        if (click != null) b.Click += (_, _) => { click(); G.Focus(); };
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
        Item("Esporta PDF", "pdf", ExportPdf, "Esporta il foglio corrente in PDF usando area di stampa, orientamento, formato carta e margini.");
        Item("Stampa", "print", PrintSheet, "Apre il foglio corrente come PDF nel programma predefinito, da cui stampare.");
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
            $"sel={r1},{c1},{r2},{c2};active={G.ActiveRow},{G.ActiveCol};edit={(Host.Editing ? 1 : 0)};dark={(Theme.Dark ? 1 : 0)};zoom={Host.ZoomPct};sheet={Cur.Name};bar={FormulaBar.Text};canvas={(int)G.ActualWidth}x{(int)G.ActualHeight};filter={(Cur.Filter != null ? 1 : 0)};{Origin()}filter2=1;frozen={Cur.FrozenRows},{Cur.FrozenCols};modified={(Wb.Modified ? 1 : 0)};rb={(int)RibbonOuter.ActualWidth};win={(int)ActualWidth};rbcompact={(_ribbonCompact ? 1 : 0)};rbovf={(_ribbonOverflowOn ? 1 : 0)};rbfull={(int)_ribbonFull};rbavail={(int)_ribbonAvail}");
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
