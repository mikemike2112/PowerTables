using System.Windows;
using System.Windows.Automation;
using System.Windows.Shell;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using PowerTables.Core;
using PowerTables.Core.Formulas;

namespace PowerTables.App;

/// <summary>Costruzione del ribbon: struttura e posizioni come l'app di riferimento, icone e codice propri.</summary>
public partial class MainWindow
{
    readonly List<(string Label, string Tab, Action? Run)> _commands = [];
    string _curTab = "Home";
    readonly Dictionary<string, FrameworkElement> _tabContent = [];
    readonly Dictionary<string, ToggleButton> _tabButtons = [];

    static readonly string[] TabNames = ["File", "Home", "Inserisci", "Layout di pagina", "Formule", "Dati", "Visualizza", "Automatizza", "Guida", "Powers"];

    Brush R(string k) => (Brush)FindResource(k);

    // ---------- elementi base ----------
    FrameworkElement Large(string label, string icon, Action? act, bool drop = false, string? tip = null)
    {
        var sp = new StackPanel { Orientation = Orientation.Vertical, HorizontalAlignment = HorizontalAlignment.Center };
        sp.Children.Add(Icons.Make(icon, 32));
        var tb = new TextBlock { Text = label + (drop ? " ▾" : ""), TextAlignment = TextAlignment.Center, TextWrapping = TextWrapping.Wrap, MaxWidth = 78, FontSize = 12, FontFamily = new FontFamily("Segoe UI"), Margin = new Thickness(0, 2, 0, 0) };
        tb.SetResourceReference(TextBlock.ForegroundProperty, "TextBrush");
        sp.Children.Add(tb); _pendingLabels.Add(tb);
        return Wrap(label, sp, act, tip, new Thickness(4, 4, 4, 2));
    }

    FrameworkElement LargeMenu(string label, string icon, Action<Button> open, string? tip = null) { Button? b = null; b = (Button)Large(label, icon, () => open(b!), true, tip); return b; }

    FrameworkElement Small(string label, string? icon, Action? act, bool drop = false, string? tip = null, double minW = 0)
    {
        var sp = new StackPanel { Orientation = Orientation.Horizontal };
        if (icon != null) sp.Children.Add(Icons.Make(icon, 16));
        if (label.Length > 0)
        {
            var tb = new TextBlock { Text = label + (drop ? " ▾" : ""), FontSize = 12, FontFamily = new FontFamily("Segoe UI"), Margin = new Thickness(icon != null ? 6 : 0, 0, 0, 0), VerticalAlignment = VerticalAlignment.Center };
            tb.SetResourceReference(TextBlock.ForegroundProperty, "TextBrush");
            sp.Children.Add(tb); if (icon != null) _pendingLabels.Add(tb);
        }
        else if (drop) sp.Children.Add(new TextBlock { Text = "▾", FontSize = 9, Margin = new Thickness(2, 0, 0, 0), VerticalAlignment = VerticalAlignment.Center, Foreground = R("TextDimBrush") });
        var el = Wrap(label.Length > 0 ? label : (tip ?? icon ?? ""), sp, act, tip, new Thickness(5, 2, 5, 2));
        if (minW > 0 && el is Button b) { b.MinWidth = minW; b.HorizontalContentAlignment = HorizontalAlignment.Left; }
        return el;
    }

    FrameworkElement Wrap(string label, UIElement content, Action? act, string? tip, Thickness pad)
    {
        var b = new Button { Content = content, Padding = pad, Margin = new Thickness(1), VerticalContentAlignment = VerticalAlignment.Center };
        b.SetValue(AutomationProperties.NameProperty, label);
        if (act == null) { b.IsEnabled = false; b.ToolTip = "Non ancora disponibile"; _commands.Add((label, _curBuildTab, null)); }
        else
        {
            b.ToolTip = tip ?? label;
            Action run = () => { try { act(); } catch (Exception ex) when (ex is FormatException or ArgumentException or InvalidOperationException) { ThemedDialog.Notify(this, ex.Message, DialogIcon.Warning); } G.Focus(); };
            b.Click += (_, _) => run();
            _commands.Add((label, _curBuildTab, run));
        }
        return b;
    }
    string _curBuildTab = "";

    FrameworkElement Check(string label, bool isChecked, Action<bool>? set)
    {
        var cb = new CheckBox { Content = label, IsChecked = isChecked, Margin = new Thickness(6, 2, 6, 2), FontSize = 12, FontFamily = new FontFamily("Segoe UI"), VerticalContentAlignment = VerticalAlignment.Center };
        cb.SetResourceReference(Control.ForegroundProperty, "TextBrush");
        cb.SetValue(AutomationProperties.NameProperty, label);
        if (set == null) { cb.IsEnabled = false; cb.ToolTip = "Non ancora disponibile"; _commands.Add((label, _curBuildTab, null)); }
        else
        {
            cb.Click += (_, _) => { set(cb.IsChecked == true); G.Focus(); };
            _commands.Add((label, _curBuildTab, () => { cb.IsChecked = cb.IsChecked != true; set(cb.IsChecked == true); }));
        }
        return cb;
    }

    FrameworkElement Combo(string text, double w, Action<Button>? open, string tip, out TextBlock label)
    {
        var tb = new TextBlock { Text = text, VerticalAlignment = VerticalAlignment.Center, FontSize = 12, FontFamily = new FontFamily("Segoe UI"), TextTrimming = TextTrimming.CharacterEllipsis };
        tb.SetResourceReference(TextBlock.ForegroundProperty, "TextBrush");
        var arrow = new TextBlock { Text = "▾", FontSize = 9, VerticalAlignment = VerticalAlignment.Center, Foreground = R("TextDimBrush") };
        DockPanel.SetDock(arrow, Dock.Right);
        var dp = new DockPanel { Width = w - 14, LastChildFill = true, Children = { arrow, tb } };
        var b = new Button { Content = dp, ToolTip = tip, Padding = new Thickness(6, 2, 6, 2), Margin = new Thickness(2, 1, 2, 1), HorizontalContentAlignment = HorizontalAlignment.Stretch, BorderBrush = R("BorderBrush2"), Width = w };
        b.SetValue(AutomationProperties.NameProperty, tip);
        if (open == null) { b.IsEnabled = false; b.ToolTip = "Non ancora disponibile"; }
        else b.Click += (_, _) => open(b);
        label = tb;
        _commands.Add((tip, _curBuildTab, open == null ? null : () => open(b)));
        return b;
    }

    static UIElement Col(params UIElement[] rows) => new StackPanel { Orientation = Orientation.Vertical, VerticalAlignment = VerticalAlignment.Top, Children = { } }.AddAll(rows);
    static UIElement RowOf(params UIElement[] items) => new StackPanel { Orientation = Orientation.Horizontal }.AddAll(items);

    UIElement Group(string name, params UIElement[] cols)
    {
        var content = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Top, Margin = new Thickness(2, 4, 2, 0) }.AddAll(cols);
        var lab = new TextBlock { Text = name, FontSize = 11, FontFamily = new FontFamily("Segoe UI"), HorizontalAlignment = HorizontalAlignment.Center, Margin = new Thickness(0, 2, 0, 3) };
        lab.SetResourceReference(TextBlock.ForegroundProperty, "TextDimBrush");
        var dock = new DockPanel();
        DockPanel.SetDock(lab, Dock.Bottom);
        dock.Children.Add(lab); dock.Children.Add(content);
        var b = new Border { BorderThickness = new Thickness(0, 0, 1, 0), Padding = new Thickness(2, 0, 2, 0), Child = dock };
        b.SetResourceReference(Border.BorderBrushProperty, "BorderBrush2");
        _groupLabels[b] = [.. _pendingLabels]; _pendingLabels.Clear();
        return b;
    }

    readonly List<TextBlock> _pendingLabels = [];
    readonly Dictionary<Border, List<TextBlock>> _groupLabels = [];
    readonly Dictionary<string, List<Border>> _tabGroups = [];

    FrameworkElement Tab(string name, params UIElement[] groups)
    {
        var sp = new StackPanel { Orientation = Orientation.Horizontal }.AddAll(groups);
        _tabGroups[name] = groups.OfType<Border>().ToList();
        var sv = new ScrollViewer { Content = sp, HorizontalScrollBarVisibility = ScrollBarVisibility.Hidden, VerticalScrollBarVisibility = ScrollBarVisibility.Disabled };
        sv.PreviewMouseWheel += (_, e) => { sv.ScrollToHorizontalOffset(sv.HorizontalOffset - e.Delta); e.Handled = true; };
        return sv;
    }

    // ---------- schede ----------
    void BuildTabs()
    {
        foreach (var n in TabNames)
        {
            var name = n;
            var tb = new ToggleButton { Content = name, Style = (Style)FindResource(name == "Powers" ? "PowersTabBtn" : "TabBtn"), IsChecked = name == "Home" };
            tb.SetValue(AutomationProperties.NameProperty, "Scheda " + name);
            tb.SetValue(AutomationProperties.AutomationIdProperty, "Tab_" + name.Replace(' ', '_'));
            tb.Click += (_, _) => { if (name == "File") { ShowBackstage(true); tb.IsChecked = false; } else ShowTab(name); };
            TabStrip.Children.Add(tb); _tabButtons[name] = tb;
        }
        _curBuildTab = "Home"; _tabContent["Home"] = BuildHome();
        _curBuildTab = "Inserisci"; _tabContent["Inserisci"] = BuildInsert();
        _curBuildTab = "Layout di pagina"; _tabContent["Layout di pagina"] = BuildLayout();
        _curBuildTab = "Formule"; _tabContent["Formule"] = BuildFormulas();
        _curBuildTab = "Dati"; _tabContent["Dati"] = BuildData();
        _curBuildTab = "Visualizza"; _tabContent["Visualizza"] = BuildView();
        _curBuildTab = "Automatizza"; _tabContent["Automatizza"] = Tab("Automatizza", Group("Automatizza", Large("Nuovo script", "func", null), Large("Visualizza script", "page", null)));
        _curBuildTab = "Guida"; _tabContent["Guida"] = Tab("Guida", Group("Guida", Large("Informazioni su PowerTables", "page", () => ThemedDialog.Notify(this, "PowerTables - foglio di calcolo open source (MIT).\nFormato nativo .xlsx.", DialogIcon.Info, "Informazioni"))));
        _curBuildTab = "Powers"; _tabContent["Powers"] = BuildPowers();
        ShowTab("Home");
    }

    void ShowTab(string name)
    {
        _curTab = name;
        RibbonHost.Content = _tabContent[name];
        foreach (var kv in _tabButtons) kv.Value.IsChecked = kv.Key == name;
        LayoutRibbon();
        // bordo del ribbon: #FF2400 con la scheda Powers
        RibbonOuter.BorderBrush = name == "Powers" ? R("PowersBrush") : R("BorderBrush2");
        RibbonOuter.BorderThickness = name == "Powers" ? new Thickness(2) : new Thickness(1);
    }

    // ---------- Home ----------
    FrameworkElement BuildHome()
    {
        var fontLabel = default(TextBlock); var sizeLabel = default(TextBlock);
        var paste = LargeMenu("Incolla", "paste", b => DropMenu(b, ("Incolla", () => DoPaste()), ("Incolla valori", () => DoPaste(1)), ("Incolla formati", () => DoPaste(2))), "Incolla (Ctrl+V)");
        var clip = Group("Appunti", paste, Col(Small("Taglia", "cut", () => DoCopy(true), tip: "Taglia (Ctrl+X)"), Small("Copia", "copy", () => DoCopy(false), tip: "Copia (Ctrl+C)"), Small("Copia formato", "fill", CopyFormat)));
        _fontBox = Combo("Calibri", 118, b => DropMenu(b, new[] { "Calibri", "Arial", "Segoe UI", "Times New Roman", "Courier New", "Verdana", "Georgia", "Consolas" }.Select(n => (n, (Action)(() => Fmt(x => x with { Font = x.Font with { Name = n } })))).ToArray()), "Tipo di carattere", out fontLabel);
        _fontLabel = fontLabel;
        _sizeBox = Combo("11", 46, b => DropMenu(b, new[] { 8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72 }.Select(n => (n.ToString(), (Action)(() => Fmt(x => x with { Font = x.Font with { Size = n } })))).ToArray()), "Dimensione carattere", out sizeLabel);
        _sizeLabel = sizeLabel;
        _bold = ToggleIcon("G", "Grassetto (Ctrl+B)", FontWeights.Bold, null, false, () => Toggle(x => x.Font.Bold, (x, v) => x with { Font = x.Font with { Bold = v } }));
        _italic = ToggleIcon("C", "Corsivo (Ctrl+I)", FontWeights.Normal, FontStyles.Italic, false, () => Toggle(x => x.Font.Italic, (x, v) => x with { Font = x.Font with { Italic = v } }));
        _underline = ToggleIcon("S", "Sottolineato (Ctrl+U)", FontWeights.Normal, null, true, () => Toggle(x => x.Font.Underline, (x, v) => x with { Font = x.Font with { Underline = v } }));
        Button? bordB = null, fillB = null, colB = null;
        bordB = (Button)Small("", "border", () => DropMenu(bordB!, ("Tutti i bordi", () => SetBorders("all")), ("Bordo esterno", () => SetBorders("outer")), ("Bordo inferiore", () => SetBorders("bottom")), ("Nessun bordo", () => SetBorders("none"))), true, "Bordi");
        fillB = (Button)Small("", "fill", () => Palette(fillB!, c => Fmt(x => x with { FillColor = c })), true, "Colore di riempimento");
        colB = (Button)Small("A", null, () => Palette(colB!, c => Fmt(x => x with { Font = x.Font with { Color = c } })), true, "Colore carattere");
        var font = Group("Carattere", Col(
            RowOf(_fontBox, _sizeBox, Small("A+", null, () => Fmt(x => x with { Font = x.Font with { Size = Math.Min(409, x.Font.Size + 1) } }), tip: "Aumenta dimensione carattere"), Small("A−", null, () => Fmt(x => x with { Font = x.Font with { Size = Math.Max(1, x.Font.Size - 1) } }), tip: "Riduci dimensione carattere")),
            RowOf(_bold, _italic, _underline, bordB, fillB, colB)));
        var align = Group("Allineamento", Col(
            RowOf(Small("", "top", () => Fmt(x => x with { VAlign = "top" }), tip: "Allinea in alto"), Small("", "middle", () => Fmt(x => x with { VAlign = "center" }), tip: "Allinea al centro"), Small("", "bottom", () => Fmt(x => x with { VAlign = "bottom" }), tip: "Allinea in basso"), Small("", "wrap", () => Fmt(x => x with { Wrap = !x.Wrap }), tip: "Testo a capo")),
            RowOf(Small("", "left", () => Fmt(x => x with { HAlign = "left" }), tip: "Allinea a sinistra"), Small("", "center", () => Fmt(x => x with { HAlign = "center" }), tip: "Centra"), Small("", "right", () => Fmt(x => x with { HAlign = "right" }), tip: "Allinea a destra"),
                Small("", "outdent", () => Fmt(x => x with { Indent = Math.Max(0, x.Indent - 1) }), tip: "Riduci rientro"), Small("", "indent", () => Fmt(x => x with { Indent = Math.Min(15, x.Indent + 1) }), tip: "Aumenta rientro"))),
            Col(Small("Unisci e centra", "merge", () => { var (r1, c1, r2, c2) = G.Selection; Wb.Merge(Cur, r1, c1, r2, c2); Fmt(x => x with { HAlign = "center" }); }, tip: "Unisci e allinea al centro"), Small("Dividi celle", "unmerge", () => { var (r1, c1, r2, c2) = G.Selection; Wb.Unmerge(Cur, r1, c1, r2, c2); })));
        Button? nfB = null;
        nfB = (Button)Combo("Generale", 120, b => DropMenu(b,
            ("Generale", () => NumFormat(0)), ("Numero", () => NumFormat(2)), ("Migliaia", () => NumFormat(4)), ("Valuta", () => NumFormat(0, "#,##0.00\\ \"€\"")), ("Percentuale", () => NumFormat(10)),
            ("Data breve", () => NumFormat(14)), ("Data estesa", () => NumFormat(0, "dddd d mmmm yyyy")), ("Ora", () => NumFormat(21)), ("Scientifico", () => NumFormat(11)), ("Testo", () => NumFormat(49))), "Formato numero", out _nfLabel);
        var numbers = Group("Numeri", Col(nfB, RowOf(Small("€", null, () => NumFormat(0, "#,##0.00\\ \"€\""), tip: "Formato valuta"), Small("%", null, () => NumFormat(9), tip: "Stile percentuale"), Small("000", null, () => NumFormat(4), tip: "Stile separatore"), Small(".00+", null, () => Decimals(1), tip: "Aumenta decimali"), Small(".0−", null, () => Decimals(-1), tip: "Diminuisci decimali"))));
        var styles = Group("Stili", LargeMenu("Formattazione condizionale", "table", CondFormatMenu), LargeMenu("Formatta come tabella", "table", TableStyleMenu), LargeMenu("Stili cella", "table", CellStyleMenu));
        Button? insB = null, delB = null, fmtB = null;
        fmtB = (Button)Small("Formato", "table", () => DropMenu(fmtB!, ("Formato celle...", FormatCellsDialog)), true);
        insB = (Button)Small("Inserisci", "rowplus", () => DropMenu(insB!, ("Inserisci righe", () => EditLines(true, true)), ("Inserisci colonne", () => EditLines(false, true))), true);
        delB = (Button)Small("Elimina", "delete", () => DropMenu(delB!, ("Elimina righe", () => EditLines(true, false)), ("Elimina colonne", () => EditLines(false, false))), true);
        AutomationProperties.SetAutomationId(insB, "Cells_Insert"); AutomationProperties.SetAutomationId(delB, "Cells_Delete");   // "Inserisci" è anche il nome di una scheda
        var cells = Group("Celle", Col(insB, delB, fmtB));
        Button? sortB = null;
        sortB = (Button)Large("Ordina e filtra", "sort", () => DropMenu(sortB!, ("Ordina dalla A alla Z", () => DoSort(true)), ("Ordina dalla Z alla A", () => DoSort(false)), ("Filtro", ToggleFilter)), true);
        var edit = Group("Modifica", Col(Small("Somma automatica", "sigma", AutoSum, tip: "Somma automatica"), Small("Riempi giù", "fill", () => FillDirection(true), tip: "Riempi in basso (Ctrl+D)"), Small("Cancella", "clear", () => G.ClearSelection(), tip: "Cancella contenuto (Canc)")), sortB, LargeMenu("Trova e seleziona", "find", b => DropMenu(b, ("Trova...", () => ShowFind(false)), ("Sostituisci...", () => ShowFind(true)), ("Vai a...", GoToDialog))));
        return Tab("Home", clip, font, align, numbers, styles, cells, edit);
    }

    TextBlock? _fontLabel, _sizeLabel, _nfLabel; FrameworkElement? _fontBox, _sizeBox;

    ToggleButton ToggleIcon(string text, string tip, FontWeight w, FontStyle? st, bool ul, Action act)
    {
        var t = new ToggleButton { Content = new TextBlock { Text = text, FontWeight = w, FontStyle = st ?? FontStyles.Normal, TextDecorations = ul ? TextDecorations.Underline : null, FontSize = 14, FontFamily = new FontFamily("Segoe UI") }, ToolTip = tip, MinWidth = 26 };
        ((TextBlock)t.Content).SetResourceReference(TextBlock.ForegroundProperty, "TextBrush");
        t.SetValue(AutomationProperties.NameProperty, tip);
        t.Click += (_, _) => { act(); G.Focus(); };
        _commands.Add((tip, _curBuildTab, () => { act(); G.Focus(); }));
        return t;
    }

    // ---------- Inserisci / Layout / Formule / Dati / Visualizza ----------
    FrameworkElement BuildInsert() => Tab("Inserisci",
        Group("Tabelle", Large("Tabella pivot", "table", null), Large("Tabella", "table", () => TableDialog())),
        Group("Illustrazioni", Large("Immagini", "image", InsertPicture), LargeMenu("Forme", "page", ShapesMenu)),
        Group("Grafici", LargeMenu("Grafici consigliati", "chart", RecommendedCharts), Large("Istogramma", "chart", () => InsertChart("col")), Large("Linee", "chart", () => InsertChart("line"))),
        Group("Collegamenti", Large("Collegamento", "page", LinkDialog)),
        Group("Commenti", Large("Commento", "comment", CommentDialog)),
        Group("Testo", Large("Casella di testo", "page", () => ShapeDialog("rect", true))));

    FrameworkElement BuildLayout() => Tab("Layout di pagina",
        Group("Temi", Large("Temi", "theme", null, true)),
        Group("Imposta pagina", LargeMenu("Margini", "page", MarginsMenu), LargeMenu("Orientamento", "page", OrientationMenu), LargeMenu("Dimensioni", "page", PaperMenu), LargeMenu("Area di stampa", "print", PrintAreaMenu)),
        Group("Opzioni del foglio", Col(Check("Griglia: visualizza", Cur?.ShowGridLines ?? true, v => { Cur!.ShowGridLines = v; Cur.LayoutDirty = true; G.InvalidateVisual(); }), Check("Griglia: stampa", false, v => Wb.SetPrintGridLines(Cur, v)))));

    FrameworkElement BuildFormulas()
    {
        FrameworkElement Cat(string cat, string label)
        {
            Button? b = null;
            b = (Button)Large(label, "func", () => DropFunctions(b!, cat), true);
            return b;
        }
        return Tab("Formule",
            Group("Libreria di funzioni", Large("Inserisci funzione", "func", () => FunctionDialog()), Large("Somma automatica", "sigma", AutoSum), Cat("Finanziarie", "Finanziarie"), Cat("Logiche", "Logiche"), Cat("Testo", "Testo"), Cat("Data e ora", "Data e ora"), Cat("Ricerca e riferimento", "Ricerca e riferimento"), Cat("Matematiche", "Matematiche"), Cat("Statistiche", "Statistiche")),
            Group("Nomi definiti", Large("Gestione nomi", "names", NameManagerDialog), Col(Small("Definisci nome", "names", DefineNameDialog))),
            Group("Verifica formule", Col(Small("Mostra formule", "func", () => { Wb.ShowFormulas = !Wb.ShowFormulas; G.InvalidateVisual(); }))),
            Group("Calcolo", Large("Calcola adesso", "calc", RecalcNow)));
    }

    /// <summary>Menu a discesa con le funzioni di una categoria: la scelta apre la modifica con =NOME( nella cella attiva.</summary>
    void DropFunctions(Button anchor, string category)
    {
        var items = FunctionCatalog.All.Where(f => f.Category == category).Select(f => (f.It != f.En ? $"{f.En}  ·  {f.It}" : f.En, (Action)(() => InsertFunction(f.En)))).ToArray();
        DropMenu(anchor, items);
    }

    /// <summary>Inserisce la funzione nella cella in modifica oppure avvia la modifica con "=NOME(".</summary>
    void InsertFunction(string name)
    {
        if (Host.Editing) Host.BeginEditWith(name + "(");
        else Host.BeginEdit(null, "=" + name + "(");
        if (Host.Editing) Host.Editor.Focus();
    }

    /// <summary>Finestra "Inserisci funzione": ricerca per nome (inglese o italiano) o categoria.</summary>
    void FunctionDialog()
    {
        var d = new ThemedDialog(this, "Inserisci funzione", "Cerca per nome o scegli una categoria.", DialogIcon.Info);
        var search = new TextBox { Margin = new Thickness(0, 12, 0, 8), Padding = new Thickness(8, 6, 8, 6), FontSize = 13, FontFamily = new FontFamily("Segoe UI"), MinWidth = 420 };
        search.SetResourceReference(Control.BackgroundProperty, "PanelAltBrush"); search.SetResourceReference(Control.ForegroundProperty, "TextBrush"); search.SetResourceReference(Control.BorderBrushProperty, "BorderBrush2");
        AutomationProperties.SetAutomationId(search, "FuncSearch"); AutomationProperties.SetName(search, "Cerca funzione");
        var list = new ListBox { Height = 220, FontFamily = new FontFamily("Segoe UI"), FontSize = 13 };
        list.SetResourceReference(Control.BackgroundProperty, "PanelAltBrush"); list.SetResourceReference(Control.ForegroundProperty, "TextBrush"); list.SetResourceReference(Control.BorderBrushProperty, "BorderBrush2");
        AutomationProperties.SetAutomationId(list, "FuncList");
        List<FunctionCatalog.Entry> shown = [];
        void Fill()
        {
            string q = search.Text.Trim();
            shown = FunctionCatalog.All.Where(f => q.Length == 0 || f.En.Contains(q, StringComparison.OrdinalIgnoreCase) || f.It.Contains(q, StringComparison.OrdinalIgnoreCase) || f.Category.Contains(q, StringComparison.OrdinalIgnoreCase)).OrderBy(f => f.Category).ThenBy(f => f.En).ToList();
            list.ItemsSource = shown.Select(f => $"{f.En}{(f.It != f.En ? "  ·  " + f.It : "")}    —    {f.Category}").ToList();
            if (shown.Count > 0) list.SelectedIndex = 0;
        }
        search.TextChanged += (_, _) => Fill();
        d.Body.Children.Add(search); d.Body.Children.Add(list);
        d.AddButton("Annulla", "Cancel", cancel: true);
        var ok = d.AddButton("Inserisci", "Insert", primary: true, close: false);
        ok.Click += (_, _) => { if (list.SelectedIndex >= 0) { d.ResultId2 = shown[list.SelectedIndex].En; d.DialogResult = true; } };
        list.MouseDoubleClick += (_, _) => ok.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        d.FocusFirst(search); Fill();
        d.ShowDialog();
        if (d.ResultId2 != null) InsertFunction(d.ResultId2);
    }

    FrameworkElement BuildData() => Tab("Dati",
        Group("Recupera e trasforma dati", Large("Da testo/CSV", "csv", ImportCsv)),
        Group("Ordina e filtra", Col(Small("Ordina dalla A alla Z", "sortaz", () => DoSort(true)), Small("Ordina dalla Z alla A", "sortza", () => DoSort(false))), Large("Filtro", "filter", ToggleFilter), Col(Small("Cancella", "clear", () => Wb.ClearFilterCriteria(Cur, Ci), tip: "Cancella i criteri del filtro (restano le frecce)"), Small("Riapplica", "filter", () => { Wb.ReapplyFilter(Cur, Ci); G.RelayoutAll(); }), Small("Avanzate", "filter", null))),
        Group("Strumenti dati", Large("Testo in colonne", "table", TextToColumnsDialog), Large("Rimuovi duplicati", "table", RemoveDuplicatesDialog), Large("Convalida dati", "table", DataValidationDialog)),
        Group("Previsione", Large("Analisi di simulazione", "chart", null, true), Large("Foglio previsione", "chart", null)));

    FrameworkElement BuildView()
    {
        Button? frz = null;
        frz = (Button)Small("Blocca riquadri", "freeze", () => DropMenu(frz!, ("Blocca riga superiore", () => Wb.SetFreeze(Cur, 1, Cur.FrozenCols)), ("Blocca prima colonna", () => Wb.SetFreeze(Cur, Cur.FrozenRows, 1)), ("Blocca riquadri fino alla selezione", () => Wb.SetFreeze(Cur, G.ActiveRow - 1, G.ActiveCol - 1)), ("Sblocca riquadri", () => Wb.SetFreeze(Cur, 0, 0))), true);
        return Tab("Visualizza",
            Group("Visualizzazioni cartella di lavoro", Large("Normale", "gridview", () => SetView(0)), Large("Anteprima interruzioni di pagina", "page", null), Small("Layout di pagina", "page", null)),
            Group("Mostra", Col(Check("Griglia", true, v => { Cur.ShowGridLines = v; Cur.LayoutDirty = true; G.InvalidateVisual(); }), Check("Barra della formula", true, v => FormulaRow.Visibility = v ? Visibility.Visible : Visibility.Collapsed), Check("Intestazioni", true, null))),
            Group("Zoom", Large("Zoom", "zoom", () => AskZoom()), Large("100%", "zoom", () => SetZoom(100))),
            Group("Finestra", Col(Small("Nuova finestra", "window", null), frz)));
    }

    FrameworkElement BuildPowers()
    {
        var sw = new ToggleButton { Style = (Style)FindResource("SwitchToggle"), IsChecked = Theme.Dark, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(8, 4, 8, 4), ToolTip = "Modalità scura" };
        sw.SetValue(AutomationProperties.NameProperty, "Modalità scura");
        sw.SetValue(AutomationProperties.AutomationIdProperty, "DarkModeToggle");
        sw.Click += (_, _) => SetDark(sw.IsChecked == true);
        _darkToggle = sw;
        var lab = new TextBlock { Text = "Modalità scura", FontSize = 12, FontFamily = new FontFamily("Segoe UI"), VerticalAlignment = VerticalAlignment.Center };
        lab.SetResourceReference(TextBlock.ForegroundProperty, "TextBrush");
        _commands.Add(("Modalità scura", "Powers", () => { sw.IsChecked = !sw.IsChecked; SetDark(sw.IsChecked == true); }));
        _pendingLabels.Clear();
        var row = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(6, 20, 6, 0) };
        row.Children.Add(Icons.Make("moon", 22)); row.Children.Add(new Border { Width = 8 }); row.Children.Add(lab); row.Children.Add(sw);
        var aspetto = Group("Aspetto", row);
        _curBuildTab = "Powers";
        var privacy = Group("Privacy", Large("Scudo privacy", "shield", PrivacyShieldDialog, tip: "Analizza il file salvato e crea una copia senza dati personali"));
        return Tab("Powers", aspetto, privacy);
    }
    ToggleButton? _darkToggle;

    void SetDark(bool dark)
    {
        Theme.Dark = dark; Theme.Apply(Application.Current.Resources); Theme.Save();
        if (_darkToggle != null) _darkToggle.IsChecked = dark;
        UpdateProbe();
    }
}

static class PanelExt
{
    public static T AddAll<T>(this T p, IEnumerable<UIElement> items) where T : Panel { foreach (var i in items) p.Children.Add(i); return p; }
}
