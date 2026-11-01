using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using Microsoft.Win32;
using PowerTables.Core;
using PowerTables.Core.Formulas;
using PowerTables.Core.IO;

namespace PowerTables.App;

public partial class MainWindow : Window
{
    Workbook Wb = null!;
    readonly GridHost Host = new();
    GridCanvas G => Host.Canvas;
    Sheet Cur => G.Sheet;
    ClipData? _clip; string? _clipText;
    readonly Dictionary<string, Panel> _tabs = [];
    CultureInfo Ci => CultureInfo.CurrentCulture;
    ToggleButton? _bold, _italic, _underline;

    public MainWindow()
    {
        InitializeComponent();
        GridArea.Children.Add(Host);
        BuildRibbon();
        SetWorkbook(XlsxWriter.CreateNew());
        G.SelectionChanged += OnSelection;
        G.CopyRequested += () => DoCopy(false); G.CutRequested += () => DoCopy(true); G.PasteRequested += DoPaste;
        G.FilterMenuRequested += ShowFilterMenu;
        Host.EditTextChanged += t => { if (!FormulaBar.IsKeyboardFocusWithin) FormulaBar.Text = t; };
        Host.EditEnded += RefreshBar;
        FormulaBar.PreviewKeyDown += (_, e) =>
        {
            if (e.Key == Key.Enter) { Host.CommitFromBar(FormulaBar.Text); G.Select(G.ActiveRow + 1, G.ActiveCol); e.Handled = true; }
            else if (e.Key == Key.Escape) { RefreshBar(); G.Focus(); e.Handled = true; }
        };
        FormulaBar.TextChanged += (_, _) => { if (Host.Editing && FormulaBar.IsKeyboardFocusWithin) Host.SetEditText(FormulaBar.Text); };
        FormulaBar.GotKeyboardFocus += (_, _) => Host.FormulaBarFocused = true;
        FormulaBar.LostKeyboardFocus += (_, _) => Host.FormulaBarFocused = false;
        NameBox.PreviewKeyDown += (_, e) =>
        {
            if (e.Key != Key.Enter) return;
            if (Addr.TryParseRange(NameBox.Text.Replace("$", ""), out var r1, out var c1, out var r2, out var c2)) { G.SelectRange(r1, c1, r2, c2); G.EnsureVisible(r1, c1); G.Focus(); }
            else MessageBox.Show(this, "Riferimento non valido.", "PowerTables");
            e.Handled = true;
        };
        Drop += (_, e) => { if (e.Data.GetData(DataFormats.FileDrop) is string[] f && f.Length > 0) OpenFile(f[0]); };
        Closing += OnClosing;
        PreviewKeyDown += (_, e) =>
        {
            if (Keyboard.Modifiers == ModifierKeys.Control)
            {
                if (e.Key == Key.S) { Save(false); e.Handled = true; }
                else if (e.Key == Key.O) { Open(); e.Handled = true; }
                else if (e.Key == Key.N) { NewDoc(); e.Handled = true; }
                else if (e.Key == Key.B && !Host.Editing) { Toggle(x => x.Font.Bold, (x, v) => x with { Font = x.Font with { Bold = v } }); e.Handled = true; }
                else if (e.Key == Key.I && !Host.Editing) { Toggle(x => x.Font.Italic, (x, v) => x with { Font = x.Font with { Italic = v } }); e.Handled = true; }
                else if (e.Key == Key.U && !Host.Editing) { Toggle(x => x.Font.Underline, (x, v) => x with { Font = x.Font with { Underline = v } }); e.Handled = true; }
                else if (e.Key == Key.D && !Host.Editing) { FillDirection(true); e.Handled = true; }
                else if (e.Key == Key.R && !Host.Editing) { FillDirection(false); e.Handled = true; }
            }
        };
        Loaded += (_, _) => G.Focus();
    }

    // ---------------- documento ----------------
    void SetWorkbook(Workbook wb)
    {
        if (Wb != null) Wb.Changed -= OnChanged;
        Wb = wb; wb.Changed += OnChanged;
        G.Attach(wb, wb.Sheets[0]);
        G.Culture = Ci;
        RebuildSheetTabs(); RefreshBar(); UpdateTitle();
    }

    void OnChanged() { UpdateTitle(); RefreshBar(); UpdateStatus(); }

    void UpdateTitle()
    {
        string n = Wb.FilePath != null ? Path.GetFileName(Wb.FilePath) : "Senza titolo";
        Title = $"{n}{(Wb.Modified ? " *" : "")} - PowerTables";
    }

    void NewDoc()
    {
        if (!ConfirmDiscard()) return;
        SetWorkbook(XlsxWriter.CreateNew());
    }

    bool ConfirmDiscard()
    {
        if (!Wb.Modified) return true;
        var r = MessageBox.Show(this, "Salvare le modifiche?", "PowerTables", MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
        if (r == MessageBoxResult.Cancel) return false;
        if (r == MessageBoxResult.Yes) return Save(false);
        return true;
    }

    void Open()
    {
        var dlg = new OpenFileDialog { Filter = "Cartelle di lavoro (*.xlsx)|*.xlsx|Tutti i file|*.*" };
        if (dlg.ShowDialog(this) == true) OpenFile(dlg.FileName);
    }

    public void OpenFile(string path)
    {
        if (!ConfirmDiscard()) return;
        try
        {
            Cursor = Cursors.Wait;
            var sw = Stopwatch.StartNew();
            var wb = path.EndsWith(".csv", StringComparison.OrdinalIgnoreCase) ? ImportCsvNew(path) : XlsxReader.Open(path);
            SetWorkbook(wb);
            StatusRight.Text = $"Aperto in {sw.Elapsed.TotalSeconds:0.00} s - {wb.Sheets.Sum(s => s.CellCount):N0} celle";
        }
        catch (XlsxException ex) { MessageBox.Show(this, ex.Message, "PowerTables", MessageBoxButton.OK, MessageBoxImage.Error); }
        catch (IOException ex) { MessageBox.Show(this, "Errore di lettura: " + ex.Message, "PowerTables", MessageBoxButton.OK, MessageBoxImage.Error); }
        finally { Cursor = null; }
    }

    Workbook ImportCsvNew(string path)
    {
        var wb = XlsxWriter.CreateNew();
        Csv.Import(wb, wb.Sheets[0], File.ReadAllText(path));
        wb.Modified = true;
        return wb;
    }

    bool Save(bool asNew)
    {
        string? path = Wb.FilePath;
        if (asNew || path == null || !path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            var dlg = new SaveFileDialog { Filter = "Cartella di lavoro (*.xlsx)|*.xlsx", FileName = path != null ? Path.GetFileNameWithoutExtension(path) + ".xlsx" : "Cartel1.xlsx" };
            if (dlg.ShowDialog(this) != true) return false;
            path = dlg.FileName;
        }
        try
        {
            Cursor = Cursors.Wait;
            var sw = Stopwatch.StartNew();
            XlsxWriter.Save(Wb, path);
            StatusRight.Text = $"Salvato in {sw.Elapsed.TotalSeconds:0.00} s";
            UpdateTitle();
            return true;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or XlsxException)
        {
            MessageBox.Show(this, "Salvataggio non riuscito, il file originale non è stato modificato:\n" + ex.Message, "PowerTables", MessageBoxButton.OK, MessageBoxImage.Error);
            return false;
        }
        finally { Cursor = null; }
    }

    void OnClosing(object? s, CancelEventArgs e) { if (!ConfirmDiscard()) e.Cancel = true; }

    // ---------------- barra, stato, schede foglio ----------------
    void OnSelection()
    {
        var (r1, c1, r2, c2) = G.Selection;
        NameBox.Text = (r1 == r2 && c1 == c2) ? Addr.A1(r1, c1) : $"{Addr.A1(r1, c1)}:{Addr.A1(r2, c2)}";
        RefreshBar(); UpdateStatus(); UpdateToggles();
    }

    void RefreshBar()
    {
        if (Host.Editing || Cur == null) return;
        FormulaBar.Text = Wb.EditText(Cur, G.ActiveRow, G.ActiveCol, Ci);
    }

    void UpdateStatus()
    {
        if (Cur == null) return;
        var (r1, c1, r2, c2) = G.Selection;
        double sum = 0; int n = 0, cnt = 0;
        if ((r1 != r2 || c1 != c2))
        {
            var rr = new RangeRef(Cur, r1, c1, r2, c2);
            foreach (var c in Fns.StoredPublic(rr)) { if (c.Kind != CellKind.Empty) cnt++; if (c.Kind == CellKind.Num) { sum += c.Num; n++; } }
            StatusLeft.Text = n > 0 ? $"Somma: {sum.ToString("#,##0.##", Ci)}    Media: {(sum / n).ToString("#,##0.##", Ci)}    Conteggio: {cnt:N0}" : $"Conteggio: {cnt:N0}";
        }
        else StatusLeft.Text = "";
        if (Wb.Cycles.Count > 0) StatusRight.Text = "Riferimento circolare: " + Addr.A1(Wb.Cycles[0].Row, Wb.Cycles[0].Col);
        else if (StatusRight.Text.StartsWith("Riferimento")) StatusRight.Text = "Pronto";
    }

    void RebuildSheetTabs()
    {
        SheetTabs.Children.Clear();
        foreach (var s in Wb.Sheets)
        {
            var sh = s;
            var b = new ToggleButton { Content = s.Name, IsChecked = ReferenceEquals(s, Cur), Padding = new Thickness(14, 4, 14, 4) };
            b.Click += (_, _) => { G.SwitchSheet(sh); RebuildSheetTabs(); RefreshBar(); UpdateStatus(); };
            b.MouseDoubleClick += (_, _) => RenameSheet(sh);
            b.ContextMenu = null;
            SheetTabs.Children.Add(b);
        }
        var add = new Button { Content = "+", ToolTip = "Nuovo foglio" };
        add.Click += (_, _) => { var s = Wb.AddSheet(); G.SwitchSheet(s); RebuildSheetTabs(); };
        SheetTabs.Children.Add(add);
    }

    void RenameSheet(Sheet s)
    {
        var name = Ask("Rinomina foglio", "Nome:", s.Name);
        if (name == null) return;
        try { Wb.RenameSheet(s, name); RebuildSheetTabs(); }
        catch (ArgumentException ex) { MessageBox.Show(this, ex.Message, "PowerTables"); }
    }

    string? Ask(string title, string prompt, string def)
    {
        var tb = new TextBox { Text = def, Margin = new Thickness(0, 6, 0, 10), MinWidth = 220 };
        string? result = null;
        var ok = new Button { Content = "OK", IsDefault = true, Padding = new Thickness(18, 4, 18, 4) };
        var cancel = new Button { Content = "Annulla", IsCancel = true, Padding = new Thickness(18, 4, 18, 4) };
        var w = new Window
        {
            Title = title, Owner = this, SizeToContent = SizeToContent.WidthAndHeight, WindowStartupLocation = WindowStartupLocation.CenterOwner,
            ResizeMode = ResizeMode.NoResize, Background = (Brush)FindResource("PanelBrush"),
            Content = new StackPanel
            {
                Margin = new Thickness(16),
                Children = { new TextBlock { Text = prompt, Foreground = (Brush)FindResource("TextBrush") }, tb, new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Children = { ok, cancel } } },
            },
        };
        ok.Click += (_, _) => { result = tb.Text; w.DialogResult = true; };
        w.Loaded += (_, _) => { tb.Focus(); tb.SelectAll(); };
        return w.ShowDialog() == true ? result : null;
    }

    // ---------------- copia / incolla ----------------
    void DoCopy(bool cut)
    {
        var (r1, c1, r2, c2) = G.Selection;
        if ((long)(r2 - r1 + 1) * (c2 - c1 + 1) > 2_000_000) { MessageBox.Show(this, "Selezione troppo grande per gli appunti.", "PowerTables"); return; }
        _clip = Wb.Copy(Cur, r1, c1, r2, c2, cut);
        var sb = new StringBuilder();
        for (int r = r1; r <= r2; r++)
        {
            for (int c = c1; c <= c2; c++) { if (c > c1) sb.Append('\t'); sb.Append(Wb.DisplayText(Cur, r, c, Ci)); }
            sb.Append("\r\n");
        }
        _clipText = sb.ToString();
        try { Clipboard.SetText(_clipText); } catch (System.Runtime.InteropServices.ExternalException) { }
        G.Marching = _clip;
    }

    void DoPaste()
    {
        try
        {
            string? text = null;
            try { text = Clipboard.ContainsText() ? Clipboard.GetText() : null; } catch (System.Runtime.InteropServices.ExternalException) { }
            (int, int) end;
            if (_clip != null && (text == null || text == _clipText)) { end = Wb.Paste(_clip, Cur, G.ActiveRow, G.ActiveCol); if (_clip.Cut) { _clip = null; G.Marching = null; } }
            else if (text != null) end = Wb.PasteText(Cur, G.ActiveRow, G.ActiveCol, text, Ci);
            else return;
            G.SelectRange(G.ActiveRow, G.ActiveCol, end.Item1, end.Item2);
        }
        catch (FormatException ex) { MessageBox.Show(this, ex.Message, "PowerTables"); }
    }

    void FillDirection(bool down)
    {
        var (r1, c1, r2, c2) = G.Selection;
        if (down && r2 > r1) Wb.AutoFill(Cur, r1, c1, r1, c2, r2, c2);
        else if (!down && c2 > c1) Wb.AutoFill(Cur, r1, c1, r2, c1, r2, c2);
    }

    // ---------------- formattazione ----------------
    XfInfo ActiveXf => Wb.Styles.Get(Cur.Get(G.ActiveRow, G.ActiveCol).Style);

    void Fmt(Func<XfInfo, XfInfo> f)
    {
        var (r1, c1, r2, c2) = G.Selection;
        if ((long)(r2 - r1 + 1) * (c2 - c1 + 1) > 500_000) { r2 = Math.Min(r2, Math.Max(Cur.MaxRow, r1)); c2 = Math.Min(c2, Math.Max(Cur.MaxCol, c1)); } // ponytail: niente stili di riga/colonna
        Wb.FormatRange(Cur, r1, c1, r2, c2, f);
        G.InvalidateVisual(); UpdateToggles();
    }

    void Toggle(Func<XfInfo, bool> get, Func<XfInfo, bool, XfInfo> set)
    {
        bool v = !get(ActiveXf);
        Fmt(x => set(x, v));
    }

    void UpdateToggles()
    {
        if (_bold == null || Cur == null) return;
        var x = ActiveXf; _bold.IsChecked = x.Font.Bold; _italic!.IsChecked = x.Font.Italic; _underline!.IsChecked = x.Font.Underline;
    }

    static string Hex(Color c) => $"FF{c.R:X2}{c.G:X2}{c.B:X2}";

    void NumFormat(int id, string? code = null) => Fmt(x => x with { NumFmtId = code != null ? 164 : id, NumFmtCode = code });

    void Decimals(int delta)
    {
        string code = Wb.Styles.FormatCode(Cur.Get(G.ActiveRow, G.ActiveCol).Style);
        var m = System.Text.RegularExpressions.Regex.Match(code, @"^([^;]*?)(?:\.(0+))?([^0#?.;]*)(;.*)?$");
        if (code == "General") code = "0.0"; // da Generale: parte da 1 decimale
        else
        {
            int cur = m.Groups[2].Success ? m.Groups[2].Length : 0; int nd = Math.Max(0, cur + delta);
            string intPart = m.Groups[1].Value; string tail = m.Groups[3].Value;
            code = intPart + (nd > 0 ? "." + new string('0', nd) : "") + tail + m.Groups[4].Value;
        }
        NumFormat(0, code);
    }

    void Palette(Button anchor, Action<string?> pick)
    {
        var pop = new Popup { PlacementTarget = anchor, Placement = PlacementMode.Bottom, StaysOpen = false, AllowsTransparency = true };
        var grid = new UniformGrid { Columns = 10, Margin = new Thickness(6) };
        string[] cols =
        [
            "000000","434343","666666","999999","B7B7B7","CCCCCC","D9D9D9","EFEFEF","F3F3F3","FFFFFF",
            "980000","FF0000","FF9900","FFFF00","00FF00","00FFFF","4A86E8","0000FF","9900FF","FF00FF",
            "E6B8AF","F4CCCC","FCE5CD","FFF2CC","D9EAD3","D0E0E3","C9DAF8","CFE2F3","D9D2E9","EAD1DC",
            "CC4125","E06666","F6B26B","FFD966","93C47D","76A5AF","6D9EEB","6FA8DC","8E7CC3","C27BA0",
            "3ECF8E","1C7A52","0E0E0E","ECE6DC","5B9BD5","ED7D31","A5A5A5","FFC000","4472C4","70AD47",
        ];
        foreach (var hex in cols)
        {
            var col = (Color)ColorConverter.ConvertFromString("#" + hex);
            var b = new Button { Width = 20, Height = 20, Margin = new Thickness(1), Padding = new Thickness(0), Background = new SolidColorBrush(col), BorderBrush = (Brush)FindResource("BorderBrush2"), ToolTip = "#" + hex };
            b.Click += (_, _) => { pop.IsOpen = false; pick("FF" + hex); };
            grid.Children.Add(b);
        }
        var none = new Button { Content = "Nessun colore", Margin = new Thickness(6, 0, 6, 6) };
        none.Click += (_, _) => { pop.IsOpen = false; pick(null); };
        pop.Child = new Border { Background = (Brush)FindResource("PanelBrush"), BorderBrush = (Brush)FindResource("AccentBrush"), BorderThickness = new Thickness(1), Child = new StackPanel { Children = { grid, none } } };
        pop.IsOpen = true;
    }

    void DropMenu(Button anchor, params (string Label, Action Act)[] items)
    {
        var pop = new Popup { PlacementTarget = anchor, Placement = PlacementMode.Bottom, StaysOpen = false, AllowsTransparency = true };
        var sp = new StackPanel { Margin = new Thickness(2) };
        foreach (var (label, act) in items)
        {
            var b = new Button { Content = label, HorizontalContentAlignment = HorizontalAlignment.Left, MinWidth = 180 };
            b.Click += (_, _) => { pop.IsOpen = false; act(); };
            sp.Children.Add(b);
        }
        pop.Child = new Border { Background = (Brush)FindResource("PanelBrush"), BorderBrush = (Brush)FindResource("AccentBrush"), BorderThickness = new Thickness(1), Child = sp };
        pop.IsOpen = true;
    }

    // ---------------- dati: ordina, filtra, riquadri ----------------
    (int R1, int C1, int R2, int C2) CurrentRegion()
    {
        var (r1, c1, r2, c2) = G.Selection;
        if (r1 != r2 || c1 != c2) return (r1, c1, r2, c2);
        bool Has(int r, int c) => r >= 1 && c >= 1 && Cur.CellMap.TryGetValue(Addr.Key(r, c), out var x) && x.Kind != CellKind.Empty;
        int top = r1, bot = r1, left = c1, right = c1; bool grew = true;
        while (grew)
        {
            grew = false;
            for (int c = left - 1; c <= right + 1; c++)
            {
                if (c < 1) continue;
                if (Has(top - 1, c) ) { top--; grew = true; break; }
                if (Has(bot + 1, c)) { bot++; grew = true; break; }
            }
            if (grew) continue;
            for (int r = top - 1; r <= bot + 1; r++)
            {
                if (r < 1) continue;
                if (Has(r, left - 1)) { left--; grew = true; break; }
                if (Has(r, right + 1)) { right++; grew = true; break; }
            }
        }
        return (top, left, bot, right);
    }

    void DoSort(bool asc)
    {
        var (r1, c1, r2, c2) = CurrentRegion();
        if (r1 == r2) return;
        bool header = Cur.Get(r1, c1).Kind == CellKind.Str && Enumerable.Range(c1, c2 - c1 + 1).Any(c => Cur.Get(r1 + 1, c).Kind is CellKind.Num or CellKind.Bool);
        int key = Math.Clamp(G.ActiveCol, c1, c2);
        Wb.Sort(Cur, r1, c1, r2, c2, [(key, asc)], header);
        G.SelectRange(r1, c1, r2, c2);
    }

    void ToggleFilter()
    {
        if (Cur.Filter != null) { Wb.ClearAutoFilter(Cur); return; }
        var (r1, c1, r2, c2) = CurrentRegion();
        if (r1 == r2 && c1 == c2 && Cur.Get(r1, c1).Kind == CellKind.Empty) return;
        Wb.SetAutoFilter(Cur, r1, c1, r2, c2);
    }

    void ShowFilterMenu(int col, Point screen)
    {
        var f = Cur.Filter!;
        var values = Wb.DistinctValues(Cur, col, Ci);
        var allowed = f.Allowed.TryGetValue(col, out var a) ? a : null;
        var list = new StackPanel();
        var boxes = new List<CheckBox>();
        foreach (var v in values)
        {
            var cb = new CheckBox { Content = v.Length == 0 ? "(vuote)" : v, Tag = v, IsChecked = allowed == null || allowed.Contains(v), Foreground = (Brush)FindResource("TextBrush"), Margin = new Thickness(4, 2, 4, 2) };
            boxes.Add(cb); list.Children.Add(cb);
        }
        var all = new Button { Content = "Tutto / nulla" };
        all.Click += (_, _) => { bool any = boxes.Any(b => b.IsChecked != true); foreach (var b in boxes) b.IsChecked = any; };
        var ok = new Button { Content = "OK", IsDefault = true };
        var clear = new Button { Content = "Cancella filtro" };
        var w = new Window
        {
            WindowStyle = WindowStyle.ToolWindow, Title = "Filtro " + Addr.ColName(col), Owner = this, Width = 240, Height = 340, Background = (Brush)FindResource("PanelBrush"),
            WindowStartupLocation = WindowStartupLocation.Manual, Left = screen.X, Top = screen.Y,
        };
        var dock = new DockPanel();
        var bar = new StackPanel { Orientation = Orientation.Horizontal, Children = { all, clear, ok } };
        DockPanel.SetDock(bar, Dock.Bottom);
        dock.Children.Add(bar); dock.Children.Add(new ScrollViewer { Content = list, VerticalScrollBarVisibility = ScrollBarVisibility.Auto });
        w.Content = dock;
        ok.Click += (_, _) =>
        {
            var sel = boxes.Where(b => b.IsChecked == true).Select(b => (string)b.Tag).ToHashSet();
            w.Close();
            Wb.SetColumnFilter(Cur, col, sel.Count == boxes.Count ? null : sel, Ci);
        };
        clear.Click += (_, _) => { w.Close(); Wb.SetColumnFilter(Cur, col, null, Ci); };
        w.ShowDialog();
    }

    // ---------------- ribbon ----------------
    UIElement Group(string label, params UIElement[] items)
    {
        var wrap = new WrapPanel { MaxWidth = 330, Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Center };
        foreach (var i in items) wrap.Children.Add(i);
        var sp = new StackPanel { Margin = new Thickness(6, 2, 6, 0), Children = { wrap, new TextBlock { Text = label, Foreground = (Brush)FindResource("TextDimBrush"), FontSize = 10, HorizontalAlignment = HorizontalAlignment.Center, FontFamily = new FontFamily("Segoe UI") } } };
        return new Border { BorderBrush = (Brush)FindResource("BorderBrush2"), BorderThickness = new Thickness(0, 0, 1, 0), Child = sp };
    }

    Button B(string text, Action act, string? tip = null, double? width = null)
    {
        var b = new Button { Content = text, ToolTip = tip ?? text };
        if (width != null) b.MinWidth = width.Value;
        b.Click += (_, _) => { try { act(); } catch (Exception ex) when (ex is FormatException or ArgumentException or InvalidOperationException) { MessageBox.Show(this, ex.Message, "PowerTables"); } G.Focus(); };
        return b;
    }

    ToggleButton T(string text, string tip, Func<XfInfo, bool> get, Func<XfInfo, bool, XfInfo> set, FontWeight? w = null, FontStyle? st = null, bool underline = false)
    {
        var t = new ToggleButton { Content = new TextBlock { Text = text, FontWeight = w ?? FontWeights.Normal, FontStyle = st ?? FontStyles.Normal, TextDecorations = underline ? TextDecorations.Underline : null }, ToolTip = tip, MinWidth = 28 };
        t.Click += (_, _) => { Toggle(get, set); G.Focus(); };
        return t;
    }

    void BuildRibbon()
    {
        var file = new StackPanel { Orientation = Orientation.Horizontal };
        file.Children.Add(Group("Documento", B("Nuovo", NewDoc), B("Apri", Open), B("Salva", () => Save(false)), B("Salva con nome", () => Save(true))));
        file.Children.Add(Group("CSV", B("Importa CSV", ImportCsv), B("Esporta CSV", ExportCsv)));

        var home = new StackPanel { Orientation = Orientation.Horizontal };
        home.Children.Add(Group("Appunti", B("Incolla", DoPaste, "Incolla (Ctrl+V)"), B("Taglia", () => DoCopy(true), "Taglia (Ctrl+X)"), B("Copia", () => DoCopy(false), "Copia (Ctrl+C)"), B("Annulla", () => Wb.Undo(), "Annulla (Ctrl+Z)"), B("Ripeti", () => Wb.Redo(), "Ripeti (Ctrl+Y)")));
        _bold = T("G", "Grassetto (Ctrl+B)", x => x.Font.Bold, (x, v) => x with { Font = x.Font with { Bold = v } }, FontWeights.Bold);
        _italic = T("C", "Corsivo (Ctrl+I)", x => x.Font.Italic, (x, v) => x with { Font = x.Font with { Italic = v } }, null, FontStyles.Italic);
        _underline = T("S", "Sottolineato (Ctrl+U)", x => x.Font.Underline, (x, v) => x with { Font = x.Font with { Underline = v } }, null, null, true);
        Button? fill = null, tcol = null, size = null, font = null, borders = null;
        font = B("Carattere", () => DropMenu(font!, new[] { "Calibri", "Arial", "Segoe UI", "Times New Roman", "Courier New", "Verdana", "Georgia" }.Select(n => (n, (Action)(() => Fmt(x => x with { Font = x.Font with { Name = n } })))).ToArray()), "Tipo di carattere");
        size = B("Dim.", () => DropMenu(size!, new[] { 8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36 }.Select(n => (n.ToString(), (Action)(() => Fmt(x => x with { Font = x.Font with { Size = n } })))).ToArray()), "Dimensione carattere");
        fill = B("Sfondo", () => Palette(fill!, c => Fmt(x => x with { FillColor = c })), "Colore di riempimento");
        tcol = B("Testo", () => Palette(tcol!, c => Fmt(x => x with { Font = x.Font with { Color = c } })), "Colore del testo");
        borders = B("Bordi", () => DropMenu(borders!,
            ("Tutti i bordi", () => SetBorders("all")), ("Bordo esterno", () => SetBorders("outer")), ("Bordo inferiore", () => SetBorders("bottom")), ("Nessun bordo", () => SetBorders("none"))), "Bordi");
        home.Children.Add(Group("Carattere", font, size, _bold, _italic, _underline, fill, tcol, borders));
        home.Children.Add(Group("Allineamento",
            B("Sx", () => Fmt(x => x with { HAlign = "left" }), "Allinea a sinistra"), B("Cx", () => Fmt(x => x with { HAlign = "center" }), "Centra"), B("Dx", () => Fmt(x => x with { HAlign = "right" }), "Allinea a destra"),
            B("A capo", () => Fmt(x => x with { Wrap = !x.Wrap }), "Testo a capo"),
            B("Unisci", () => { var (r1, c1, r2, c2) = G.Selection; Wb.Merge(Cur, r1, c1, r2, c2); }, "Unisci celle"), B("Dividi", () => { var (r1, c1, r2, c2) = G.Selection; Wb.Unmerge(Cur, r1, c1, r2, c2); }, "Dividi celle")));
        Button? nf = null;
        nf = B("Formato numero", () => DropMenu(nf!,
            ("Generale", () => NumFormat(0)), ("Numero 0,00", () => NumFormat(2)), ("Migliaia 1.234,00", () => NumFormat(4)),
            ("Valuta €", () => NumFormat(0, "#,##0.00\\ \"€\"")), ("Percentuale", () => NumFormat(10)), ("Data breve", () => NumFormat(14)),
            ("Data lunga", () => NumFormat(0, "dddd d mmmm yyyy")), ("Ora", () => NumFormat(21)), ("Scientifico", () => NumFormat(11)), ("Testo", () => NumFormat(49))), "Formato numerico");
        home.Children.Add(Group("Numeri", nf, B("%", () => NumFormat(9), "Percentuale"), B(".00+", () => Decimals(1), "Aumenta decimali"), B(".0-", () => Decimals(-1), "Riduci decimali")));
        home.Children.Add(Group("Celle",
            B("Larghezza colonna", () => { var s = Ask("Larghezza colonna", "Caratteri:", (Cur.ColWidth(G.ActiveCol) - 5.0 / 7).ToString("0.##", Ci)); if (s != null && double.TryParse(s, NumberStyles.Float, Ci, out var w) && w is > 0 and <= 255) { var (_, c1, _, c2) = G.Selection; for (int c = c1; c <= c2; c++) Wb.SetColWidth(Cur, c, w); } }),
            B("Altezza riga", () => { var s = Ask("Altezza riga", "Punti:", Cur.RowHeight(G.ActiveRow).ToString("0.##", Ci)); if (s != null && double.TryParse(s, NumberStyles.Float, Ci, out var h) && h is > 0 and <= 409) { var (r1, _, r2, _) = G.Selection; for (int r = r1; r <= r2; r++) Wb.SetRowHeight(Cur, r, h); } }),
            B("Adatta colonna", () => G.AutoFitColumn(G.ActiveCol))));
        home.Children.Add(Group("Modifica",
            B("Σ Somma", AutoSum, "Somma automatica"), B("Cancella", () => G.ClearSelection(), "Cancella contenuto (Canc)"),
            B("Riempi giù", () => FillDirection(true), "Ctrl+D"), B("Riempi destra", () => FillDirection(false), "Ctrl+R")));

        var data = new StackPanel { Orientation = Orientation.Horizontal };
        data.Children.Add(Group("Ordina e filtra", B("A → Z", () => DoSort(true)), B("Z → A", () => DoSort(false)), B("Filtro", ToggleFilter, "Attiva/disattiva filtro automatico")));
        data.Children.Add(Group("Riquadri", B("Blocca riga 1", () => Wb.SetFreeze(Cur, 1, Cur.FrozenCols)), B("Blocca col. A", () => Wb.SetFreeze(Cur, Cur.FrozenRows, 1)),
            B("Blocca fino a selezione", () => Wb.SetFreeze(Cur, G.ActiveRow - 1, G.ActiveCol - 1)), B("Sblocca", () => Wb.SetFreeze(Cur, 0, 0))));
        data.Children.Add(Group("Calcolo", B("Ricalcola tutto", () => { var sw = Stopwatch.StartNew(); Wb.RecalcAll(); G.InvalidateVisual(); StatusRight.Text = $"Ricalcolo: {sw.ElapsedMilliseconds} ms ({Wb.LastRecalcCount:N0} formule)"; UpdateStatus(); })));

        var view = new StackPanel { Orientation = Orientation.Horizontal };
        view.Children.Add(Group("Tema", B("Chiaro / scuro", ToggleTheme)));
        view.Children.Add(Group("Mostra", B("Griglia", () => { Cur.ShowGridLines = !Cur.ShowGridLines; G.InvalidateVisual(); })));

        _tabs["File"] = file; _tabs["Home"] = home; _tabs["Dati"] = data; _tabs["Visualizza"] = view;
        foreach (var name in _tabs.Keys)
        {
            var n = name;
            var tb = new ToggleButton { Content = n, Style = (Style)FindResource("TabBtn"), IsChecked = n == "Home" };
            tb.Click += (_, _) => ShowTab(n);
            TabStrip.Children.Add(tb);
        }
        ShowTab("Home");
    }

    void ShowTab(string name)
    {
        RibbonHost.Child = _tabs[name];
        foreach (ToggleButton t in TabStrip.Children) t.IsChecked = (string)t.Content == name;
    }

    void ToggleTheme()
    {
        Theme.Dark = !Theme.Dark; Theme.Apply(Application.Current.Resources);
        G.RelayoutAll();
    }

    void SetBorders(string kind)
    {
        var (r1, c1, r2, c2) = G.Selection;
        var line = new BorderSide("thin", null); var none = new BorderSide(null, null);
        Wb.Do("Bordi", () =>
        {
            for (int r = r1; r <= r2; r++)
                for (int c = c1; c <= c2; c++)
                {
                    var cell = Cur.Get(r, c); var xf = Wb.Styles.Get(cell.Style); var b = xf.Border;
                    b = kind switch
                    {
                        "all" => new BorderInfo(line, line, line, line),
                        "none" => BorderInfo.None,
                        "bottom" => r == r2 ? b with { Bottom = line } : b,
                        _ => b with { Left = c == c1 ? line : b.Left, Right = c == c2 ? line : b.Right, Top = r == r1 ? line : b.Top, Bottom = r == r2 ? line : b.Bottom },
                    };
                    cell.Style = Wb.Styles.Derive(cell.Style, x => x with { Border = b });
                    Wb.Write(Cur, r, c, cell);
                }
        });
    }

    void AutoSum()
    {
        int r = G.ActiveRow, c = G.ActiveCol, top = r - 1;
        while (top >= 1 && Cur.Get(top, c).Kind == CellKind.Num) top--;
        if (top == r - 1) return;
        Wb.Do("Somma", () => Wb.SetFormula(Cur, r, c, $"SUM({Addr.A1(top + 1, c)}:{Addr.A1(r - 1, c)})"));
    }

    void ImportCsv()
    {
        var dlg = new OpenFileDialog { Filter = "CSV (*.csv;*.txt)|*.csv;*.txt|Tutti i file|*.*" };
        if (dlg.ShowDialog(this) != true) return;
        try { Csv.Import(Wb, Cur, File.ReadAllText(dlg.FileName)); }
        catch (IOException ex) { MessageBox.Show(this, ex.Message, "PowerTables"); }
    }

    void ExportCsv()
    {
        var dlg = new SaveFileDialog { Filter = "CSV (*.csv)|*.csv", FileName = Cur.Name + ".csv" };
        if (dlg.ShowDialog(this) != true) return;
        try { Csv.Export(Wb, Cur, dlg.FileName); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { MessageBox.Show(this, ex.Message, "PowerTables"); }
    }

    // ---------------- benchmark ----------------
    public void RunBench(string file, string outPath)
    {
        Dispatcher.BeginInvoke(DispatcherPriority.ApplicationIdle, new Action(() =>
        {
            var res = new Dictionary<string, object>();
            var proc = Process.GetCurrentProcess();
            GC.Collect(); proc.Refresh();
            res["workingSetBeforeMB"] = proc.WorkingSet64 / 1048576.0;
            long baseMem = GC.GetTotalMemory(true);
            var sw = Stopwatch.StartNew();
            var wb = XlsxReader.Open(file);
            res["openMs"] = sw.ElapsedMilliseconds;
            SetWorkbook(wb);
            UpdateLayout();
            Dispatcher.Invoke(DispatcherPriority.Render, () => { });
            res["openToFirstPaintMs"] = sw.ElapsedMilliseconds;
            GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect(); proc.Refresh();
            res["managedMB"] = (GC.GetTotalMemory(true) - baseMem) / 1048576.0;
            res["workingSetMB"] = proc.WorkingSet64 / 1048576.0;
            res["cells"] = wb.Sheets[0].CellCount;
            res["maxRow"] = wb.Sheets[0].MaxRow;
            // scorrimento: salti casuali e passi sequenziali, misurando il rendering completo
            var times = new List<double>();
            var rnd = new Random(42);
            for (int i = 0; i < 100; i++)
            {
                G.SetScroll(0, rnd.NextDouble() * G.MaxScrollY);
                var t = Stopwatch.StartNew();
                G.InvalidateVisual();
                Dispatcher.Invoke(DispatcherPriority.Render, () => { });
                times.Add(t.Elapsed.TotalMilliseconds);
            }
            var seq = new List<double>();
            G.SetScroll(0, 0);
            for (int i = 0; i < 300; i++)
            {
                G.SetScroll(0, i * 60);
                var t = Stopwatch.StartNew();
                G.InvalidateVisual();
                Dispatcher.Invoke(DispatcherPriority.Render, () => { });
                seq.Add(t.Elapsed.TotalMilliseconds);
            }
            res["scrollJumpAvgMs"] = times.Average(); res["scrollJumpMaxMs"] = times.Max();
            res["scrollSeqAvgMs"] = seq.Average(); res["scrollSeqMaxMs"] = seq.Max();
            var s0 = wb.Sheets[0];
            sw.Restart(); wb.RecalcAll(); res["fullRecalcMs"] = sw.ElapsedMilliseconds; res["fullRecalcFormulas"] = wb.LastRecalcCount;
            sw.Restart(); wb.Do("edit", () => wb.SetValue(s0, 2, 1, 12345.0)); res["incrementalRecalcMs"] = sw.Elapsed.TotalMilliseconds; res["incrementalRecalcFormulas"] = wb.LastRecalcCount;
            proc.Refresh(); res["workingSetAfterMB"] = proc.WorkingSet64 / 1048576.0;
            File.WriteAllText(outPath, JsonSerializer.Serialize(res, new JsonSerializerOptions { WriteIndented = true }));
            Wb.Modified = false;
            Close();
        }));
    }
}
