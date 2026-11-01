using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using PowerTables.Core;

namespace PowerTables.App;

/// <summary>Contiene la griglia, le barre di scorrimento e l'editor in cella.</summary>
public sealed class GridHost : Grid
{
    public readonly GridCanvas Canvas = new();
    readonly Grid _zoomGrid = new() { ClipToBounds = true };
    public int ZoomPct { get; private set; } = 100;
    public void SetZoom(int pct) { ZoomPct = Math.Clamp(pct, 10, 400); _zoomGrid.LayoutTransform = new ScaleTransform(ZoomPct / 100.0, ZoomPct / 100.0); }
    readonly ScrollBar _v = new() { Orientation = Orientation.Vertical }, _h = new() { Orientation = Orientation.Horizontal };
    readonly Canvas _overlay = new() { IsHitTestVisible = true };
    public readonly TextBox Editor = new() { Visibility = Visibility.Collapsed, BorderThickness = new Thickness(2), Padding = new Thickness(1, 0, 1, 0), AcceptsReturn = false, VerticalContentAlignment = VerticalAlignment.Bottom };
    public bool Editing { get; private set; }
    public event Action<string>? EditTextChanged;
    public event Action? EditEnded;
    bool _sync;
    int _editR, _editC;

    public GridHost()
    {
        RowDefinitions.Add(new RowDefinition()); RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        ColumnDefinitions.Add(new ColumnDefinition()); ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        SetRow(_v, 0); SetColumn(_v, 1); SetRow(_h, 1); SetColumn(_h, 0);
        _zoomGrid.Children.Add(Canvas); _zoomGrid.Children.Add(_overlay);
        SetRow(_zoomGrid, 0); SetColumn(_zoomGrid, 0);
        Children.Add(_zoomGrid); Children.Add(_v); Children.Add(_h);
        _overlay.Children.Add(Editor);
        _overlay.Background = null;
        _overlay.IsHitTestVisible = false;
        Canvas.ViewChanged += UpdateBars;
        _v.Scroll += (_, _) => { if (!_sync) Canvas.SetScroll(Canvas.ScrollX, _v.Value); };
        _h.Scroll += (_, _) => { if (!_sync) Canvas.SetScroll(_h.Value, Canvas.ScrollY); };
        SizeChanged += (_, _) => UpdateBars();
        Editor.PreviewKeyDown += EditorKey;
        Editor.TextChanged += (_, _) => { if (Editing) EditTextChanged?.Invoke(Editor.Text); };
        Editor.LostKeyboardFocus += (_, e) => { if (Editing && !IsAncestorOf(e.NewFocus as DependencyObject) && !FormulaBarFocused) Commit(0, 0); };
        Canvas.EditRequested += c => BeginEdit(c);
        Canvas.ViewChanged += PositionEditor;
    }

    public bool FormulaBarFocused;

    void UpdateBars()
    {
        if (Canvas.Sheet == null) return;
        _sync = true;
        _v.Maximum = Canvas.MaxScrollY; _v.ViewportSize = Canvas.ViewH; _v.Value = Math.Min(Canvas.ScrollY, _v.Maximum);
        _v.SmallChange = 20; _v.LargeChange = Math.Max(20, Canvas.ViewH);
        _h.Maximum = Canvas.MaxScrollX; _h.ViewportSize = Canvas.ViewW; _h.Value = Math.Min(Canvas.ScrollX, _h.Maximum);
        _h.SmallChange = 32; _h.LargeChange = Math.Max(20, Canvas.ViewW);
        _sync = false;
    }

    void PositionEditor()
    {
        if (!Editing) return;
        var rc = Canvas.CellRect(_editR, _editC);
        Editor.Visibility = rc.Right < GridCanvas.HeaderW || rc.Bottom < GridCanvas.HeaderH ? Visibility.Collapsed : Visibility.Visible;
        Canvas2Set(Editor, rc);
    }

    void Canvas2Set(TextBox tb, Rect rc)
    {
        System.Windows.Controls.Canvas.SetLeft(tb, rc.X); System.Windows.Controls.Canvas.SetTop(tb, rc.Y);
        tb.MinWidth = rc.Width; tb.MinHeight = rc.Height;
        var xf = Canvas.Wb.Styles.Get(Canvas.Sheet.Get(_editR, _editC).Style);
        tb.FontFamily = new FontFamily(xf.Font.Name); tb.FontSize = xf.Font.Size * 96 / 72;
        tb.FontWeight = xf.Font.Bold ? FontWeights.Bold : FontWeights.Normal; tb.FontStyle = xf.Font.Italic ? FontStyles.Italic : FontStyles.Normal;
    }

    /// <summary>Avvia la modifica; ch = primo carattere digitato (sostituisce il contenuto) o null per modificare l'esistente.</summary>
    public void BeginEdit(char? ch)
    {
        if (Canvas.Sheet == null) return;
        _editR = Canvas.ActiveRow; _editC = Canvas.ActiveCol;
        var m = Canvas.Sheet.MergeAt(_editR, _editC); if (m != null) { _editR = m.Value.R1; _editC = m.Value.C1; }
        Editor.Background = Theme.Brush(Theme.CellBg); Editor.Foreground = Theme.Brush(Theme.CellText); Editor.BorderBrush = Theme.Brush(Theme.Accent);
        Editor.CaretBrush = Theme.Brush(Theme.CellText);
        Editing = true;
        Editor.Text = ch is { } c ? c.ToString() : Canvas.Wb.EditText(Canvas.Sheet, _editR, _editC, Canvas.Culture);
        _overlay.IsHitTestVisible = true;
        PositionEditor();
        Editor.Visibility = Visibility.Visible;
        Editor.Focus();
        Editor.CaretIndex = Editor.Text.Length;
        EditTextChanged?.Invoke(Editor.Text);
    }

    public void SetEditText(string t) { if (Editing && Editor.Text != t) { Editor.Text = t; Editor.CaretIndex = t.Length; } }

    public void Commit(int dr, int dc)
    {
        if (!Editing) return;
        string text = Editor.Text;
        EndEditUi();
        ApplyInput(_editR, _editC, text);
        if (dr != 0 || dc != 0) Canvas.Select(Canvas.ActiveRow + dr, Canvas.ActiveCol + dc);
        Canvas.Focus();
        EditEnded?.Invoke();
    }

    public void CommitFromBar(string text)
    {
        EndEditUi();
        ApplyInput(Canvas.ActiveRow, Canvas.ActiveCol, text);
        Canvas.Focus();
        EditEnded?.Invoke();
    }

    void ApplyInput(int r, int c, string text)
    {
        try { Canvas.Wb.Do("Modifica cella", () => Canvas.Wb.SetInput(Canvas.Sheet, r, c, text, Canvas.Culture)); }
        catch (FormatException ex) { MessageBox.Show(Window.GetWindow(this), ex.Message, "PowerTables", MessageBoxButton.OK, MessageBoxImage.Warning); }
    }

    public void Cancel()
    {
        if (!Editing) return;
        EndEditUi(); Canvas.Focus(); EditEnded?.Invoke();
    }

    void EndEditUi() { Editing = false; Editor.Visibility = Visibility.Collapsed; _overlay.IsHitTestVisible = false; }

    void EditorKey(object s, KeyEventArgs e)
    {
        switch (e.Key)
        {
            case Key.Enter: Commit((Keyboard.Modifiers & ModifierKeys.Shift) != 0 ? -1 : 1, 0); e.Handled = true; break;
            case Key.Tab: Commit(0, (Keyboard.Modifiers & ModifierKeys.Shift) != 0 ? -1 : 1); e.Handled = true; break;
            case Key.Escape: Cancel(); e.Handled = true; break;
            case Key.Up or Key.Down when !Editor.Text.StartsWith('='): Commit(e.Key == Key.Up ? -1 : 1, 0); e.Handled = true; break;
        }
    }
}
