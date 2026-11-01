using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using PowerTables.Core;
using PowerTables.Core.Formulas;

namespace PowerTables.App;

/// <summary>Modifica di formule: elenco funzioni con Tab e selezione dinamica dei riferimenti (clic, trascinamento, tastiera).</summary>
public sealed partial class GridHost
{
    // ---------------- elenco funzioni ----------------
    readonly Popup _ac = new() { StaysOpen = false, AllowsTransparency = true, Placement = PlacementMode.Bottom };
    readonly ListBox _acList = new() { Focusable = false, MaxHeight = 260, MinWidth = 240, FontFamily = new FontFamily("Segoe UI"), FontSize = 13, BorderThickness = new Thickness(0) };
    List<(string Show, string Insert)> _acItems = [];
    int _acStart;

    // ---------------- selezione dinamica ----------------
    int _pointStart, _pointLen;                 // riferimento inserito dal puntamento (sessione attiva se _pointLen > 0)
    (int R, int C) _pointAnchor, _pointCur;
    bool _own;                                  // modifica del testo fatta dal codice (non resetta la sessione)
    public bool PointSessionActive => _pointLen > 0;
    public bool EditingFormula => Editing && Editor.Text.StartsWith('=');

    static readonly Color[] RefColors = [Color.FromRgb(0x4f, 0x8c, 0xff), Color.FromRgb(0xe5, 0x5b, 0x5b), Color.FromRgb(0x9b, 0x6b, 0xe0), Color.FromRgb(0xf0, 0xa8, 0x30), Color.FromRgb(0x2f, 0xb3, 0x9a), Color.FromRgb(0xd0, 0x5f, 0xb0)];
    static readonly Regex RefTail = new(@"(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)$", RegexOptions.Compiled);

    void InitFormulaEditing()
    {
        _acList.SetResourceReference(Control.BackgroundProperty, "PanelBrush");
        _acList.SetResourceReference(Control.ForegroundProperty, "TextBrush");
        AutomationProperties.SetAutomationId(_acList, "FormulaAutocomplete");
        var border = new Border { Child = _acList, BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(6), Margin = new Thickness(0, 2, 0, 0) };
        border.SetResourceReference(Border.BorderBrushProperty, "AccentBrush");
        border.SetResourceReference(Border.BackgroundProperty, "PanelBrush");
        _ac.Child = border; _ac.PlacementTarget = Editor;
        _acList.PreviewMouseLeftButtonUp += (_, _) => { if (_acList.SelectedIndex >= 0) CompleteFunction(); Editor.Focus(); };
        // il foglio disegna i riferimenti della formula in modifica e inoltra clic/trascinamento
        Canvas.PointClick = HandlePointClick;
        Canvas.PointDrag = HandlePointDrag;
        Canvas.EditingFormula = () => EditingFormula;
    }

    void EndFormulaEditing() { _ac.IsOpen = false; _pointLen = 0; Canvas.EditRefs = []; }

    // ---------------- testo cambiato ----------------
    void OnEditorTextChanged()
    {
        if (!_own) _pointLen = 0;
        RefreshEditRefs();
        UpdateAutocomplete();
    }

    static bool IsOperatorBefore(string text, int index)
    {
        int i = index - 1;
        while (i >= 0 && text[i] == ' ') i--;
        return i >= 0 && "=(,;+-*/^&<>:{".IndexOf(text[i]) >= 0;
    }

    // ---------------- autocompletamento ----------------
    void UpdateAutocomplete()
    {
        string t = Editor.Text; int caret = Editor.CaretIndex;
        if (!t.StartsWith('=') || caret < 1 || caret > t.Length) { _ac.IsOpen = false; return; }
        int start = caret;
        while (start > 0 && (char.IsLetterOrDigit(t[start - 1]) || t[start - 1] is '.' or '_')) start--;
        string token = t[start..caret];
        bool exprStart = IsOperatorBefore(t, start);
        if (!exprStart || (token.Length > 0 && !char.IsLetter(token[0])) || (token.Length == 0 && !(start == 1))) { _ac.IsOpen = false; return; }
        bool inString = t[..start].Count(c => c == '"') % 2 == 1;
        if (inString) { _ac.IsOpen = false; return; }
        var items = new List<(string, string)>();
        foreach (var f in FunctionCatalog.All)
        {
            bool en = f.En.StartsWith(token, StringComparison.OrdinalIgnoreCase), it = f.It != f.En && f.It.StartsWith(token, StringComparison.OrdinalIgnoreCase);
            if (token.Length == 0 || en) items.Add((f.It != f.En ? $"{f.En}    {f.It}" : f.En, f.En));
            else if (it) items.Add(($"{f.It}    {f.En}", f.It));
        }
        // se il testo è già un nome completo e unico (es. "SUM(") non serve l'elenco
        if (items.Count == 0 || (items.Count == 1 && string.Equals(items[0].Item2, token, StringComparison.OrdinalIgnoreCase))) { _ac.IsOpen = false; return; }
        _acItems = items.Take(40).ToList(); _acStart = start;
        _acList.ItemsSource = _acItems.Select(i => i.Show).ToList();
        _acList.SelectedIndex = 0;
        if (!_ac.IsOpen) _ac.IsOpen = true;
        UpdateAcPosition();
    }

    void UpdateAcPosition()
    {
        _ac.HorizontalOffset = 0; _ac.VerticalOffset = 0;
    }

    void CompleteFunction()
    {
        int i = _acList.SelectedIndex; if (i < 0 || i >= _acItems.Count) return;
        string name = _acItems[i].Insert;
        int caret = Editor.CaretIndex;
        _own = true;
        Editor.Text = Editor.Text[.._acStart] + name + "(" + Editor.Text[caret..];
        Editor.CaretIndex = _acStart + name.Length + 1;
        _own = false;
        _ac.IsOpen = false;
    }

    // ---------------- tastiera ----------------
    bool FormulaKey(KeyEventArgs e)
    {
        bool shift = (Keyboard.Modifiers & ModifierKeys.Shift) != 0;
        if (_ac.IsOpen)
        {
            switch (e.Key)
            {
                case Key.Down: _acList.SelectedIndex = Math.Min(_acItems.Count - 1, _acList.SelectedIndex + 1); _acList.ScrollIntoView(_acList.SelectedItem); return true;
                case Key.Up: _acList.SelectedIndex = Math.Max(0, _acList.SelectedIndex - 1); _acList.ScrollIntoView(_acList.SelectedItem); return true;
                case Key.Tab when !shift: CompleteFunction(); return true;
                case Key.Escape: _ac.IsOpen = false; return true;
            }
        }
        if (!EditingFormula || e.Key is not (Key.Left or Key.Right or Key.Up or Key.Down)) return false;
        if ((Keyboard.Modifiers & ModifierKeys.Control) != 0) return false;
        bool ctx = PointSessionActive || IsOperatorBefore(Editor.Text, Editor.CaretIndex) && Editor.SelectionLength == 0 && InsideEnd();
        if (!ctx) return false;
        int dr = e.Key == Key.Up ? -1 : e.Key == Key.Down ? 1 : 0, dc = e.Key == Key.Left ? -1 : e.Key == Key.Right ? 1 : 0;
        if (!PointSessionActive) { _pointCur = _pointAnchor = (Canvas.ActiveRow, Canvas.ActiveCol); }
        var next = (R: Math.Clamp(_pointCur.R + dr, 1, Addr.MaxRow), C: Math.Clamp(_pointCur.C + dc, 1, Addr.MaxCol));
        if (!PointSessionActive) { _pointAnchor = next; _pointCur = next; PointInsert(next.R, next.C, next.R, next.C, replaceExisting: false); }
        else
        {
            _pointCur = next; if (!shift) _pointAnchor = next;
            PointInsert(_pointAnchor.R, _pointAnchor.C, _pointCur.R, _pointCur.C, replaceExisting: true);
        }
        Canvas.EnsureVisible(next.R, next.C);
        return true;
    }

    bool InsideEnd() => true;

    // ---------------- clic e trascinamento ----------------
    /// <summary>Clic in una cella mentre si modifica: restituisce true se il clic è stato assorbito (formula in modifica).</summary>
    bool HandlePointClick(int r, int c)
    {
        if (!EditingFormula) return false;
        _ac.IsOpen = false;
        bool replaceTyped = !PointSessionActive && !IsOperatorBefore(Editor.Text, Editor.CaretIndex) && RefTail.IsMatch(Editor.Text[..Editor.CaretIndex]) && IsOperatorBefore(Editor.Text, Editor.CaretIndex - RefTail.Match(Editor.Text[..Editor.CaretIndex]).Length);
        bool canInsert = IsOperatorBefore(Editor.Text, Editor.CaretIndex) || PointSessionActive || replaceTyped;
        if (!canInsert) return true;                       // clic in un punto dove non va un riferimento: ignorato, la modifica resta aperta
        _pointAnchor = _pointCur = (r, c);
        PointInsert(r, c, r, c, replaceExisting: PointSessionActive || replaceTyped, replaceTyped);
        Editor.Focus();
        return true;
    }

    void HandlePointDrag(int r, int c)
    {
        if (!EditingFormula || !PointSessionActive) return;
        _pointCur = (r, c);
        PointInsert(_pointAnchor.R, _pointAnchor.C, r, c, replaceExisting: true);
    }

    /// <summary>Scrive il riferimento (A1 oppure A1:B2, sempre relativo) nel punto giusto della formula.</summary>
    void PointInsert(int r1, int c1, int r2, int c2, bool replaceExisting, bool replaceTyped = false)
    {
        int ra = Math.Min(r1, r2), rb = Math.Max(r1, r2), ca = Math.Min(c1, c2), cb = Math.Max(c1, c2);
        string refText = (ra == rb && ca == cb) ? Addr.A1(ra, ca) : Addr.A1(ra, ca) + ":" + Addr.A1(rb, cb);
        string t = Editor.Text; int caret = Editor.CaretIndex;
        int start, len;
        if (PointSessionActive && replaceExisting) { start = _pointStart; len = _pointLen; }
        else if (replaceTyped) { var m = RefTail.Match(t[..caret]); start = m.Index; len = m.Length; }
        else { start = caret; len = 0; }
        _own = true;
        Editor.Text = t[..start] + refText + t[(start + len)..];
        Editor.CaretIndex = start + refText.Length;
        _own = false;
        _pointStart = start; _pointLen = refText.Length;
        RefreshEditRefs();
    }

    // ---------------- evidenziazione dei riferimenti ----------------
    void RefreshEditRefs()
    {
        var list = new List<(int, int, int, int, Color)>();
        if (EditingFormula)
        {
            try
            {
                int k = 0;
                foreach (var tk in Lexer.Tokenize(Editor.Text[1..]).Where(t => t.Kind == T.Ref))
                {
                    var rt = tk.Ref!; if (rt.Sheet != null && !string.Equals(rt.Sheet, Canvas.Sheet.Name, StringComparison.OrdinalIgnoreCase)) continue;
                    int a = rt.A.Row, b = rt.A.Col, c = rt.IsRange ? rt.B.Row : a, d = rt.IsRange ? rt.B.Col : b;
                    if (a == 0 || b == 0) continue;
                    list.Add((Math.Min(a, c), Math.Min(b, d), Math.Max(a, c), Math.Max(b, d), RefColors[k++ % RefColors.Length]));
                }
            }
            catch (FormatException) { }
        }
        Canvas.EditRefs = list;
    }

    /// <summary>Apre la modifica nella cella attiva con il testo indicato (es. "=SUM(").</summary>
    public void BeginEditWith(string text)
    {
        if (Editing) { _own = true; int c = Editor.CaretIndex; Editor.Text = Editor.Text[..c] + text + Editor.Text[c..]; Editor.CaretIndex = c + text.Length; _own = false; Editor.Focus(); return; }
        BeginEdit(null, text);
    }
}
