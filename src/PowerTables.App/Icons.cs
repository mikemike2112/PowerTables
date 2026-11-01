using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;

namespace PowerTables.App;

/// <summary>Set di icone proprio (geometrie vettoriali su griglia 24x24, tratto lineare). Nessuna risorsa di terzi.</summary>
public static class Icons
{
    static readonly Dictionary<string, string> D = new()
    {
        ["paste"] = "M8 4h8v3H8z M6 5.5H5v15h14v-15h-1",
        ["cut"] = "M6 3l9 13 M18 3L9 16 M7 17.5a2.5 2.5 0 1 0 .01 0 M17 17.5a2.5 2.5 0 1 0 .01 0",
        ["copy"] = "M9 8h10v13H9z M6 17V3h10",
        ["undo"] = "M9 6L4 11l5 5 M4 11h10a5 5 0 0 1 0 10h-3",
        ["redo"] = "M15 6l5 5-5 5 M20 11H10a5 5 0 0 0 0 10h3",
        ["save"] = "M4 4h13l3 3v13H4z M8 4v5h8V4 M8 20v-6h8v6",
        ["new"] = "M6 3h8l4 4v14H6z M14 3v4h4",
        ["open"] = "M3 7h7l2 2h9v10H3z",
        ["export"] = "M12 4v11 M8 11l4 4 4-4 M5 20h14",
        ["import"] = "M12 15V4 M8 8l4-4 4 4 M5 20h14",
        ["print"] = "M7 9V3h10v6 M7 17H4V9h16v8h-3 M7 14h10v7H7z",
        ["close"] = "M6 6l12 12 M18 6L6 18",
        ["sort"] = "M6 4v15 M3 16l3 4 3-4 M13 6h8 M13 11h6 M13 16h4",
        ["sortaz"] = "M7 4v15 M4 16l3 4 3-4 M13 4l3 7 3-7 M13 20h6l-6-6h6",
        ["sortza"] = "M7 4v15 M4 8l3-4 3 4 M13 4h6l-6 6h6 M13 20l3-7 3 7",
        ["filter"] = "M4 5h16l-6 8v6l-4-2v-4z",
        ["merge"] = "M3 7h18v10H3z M7 12h10 M7 12l2.5-2.5 M7 12l2.5 2.5 M17 12l-2.5-2.5 M17 12l-2.5 2.5",
        ["unmerge"] = "M3 7h18v10H3z M12 7v10 M8 12h8",
        ["fill"] = "M5 12l7-8 7 8-7 8z M5 12h14 M21 15c1 2 1 3 0 4",
        ["border"] = "M4 4h16v16H4z M4 12h16 M12 4v16",
        ["wrap"] = "M4 6h16 M4 12h12a3.5 3.5 0 0 1 0 7h-5 M13 16l-3 3 3 3 M4 18h3",
        ["left"] = "M4 6h16 M4 10h10 M4 14h16 M4 18h10",
        ["center"] = "M4 6h16 M7 10h10 M4 14h16 M7 18h10",
        ["right"] = "M4 6h16 M10 10h10 M4 14h16 M10 18h10",
        ["top"] = "M4 4h16 M8 8h8v6H8z",
        ["middle"] = "M4 12h16 M8 8h8v8H8z",
        ["bottom"] = "M4 20h16 M8 10h8v6H8z",
        ["indent"] = "M4 5h16 M10 10h10 M10 14h10 M4 19h16 M3 10l4 2-4 2z",
        ["outdent"] = "M4 5h16 M10 10h10 M10 14h10 M4 19h16 M7 10l-4 2 4 2z",
        ["freeze"] = "M3 5h18v14H3z M3 10h18 M9 5v14",
        ["find"] = "M10.5 4a6.5 6.5 0 1 0 .01 0 M15.5 15.5L21 21",
        ["rowplus"] = "M3 6h18v5H3z M3 15h8 M7 13v4 M15 15h6 M18 13v4",
        ["colplus"] = "M6 3v18 M11 3v18 M3 11h4 M5 9v4",
        ["delete"] = "M4 6h16v12H4z M9 10l6 4 M15 10l-6 4",
        ["clear"] = "M4 16l8-9 8 7-5 5H8z M10 11l7 6",
        ["sigma"] = "M18 5H6l6 7-6 7h12",
        ["moon"] = "M20 14.5A8.5 8.5 0 1 1 9.5 4 7 7 0 0 0 20 14.5z",
        ["sun"] = "M12 8a4 4 0 1 0 .01 0 M12 2v3 M12 19v3 M2 12h3 M19 12h3 M5 5l2 2 M17 17l2 2 M19 5l-2 2 M7 17l-2 2",
        ["table"] = "M3 4h18v16H3z M3 9h18 M3 14h18 M9 4v16 M15 4v16",
        ["page"] = "M6 3h9l4 4v14H6z M9 12h7 M9 16h7",
        ["chart"] = "M4 20V4 M4 20h16 M8 16v-5 M12 16V8 M16 16v-8",
        ["image"] = "M3 5h18v14H3z M3 16l5-5 5 5 3-3 5 5 M16 9a1.5 1.5 0 1 0 .01 0",
        ["comment"] = "M4 4h16v12H9l-5 4z",
        ["calc"] = "M5 3h14v18H5z M8 7h8 M8 11h2 M12 11h2 M8 15h2 M12 15h2 M8 18h2",
        ["func"] = "M13 4c-3 0-3 2-3 4v8c0 2 0 4-3 4 M7 12h6 M15 10l4 6 M19 10l-4 6",
        ["names"] = "M4 4h16v16H4z M8 9h8 M8 13h5",
        ["zoom"] = "M10.5 4a6.5 6.5 0 1 0 .01 0 M15.5 15.5L21 21 M8 10.5h5 M10.5 8v5",
        ["gridview"] = "M3 3h18v18H3z M3 9h18 M3 15h18 M9 3v18 M15 3v18",
        ["window"] = "M3 5h18v14H3z M3 9h18",
        ["theme"] = "M12 3a9 9 0 1 0 .01 0 M12 3v18 M12 3a9 9 0 0 1 0 18z",
        ["csv"] = "M6 3h8l4 4v14H6z M14 3v4h4 M9 13h6 M9 17h6",
        ["plus"] = "M12 5v14 M5 12h14",
        ["minus"] = "M5 12h14",
        ["search"] = "M10.5 4a6.5 6.5 0 1 0 .01 0 M15.5 15.5L21 21",
        ["pdf"] = "M6 3h8l4 4v14H6z M14 3v4h4 M8.5 17v-4h2a1.5 1.5 0 0 1 0 3h-2",
    };

    public static bool Has(string key) => D.ContainsKey(key);

    /// <summary>Icona di dimensione size; colore dalla risorsa (DynamicResource) o dal pennello indicato.</summary>
    public static FrameworkElement Make(string key, double size = 20, string brushKey = "TextBrush")
    {
        var p = new Path { Data = Geometry.Parse(D.GetValueOrDefault(key, "M4 4h16v16H4z")), StrokeThickness = 1.5, StrokeLineJoin = PenLineJoin.Round, StrokeStartLineCap = PenLineCap.Round, StrokeEndLineCap = PenLineCap.Round, Fill = Brushes.Transparent, Stretch = Stretch.None };
        p.SetResourceReference(Shape.StrokeProperty, brushKey);
        return new Viewbox { Width = size, Height = size, Child = new Canvas { Width = 24, Height = 24, Children = { p } }, IsHitTestVisible = false };
    }
}
