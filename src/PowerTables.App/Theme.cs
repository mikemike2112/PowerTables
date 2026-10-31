using System.Windows;
using System.Windows.Media;

namespace PowerTables.App;

/// <summary>Colori dei temi scuro (predefinito) e chiaro. Accento PowerTables #3ecf8e.</summary>
public static class Theme
{
    public static bool Dark = true;
    public static readonly Color Accent = Color.FromRgb(0x3e, 0xcf, 0x8e);

    public static Color Bg => Dark ? C("#0e0e0e") : C("#f4f4f2");
    public static Color Panel => Dark ? C("#161616") : C("#ffffff");
    public static Color PanelAlt => Dark ? C("#1d1d1d") : C("#ebebe8");
    public static Color Border => Dark ? C("#2c2c2c") : C("#cfcfca");
    public static Color Text => Dark ? C("#ece6dc") : C("#1b1b1b");
    public static Color TextDim => Dark ? C("#8f8a82") : C("#6a6a66");
    public static Color CellBg => Dark ? C("#141414") : C("#ffffff");
    public static Color CellText => Dark ? C("#ece6dc") : C("#000000");
    public static Color GridLine => Dark ? C("#2a2a2a") : C("#dadad6");
    public static Color HeaderBg => Dark ? C("#1a1a1a") : C("#eeeeea");
    public static Color HeaderSel => Dark ? C("#24382f") : C("#d3eadf");

    static Color C(string s) => (Color)ColorConverter.ConvertFromString(s);

    public static void Apply(ResourceDictionary r)
    {
        void Set(string k, Color c) { var b = new SolidColorBrush(c); b.Freeze(); r[k] = b; }
        Set("BgBrush", Bg); Set("PanelBrush", Panel); Set("PanelAltBrush", PanelAlt); Set("BorderBrush2", Border);
        Set("TextBrush", Text); Set("TextDimBrush", TextDim); Set("AccentBrush", Accent);
        Set("HoverBrush", Dark ? C("#262626") : C("#e0e0dc"));
    }

    public static Brush Brush(Color c) { var b = new SolidColorBrush(c); b.Freeze(); return b; }
}
