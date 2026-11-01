using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Media;

namespace PowerTables.App;

/// <summary>Temi scuro (predefinito) e chiaro. Accento app #46B862 (campionato dall'icona); scheda "Powers" rosso fuoco #FF2400. Il contenuto del foglio non cambia con il tema.</summary>
public static class Theme
{
    public static bool Dark = true;
    public static readonly Color Accent = Color.FromRgb(0x46, 0xb8, 0x62);
    public static readonly Color Powers = Color.FromRgb(0xFF, 0x24, 0x00);
    public static event Action? Changed;

    public static Color Bg => Dark ? C("#0e0e0e") : C("#efe9e8");
    public static Color Panel => Dark ? C("#181818") : C("#fbfafa");
    public static Color PanelAlt => Dark ? C("#222222") : C("#ffffff");
    public static Color Border => Dark ? C("#303030") : C("#d9d3d1");
    public static Color Text => Dark ? C("#ece6dc") : C("#1b1b1b");
    public static Color TextDim => Dark ? C("#9a948a") : C("#666666");
    public static Color Hover => Dark ? C("#2a2a2a") : C("#e6e0de");
    // il foglio resta com'è nel file: sfondo bianco, testo automatico nero, in entrambi i temi
    public static Color CellBg => Colors.White;
    public static Color CellText => Colors.Black;
    public static Color GridLine => C("#d4d4d4");
    public static Color HeaderBg => Dark ? C("#1a1a1a") : C("#f3efee");
    public static Color HeaderSel => Dark ? C("#24382f") : C("#d3eadf");

    static Color C(string s) => (Color)ColorConverter.ConvertFromString(s);

    public static void Apply(ResourceDictionary r)
    {
        void Set(string k, Color c) { var b = new SolidColorBrush(c); b.Freeze(); r[k] = b; }
        Set("BgBrush", Bg); Set("PanelBrush", Panel); Set("PanelAltBrush", PanelAlt); Set("BorderBrush2", Border);
        Set("TextBrush", Text); Set("TextDimBrush", TextDim); Set("AccentBrush", Accent); Set("PowersBrush", Powers);
        Set("HoverBrush", Hover);
        Changed?.Invoke();
    }

    public static Brush Brush(Color c) { var b = new SolidColorBrush(c); b.Freeze(); return b; }

    // ---- impostazioni persistenti: %APPDATA%\PowerTables\settings.json ----
    public static string SettingsPath => Environment.GetEnvironmentVariable("PT_SETTINGS") ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PowerTables", "settings.json");

    /// <summary>Legge le impostazioni; file mancante o corrotto → predefiniti (scuro), senza eccezioni.</summary>
    public static void Load()
    {
        Dark = true;
        try
        {
            if (!File.Exists(SettingsPath)) return;
            using var d = JsonDocument.Parse(File.ReadAllText(SettingsPath));
            if (d.RootElement.ValueKind == JsonValueKind.Object && d.RootElement.TryGetProperty("dark", out var v) && v.ValueKind is JsonValueKind.True or JsonValueKind.False) Dark = v.GetBoolean();
        }
        catch (Exception e) when (e is IOException or JsonException or UnauthorizedAccessException or InvalidOperationException) { Dark = true; }
    }

    public static void Save()
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath)!);
            File.WriteAllText(SettingsPath, JsonSerializer.Serialize(new { dark = Dark }));
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { /* impostazioni non salvabili: si continua */ }
    }
}
