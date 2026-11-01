namespace PowerTables.Core.IO;

/// <summary>Logica pura della richiesta di salvataggio (testabile senza interfaccia).</summary>
public static class SaveLogic
{
    public enum Choice { Save, DontSave, Cancel }

    /// <summary>Esito di una scelta: Proceed = l'azione che ha causato la richiesta (chiusura, Nuovo, Apri) può continuare; WriteFile = bisogna scrivere il file prima.</summary>
    public readonly record struct Outcome(bool Proceed, bool WriteFile);

    public static Outcome Map(Choice c) => c switch
    {
        Choice.Save => new(true, true),
        Choice.DontSave => new(true, false),
        _ => new(false, false),
    };

    static readonly string[] Reserved = ["CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"];

    /// <summary>Messaggio d'errore se il nome file non è valido, altrimenti null.</summary>
    public static string? ValidateName(string? name)
    {
        if (string.IsNullOrWhiteSpace(name)) return "Inserisci un nome per il file.";
        name = name.Trim();
        if (name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) return "Il nome non può contenere i caratteri \\ / : * ? \" < > |";
        if (name.EndsWith('.') || name.EndsWith(' ')) return "Il nome non può terminare con un punto o uno spazio.";
        if (name.Length > 200) return "Il nome è troppo lungo.";
        string stem = name.Split('.')[0].ToUpperInvariant();
        if (Reserved.Contains(stem)) return "Questo nome è riservato dal sistema.";
        return null;
    }

    public static string? ValidateFolder(string? folder)
    {
        if (string.IsNullOrWhiteSpace(folder)) return "Scegli una cartella.";
        try
        {
            if (!Path.IsPathRooted(folder)) return "Il percorso deve essere completo.";
            if (!Directory.Exists(folder)) return "La cartella non esiste.";
        }
        catch (ArgumentException) { return "Il percorso non è valido."; }
        return null;
    }

    /// <summary>Percorso finale: aggiunge l'estensione se manca (ext con punto, es. ".xlsx").</summary>
    public static string BuildPath(string folder, string name, string ext)
    {
        name = name.Trim();
        if (!name.EndsWith(ext, StringComparison.OrdinalIgnoreCase)) name += ext;
        return Path.Combine(folder, name);
    }

    /// <summary>Un file esistente non va mai sovrascritto senza conferma esplicita.</summary>
    public static bool NeedsOverwriteConfirmation(string path) => File.Exists(path);
}
