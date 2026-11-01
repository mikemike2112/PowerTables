namespace PowerTables.Core.Formulas;

/// <summary>Elenco delle funzioni con categoria e nome italiano (solo per interfaccia e digitazione: nel file restano sempre i nomi inglesi).</summary>
public static class FunctionCatalog
{
    public sealed record Entry(string En, string It, string Category);

    static readonly (string Category, string Names)[] Cats =
    [
        ("Finanziarie", "PMT FV PV NPER NPV"),
        ("Logiche", "IF IFS AND OR NOT XOR IFERROR IFNA SWITCH TRUE FALSE CHOOSE"),
        ("Testo", "LEN LEFT RIGHT MID UPPER LOWER PROPER TRIM CLEAN EXACT REPT CHAR UNICHAR CODE UNICODE VALUE NUMBERVALUE CONCATENATE CONCAT TEXTJOIN FIND SEARCH SUBSTITUTE REPLACE TEXT T TEXTBEFORE TEXTAFTER DOLLAR FIXED"),
        ("Data e ora", "DATE YEAR MONTH DAY HOUR MINUTE SECOND TIME TODAY NOW WEEKDAY EDATE EOMONTH DAYS DATEDIF NETWORKDAYS WORKDAY WEEKNUM ISOWEEKNUM DATEVALUE TIMEVALUE"),
        ("Ricerca e riferimento", "VLOOKUP HLOOKUP LOOKUP XLOOKUP MATCH XMATCH INDEX OFFSET INDIRECT ADDRESS TRANSPOSE ROW COLUMN ROWS COLUMNS"),
        ("Matematiche", "SUM PRODUCT SUMIF SUMIFS SUMPRODUCT SUBTOTAL ABS SIGN INT SQRT EXP LN LOG LOG10 PI SIN COS TAN ASIN ACOS ATAN ATAN2 DEGREES RADIANS ROUND ROUNDUP ROUNDDOWN TRUNC MOD QUOTIENT POWER CEILING FLOOR MROUND EVEN ODD FACT GCD LCM RAND RANDBETWEEN"),
        ("Statistiche", "AVERAGE AVERAGEIF AVERAGEIFS COUNT COUNTA COUNTBLANK COUNTIF COUNTIFS MAX MIN MAXIFS MINIFS MEDIAN MODE MODE.SNGL LARGE SMALL RANK RANK.EQ STDEV STDEV.S STDEV.P STDEVP VAR VAR.S VAR.P VARP PERCENTILE PERCENTILE.INC CORREL"),
        ("Informative", "ISNUMBER ISTEXT ISLOGICAL ISERROR ISERR ISNA ISBLANK ISNONTEXT ISEVEN ISODD ISREF ISFORMULA FORMULATEXT NA N TYPE ERROR.TYPE SHEET SHEETS"),
    ];

    // nomi italiani di Excel (separatore ';' e virgola decimale)
    static readonly Dictionary<string, string> It = new()
    {
        ["SUM"] = "SOMMA", ["AVERAGE"] = "MEDIA", ["IF"] = "SE", ["COUNT"] = "CONTA.NUMERI", ["COUNTA"] = "CONTA.VALORI", ["COUNTBLANK"] = "CONTA.VUOTE",
        ["COUNTIF"] = "CONTA.SE", ["COUNTIFS"] = "CONTA.PIÙ.SE", ["SUMIF"] = "SOMMA.SE", ["SUMIFS"] = "SOMMA.PIÙ.SE", ["AVERAGEIF"] = "MEDIA.SE", ["AVERAGEIFS"] = "MEDIA.PIÙ.SE",
        ["MAXIFS"] = "MAX.PIÙ.SE", ["MINIFS"] = "MIN.PIÙ.SE", ["ABS"] = "ASS", ["ROUND"] = "ARROTONDA", ["ROUNDUP"] = "ARROTONDA.PER.ECC", ["ROUNDDOWN"] = "ARROTONDA.PER.DIF",
        ["MOD"] = "RESTO", ["POWER"] = "POTENZA", ["SQRT"] = "RADQ", ["PRODUCT"] = "PRODOTTO", ["SUMPRODUCT"] = "MATR.SOMMA.PRODOTTO", ["SIGN"] = "SEGNO",
        ["CEILING"] = "ARROTONDA.ECCESSO", ["FLOOR"] = "ARROTONDA.DIFETTO", ["TRUNC"] = "TRONCA", ["PI"] = "PI.GRECO", ["RAND"] = "CASUALE", ["RANDBETWEEN"] = "CASUALE.TRA",
        ["FACT"] = "FATTORIALE", ["GCD"] = "MCD", ["LCM"] = "MCM", ["EVEN"] = "PARI", ["ODD"] = "DISPARI", ["MEDIAN"] = "MEDIANA", ["MODE"] = "MODA", ["MODE.SNGL"] = "MODA.SNGL",
        ["LARGE"] = "GRANDE", ["SMALL"] = "PICCOLO", ["RANK"] = "RANGO", ["RANK.EQ"] = "RANGO.UG", ["STDEV"] = "DEV.ST", ["STDEV.S"] = "DEV.ST.C", ["STDEV.P"] = "DEV.ST.P",
        ["VAR.S"] = "VAR.C", ["CORREL"] = "CORRELAZIONE", ["SIN"] = "SEN", ["ASIN"] = "ARCSEN", ["ACOS"] = "ARCCOS", ["ATAN"] = "ARCTAN", ["ATAN2"] = "ARCTAN.2",
        ["DEGREES"] = "GRADI", ["RADIANS"] = "RADIANTI", ["QUOTIENT"] = "QUOZIENTE", ["MROUND"] = "ARROTONDA.MULTIPLO", ["SUBTOTAL"] = "SUBTOTALE",
        ["AND"] = "E", ["OR"] = "O", ["NOT"] = "NON", ["IFERROR"] = "SE.ERRORE", ["IFNA"] = "SE.NON.DISP.", ["IFS"] = "PIÙ.SE", ["CHOOSE"] = "SCEGLI",
        ["TRUE"] = "VERO", ["FALSE"] = "FALSO", ["ISNUMBER"] = "VAL.NUMERO", ["ISTEXT"] = "VAL.TESTO", ["ISLOGICAL"] = "VAL.LOGICO", ["ISERROR"] = "VAL.ERRORE",
        ["ISERR"] = "VAL.ERR", ["ISNA"] = "VAL.NON.DISP", ["ISBLANK"] = "VAL.VUOTO", ["ISNONTEXT"] = "VAL.NON.TESTO", ["ISEVEN"] = "VAL.PARI", ["ISODD"] = "VAL.DISPARI",
        ["NA"] = "NON.DISP", ["N"] = "NUM", ["TYPE"] = "TIPO", ["ERROR.TYPE"] = "ERRORE.TIPO", ["ROW"] = "RIF.RIGA", ["COLUMN"] = "RIF.COLONNA", ["ROWS"] = "RIGHE", ["COLUMNS"] = "COLONNE",
        ["LEN"] = "LUNGHEZZA", ["LEFT"] = "SINISTRA", ["RIGHT"] = "DESTRA", ["MID"] = "STRINGA.ESTRAI", ["UPPER"] = "MAIUSC", ["LOWER"] = "MINUSC", ["PROPER"] = "MAIUSC.INIZ",
        ["TRIM"] = "ANNULLA.SPAZI", ["CLEAN"] = "LIBERA", ["EXACT"] = "IDENTICO", ["REPT"] = "RIPETI", ["CHAR"] = "CODICE.CARATT", ["CODE"] = "CODICE", ["VALUE"] = "VALORE",
        ["CONCATENATE"] = "CONCATENA", ["TEXTJOIN"] = "TESTO.UNISCI", ["FIND"] = "TROVA", ["SEARCH"] = "RICERCA", ["SUBSTITUTE"] = "SOSTITUISCI", ["REPLACE"] = "RIMPIAZZA",
        ["TEXT"] = "TESTO", ["MATCH"] = "CONFRONTA", ["VLOOKUP"] = "CERCA.VERT", ["HLOOKUP"] = "CERCA.ORIZZ", ["LOOKUP"] = "CERCA", ["XLOOKUP"] = "CERCA.X", ["INDEX"] = "INDICE",
        ["OFFSET"] = "SCARTO", ["INDIRECT"] = "INDIRETTO", ["ADDRESS"] = "INDIRIZZO", ["TRANSPOSE"] = "MATR.TRASPOSTA", ["DATE"] = "DATA", ["YEAR"] = "ANNO", ["MONTH"] = "MESE",
        ["DAY"] = "GIORNO", ["HOUR"] = "ORA", ["MINUTE"] = "MINUTO", ["SECOND"] = "SECONDO", ["TIME"] = "ORARIO", ["TODAY"] = "OGGI", ["NOW"] = "ADESSO", ["WEEKDAY"] = "GIORNO.SETTIMANA",
        ["EDATE"] = "DATA.MESE", ["EOMONTH"] = "FINE.MESE", ["DAYS"] = "GIORNI", ["DATEDIF"] = "DATA.DIFF", ["NETWORKDAYS"] = "GIORNI.LAVORATIVI.TOT", ["WORKDAY"] = "GIORNO.LAVORATIVO",
        ["WEEKNUM"] = "NUM.SETTIMANA", ["ISOWEEKNUM"] = "NUM.SETTIMANA.ISO", ["DATEVALUE"] = "DATA.VALORE", ["TIMEVALUE"] = "ORARIO.VALORE",
        ["PMT"] = "RATA", ["FV"] = "VAL.FUT", ["PV"] = "VA", ["NPER"] = "NUM.RATE", ["NPV"] = "VAN",
    };

    static readonly Lazy<List<Entry>> AllLazy = new(() =>
    {
        var cat = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var (c, names) in Cats) foreach (var n in names.Split(' ')) cat[n] = c;
        return Fns.Names.Distinct().OrderBy(n => n, StringComparer.Ordinal).Select(n => new Entry(n, It.GetValueOrDefault(n, n), cat.GetValueOrDefault(n, "Altre"))).ToList();
    });

    public static IReadOnlyList<Entry> All => AllLazy.Value;
    public static IEnumerable<string> Categories => Cats.Select(c => c.Category);

    static readonly Lazy<Dictionary<string, string>> ItToEn = new(() =>
    {
        var d = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var kv in It) d[kv.Value] = kv.Key;
        return d;
    });

    /// <summary>Nome inglese di una funzione scritta in italiano; null se non è un nome italiano noto.</summary>
    public static string? EnglishOf(string italian) => ItToEn.Value.GetValueOrDefault(italian);

    public static string ItalianOf(string english) => It.GetValueOrDefault(english, english);
}
