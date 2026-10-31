using System.Globalization;
using System.Text;
using PowerTables.Core.Formulas;

namespace PowerTables.Core.IO;

/// <summary>Import/export CSV (RFC 4180, UTF-8; separatore rilevato).</summary>
public static class Csv
{
    public static char DetectDelimiter(string text)
    {
        int end = text.IndexOf('\n'); if (end < 0) end = text.Length;
        var line = text.AsSpan(0, Math.Min(end, 4096));
        bool q = false; int comma = 0, semi = 0, tab = 0;
        foreach (var ch in line) { if (ch == '"') q = !q; else if (!q) { if (ch == ',') comma++; else if (ch == ';') semi++; else if (ch == '\t') tab++; } }
        return tab > comma && tab > semi ? '\t' : semi > comma ? ';' : ',';
    }

    public static List<string[]> Parse(string text, char delim)
    {
        var rows = new List<string[]>(); var row = new List<string>(); var sb = new StringBuilder();
        bool q = false, any = false;
        for (int i = 0; i < text.Length; i++)
        {
            char c = text[i];
            if (q)
            {
                if (c == '"') { if (i + 1 < text.Length && text[i + 1] == '"') { sb.Append('"'); i++; } else q = false; }
                else sb.Append(c);
            }
            else if (c == '"' && sb.Length == 0) { q = true; any = true; }
            else if (c == delim) { row.Add(sb.ToString()); sb.Clear(); any = true; }
            else if (c == '\n' || c == '\r')
            {
                if (c == '\r' && i + 1 < text.Length && text[i + 1] == '\n') i++;
                if (any || sb.Length > 0 || row.Count > 0) { row.Add(sb.ToString()); rows.Add([.. row]); }
                row.Clear(); sb.Clear(); any = false;
            }
            else { sb.Append(c); any = true; }
        }
        if (any || sb.Length > 0 || row.Count > 0) { row.Add(sb.ToString()); rows.Add([.. row]); }
        return rows;
    }

    /// <summary>Importa in un foglio (dalla cella A1). Le celle che iniziano con '=' restano testo (niente formule da dati non fidati).</summary>
    public static void Import(Workbook wb, Sheet s, string text, char? delimiter = null, CultureInfo? ci = null)
    {
        if (text.Length > 0 && text[0] == '﻿') text = text[1..];
        char d = delimiter ?? DetectDelimiter(text);
        ci ??= d == ';' ? CultureInfo.CurrentCulture : CultureInfo.InvariantCulture;
        var rows = Parse(text, d);
        wb.Do("Importa CSV", () =>
        {
            for (int i = 0; i < rows.Count && i < Addr.MaxRow; i++)
                for (int j = 0; j < rows[i].Length && j < Addr.MaxCol; j++)
                {
                    string f = rows[i][j];
                    if (f.Length == 0) continue;
                    if (f[0] == '=') { wb.SetValue(s, i + 1, j + 1, f); continue; }
                    var (v, fmt) = Workbook.ParseInput(f, ci);
                    var cell = new Cell(); cell.SetValue(v);
                    if (fmt != 0) cell.Style = wb.Styles.Derive(0, x => x with { NumFmtId = fmt });
                    wb.Write(s, i + 1, j + 1, cell);
                }
        });
    }

    public static void Export(Workbook wb, Sheet s, string path, char delim = ',')
    {
        var sb = new StringBuilder();
        for (int r = 1; r <= s.MaxRow; r++)
        {
            for (int c = 1; c <= s.MaxCol; c++)
            {
                if (c > 1) sb.Append(delim);
                var cell = s.Get(r, c);
                string f = cell.Kind switch
                {
                    CellKind.Num => NumFmt.IsDate(wb.Styles.FormatCode(cell.Style)) && cell.Style != 0 ? NumFmt.Format(cell.Num, "yyyy-mm-dd" + (cell.Num % 1 != 0 ? " hh:mm:ss" : ""), CultureInfo.InvariantCulture) : Fmt.NumToText(cell.Num),
                    CellKind.Str => (string)cell.Obj!,
                    CellKind.Bool => cell.Num != 0 ? "TRUE" : "FALSE",
                    CellKind.Err => Errors.Text((CellError)(int)cell.Num),
                    _ => "",
                };
                if (f.IndexOfAny([delim, '"', '\n', '\r']) >= 0 || (f.Length > 0 && (f[0] == ' ' || f[^1] == ' '))) sb.Append('"').Append(f.Replace("\"", "\"\"")).Append('"');
                else sb.Append(f);
            }
            sb.Append("\r\n");
        }
        string tmp = path + ".tmp~";
        File.WriteAllText(tmp, sb.ToString(), new UTF8Encoding(true));
        File.Move(tmp, path, true);
    }
}
