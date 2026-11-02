using System.Globalization;

namespace PowerTables.Core;

public sealed partial class Workbook
{
    /// <summary>Ricerca obiettivo: cambia la cella <paramref name="change"/> (costante numerica) finché la formula in <paramref name="set"/> vale <paramref name="target"/>.
    /// Restituisce il valore trovato oppure null (nessuna modifica) se non converge. Un solo passo di annullamento.</summary>
    public double? GoalSeek(Sheet s, int setR, int setC, double target, int chR, int chC, double tolerance = 1e-7)
    {
        var sc = s.Get(setR, setC);
        if (sc.F == null) throw new ArgumentException("La cella da impostare deve contenere una formula.");
        var cc = s.Get(chR, chC);
        if (cc.F != null || cc.Kind is not (CellKind.Num or CellKind.Empty)) throw new ArgumentException("La cella da cambiare deve contenere un numero, non una formula o un testo.");
        if (setR == chR && setC == chC) throw new ArgumentException("Le due celle devono essere diverse.");
        double orig = cc.Kind == CellKind.Num ? cc.Num : 0; Cell? origCell = cc.Kind == CellKind.Empty ? null : cc;
        double? found = null;
        Do("Ricerca obiettivo", () =>
        {
            double? F(double x)
            {
                SetValue(s, chR, chC, x); Recalc();
                return s.GetValue(setR, setC) is double d && !double.IsNaN(d) && !double.IsInfinity(d) ? d - target : null;
            }
            double? x0 = orig, f0 = F(x0.Value);
            if (f0 != null && Math.Abs(f0.Value) <= tolerance) found = x0;
            // secante, poi ricerca di un intervallo con cambio di segno e bisezione
            if (found == null && f0 != null)
            {
                double x1 = orig + Math.Max(1, Math.Abs(orig) * 0.1); var f1 = F(x1);
                for (int i = 0; i < 100 && f1 != null && found == null; i++)
                {
                    if (Math.Abs(f1.Value) <= tolerance) { found = x1; break; }
                    if (f1.Value == f0.Value) break;
                    double x2 = x1 - f1.Value * (x1 - x0.Value) / (f1.Value - f0.Value); var f2 = F(x2);
                    if (f2 == null) break;
                    (x0, f0, x1, f1) = (x1, f1, x2, f2);
                }
            }
            if (found == null)
            {
                double step = Math.Max(1, Math.Abs(orig));
                for (int k = 0; k < 40 && found == null; k++)
                {
                    double a = orig - step, b = orig + step; var fa = F(a); var fb = F(b);
                    if (fa != null && fb != null && Math.Sign(fa.Value) != Math.Sign(fb.Value))
                    {
                        for (int i = 0; i < 200; i++)
                        {
                            double m = (a + b) / 2; var fm = F(m); if (fm == null) break;
                            if (Math.Abs(fm.Value) <= tolerance) { found = m; break; }
                            if (Math.Sign(fm.Value) == Math.Sign(fa.Value)) { a = m; fa = fm; } else b = m;
                        }
                        break;
                    }
                    step *= 2;
                }
            }
            if (found != null) SetValue(s, chR, chC, found.Value);
            else if (origCell != null) SetValue(s, chR, chC, orig); else SetValue(s, chR, chC, null);
        });
        return found;
    }

    static double TQuantile975(int df)
    {
        const double z = 1.959964;
        double v = df;
        return z + (Math.Pow(z, 3) + z) / (4 * v) + (5 * Math.Pow(z, 5) + 16 * Math.Pow(z, 3) + 3 * z) / (96 * v * v);
    }

    /// <summary>Foglio previsione (regressione lineare, non il modello ETS di Excel): crea un foglio con cronologia, valori, previsione e limiti al 95% e un grafico a linee. Un solo passo di annullamento.</summary>
    public Sheet CreateForecastSheet(Sheet s, (int R1, int C1, int R2, int C2) values, (int R1, int C1, int R2, int C2) timeline, int periods)
    {
        if (values.C1 != values.C2 || timeline.C1 != timeline.C2) throw new ArgumentException("Valori e cronologia devono essere intervalli di una sola colonna.");
        int n = values.R2 - values.R1 + 1;
        if (n != timeline.R2 - timeline.R1 + 1) throw new ArgumentException("Valori e cronologia devono avere lo stesso numero di righe.");
        if (n < 3) throw new ArgumentException("Servono almeno 3 punti.");
        if (periods < 1 || periods > 1000) throw new ArgumentException("Numero di periodi da 1 a 1000.");
        var xs = new double[n]; var ys = new double[n];
        for (int i = 0; i < n; i++)
        {
            if (s.GetValue(timeline.R1 + i, timeline.C1) is not double x || s.GetValue(values.R1 + i, values.C1) is not double y) throw new ArgumentException("Cronologia e valori devono contenere solo numeri o date.");
            xs[i] = x; ys[i] = y;
            if (i > 0 && xs[i] <= xs[i - 1]) throw new ArgumentException("La cronologia deve essere crescente.");
        }
        double step = (xs[^1] - xs[0]) / (n - 1), xm = xs.Average(), ym = ys.Average();
        double sxx = xs.Sum(x => (x - xm) * (x - xm)), sxy = xs.Zip(ys, (x, y) => (x - xm) * (y - ym)).Sum();
        double slope = sxy / sxx, icpt = ym - slope * xm;
        double sse = xs.Zip(ys, (x, y) => Math.Pow(y - (icpt + slope * x), 2)).Sum(), se = Math.Sqrt(sse / (n - 2)), t = TQuantile975(n - 2);
        double Lim(double x) => t * se * Math.Sqrt(1 + 1.0 / n + (x - xm) * (x - xm) / sxx);
        Sheet? ns = null;
        string baseName = "Previsione"; string name = baseName; int k = 1;
        while (Sheets.Any(x => string.Equals(x.Name, name, StringComparison.OrdinalIgnoreCase))) name = baseName + " " + ++k;
        Do("Foglio previsione", () =>
        {
            ns = InsertSheet(name);
            string[] h = ["Cronologia", "Valori", "Previsione", "Limite inferiore", "Limite superiore"];
            for (int c = 0; c < h.Length; c++) SetValue(ns, 1, c + 1, h[c]);
            uint tstyle = s.Get(timeline.R1, timeline.C1).Style;
            void TimeCell(int r, double x) { var cell = new Cell { Style = tstyle }; cell.SetValue(x); Write(ns, r, 1, cell); }
            for (int i = 0; i < n; i++) { TimeCell(i + 2, xs[i]); SetValue(ns, i + 2, 2, ys[i]); }
            // il primo punto della previsione coincide con l'ultimo valore reale, così le linee si congiungono
            SetValue(ns, n + 1, 3, ys[^1]); SetValue(ns, n + 1, 4, ys[^1]); SetValue(ns, n + 1, 5, ys[^1]);
            for (int j = 1; j <= periods; j++)
            {
                double x = xs[^1] + step * j, yh = icpt + slope * x, l = Lim(x); int r = n + 1 + j;
                TimeCell(r, x); SetValue(ns, r, 3, yh); SetValue(ns, r, 4, yh - l); SetValue(ns, r, 5, yh + l);
            }
            int last = n + 1 + periods; string q = Formulas.Lexer.QuoteSheet(name); string cat = $"{q}!$A$2:$A${last}";
            var spec = new ChartSpec("line", "Previsione", new[] { "B", "C", "D", "E" }.Select((col, i) => new ChartSeries(h[i + 1], $"{q}!${col}$2:${col}${last}", cat)).ToList());
            AddObject(ns, new SheetObject { Kind = ObjKind.Chart, Chart = spec, R1 = 2, C1 = 7, R2 = 17, C2 = 14 });
        });
        return ns!;
    }
}
