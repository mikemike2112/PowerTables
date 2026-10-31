using System.Globalization;

namespace PowerTables.Core.Formulas;

public static partial class Fns
{
    static double Serial(object? v)
    {
        double d = V.Num(v);
        if (d < 0 || d >= XlDate.MaxSerial + 1) throw new ExcelEx(CellError.Num);
        return d;
    }

    static object? SerialOrNum(double s) => double.IsNaN(s) ? CellError.Num : s;

    static double AddMonths(double serial, int months)
    {
        var (y, m, d) = XlDate.FromSerial(serial);
        long tm = (long)y * 12 + (m - 1) + months;
        int ny = (int)Math.Floor(tm / 12.0), nm = (int)(tm - (long)ny * 12) + 1;
        int dim = nm == 2 ? (DateTime.IsLeapYear(ny) ? 29 : 28) : nm is 4 or 6 or 9 or 11 ? 30 : 31;
        if (ny == 1900 && nm == 2) dim = 29; // 1900 "bisestile" per Excel
        return XlDate.ToSerial(ny, nm, Math.Min(d, dim));
    }

    static bool IsWeekend(double s) { int w = XlDate.DayOfWeek(s); return w == 1 || w == 7; }

    static HashSet<long> Holidays(object? h) => h == null || h is Missing ? [] : Flat(h).OfType<double>().Select(d => (long)Math.Floor(d)).ToHashSet();

    static void RegisterDate()
    {
        S("DATE", 3, 3, a => SerialOrNum(XlDate.ToSerial((long)Math.Truncate(V.Num(a[0])), (long)Math.Truncate(V.Num(a[1])), (long)Math.Truncate(V.Num(a[2])))), true);
        S("YEAR", 1, 1, a => (double)XlDate.FromSerial(Serial(a[0])).y, true);
        S("MONTH", 1, 1, a => (double)XlDate.FromSerial(Serial(a[0])).m, true);
        S("DAY", 1, 1, a => (double)XlDate.FromSerial(Serial(a[0])).d, true);
        S("HOUR", 1, 1, a => (double)(XlDate.SecondsOfDay(Serial(a[0])) % 86400 / 3600), true);
        S("MINUTE", 1, 1, a => (double)(XlDate.SecondsOfDay(Serial(a[0])) % 3600 / 60), true);
        S("SECOND", 1, 1, a => (double)(XlDate.SecondsOfDay(Serial(a[0])) % 60), true);
        S("TIME", 3, 3, a =>
        {
            double h = Math.Truncate(V.Num(a[0])), m = Math.Truncate(V.Num(a[1])), s = Math.Truncate(V.Num(a[2]));
            if (h < 0 || m < 0 || s < 0 || h > 32767 || m > 32767 || s > 32767) return CellError.Num;
            double t = h * 3600 + m * 60 + s;
            return (t % 86400) / 86400;
        }, true);
        L("TODAY", 0, 0, (_, _) => Math.Floor(Now()));
        L("NOW", 0, 0, (_, _) => Now());
        S("WEEKDAY", 1, 2, a =>
        {
            double s = Serial(a[0]); int type = a.Length > 1 ? In(a[1]) : 1;
            int w = XlDate.DayOfWeek(s); // 1=Dom
            return type switch
            {
                1 => (double)w,
                2 => (double)((w + 5) % 7 + 1),
                3 => (double)((w + 5) % 7),
                >= 11 and <= 17 => (double)((w - (type == 17 ? 1 : type - 9) + 7) % 7 + 1),
                _ => CellError.Num,
            };
        }, true);
        S("EDATE", 2, 2, a => SerialOrNum(AddMonths(Math.Floor(Serial(a[0])), In(a[1]))), true);
        S("EOMONTH", 2, 2, a =>
        {
            double s = AddMonths(Math.Floor(Serial(a[0])), In(a[1]));
            if (double.IsNaN(s)) return CellError.Num;
            var (y, m, _) = XlDate.FromSerial(s);
            return SerialOrNum(XlDate.ToSerial(y, m + 1, 0));
        }, true);
        S("DAYS", 2, 2, a => Math.Floor(Serial(a[0])) - Math.Floor(Serial(a[1])), true);
        S("DATEDIF", 3, 3, a =>
        {
            double s = Math.Floor(Serial(a[0])), e = Math.Floor(Serial(a[1]));
            if (s > e) return CellError.Num;
            var (y1, m1, d1) = XlDate.FromSerial(s); var (y2, m2, d2) = XlDate.FromSerial(e);
            switch (St(a[2]).ToUpperInvariant())
            {
                case "D": return e - s;
                case "M": return (double)((y2 - y1) * 12 + m2 - m1 - (d2 < d1 ? 1 : 0));
                case "Y": return (double)(y2 - y1 - (m2 < m1 || (m2 == m1 && d2 < d1) ? 1 : 0));
                case "YM": return (double)(((m2 - m1 - (d2 < d1 ? 1 : 0)) % 12 + 12) % 12);
                case "MD":
                    {
                        if (d2 >= d1) return (double)(d2 - d1);
                        var (py, pm, _) = XlDate.FromSerial(XlDate.ToSerial(y2, m2, 0));
                        int dim = (int)(XlDate.ToSerial(y2, m2, 0) - XlDate.ToSerial(py, pm, 0));
                        return (double)(dim - d1 + d2);
                    }
                case "YD":
                    {
                        double ref1 = XlDate.ToSerial(y2, m1, d1);
                        if (ref1 > e) ref1 = XlDate.ToSerial(y2 - 1, m1, d1);
                        return e - ref1;
                    }
                default: return CellError.Num;
            }
        });
        R("NETWORKDAYS", 2, 3, (cx, a) =>
        {
            double s = Math.Floor(Serial(V.Scalar(a[0], cx))), e = Math.Floor(Serial(V.Scalar(a[1], cx)));
            var hol = Holidays(a.Length > 2 ? a[2] : null);
            int sign = s <= e ? 1 : -1; if (sign < 0) (s, e) = (e, s);
            double n = 0;
            for (double d = s; d <= e; d++) if (!IsWeekend(d) && !hol.Contains((long)d)) n++;
            return n * sign;
        });
        R("WORKDAY", 2, 3, (cx, a) =>
        {
            double s = Math.Floor(Serial(V.Scalar(a[0], cx))); int days = In(V.Scalar(a[1], cx));
            var hol = Holidays(a.Length > 2 ? a[2] : null);
            int step = days < 0 ? -1 : 1; int left = Math.Abs(days);
            while (left > 0) { s += step; if (!IsWeekend(s) && !hol.Contains((long)s)) left--; }
            return s;
        });
        S("WEEKNUM", 1, 2, a =>
        {
            double s = Math.Floor(Serial(a[0])); int type = a.Length > 1 ? In(a[1]) : 1;
            if (type == 21) return IsoWeek(s);
            int startDow = type is 2 or 11 ? 2 : type == 1 ? 1 : type >= 12 && type <= 17 ? type - 10 : -1;
            if (startDow < 0) return CellError.Num;
            var (y, _, _) = XlDate.FromSerial(s);
            double jan1 = XlDate.ToSerial(y, 1, 1);
            int off = ((XlDate.DayOfWeek(jan1) - startDow) % 7 + 7) % 7;
            return Math.Floor((s - jan1 + off) / 7) + 1;
        }, true);
        S("ISOWEEKNUM", 1, 1, a => IsoWeek(Math.Floor(Serial(a[0]))), true);
        S("DATEVALUE", 1, 1, a =>
        {
            string t = St(a[0]).Trim();
            if (DateTime.TryParseExact(t, ["yyyy-MM-dd", "yyyy/MM/dd", "d-MMM-yyyy", "MMMM d, yyyy", "yyyy-MM-ddTHH:mm:ss", "yyyy-MM-dd HH:mm:ss"], CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
                return XlDate.ToSerial(dt.Year, dt.Month, dt.Day);
            return CellError.Value;
        }, true);
        S("TIMEVALUE", 1, 1, a =>
        {
            if (DateTime.TryParseExact(St(a[0]).Trim(), ["H:mm", "H:mm:ss", "h:mm tt", "h:mm:ss tt"], CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
                return dt.TimeOfDay.TotalSeconds / 86400;
            return CellError.Value;
        }, true);
    }

    static double IsoWeek(double s)
    {
        int dow = (XlDate.DayOfWeek(s) + 5) % 7; // lun=0
        double thursday = s - dow + 3;
        var (ty, _, _) = XlDate.FromSerial(thursday);
        double jan1 = XlDate.ToSerial(ty, 1, 1);
        return Math.Floor((thursday - jan1) / 7) + 1;
    }

    public static Func<double> Clock = () => XlDate.ToSerial(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day) + DateTime.Now.TimeOfDay.TotalDays;
    static double Now() => Clock();

    static void RegisterFinance()
    {
        S("PMT", 3, 5, a =>
        {
            double rate = V.Num(a[0]), n = V.Num(a[1]), pv = V.Num(a[2]), fv = a.Length > 3 ? V.Num(a[3]) : 0, type = a.Length > 4 ? V.Num(a[4]) : 0;
            if (n == 0) return CellError.Num;
            if (rate == 0) return -(pv + fv) / n;
            double f = Math.Pow(1 + rate, n);
            return -(pv * f + fv) * rate / ((f - 1) * (1 + rate * (type != 0 ? 1 : 0)));
        });
        S("FV", 3, 5, a =>
        {
            double rate = V.Num(a[0]), n = V.Num(a[1]), pmt = V.Num(a[2]), pv = a.Length > 3 ? V.Num(a[3]) : 0, type = a.Length > 4 ? V.Num(a[4]) : 0;
            if (rate == 0) return -(pv + pmt * n);
            double f = Math.Pow(1 + rate, n);
            return -(pv * f + pmt * (1 + rate * (type != 0 ? 1 : 0)) * (f - 1) / rate);
        });
        S("PV", 3, 5, a =>
        {
            double rate = V.Num(a[0]), n = V.Num(a[1]), pmt = V.Num(a[2]), fv = a.Length > 3 ? V.Num(a[3]) : 0, type = a.Length > 4 ? V.Num(a[4]) : 0;
            if (rate == 0) return -(fv + pmt * n);
            double f = Math.Pow(1 + rate, n);
            return -(fv + pmt * (1 + rate * (type != 0 ? 1 : 0)) * (f - 1) / rate) / f;
        });
        S("NPER", 3, 5, a =>
        {
            double rate = V.Num(a[0]), pmt = V.Num(a[1]), pv = V.Num(a[2]), fv = a.Length > 3 ? V.Num(a[3]) : 0, type = a.Length > 4 ? V.Num(a[4]) : 0;
            if (rate == 0) return pmt == 0 ? CellError.Num : -(pv + fv) / pmt;
            double k = pmt * (1 + rate * (type != 0 ? 1 : 0)) / rate;
            double num = k - fv, den = k + pv;
            if (den == 0 || num / den <= 0) return CellError.Num;
            return Math.Log(num / den) / Math.Log(1 + rate);
        });
        R("NPV", 2, 255, (cx, a) =>
        {
            double rate = V.Num(V.Scalar(a[0], cx)), s = 0; int i = 1;
            foreach (var d in Nums(a.Skip(1).ToArray())) s += d / Math.Pow(1 + rate, i++);
            return s;
        });
    }
}
