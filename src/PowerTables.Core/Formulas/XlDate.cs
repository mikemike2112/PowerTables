namespace PowerTables.Core.Formulas;

/// <summary>Date seriali di Excel, sistema 1900 con il bug del 29/02/1900 (serial 60).</summary>
public static class XlDate
{
    public const double MaxSerial = 2958465; // 31/12/9999

    // algoritmo di Howard Hinnant (calendario gregoriano prolettico)
    public static long DaysFromCivil(long y, int m, int d)
    {
        y -= m <= 2 ? 1 : 0;
        long era = (y >= 0 ? y : y - 399) / 400;
        long yoe = y - era * 400;
        long doy = (153 * (m + (m > 2 ? -3 : 9)) + 2) / 5 + d - 1;
        long doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
        return era * 146097 + doe - 719468;
    }

    public static (int y, int m, int d) CivilFromDays(long z)
    {
        z += 719468;
        long era = (z >= 0 ? z : z - 146096) / 146097;
        long doe = z - era * 146097;
        long yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
        long y = yoe + era * 400;
        long doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
        long mp = (5 * doy + 2) / 153;
        int d = (int)(doy - (153 * mp + 2) / 5 + 1);
        int m = (int)(mp < 10 ? mp + 3 : mp - 9);
        return ((int)(y + (m <= 2 ? 1 : 0)), m, d);
    }

    static readonly long Epoch = DaysFromCivil(1899, 12, 30);

    /// <summary>Serial da anno/mese/giorno (mese e giorno possono essere fuori intervallo). NaN se non valido.</summary>
    public static double ToSerial(long y, long m, long d)
    {
        if (y < 0 || y > 9999) return double.NaN;
        if (y < 1900) y += 1900;
        long mm = m - 1;
        y += (long)Math.Floor(mm / 12.0);
        m = ((mm % 12) + 12) % 12 + 1;
        if (y > 9999 || y < 1900) return double.NaN;
        double s = y == 1900 && m <= 2
            ? (m == 1 ? 0 : 31) + d
            : DaysFromCivil(y, (int)m, 1) - Epoch + d - 1;
        return s is < 0 or > MaxSerial ? double.NaN : s;
    }

    public static (int y, int m, int d) FromSerial(double serial)
    {
        long n = (long)Math.Floor(serial);
        if (n < 1) return (1900, 1, 0);
        if (n < 60) return n <= 31 ? (1900, 1, (int)n) : (1900, 2, (int)(n - 31));
        if (n == 60) return (1900, 2, 29);
        return CivilFromDays(Epoch + n);
    }

    /// <summary>1=domenica … 7=sabato (il 1/1/1900 è "domenica" secondo Excel).</summary>
    public static int DayOfWeek(double serial)
    {
        long n = (long)Math.Floor(serial);
        return (int)(((n - 1) % 7 + 7) % 7) + 1;
    }

    /// <summary>Secondi nella giornata (0..86400), arrotondati al secondo.</summary>
    public static long SecondsOfDay(double serial)
    {
        double frac = serial - Math.Floor(serial);
        return (long)Math.Round(frac * 86400, MidpointRounding.AwayFromZero);
    }
}
