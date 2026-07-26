using System;
using System.Collections.Generic;

namespace XLibur.Excel.CalcEngine.Functions;

/// <summary>
/// The unit table behind CONVERT. Every unit is stored as a factor to its measure's base unit, so a
/// conversion is a multiply and a divide; temperature is the exception, being affine rather than
/// proportional, and is handled separately.
/// <para>
/// Unit names are case sensitive, deliberately: Excel distinguishes "Pica" (a point, 1/72 inch)
/// from "pica" (1/6 inch), and "T" (tesla) from "t" — and the metric prefixes need the same
/// distinction to tell milli from mega.
/// </para>
/// </summary>
internal static class UnitConversion
{
    private enum Measure
    {
        Mass,
        Distance,
        Time,
        Pressure,
        Force,
        Energy,
        Power,
        Magnetism,
        Temperature,
        Volume,
        Area,
        Information,
        Speed,
    }

    /// <param name="Measure">Which units this one can be converted to.</param>
    /// <param name="ToBase">Multiplier that takes a value in this unit to the measure's base unit.</param>
    /// <param name="Prefixes">Which families of prefix the unit accepts.</param>
    private readonly record struct Unit(Measure Measure, double ToBase, PrefixKind Prefixes);

    [Flags]
    private enum PrefixKind
    {
        None = 0,
        Metric = 1,
        Binary = 2,
    }

    private const double Inch = 0.0254;
    private const double Foot = 12 * Inch;
    private const double Yard = 36 * Inch;
    private const double Mile = 5280 * Foot;
    private const double NauticalMile = 1852;
    private const double LightYear = 9460730472580800d;
    private const double Parsec = 30856775814913672.789139379577965; // 648000/π AU.
    private const double Pound = 453.59237;
    private const double Gallon = 3.785411784; // Litres.
    private const double Atmosphere = 101325;

    /// <summary>
    /// Metric prefixes. "e" is deka rather than an SI symbol — Excel's spelling, kept for
    /// compatibility — and both "u" and "µ" are accepted for micro.
    /// </summary>
    private static readonly Dictionary<string, double> MetricPrefixes = new(StringComparer.Ordinal)
    {
        ["Y"] = 1e24,
        ["Z"] = 1e21,
        ["E"] = 1e18,
        ["P"] = 1e15,
        ["T"] = 1e12,
        ["G"] = 1e9,
        ["M"] = 1e6,
        ["k"] = 1e3,
        ["h"] = 1e2,
        ["da"] = 1e1,
        ["e"] = 1e1,
        ["d"] = 1e-1,
        ["c"] = 1e-2,
        ["m"] = 1e-3,
        ["u"] = 1e-6,
        ["µ"] = 1e-6,
        ["n"] = 1e-9,
        ["p"] = 1e-12,
        ["f"] = 1e-15,
        ["a"] = 1e-18,
        ["z"] = 1e-21,
        ["y"] = 1e-24,
    };

    /// <summary>Binary prefixes, which Excel allows only on the information units.</summary>
    private static readonly Dictionary<string, double> BinaryPrefixes = new(StringComparer.Ordinal)
    {
        ["Yi"] = 1024d * 1024 * 1024 * 1024 * 1024 * 1024 * 1024 * 1024,
        ["Zi"] = 1024d * 1024 * 1024 * 1024 * 1024 * 1024 * 1024,
        ["Ei"] = 1024d * 1024 * 1024 * 1024 * 1024 * 1024,
        ["Pi"] = 1024d * 1024 * 1024 * 1024 * 1024,
        ["Ti"] = 1024d * 1024 * 1024 * 1024,
        ["Gi"] = 1024d * 1024 * 1024,
        ["Mi"] = 1024d * 1024,
        ["ki"] = 1024d,
    };

    private static readonly Dictionary<string, Unit> Units = BuildUnits();

    /// <summary>
    /// Convert <paramref name="value"/> from one unit to another. Returns false when either unit is
    /// unknown or when the two measure different things, which is #N/A in Excel.
    /// </summary>
    internal static bool TryConvert(double value, string fromUnit, string toUnit, out double result)
    {
        result = 0;
        if (!TryResolve(fromUnit, out var from, out var fromFactor) ||
            !TryResolve(toUnit, out var to, out var toFactor))
        {
            return false;
        }

        if (from.Measure != to.Measure)
            return false;

        if (from.Measure == Measure.Temperature)
        {
            // A temperature scale has an offset as well as a step, so the prefixes that scale every
            // other measure have no meaning here and are not accepted.
            result = FromCelsius(ToCelsius(value, fromUnit), toUnit);
            return true;
        }

        result = value * (from.ToBase * fromFactor) / (to.ToBase * toFactor);
        return true;
    }

    /// <summary>
    /// Look a unit name up, falling back to splitting a prefix off the front. The exact name always
    /// wins, which is what stops "m" (metre) being read as milli-nothing and "T" (tesla) as tera.
    /// </summary>
    private static bool TryResolve(string name, out Unit unit, out double prefixFactor)
    {
        prefixFactor = 1;
        if (Units.TryGetValue(name, out unit))
            return true;

        // Longest prefix first, so "da" is preferred over "d" and "ki" over "k".
        for (var length = 2; length >= 1; length--)
        {
            if (name.Length <= length)
                continue;

            var prefix = name[..length];
            var rest = name[length..];
            if (!Units.TryGetValue(rest, out var candidate))
                continue;

            if (candidate.Prefixes.HasFlag(PrefixKind.Binary) && BinaryPrefixes.TryGetValue(prefix, out var binary))
            {
                unit = candidate;
                prefixFactor = binary;
                return true;
            }

            if (candidate.Prefixes.HasFlag(PrefixKind.Metric) && MetricPrefixes.TryGetValue(prefix, out var metric))
            {
                unit = candidate;
                prefixFactor = metric;
                return true;
            }
        }

        unit = default;
        return false;
    }

    private static double ToCelsius(double value, string unit) => unit switch
    {
        "C" or "cel" => value,
        "F" or "fah" => (value - 32) * 5 / 9,
        "K" or "kel" => value - 273.15,
        "Rank" => (value - 491.67) * 5 / 9,
        _ => value * 5 / 4, // Réaumur.
    };

    private static double FromCelsius(double celsius, string unit) => unit switch
    {
        "C" or "cel" => celsius,
        "F" or "fah" => celsius * 9 / 5 + 32,
        "K" or "kel" => celsius + 273.15,
        "Rank" => (celsius + 273.15) * 9 / 5,
        _ => celsius * 4 / 5, // Réaumur.
    };

    private static Dictionary<string, Unit> BuildUnits()
    {
        var units = new Dictionary<string, Unit>(StringComparer.Ordinal);

        void Add(Measure measure, double toBase, PrefixKind prefixes, params string[] names)
        {
            foreach (var name in names)
                units[name] = new Unit(measure, toBase, prefixes);
        }

        // Mass — base gram.
        Add(Measure.Mass, 1, PrefixKind.Metric, "g");
        Add(Measure.Mass, 14593.9029372064, PrefixKind.None, "sg");
        Add(Measure.Mass, Pound, PrefixKind.None, "lbm");
        Add(Measure.Mass, 1.66053886e-24, PrefixKind.Metric, "u");
        Add(Measure.Mass, Pound / 16, PrefixKind.None, "ozm");
        Add(Measure.Mass, Pound / 7000, PrefixKind.None, "grain");
        Add(Measure.Mass, Pound * 100, PrefixKind.None, "cwt", "shweight");
        Add(Measure.Mass, Pound * 112, PrefixKind.None, "uk_cwt", "lcwt", "hweight");
        Add(Measure.Mass, Pound * 14, PrefixKind.None, "stone");
        Add(Measure.Mass, Pound * 2000, PrefixKind.None, "ton");
        Add(Measure.Mass, Pound * 2240, PrefixKind.None, "uk_ton", "LTON", "brton");

        // Distance — base metre.
        Add(Measure.Distance, 1, PrefixKind.Metric, "m");
        Add(Measure.Distance, Mile, PrefixKind.None, "mi");
        Add(Measure.Distance, NauticalMile, PrefixKind.None, "Nmi");
        Add(Measure.Distance, Inch, PrefixKind.None, "in");
        Add(Measure.Distance, Foot, PrefixKind.None, "ft");
        Add(Measure.Distance, Yard, PrefixKind.None, "yd");
        Add(Measure.Distance, 1e-10, PrefixKind.Metric, "ang");
        Add(Measure.Distance, 1.143, PrefixKind.None, "ell");
        Add(Measure.Distance, LightYear, PrefixKind.None, "ly");
        Add(Measure.Distance, Parsec, PrefixKind.None, "parsec", "pc");
        Add(Measure.Distance, Inch / 72, PrefixKind.None, "Picapt", "Pica");
        Add(Measure.Distance, Inch / 6, PrefixKind.None, "pica");
        Add(Measure.Distance, 1609.347218694437, PrefixKind.None, "survey_mi");

        // Time — base second.
        Add(Measure.Time, 31557600, PrefixKind.None, "yr");
        Add(Measure.Time, 86400, PrefixKind.None, "day", "d");
        Add(Measure.Time, 3600, PrefixKind.None, "hr");
        Add(Measure.Time, 60, PrefixKind.None, "mn", "min");
        Add(Measure.Time, 1, PrefixKind.Metric, "sec", "s");

        // Pressure — base pascal. Excel treats mmHg and torr as the same unit, 1/760 atm.
        Add(Measure.Pressure, 1, PrefixKind.Metric, "Pa", "p");
        Add(Measure.Pressure, Atmosphere, PrefixKind.Metric, "atm", "at");
        Add(Measure.Pressure, Atmosphere / 760, PrefixKind.Metric, "mmHg", "Torr");
        Add(Measure.Pressure, 6894.75729316836, PrefixKind.None, "psi");

        // Force — base newton.
        Add(Measure.Force, 1, PrefixKind.Metric, "N");
        Add(Measure.Force, 1e-5, PrefixKind.Metric, "dyn", "dy");
        Add(Measure.Force, 4.4482216152605, PrefixKind.None, "lbf");
        Add(Measure.Force, 9.80665e-3, PrefixKind.Metric, "pond");

        // Energy — base joule.
        Add(Measure.Energy, 1, PrefixKind.Metric, "J");
        Add(Measure.Energy, 1e-7, PrefixKind.Metric, "e");
        Add(Measure.Energy, 4.184, PrefixKind.Metric, "c");
        Add(Measure.Energy, 4.1868, PrefixKind.Metric, "cal");
        Add(Measure.Energy, 1.60217646e-19, PrefixKind.Metric, "eV", "ev");
        Add(Measure.Energy, 2684519.53769617, PrefixKind.None, "HPh", "hh");
        Add(Measure.Energy, 3600, PrefixKind.Metric, "Wh", "wh");
        Add(Measure.Energy, 1.3558179483314, PrefixKind.None, "flb");
        Add(Measure.Energy, 1055.05585262, PrefixKind.None, "BTU", "btu");

        // Power — base watt.
        Add(Measure.Power, 745.69987158227, PrefixKind.None, "HP", "h");
        Add(Measure.Power, 735.49875, PrefixKind.None, "PS");
        Add(Measure.Power, 1, PrefixKind.Metric, "W", "w");

        // Magnetism — base tesla.
        Add(Measure.Magnetism, 1, PrefixKind.Metric, "T");
        Add(Measure.Magnetism, 1e-4, PrefixKind.Metric, "ga");

        // Temperature — the factor is unused; the conversion is affine.
        Add(Measure.Temperature, 1, PrefixKind.None, "C", "cel", "F", "fah", "K", "kel", "Rank", "Reau");

        // Volume — base litre.
        Add(Measure.Volume, Gallon / 768, PrefixKind.None, "tsp");
        Add(Measure.Volume, 5e-3, PrefixKind.None, "tspm");
        Add(Measure.Volume, Gallon / 256, PrefixKind.None, "tbs");
        Add(Measure.Volume, Gallon / 128, PrefixKind.None, "oz");
        Add(Measure.Volume, Gallon / 16, PrefixKind.None, "cup");
        Add(Measure.Volume, Gallon / 8, PrefixKind.None, "pt", "us_pt");
        Add(Measure.Volume, 4.54609 / 8, PrefixKind.None, "uk_pt");
        Add(Measure.Volume, Gallon / 4, PrefixKind.None, "qt");
        Add(Measure.Volume, 4.54609 / 4, PrefixKind.None, "uk_qt");
        Add(Measure.Volume, Gallon, PrefixKind.None, "gal");
        Add(Measure.Volume, 4.54609, PrefixKind.None, "uk_gal");
        Add(Measure.Volume, 1, PrefixKind.Metric, "l", "L", "lt");
        Add(Measure.Volume, 1e-27, PrefixKind.None, "ang3", "ang^3");
        Add(Measure.Volume, 158.987294928, PrefixKind.None, "barrel");
        Add(Measure.Volume, 35.23907016688, PrefixKind.None, "bushel");
        Add(Measure.Volume, Cube(Foot) * 1000, PrefixKind.None, "ft3", "ft^3");
        Add(Measure.Volume, Cube(Inch) * 1000, PrefixKind.None, "in3", "in^3");
        Add(Measure.Volume, Cube(LightYear) * 1000, PrefixKind.None, "ly3", "ly^3");
        Add(Measure.Volume, 1000, PrefixKind.None, "m3", "m^3");
        Add(Measure.Volume, Cube(Mile) * 1000, PrefixKind.None, "mi3", "mi^3");
        Add(Measure.Volume, Cube(Yard) * 1000, PrefixKind.None, "yd3", "yd^3");
        Add(Measure.Volume, Cube(NauticalMile) * 1000, PrefixKind.None, "Nmi3", "Nmi^3");
        Add(Measure.Volume, Cube(Inch / 72) * 1000, PrefixKind.None, "Picapt3", "Picapt^3", "Pica3", "Pica^3");
        Add(Measure.Volume, 2831.6846592, PrefixKind.None, "GRT", "regton");
        Add(Measure.Volume, 1132.67386368, PrefixKind.None, "MTON");

        // Area — base square metre.
        Add(Measure.Area, 4046.8564224, PrefixKind.None, "uk_acre");
        Add(Measure.Area, 4046.87260987425, PrefixKind.None, "us_acre");
        Add(Measure.Area, 1e-20, PrefixKind.None, "ang2", "ang^2");
        Add(Measure.Area, 100, PrefixKind.None, "ar");
        Add(Measure.Area, Square(Foot), PrefixKind.None, "ft2", "ft^2");
        Add(Measure.Area, 10000, PrefixKind.None, "ha");
        Add(Measure.Area, Square(Inch), PrefixKind.None, "in2", "in^2");
        Add(Measure.Area, Square(LightYear), PrefixKind.None, "ly2", "ly^2");
        Add(Measure.Area, 1, PrefixKind.None, "m2", "m^2");
        Add(Measure.Area, 2500, PrefixKind.None, "Morgen");
        Add(Measure.Area, Square(Mile), PrefixKind.None, "mi2", "mi^2");
        Add(Measure.Area, Square(NauticalMile), PrefixKind.None, "Nmi2", "Nmi^2");
        Add(Measure.Area, Square(Inch / 72), PrefixKind.None, "Picapt2", "Picapt^2", "Pica2", "Pica^2");
        Add(Measure.Area, Square(Yard), PrefixKind.None, "yd2", "yd^2");

        // Information — base bit; these are the only units that take a binary prefix.
        Add(Measure.Information, 1, PrefixKind.Metric | PrefixKind.Binary, "bit");
        Add(Measure.Information, 8, PrefixKind.Metric | PrefixKind.Binary, "byte");

        // Speed — base metre per second.
        Add(Measure.Speed, 0.514773333333333, PrefixKind.None, "admkn");
        Add(Measure.Speed, NauticalMile / 3600, PrefixKind.None, "kn");
        Add(Measure.Speed, 1 / 3600d, PrefixKind.Metric, "m/h", "m/hr");
        Add(Measure.Speed, 1, PrefixKind.Metric, "m/s", "m/sec");
        Add(Measure.Speed, Mile / 3600, PrefixKind.None, "mph");

        return units;
    }

    private static double Square(double value) => value * value;

    private static double Cube(double value) => value * value * value;
}
