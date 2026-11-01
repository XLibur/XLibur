using System.Text.RegularExpressions;

namespace PerfReport;

/// <summary>
/// Orders records by release version, not by when they were recorded. A re-run of an old release or a
/// hotfix tag cut after a newer release then still lands where it belongs in the history.
/// </summary>
public static partial class VersionOrder
{
    [GeneratedRegex(@"^v?(?<major>\d+)\.(?<minor>\d+)\.(?<patch>\d+)(?:-(?<pre>[0-9A-Za-z.-]+))?(?:\+[0-9A-Za-z.-]+)?$",
        RegexOptions.CultureInvariant, matchTimeoutMilliseconds: 1000)]
    private static partial Regex SemVer();

    /// <summary>Compares two records: parseable versions by semver, then the rest by recording time.</summary>
    public static int Compare(ReleaseRecord a, ReleaseRecord b)
    {
        var pa = Parse(a.Version);
        var pb = Parse(b.Version);
        if (pa is not null && pb is not null)
        {
            var c = CompareParsed(pa.Value, pb.Value);
            if (c != 0)
                return c;
        }
        else if (pa is not null || pb is not null)
        {
            return pa is not null ? -1 : 1;
        }

        var byTime = string.CompareOrdinal(a.RecordedAt, b.RecordedAt);
        return byTime != 0 ? byTime : string.CompareOrdinal(a.Version, b.Version);
    }

    /// <summary>True when <paramref name="a"/> is a lower version than <paramref name="b"/>. False if either is unparseable.</summary>
    public static bool IsLower(ReleaseRecord a, ReleaseRecord b)
    {
        var pa = Parse(a.Version);
        var pb = Parse(b.Version);
        return pa is not null && pb is not null && CompareParsed(pa.Value, pb.Value) < 0;
    }

    public static bool IsParseable(string version) => Parse(version) is not null;

    public static List<ReleaseRecord> Sort(IEnumerable<ReleaseRecord> records)
    {
        var list = records.ToList();
        list.Sort(Compare);
        return list;
    }

    private readonly record struct Parsed(long Major, long Minor, long Patch, string[] Pre);

    private static Parsed? Parse(string version)
    {
        var m = SemVer().Match(version);
        if (!m.Success)
            return null;
        if (!long.TryParse(m.Groups["major"].Value, out var major)
            || !long.TryParse(m.Groups["minor"].Value, out var minor)
            || !long.TryParse(m.Groups["patch"].Value, out var patch))
        {
            return null;
        }

        var pre = m.Groups["pre"].Success ? m.Groups["pre"].Value.Split('.') : [];
        return new Parsed(major, minor, patch, pre);
    }

    private static int CompareParsed(Parsed a, Parsed b)
    {
        var c = a.Major.CompareTo(b.Major);
        if (c != 0) return c;
        c = a.Minor.CompareTo(b.Minor);
        if (c != 0) return c;
        c = a.Patch.CompareTo(b.Patch);
        if (c != 0) return c;

        // A version with a pre-release part is lower than the same version without one.
        if (a.Pre.Length == 0 || b.Pre.Length == 0)
            return a.Pre.Length == b.Pre.Length ? 0 : a.Pre.Length == 0 ? 1 : -1;

        for (var i = 0; i < Math.Min(a.Pre.Length, b.Pre.Length); i++)
        {
            var an = long.TryParse(a.Pre[i], out var x);
            var bn = long.TryParse(b.Pre[i], out var y);
            c = an && bn ? x.CompareTo(y)
                : an ? -1
                : bn ? 1
                : string.CompareOrdinal(a.Pre[i], b.Pre[i]);
            if (c != 0) return c;
        }

        return a.Pre.Length.CompareTo(b.Pre.Length);
    }
}
