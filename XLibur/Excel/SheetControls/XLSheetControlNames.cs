using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace XLibur.Excel;

/// <summary>
/// The naming and placement rules slicers and timelines share: the control names Excel shows in the
/// selection pane, the cache names behind them, and where a new control goes when the caller has not
/// said.
/// </summary>
/// <remarks>
/// Slicers and timelines are otherwise unrelated types with unrelated parts, so these are plain
/// helpers rather than a base class. They exist so that a rule both kinds must obey — above all the
/// one name namespace the two share — is written once rather than kept in step by hand.
/// </remarks>
internal static class XLSheetControlNames
{
    /// <summary>
    /// A control name not already taken, in the shape Excel uses: <c>Region</c>, then
    /// <c>Region 1</c>. Control names are unique across the workbook, not just the sheet.
    /// </summary>
    /// <remarks>
    /// Slicers and timelines share one name namespace in Excel's selection pane, so this scans both —
    /// otherwise a slicer could take a name a timeline already has, or the other way round.
    /// </remarks>
    internal static string NextControlName(XLWorkbook workbook, string sourceName)
    {
        var taken = new HashSet<string>(XLHelper.NameComparer);
        foreach (var worksheet in workbook.WorksheetsInternal)
        {
            foreach (var slicer in worksheet.SlicersInternal.Items)
                taken.Add(slicer.Name);
            foreach (var timeline in worksheet.TimelinesInternal.Items)
                taken.Add(timeline.Name);
        }

        return FirstFree(taken, sourceName, " ");
    }

    /// <summary>
    /// A cache name not already taken, in the shape Excel uses: the prefix and the sanitised source
    /// name (<c>Slicer_Region</c>, <c>NativeTimeline_Date</c>), then the same with a number appended
    /// (<c>Slicer_Region1</c>).
    /// </summary>
    /// <remarks>
    /// The name is not decoration. A control refers to its cache by it, and a <c>#N/A</c> defined name
    /// is written under the same name, so it has to be a legal defined name: no spaces, and nothing
    /// that would parse as a cell reference. Every defined name in the workbook counts as taken, as
    /// does every name <paramref name="cacheNamesOf"/> gives for a worksheet — the caller's own kind
    /// of cache.
    /// </remarks>
    internal static string NextCacheName(
        XLWorkbook workbook, string prefix, string sourceName, Func<XLWorksheet, IEnumerable<string>> cacheNamesOf)
    {
        var taken = WorkbookCacheNames(workbook, cacheNamesOf);
        return FirstFree(taken, prefix + Sanitise(sourceName), string.Empty);
    }

    /// <summary>
    /// Where a new control goes when the caller has not said: two columns to the right of whatever it
    /// filters, at that thing's top row.
    /// </summary>
    /// <remarks>
    /// <para>
    /// A default is not optional here, and it is worth saying why rather than leaving it to the layer
    /// below. <c>DrawingAnchorFactory</c> documents that a drawing handed no marker gets one at A1 —
    /// silently, with no exception and no missing element. For a picture that is a reasonable
    /// default. For a slicer or a timeline it would drop the panel over the top-left of the sheet,
    /// covering the very data it filters, and the caller would have no idea why.
    /// </para>
    /// <para>
    /// So every control XLibur creates is given a marker before the factory sees it, and that fallback
    /// stays unreachable from here. Two columns of clearance keeps the panel off the source without
    /// guessing at column widths.
    /// </para>
    /// </remarks>
    internal static XLCell DefaultPositionBeside(XLWorksheet worksheet, int topRow, int rightmostColumn) =>
        worksheet.Cell(
            Math.Max(1, topRow),
            Math.Min(XLHelper.MaxColumnNumber, rightmostColumn + 2));

    private static HashSet<string> WorkbookCacheNames(
        XLWorkbook workbook, Func<XLWorksheet, IEnumerable<string>> cacheNamesOf)
    {
        var taken = new HashSet<string>(XLHelper.NameComparer);
        foreach (var worksheet in workbook.WorksheetsInternal)
            taken.UnionWith(cacheNamesOf(worksheet));

        // A defined name already using the stem would collide with the one written for the cache.
        foreach (var definedName in workbook.DefinedNamesInternal)
            taken.Add(definedName.Name);

        return taken;
    }

    /// <summary>
    /// The stem itself if free, otherwise the stem, the separator and the lowest free suffix from 1.
    /// </summary>
    private static string FirstFree(HashSet<string> taken, string stem, string separator)
    {
        if (!taken.Contains(stem))
            return stem;

        // Bounded by `taken`, not by the counter: the set is finite, so some suffix is always free
        // and the loop returns within `taken.Count + 1` iterations.
#pragma warning disable S1994
        for (var suffix = 1; ; suffix++)
        {
            var candidate = stem + separator + suffix.ToString(CultureInfo.InvariantCulture);
            if (!taken.Contains(candidate))
                return candidate;
        }
#pragma warning restore S1994
    }

    private static string Sanitise(string sourceName)
    {
        var builder = new StringBuilder(sourceName.Length);
        foreach (var c in sourceName)
            builder.Append(char.IsLetterOrDigit(c) || c == '_' ? c : '_');

        return builder.Length > 0 ? builder.ToString() : "Field";
    }
}
