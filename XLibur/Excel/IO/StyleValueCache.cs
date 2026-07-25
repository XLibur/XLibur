using System.Collections.Generic;

namespace XLibur.Excel.IO;

/// <summary>
/// Caches the <see cref="XLStyleValue"/> resolved for each <c>cellXfs</c> style index encountered
/// while loading cells.
/// <para>
/// Style indexes are dense and bounded by the number of <c>cellXfs</c> entries, so the cache is a
/// flat array rather than a dictionary: resolving a cell's style is a bounds check and an array read
/// instead of a hash and probe, and every cell in a sheet performs one such lookup. Indexes outside
/// the declared range (malformed files) fall back to a lazily created dictionary so behaviour is
/// unchanged for them.
/// </para>
/// <para>
/// The cache is workbook-scoped: <c>styles.xml</c> is workbook-global and
/// <see cref="WorksheetSheetDataReader.ResolveStyleValue"/> is a pure function of the style index,
/// so entries resolved for one worksheet are reused by the rest.
/// </para>
/// </summary>
internal sealed class StyleValueCache
{
    private readonly StylesheetData _styles;
    private readonly XLStyleValue?[] _byIndex;
    private Dictionary<int, XLStyleValue>? _outOfRange;

    internal StyleValueCache(StylesheetData styles)
    {
        _styles = styles;

        // ChildElements.Count is the actual number of <xf> elements; CellFormats.Count is the
        // declared count attribute, which is not authoritative.
        var cellFormatCount = styles.Stylesheet?.CellFormats?.ChildElements.Count ?? 0;
        _byIndex = cellFormatCount > 0 ? new XLStyleValue?[cellFormatCount] : [];
    }

    internal XLStyleValue Resolve(int styleIndex)
    {
        if ((uint)styleIndex < (uint)_byIndex.Length)
        {
            var cached = _byIndex[styleIndex];
            if (cached is not null)
                return cached;

            var resolved = WorksheetSheetDataReader.ResolveStyleValue(styleIndex, _styles);
            _byIndex[styleIndex] = resolved;
            return resolved;
        }

        return ResolveOutOfRange(styleIndex);
    }

    /// <summary>
    /// Handles a style index beyond the declared <c>cellXfs</c> range. This happens for workbooks
    /// with no stylesheet at all (every index resolves to the default style) and for malformed
    /// files; <see cref="WorksheetSheetDataReader.ResolveStyleValue"/> keeps whatever behaviour
    /// those cases had before.
    /// </summary>
    private XLStyleValue ResolveOutOfRange(int styleIndex)
    {
        _outOfRange ??= new Dictionary<int, XLStyleValue>();
        if (_outOfRange.TryGetValue(styleIndex, out var cached))
            return cached;

        var resolved = WorksheetSheetDataReader.ResolveStyleValue(styleIndex, _styles);
        _outOfRange[styleIndex] = resolved;
        return resolved;
    }
}
