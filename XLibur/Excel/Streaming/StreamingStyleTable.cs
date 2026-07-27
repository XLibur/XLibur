using System.Collections.Generic;

namespace XLibur.Excel.Streaming;

/// <summary>
/// Interns the styles used by a streaming write, handing out the cell style id at the moment
/// a style is first used.
/// </summary>
/// <remarks>
/// The id has to be final immediately: it is written into cell XML that is in the package long
/// before the styles part exists. So the order styles are interned in <em>is</em> the
/// <c>cellXfs</c> order, and <c>WorkbookStylesPartWriter.GenerateStreamingContent</c> emits one
/// <c>cellXf</c> per entry of <see cref="OrderedStyles"/>, in order, at the end of the write.
/// </remarks>
internal sealed class StreamingStyleTable
{
    private readonly Dictionary<XLStyleValue, uint> _ids = [];
    private readonly List<XLStyleValue> _ordered = [];

    internal StreamingStyleTable()
    {
        // Id 0 must be the default style: it is what every cell written without an explicit
        // style refers to, and what Excel treats as the workbook default.
        GetOrAdd(XLStyleValue.Default);
    }

    internal IReadOnlyList<XLStyleValue> OrderedStyles => _ordered;

    /// <summary>
    /// The cell style id for a style, or 0 when <paramref name="style"/> is <c>null</c>.
    /// </summary>
    internal uint GetOrAdd(IXLStyle? style)
    {
        if (style is null)
            return 0;

        var key = XLStyle.GenerateKey(style);
        return GetOrAdd(XLStyleValue.FromKey(ref key));
    }

    internal uint GetOrAdd(XLStyleValue value)
    {
        if (_ids.TryGetValue(value, out var id))
            return id;

        id = (uint)_ordered.Count;
        _ordered.Add(value);
        _ids.Add(value, id);
        return id;
    }
}
