using System.Collections.Generic;
using System.Globalization;
using System.Xml;
using XLibur.Excel.IO;
using XLibur.Utils;
using static XLibur.Excel.IO.OpenXmlConst;

namespace XLibur.Excel.Streaming;

/// <summary>
/// The shared string table of a streaming write: an append-only, insertion-ordered dictionary
/// of distinct strings.
/// </summary>
/// <remarks>
/// Deliberately not <see cref="SharedStringTable"/>. That one is reference counted, so ids can
/// be freed and reused, which leaves gaps that the normal save path closes with a remap
/// (<c>SaveContext.SstMap</c>) after the sheets are in memory. A streaming write cannot remap
/// anything - the index goes into cell XML that is on disk before the table is written - so ids
/// here are handed out densely and never released, and the part is emitted in that same order.
/// </remarks>
internal sealed class StreamingSharedStringTable
{
    private readonly Dictionary<string, int> _ids = new(System.StringComparer.Ordinal);
    private readonly List<string> _texts = [];

    /// <summary>
    /// Total number of cells that referenced the table, which is what the <c>count</c> attribute
    /// reports - as against <see cref="Count"/>, the number of <c>si</c> entries.
    /// </summary>
    private int _referenceCount;

    internal int Count => _texts.Count;

    /// <summary>
    /// The index of a text, adding it to the table if this is its first use.
    /// </summary>
    internal int GetOrAdd(string text)
    {
        _referenceCount++;

        if (_ids.TryGetValue(text, out var id))
            return id;

        id = _texts.Count;
        _texts.Add(text);
        _ids.Add(text, id);
        return id;
    }

    internal void Write(XmlWriter xml)
    {
        xml.WriteStartDocument();
        xml.WriteStartElement("x", "sst", Main2006SsNs);
        // count is how many cells point at the table; uniqueCount is how many entries it holds.
        xml.WriteAttributeString("count", _referenceCount.ToString(CultureInfo.InvariantCulture));
        xml.WriteAttributeString("uniqueCount", _texts.Count.ToString(CultureInfo.InvariantCulture));

        foreach (var text in _texts)
        {
            xml.WriteStartElement("si", Main2006SsNs);
            CellXmlWriter.WriteTextElement(xml, XmlEncoder.EncodeString(text));
            xml.WriteEndElement(); // si
        }

        xml.WriteEndElement(); // sst
        xml.WriteEndDocument();
    }
}
