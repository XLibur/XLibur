using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Utils;

namespace XLibur.Excel.IO;

/// <summary>
/// A shared string table entry that is either a plain text string or a rich text <see cref="RstType"/> element.
/// Plain text entries (the vast majority) are stored as simple strings to avoid retaining DOM objects.
/// Rich text entries retain the DOM element for formatting extraction during cell loading.
/// </summary>
internal readonly struct SharedStringEntry
{
    /// <summary>
    /// Either a <see cref="string"/> (plain text) or a <see cref="RstType"/> (rich text with runs/phonetics).
    /// </summary>
    private readonly object? _value;

    private SharedStringEntry(object? value) => _value = value;

    internal static SharedStringEntry Plain(string text) => new(text);
    internal static SharedStringEntry Rich(RstType element) => new(element);

    internal bool IsRichText => _value is RstType;

    internal string PlainText => (string)(_value ?? string.Empty);

    internal RstType RichText => (RstType)_value!;
}

/// <summary>
/// Reads the shared string table from an SST part with a raw <see cref="XmlReader"/>.
/// <para>
/// The overwhelmingly common entry is a plain <c>&lt;si&gt;&lt;t&gt;text&lt;/t&gt;&lt;/si&gt;</c>, for which only
/// the decoded string is retained — no DOM node is ever built. Materializing the whole
/// <see cref="SharedStringTablePart.SharedStringTable"/> DOM instead costs two OpenXml elements plus an
/// attribute collection per entry, all of which become garbage immediately after the text is extracted;
/// for a string-heavy workbook the table can hold hundreds of thousands of entries.
/// </para>
/// <para>
/// Entries with runs or phonetic data are rare and still need the DOM for formatting extraction via
/// <see cref="WorksheetSheetDataReader.SetCellText"/>, so their subtree is rebuilt into a
/// <see cref="SharedStringItem"/>. Richness can only be established after the first child has been consumed
/// (a leading <c>&lt;t&gt;</c> may be followed by <c>&lt;rPh&gt;</c>), hence the re-serialization in
/// <see cref="ReadRichItem"/> rather than a plain <c>ReadOuterXml</c>.
/// </para>
/// </summary>
internal static class SharedStringReader
{
    internal static SharedStringEntry[] Read(SharedStringTablePart part)
    {
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using var reader = XmlReader.Create(stream, new XmlReaderSettings
        {
            // Whitespace must NOT be ignored: a <t> holding only spaces is legitimate text content.
            IgnoreWhitespace = false,
            IgnoreComments = true,
            IgnoreProcessingInstructions = true,
            CloseInput = false
        });

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element
                && reader.LocalName == "sst"
                && reader.NamespaceURI == OpenXmlConst.Main2006SsNs)
            {
                return ReadSst(reader);
            }
        }

        return [];
    }

#pragma warning disable S3776 // A hand-rolled XmlReader walk; the node-type tests are the parser
    private static SharedStringEntry[] ReadSst(XmlReader reader)
    {
        // Pre-allocate from the sst's uniqueCount attribute to avoid resize+copy for large tables.
        // Only uniqueCount (number of unique <si> entries) is usable here, not count (total reference
        // count including duplicates), which would over-allocate.
        var uniqueCount = ReadUniqueCount(reader);

        if (reader.IsEmptyElement)
            return [];

        var entries = uniqueCount > 0 ? new SharedStringEntry[uniqueCount] : [];
        var count = 0;

        reader.Read(); // Move into <sst> (first child or </sst>).

        while (true)
        {
            if (reader.NodeType == XmlNodeType.Element)
            {
                if (reader.LocalName == "si" && reader.NamespaceURI == OpenXmlConst.Main2006SsNs)
                {
                    var entry = ReadSharedStringItem(reader); // Leaves reader after </si>.

                    if (count == entries.Length)
                    {
                        // uniqueCount was absent or understated — grow geometrically.
                        Array.Resize(ref entries, entries.Length == 0 ? 16 : entries.Length * 2);
                    }

                    entries[count++] = entry;
                    continue;
                }

                reader.Skip(); // Unknown element (e.g. <extLst>).
                continue;
            }

            if (reader.NodeType == XmlNodeType.EndElement || reader.EOF)
                break;

            if (!reader.Read())
                break;
        }

        if (count != entries.Length)
            Array.Resize(ref entries, count);

        return entries;
    }
#pragma warning restore S3776

    private static int ReadUniqueCount(XmlReader reader)
    {
        var uniqueCount = 0;
        if (reader.HasAttributes)
        {
            var value = reader.GetAttribute("uniqueCount");
            if (value is not null && int.TryParse(value, out var parsed) && parsed > 0)
                uniqueCount = parsed;

            reader.MoveToElement();
        }

        return uniqueCount;
    }

    /// <summary>
    /// Reads a single <c>&lt;si&gt;</c>. The reader enters positioned on the start element and returns
    /// positioned on the node immediately after <c>&lt;/si&gt;</c>.
    /// </summary>
    private static SharedStringEntry ReadSharedStringItem(XmlReader reader)
    {
        if (reader.IsEmptyElement)
        {
            reader.Read(); // Move past <si/>.
            return SharedStringEntry.Plain(string.Empty);
        }

        reader.Read(); // Move into <si> (first child or </si>).
        SkipToContent(reader);

        string? plainText = null;
        if (IsMainElement(reader, "t"))
        {
            plainText = reader.ReadElementContentAsString(); // Reads <t> text and moves past </t>.
            SkipToContent(reader);
        }

        // A lone <t> is the plain case; anything left before </si> means runs or phonetic data.
        if (reader.NodeType != XmlNodeType.Element)
        {
            reader.Read(); // Move past </si>.

            // Decode _xHHHH_ escapes (e.g. _x0018_ → ), matching the SetCellText path.
            return SharedStringEntry.Plain(XmlEncoder.DecodeString(plainText ?? string.Empty));
        }

        var rich = ReadRichItem(reader, plainText);
        reader.Read(); // Move past </si>.
        return rich;
    }

    /// <summary>
    /// Rebuilds the remainder of an <c>&lt;si&gt;</c> subtree (plus any <c>&lt;t&gt;</c> already consumed)
    /// into a <see cref="SharedStringItem"/> DOM element. The reader enters on the first remaining child
    /// element and returns on <c>&lt;/si&gt;</c>.
    /// </summary>
    private static SharedStringEntry ReadRichItem(XmlReader reader, string? leadingText)
    {
        var sb = new StringBuilder();
        using (var writer = XmlWriter.Create(sb, new XmlWriterSettings
        {
            OmitXmlDeclaration = true,
            ConformanceLevel = ConformanceLevel.Fragment,
            CheckCharacters = false
        }))
        {
            writer.WriteStartElement("si", OpenXmlConst.Main2006SsNs);

            if (leadingText is not null)
            {
                writer.WriteStartElement("t", OpenXmlConst.Main2006SsNs);
                writer.WriteAttributeString("space", OpenXmlConst.Xml1998Ns, "preserve");
                writer.WriteString(leadingText);
                writer.WriteEndElement();
            }

            while (true)
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    // ReadOuterXml emits the in-scope default namespace and moves past the child.
                    writer.WriteRaw(reader.ReadOuterXml());
                    continue;
                }

                if (reader.NodeType == XmlNodeType.EndElement || reader.EOF)
                    break;

                if (!reader.Read())
                    break;
            }

            writer.WriteEndElement();
        }

        return SharedStringEntry.Rich(new SharedStringItem(sb.ToString()));
    }

    private static bool IsMainElement(XmlReader reader, string localName)
        => reader.NodeType == XmlNodeType.Element
           && reader.LocalName == localName
           && reader.NamespaceURI == OpenXmlConst.Main2006SsNs;

    /// <summary>
    /// Advances over whitespace and other non-structural nodes so the reader lands on the next
    /// element start or end tag.
    /// </summary>
    private static void SkipToContent(XmlReader reader)
    {
        while (reader.NodeType is not (XmlNodeType.Element or XmlNodeType.EndElement))
        {
            if (reader.EOF || !reader.Read())
                return;
        }
    }
}
