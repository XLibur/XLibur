using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Utils;

namespace XLibur.Excel.IO;

/// <summary>
/// A shared string table entry that is either a plain text string or a rich text entry.
/// Plain text entries (the vast majority) are stored as simple strings to avoid retaining DOM objects.
/// Rich text entries retain the DOM element for formatting extraction during cell loading, and the
/// rich text built from it (<see cref="RichSharedString"/>).
/// </summary>
internal readonly struct SharedStringEntry
{
    /// <summary>
    /// Either a <see cref="string"/> (plain text) or a <see cref="RichSharedString"/> (rich text with runs/phonetics).
    /// </summary>
    private readonly object? _value;

    private SharedStringEntry(object? value) => _value = value;

    internal static SharedStringEntry Plain(string text) => new(text);
    internal static SharedStringEntry Rich(RstType element) => new(new RichSharedString(element));

    internal bool IsRichText => _value is RichSharedString;

    internal string PlainText => (string)(_value ?? string.Empty);

    internal RichSharedString RichText => (RichSharedString)_value!;
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
/// (a leading <c>&lt;t&gt;</c> may be followed by <c>&lt;rPh&gt;</c>), so <see cref="ReadRichItem"/>
/// takes the text already read as well as the rest of the subtree.
/// </para>
/// </summary>
internal static class SharedStringReader
{
    internal static SharedStringEntry[] Read(SharedStringTablePart part)
    {
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        // Whitespace must NOT be ignored: a <t> holding only spaces is legitimate text content.
        using var reader = PartXmlReader.Create(stream, ignoreWhitespace: false);

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

        ReadItems(reader, ref entries, ref count);

        if (count != entries.Length)
            Array.Resize(ref entries, count);

        return entries;
    }

    /// <summary>
    /// Reads every <c>&lt;si&gt;</c> child of <c>&lt;sst&gt;</c> into <paramref name="entries"/>. The reader
    /// enters on the first child (or <c>&lt;/sst&gt;</c>) and returns on <c>&lt;/sst&gt;</c>.
    /// </summary>
    private static void ReadItems(XmlReader reader, ref SharedStringEntry[] entries, ref int count)
    {
        while (true)
        {
            if (reader.NodeType == XmlNodeType.Element)
            {
                if (IsMainElement(reader, "si"))
                    Append(ref entries, ref count, ReadSharedStringItem(reader)); // Leaves reader after </si>.
                else
                    reader.Skip(); // Unknown element (e.g. <extLst>).

                continue;
            }

            if (reader.NodeType == XmlNodeType.EndElement || reader.EOF || !reader.Read())
                break;
        }
    }

    private static void Append(ref SharedStringEntry[] entries, ref int count, SharedStringEntry entry)
    {
        if (count == entries.Length)
        {
            // uniqueCount was absent or understated — grow geometrically.
            Array.Resize(ref entries, entries.Length == 0 ? 16 : entries.Length * 2);
        }

        entries[count++] = entry;
    }

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
    /// Reads the remainder of an <c>&lt;si&gt;</c> (plus any <c>&lt;t&gt;</c> already consumed) into a
    /// <see cref="SharedStringItem"/>. The reader enters on the first remaining child element and
    /// returns on <c>&lt;/si&gt;</c>.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The elements are built straight from the reader, holding the attributes as the file has
    /// them, so the SDK parses their values as it would have parsed the markup. Only what
    /// <see cref="WorksheetSheetDataReader.SetCellText"/> reads is built: the runs with their
    /// properties and text, the phonetic runs and the phonetic properties. Anything else is skipped,
    /// as that reader skips it.
    /// </para>
    /// <para>
    /// The item used to be written back out through an <see cref="XmlWriter"/>, a
    /// <c>ReadOuterXml</c> for each child, and the resulting markup parsed again into the SDK's DOM.
    /// Each writer and each outer read took buffers of its own, which made half of everything a
    /// load allocated for a workbook rich in formatted text.
    /// </para>
    /// </remarks>
    private static SharedStringEntry ReadRichItem(XmlReader reader, string? leadingText)
    {
        var item = new SharedStringItem();
        if (leadingText is not null)
            item.AppendChild(new Text(leadingText));

        while (true)
        {
            if (reader.NodeType == XmlNodeType.Element)
            {
                OpenXmlElement? child = !IsMainNamespace(reader) ? null : reader.LocalName switch
                {
                    "r" => ReadRun(reader),
                    "rPh" => ReadElement(reader, new PhoneticRun(), static (r, _) =>
                        r.LocalName == "t" ? new Text() : null),
                    "phoneticPr" => ReadElement(reader, new PhoneticProperties(), static (_, _) => null),
                    "t" => ReadText(reader),
                    _ => null,
                };

                if (child is null)
                    reader.Skip();
                else
                    item.AppendChild(child);

                continue;
            }

            if (reader.NodeType == XmlNodeType.EndElement || reader.EOF)
                break;

            if (!reader.Read())
                break;
        }

        return SharedStringEntry.Rich(item);
    }

    /// <summary>A run: its properties, each child element of them, and its text.</summary>
    private static Run ReadRun(XmlReader reader) =>
        ReadElement(reader, new Run(), static (r, _) => r.LocalName switch
        {
            "rPr" => new RunProperties(),
            "t" => new Text(),
            _ => null,
        });

    /// <summary>
    /// The children of <c>&lt;rPr&gt;</c> that <see cref="StyleDecoder.RunFontKey"/> reads, by local
    /// name.
    /// </summary>
    private static OpenXmlElement? RunPropertyElement(string localName) => localName switch
    {
        "b" => new Bold(),
        "i" => new Italic(),
        "strike" => new Strike(),
        "condense" => new Condense(),
        "extend" => new Extend(),
        "outline" => new Outline(),
        "shadow" => new Shadow(),
        "u" => new Underline(),
        "vertAlign" => new VerticalTextAlignment(),
        "sz" => new FontSize(),
        "color" => new Color(),
        "rFont" => new RunFont(),
        "family" => new FontFamily(),
        "charset" => new RunPropertyCharSet(),
        "scheme" => new FontScheme(),
        _ => null,
    };

    /// <summary>
    /// Fills <paramref name="element"/> from the element the reader is on: its attributes, then its
    /// text when it is a <see cref="Text"/>, or else each child <paramref name="childOf"/> makes an
    /// element for, recursively. A child it makes none for is skipped. The reader returns past the
    /// element's end.
    /// </summary>
    private static T ReadElement<T>(XmlReader reader, T element, Func<XmlReader, OpenXmlElement, OpenXmlElement?> childOf)
        where T : OpenXmlElement
    {
        CopyAttributes(reader, element);

        if (reader.IsEmptyElement)
        {
            reader.Read();
            return element;
        }

        reader.Read();
        while (true)
        {
            if (reader.NodeType == XmlNodeType.Element)
            {
                var child = IsMainNamespace(reader)
                    ? element is RunProperties ? RunPropertyElement(reader.LocalName) : childOf(reader, element)
                    : null;

                switch (child)
                {
                    case null:
                        reader.Skip();
                        break;
                    case Text text:
                        element.AppendChild(ReadText(reader, text));
                        break;
                    case RunProperties properties:
                        element.AppendChild(ReadElement(reader, properties, childOf));
                        break;
                    default:
                        element.AppendChild(ReadElement(reader, child, static (_, _) => null));
                        break;
                }

                continue;
            }

            if (reader.NodeType == XmlNodeType.EndElement)
            {
                reader.Read();
                return element;
            }

            if (!reader.Read())
                return element;
        }
    }

    /// <summary>A <c>&lt;t&gt;</c>: its text, which the reader has already decoded. The reader returns past it.</summary>
    private static Text ReadText(XmlReader reader, Text? text = null)
    {
        text ??= new Text();
        CopyAttributes(reader, text);
        text.Text = reader.ReadElementContentAsString();
        return text;
    }

    /// <summary>
    /// Gives <paramref name="element"/> the attributes the reader's element has, namespace
    /// declarations aside, as their text: the SDK parses each when it is read, as it does markup.
    /// </summary>
    private static void CopyAttributes(XmlReader reader, OpenXmlElement element)
    {
        if (!reader.MoveToFirstAttribute())
            return;

        do
        {
            if (reader.NamespaceURI == XmlnsNamespace)
                continue;

            element.SetAttribute(new OpenXmlAttribute(reader.Prefix, reader.LocalName, reader.NamespaceURI, reader.Value));
        }
        while (reader.MoveToNextAttribute());

        reader.MoveToElement();
    }

    private const string XmlnsNamespace = "http://www.w3.org/2000/xmlns/";

    private static bool IsMainNamespace(XmlReader reader) => reader.NamespaceURI == OpenXmlConst.Main2006SsNs;

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
