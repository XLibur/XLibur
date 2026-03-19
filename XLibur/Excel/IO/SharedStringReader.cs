using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Linq;
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
/// Reads the shared string table from an SST part. For plain text entries (the vast majority),
/// only the decoded text string is retained and the DOM element is released for GC. For rich
/// text entries (with runs or phonetic data), the <see cref="SharedStringItem"/> is kept for
/// later formatting extraction via <see cref="WorksheetSheetDataReader.SetCellText"/>.
/// </summary>
internal static class SharedStringReader
{
    internal static SharedStringEntry[] Read(SharedStringTablePart part)
    {
        var sst = part.SharedStringTable;
        if (sst is null)
            return [];

        // Pre-allocate from the SST's UniqueCount attribute to avoid
        // List<T> resize+copy overhead for large shared string tables.
        // Only use UniqueCount (number of unique <si> entries), not Count
        // (total reference count including duplicates) which would over-allocate.
        var uniqueCount = sst.UniqueCount?.Value;
        if (uniqueCount is not null and > 0)
        {
            var entries = new SharedStringEntry[(int)uniqueCount.Value];
            var idx = 0;
            foreach (var item in sst.Elements<SharedStringItem>())
            {
                var entry = ReadEntry(item);
                if (idx < entries.Length)
                    entries[idx++] = entry;
                else
                {
                    // Count attribute was wrong — fall back to growing
                    Array.Resize(ref entries, entries.Length * 2);
                    entries[idx++] = entry;
                }
            }

            // Trim if the declared count was larger than actual entries
            if (idx < entries.Length)
                Array.Resize(ref entries, idx);

            return entries;
        }

        // Fallback: no count attribute, use list
        var list = new List<SharedStringEntry>();
        foreach (var item in sst.Elements<SharedStringItem>())
            list.Add(ReadEntry(item));

        return list.ToArray();
    }

    private static SharedStringEntry ReadEntry(SharedStringItem item)
    {
        // Schema: <si> contains either (t, rPh*, phoneticPr?) or (r+, rPh*, phoneticPr?).
        // Pure plain text: a single <t> child with no runs and no phonetic data.
        var text = item.Text;
        if (text is not null && text == item.FirstChild && text == item.LastChild)
        {
            // Decode _xHHHH_ escapes (e.g. _x0018_ → \u0018) matching the original
            // SetCellText code path.
            var decoded = XmlEncoder.DecodeString(text.InnerText);
            return SharedStringEntry.Plain(decoded);
        }

        return SharedStringEntry.Rich(item);
    }
}
