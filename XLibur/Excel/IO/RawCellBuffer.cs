using System;
using System.Globalization;
using System.Xml;
using XLibur.Extensions;
using static XLibur.Excel.IO.OpenXmlConst;

namespace XLibur.Excel.IO;

/// <summary>
/// Formats the common kinds of <c>&lt;c&gt;</c> element straight into characters and hands a
/// row's worth to the <see cref="XmlWriter"/> in one <see cref="XmlWriter.WriteRaw(char[], int, int)"/>.
/// </summary>
/// <remarks>
/// Written through the element and attribute API, a value cell is eight or so writer calls, and
/// each of them runs the well-formedness state machine and resolves the namespace prefix again.
/// On a sheet of plain numbers that was most of the cost of saving it. The markup built here is
/// the exact text those calls produce — the same prefix, attribute order, quoting and number
/// format — so the part does not change by a byte.
/// <para>
/// Only cells whose markup needs no escaping are taken: a value without a formula whose content
/// is a number, date, time, boolean, error or shared-string index. Anything else — formulas,
/// inline and rich text, blank styled cells — is refused, and the caller flushes this buffer and
/// writes that cell through the writer as before.
/// </para>
/// </remarks>
internal sealed class RawCellBuffer
{
    /// <summary>
    /// Room for the longest cell this writes: the tags, a 10-character reference, a style id, the
    /// three optional metadata attributes and a G15 number, with margin.
    /// </summary>
    private const int MaxCellLength = 256;

    private readonly string _cellOpen;
    private readonly string _valueOpen;
    private readonly string _cellClose;
    private char[] _buffer = new char[4096];
    private int _length;

    private RawCellBuffer(string prefix)
    {
        var qualifier = prefix.Length == 0 ? string.Empty : prefix + ":";
        _cellOpen = "<" + qualifier + "c r=\"";
        _valueOpen = "><" + qualifier + "v>";
        _cellClose = "</" + qualifier + "v></" + qualifier + "c>";
    }

    /// <summary>
    /// A buffer for cells written inside the element the writer has open, or null when the main
    /// namespace has no prefix in scope there, which leaves every cell to the writer.
    /// </summary>
    internal static RawCellBuffer? Create(XmlWriter writer)
    {
        var prefix = writer.LookupPrefix(Main2006SsNs);
        return prefix is null ? null : new RawCellBuffer(prefix);
    }

    /// <summary>
    /// Appends a value cell, or returns false, having appended nothing, for a cell this does not
    /// write.
    /// </summary>
    internal bool TryAppendValueCell(ReadOnlySpan<char> reference, uint styleId, XLCellValue value,
        bool shareString, int sharedStringId, in XLMiscSliceContent misc, bool use1904DateSystem)
    {
        string? dataType;
        switch (value.Type)
        {
            case XLDataType.Number or XLDataType.DateTime or XLDataType.TimeSpan:
                dataType = null;
                break;
            case XLDataType.Boolean:
                dataType = "b";
                break;
            case XLDataType.Error:
                dataType = "e";
                break;
            case XLDataType.Text when shareString:
                dataType = "s";
                break;
            default:
                return false;
        }

        EnsureCapacity();

        Append(_cellOpen);
        Append(reference);
        Append('"');

        // Style 0 is what a missing s means, and Excel does not write it.
        if (styleId != 0)
        {
            Append(" s=\"");
            AppendNumber(styleId);
            Append('"');
        }

        if (dataType is not null)
        {
            Append(" t=\"");
            Append(dataType);
            Append('"');
        }

        if (misc.HasPhonetic)
        {
            Append(" ph=\"");
            Append(TrueValue);
            Append('"');
        }

        if (misc.CellMetaIndex is { } cellMetaIndex)
        {
            Append(" cm=\"");
            AppendNumber(cellMetaIndex);
            Append('"');
        }

        if (misc.ValueMetaIndex is { } valueMetaIndex)
        {
            Append(" vm=\"");
            AppendNumber(valueMetaIndex);
            Append('"');
        }

        Append(_valueOpen);
        switch (value.Type)
        {
            case XLDataType.Number:
                AppendNumber(value.GetNumber());
                break;
            case XLDataType.TimeSpan:
                AppendNumber(value.GetUnifiedNumber());
                break;
            case XLDataType.DateTime:
                AppendNumber(CellXmlWriter.ToSerialDateTime(value.GetDateTime(), use1904DateSystem));
                break;
            case XLDataType.Boolean:
                Append(value.GetBoolean() ? TrueValue : FalseValue);
                break;
            case XLDataType.Error:
                // #N/A, #DIV/0!, #VALUE! and the like: nothing in them needs escaping.
                Append(value.GetError().ToDisplayString());
                break;
            default:
                AppendNumber(sharedStringId);
                break;
        }

        Append(_cellClose);
        return true;
    }

    /// <summary>Writes out whatever has been appended.</summary>
    internal void Flush(XmlWriter writer)
    {
        if (_length == 0)
            return;

        writer.WriteRaw(_buffer, 0, _length);
        _length = 0;
    }

    private void EnsureCapacity()
    {
        if (_buffer.Length - _length < MaxCellLength)
            Array.Resize(ref _buffer, _buffer.Length * 2);
    }

    private void Append(char c) => _buffer[_length++] = c;

    private void Append(ReadOnlySpan<char> text)
    {
        text.CopyTo(_buffer.AsSpan(_length));
        _length += text.Length;
    }

    private void AppendNumber(uint value)
    {
        value.TryFormat(_buffer.AsSpan(_length), out var written, provider: CultureInfo.InvariantCulture);
        _length += written;
    }

    private void AppendNumber(int value)
    {
        value.TryFormat(_buffer.AsSpan(_length), out var written, provider: CultureInfo.InvariantCulture);
        _length += written;
    }

    /// <summary>
    /// The same "G15" invariant text <c>XmlWriterExtensions.WriteNumberValue(double)</c> writes.
    /// </summary>
    private void AppendNumber(double value)
    {
        if (!value.TryFormat(_buffer.AsSpan(_length), out var written, "G15", CultureInfo.InvariantCulture))
        {
            Append(value.ToString("G15", CultureInfo.InvariantCulture));
            return;
        }

        _length += written;
    }
}
