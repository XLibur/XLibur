using System;
using System.Xml;
using XLibur.Extensions;
using static XLibur.Excel.IO.OpenXmlConst;

namespace XLibur.Excel.IO;

/// <summary>
/// Leaf-level <c>sheetData</c> XML primitives, shared by the two producers of that element:
/// <see cref="SheetDataWriter"/>, which walks a fully materialised <see cref="XLWorksheet"/>,
/// and the forward-only writer in <c>XLibur.Excel.Streaming</c>, whose enumeration is driven
/// by the caller and has no slice storage to read from.
/// </summary>
/// <remarks>
/// Only the leaves are shared. Neither producer's row/cell loop is expressed in terms of the
/// other, deliberately: the worksheet loop is a single pass over a struct slice enumerator with
/// no per-cell allocation, and routing it through a row/cell abstraction wide enough for the
/// streaming side (formula, misc metadata, rich text, table-totals membership) would cost an
/// interface dispatch per cell for no gain.
/// </remarks>
internal static class CellXmlWriter
{
    /// <summary>
    /// Day offset between the 1900 and 1904 date systems used by Excel.
    /// </summary>
    internal const int Date1904OffsetDays = 1462;

    /// <summary>
    /// Enough to hold the longest cell reference (<c>XFD1048576</c>) as chars.
    /// </summary>
    internal const int CellRefBufferLength = 10;

    /// <summary>
    /// An array to convert data type for a formula cell. Key is <see cref="XLDataType"/>.
    /// It saves some performance through direct indexation instead of switch.
    /// </summary>
    private static readonly string?[] FormulaDataType =
    [
        null, // blank
        "b", // boolean
        null, // number, default value, no need to save type
        "str", // text, formula can only save this type, no inline or shared string
        "e", // error
        null, // datetime, saved as serialized date-time
        null // timespan, saved as serialized date-time
    ];

    /// <summary>
    /// An array to convert a data type for a cell that only contains a value. Key is <see cref="XLDataType"/>.
    /// It saves some performance through direct indexation instead of switch.
    /// </summary>
    private static readonly string?[] ValueDataType =
    [
        null, // blank
        "b", // boolean
        null, // number, default value, no need to save type
        "s", // text, the default is a shared string, but there also can be inline string depending on ShareString property
        "e", // error
        null, // datetime, saved as serialized date-time
        null // timespan, saved as serialized date-time
    ];

    /// <summary>
    /// The <c>t</c> attribute value for a formula cell whose cached value has the passed type,
    /// or <c>null</c> when the default (number) applies and the attribute can be omitted.
    /// </summary>
    internal static string? GetFormulaCellType(XLDataType dataType) => FormulaDataType[(int)dataType];

    /// <summary>
    /// The <c>t</c> attribute value for a cell that only carries a value, or <c>null</c> when
    /// the default (number) applies and the attribute can be omitted.
    /// </summary>
    internal static string? GetValueCellType(XLDataType dataType, bool shareString)
    {
        if (dataType == XLDataType.Text && !shareString)
            return "inlineStr";
        return ValueDataType[(int)dataType];
    }

    /// <summary>
    /// Open a <c>&lt;row&gt;</c> element and write its <c>r</c> and (when known) <c>spans</c>
    /// attributes. The element is left open so the caller can add its own attributes.
    /// </summary>
    internal static void WriteRowStart(XmlWriter w, int rowNumber, int maxColumn)
    {
        w.WriteStartElement("row", Main2006SsNs);

        w.WriteStartAttribute("r");
        w.WriteNumberValue(rowNumber);
        w.WriteEndAttribute();

        if (maxColumn > 0)
        {
            w.WriteStartAttribute("spans");
            w.WriteString("1:");
            w.WriteNumberValue(maxColumn);
            w.WriteEndAttribute();
        }
    }

    /// <summary>
    /// Open a <c>&lt;c&gt;</c> element with its reference, style and type attributes. The
    /// element is left open so the caller can add optional attributes and the value.
    /// <c>reference</c> holds the cell reference as written by <c>Point.Format</c>, of which
    /// <c>referenceLength</c> chars are used.
    /// </summary>
    internal static void WriteCellStart(XmlWriter w, char[] reference, int referenceLength, string? dataType,
        uint styleId)
    {
        w.WriteStartElement("c", Main2006SsNs);

        w.WriteStartAttribute("r");
        w.WriteRaw(reference, 0, referenceLength);
        w.WriteEndAttribute();

        w.WriteAttribute("s", styleId);

        if (dataType is not null)
            w.WriteAttributeString("t", dataType);
    }

    /// <summary>
    /// The optional cell attributes carried by the misc slice: phonetic flag and the cell and
    /// value metadata indexes. Written straight after <see cref="WriteCellStart"/>.
    /// </summary>
    internal static void WriteCellMetaAttributes(XmlWriter w, bool hasPhonetic, uint? cellMetaIndex,
        uint? valueMetaIndex)
    {
        if (hasPhonetic)
            w.WriteAttributeString("ph", TrueValue);

        if (cellMetaIndex is not null)
            w.WriteAttribute("cm", cellMetaIndex.Value);

        if (valueMetaIndex is not null)
            w.WriteAttribute("vm", valueMetaIndex.Value);
    }

    /// <summary>
    /// Write a <c>&lt;v&gt;</c> element holding a shared string index.
    /// </summary>
    internal static void WriteSharedStringValue(XmlWriter w, int sharedStringId)
    {
        w.WriteStartElement("v", Main2006SsNs);
        w.WriteNumberValue(sharedStringId);
        w.WriteEndElement();
    }

    /// <summary>
    /// Write a <c>&lt;v&gt;</c> element holding verbatim text.
    /// </summary>
    internal static void WriteStringValue(XmlWriter w, string text)
    {
        w.WriteStartElement("v", Main2006SsNs);
        w.WriteString(text);
        w.WriteEndElement();
    }

    /// <summary>
    /// Write a <c>&lt;v&gt;</c> element holding a number.
    /// </summary>
    internal static void WriteNumberValue(XmlWriter w, double value)
    {
        w.WriteStartElement("v", Main2006SsNs);
        w.WriteNumberValue(value);
        w.WriteEndElement();
    }

    /// <summary>
    /// Write an <c>&lt;is&gt;&lt;t&gt;</c> inline string, preserving leading/trailing spaces
    /// when needed.
    /// </summary>
    internal static void WriteInlineString(XmlWriter w, string text)
    {
        w.WriteStartElement("is", Main2006SsNs);
        WriteTextElement(w, text);
        w.WriteEndElement(); // is
    }

    /// <summary>
    /// Write a <c>&lt;t&gt;</c> element, preserving leading/trailing spaces when needed.
    /// </summary>
    internal static void WriteTextElement(XmlWriter w, string text)
    {
        w.WriteStartElement("t", Main2006SsNs);
        if (text.PreserveSpaces())
            w.WritePreserveSpaceAttr();

        w.WriteString(text);
        w.WriteEndElement(); // t
    }

    /// <summary>
    /// Write the <c>&lt;v&gt;</c> element for every value type except
    /// <see cref="XLDataType.Blank"/> (which has no value element) and
    /// <see cref="XLDataType.Text"/>, whose representation differs per caller: a shared string
    /// index, an inline string, rich text, or - for a formula's cached value - verbatim text.
    /// </summary>
    internal static void WriteNonTextValue(XmlWriter w, XLCellValue cellValue, bool use1904DateSystem)
    {
        switch (cellValue.Type)
        {
            case XLDataType.TimeSpan:
                WriteNumberValue(w, cellValue.GetUnifiedNumber());
                break;
            case XLDataType.Number:
                WriteNumberValue(w, cellValue.GetNumber());
                break;
            case XLDataType.DateTime:
                WriteNumberValue(w, ToSerialDateTime(cellValue.GetDateTime(), use1904DateSystem));
                break;
            case XLDataType.Boolean:
                WriteStringValue(w, cellValue.GetBoolean() ? TrueValue : FalseValue);
                break;
            case XLDataType.Error:
                WriteStringValue(w, cellValue.GetError().ToDisplayString());
                break;
            default:
                throw new InvalidOperationException();
        }
    }

    /// <summary>
    /// Convert a date to its Excel serial number, shifting it into the 1904 date system when
    /// the workbook uses one.
    /// </summary>
    internal static double ToSerialDateTime(DateTime date, bool use1904DateSystem)
    {
        if (use1904DateSystem)
            date = date.AddDays(-Date1904OffsetDays);

        return date.ToSerialDateTime();
    }
}
