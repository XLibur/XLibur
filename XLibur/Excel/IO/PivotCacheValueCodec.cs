using System;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Extensions;
using XLibur.Utils;

namespace XLibur.Excel.IO;

/// <summary>
/// How one pivot cache value is spelled in XML: <c>m</c> for a missing value, <c>n</c> a number,
/// <c>b</c> a boolean, <c>e</c> an error, <c>s</c> a string and <c>d</c> a date. The shared items of
/// a cache field and the records of the cache write and read a value the same way, so the mapping
/// between the type of a value and the name of its element lives here and nowhere else (spec 41).
/// </summary>
/// <remarks>
/// A record can also hold <c>x</c>, an index into the shared items of its field. That is not a
/// value, so the records part reads and writes it itself.
/// </remarks>
internal static class PivotCacheValueCodec
{
    /// <summary>
    /// Adds the value an item element holds to <paramref name="sink"/>.
    /// </summary>
    /// <exception cref="PartStructureException">The element is not one of the value elements.</exception>
    internal static void ReadValue(IXLPivotCacheValueSink sink, OpenXmlElement item)
    {
        switch (item)
        {
            case MissingItem:
                sink.AddMissing();
                break;
            case NumberItem numberItem:
                sink.AddNumber(numberItem.Val?.Value ?? throw PartStructureException.MissingAttribute());
                break;
            case BooleanItem booleanItem:
                sink.AddBoolean(booleanItem.Val?.Value ?? throw PartStructureException.MissingAttribute());
                break;
            case ErrorItem errorItem:
                sink.AddError(GetErrorValue(errorItem));
                break;
            case StringItem stringItem:
                sink.AddString(stringItem.Val?.Value ?? throw PartStructureException.MissingAttribute());
                break;
            case DateTimeItem dateTimeItem:
                sink.AddDateTime(dateTimeItem.Val?.Value ?? throw PartStructureException.MissingAttribute());
                break;
            default:
                throw PartStructureException.ExpectedElementNotFound();
        }
    }

    private static XLError GetErrorValue(ErrorItem errorItem)
    {
        var errorText = errorItem.Val?.Value ?? throw PartStructureException.MissingAttribute();
        if (!XLErrorParser.TryParseError(errorText, out var error))
            throw PartStructureException.IncorrectAttributeFormat();

        return error;
    }

    /// <summary>
    /// The element a shared item of <paramref name="value"/> is written as. The cache value is written
    /// as it is: going through <see cref="XLCellValue"/> would send a date through the cell serial
    /// conversion, which refuses a date on 1900-02-28 with a time part.
    /// </summary>
    /// <param name="value">The value of the shared item. Not an index.</param>
    /// <param name="text">The text of <paramref name="value"/>, when it is a string.</param>
    internal static OpenXmlElement ToSharedItem(XLPivotCacheValue value, string? text)
    {
        return value.Type switch
        {
            XLPivotCacheValueType.Missing => new MissingItem(),
            XLPivotCacheValueType.Boolean => new BooleanItem { Val = value.GetBoolean() },
            XLPivotCacheValueType.Number => new NumberItem { Val = value.GetNumber() },
            XLPivotCacheValueType.String => new StringItem { Val = text },
            XLPivotCacheValueType.Error => new ErrorItem { Val = value.GetError().ToDisplayString() },
            XLPivotCacheValueType.DateTime => new DateTimeItem { Val = value.GetDateTime() },
            _ => throw new InvalidOperationException(),
        };
    }

    /// <summary>
    /// Writes the element of a record that holds <paramref name="value"/>: a value as the element its
    /// type is spelled with, or an index into the shared items of its field.
    /// </summary>
    /// <param name="xml">The writer of the records part.</param>
    /// <param name="value">The value of the record.</param>
    /// <param name="text">The text of <paramref name="value"/>, when it is a string.</param>
    internal static void WriteRecordValue(XmlWriter xml, XLPivotCacheValue value, string? text)
    {
        switch (value.Type)
        {
            case XLPivotCacheValueType.Missing:
                xml.WriteEmptyElement("m");
                break;
            case XLPivotCacheValueType.Number:
                xml.WriteStartElement("n");
                // As a cell writes it, so the record matches its source cell (#741).
                xml.WriteAttribute("v", value.GetNumber(), XmlWriterExtensions.RoundTripNumberFormat);
                xml.WriteEndElement();
                break;
            case XLPivotCacheValueType.Boolean:
                xml.WriteStartElement("b");
                xml.WriteAttribute("v", value.GetBoolean());
                xml.WriteEndElement();
                break;
            case XLPivotCacheValueType.Error:
                xml.WriteStartElement("e");
                xml.WriteAttribute("v", value.GetError().ToDisplayString());
                xml.WriteEndElement();
                break;
            case XLPivotCacheValueType.String:
                xml.WriteStartElement("s");
                xml.WriteAttribute("v", text ?? throw new InvalidOperationException());
                xml.WriteEndElement();
                break;
            case XLPivotCacheValueType.DateTime:
                xml.WriteStartElement("d");
                xml.WriteAttribute("v", value.GetDateTime());
                xml.WriteEndElement();
                break;
            case XLPivotCacheValueType.Index:
                xml.WriteStartElement("x");
                xml.WriteAttribute("v", value.GetIndex());
                xml.WriteEndElement();
                break;
            default:
                throw new NotSupportedException();
        }
    }
}
