using System;
using System.Collections.ObjectModel;
using System.Globalization;
using XLibur.Excel;
using XLibur.Excel.IO;
using DocumentFormat.OpenXml;
using XLibur.Excel.Coordinates;

namespace XLibur.Extensions;

internal static class OpenXmlPartReaderExtensions
{
    extension(OpenXmlPartReader reader)
    {
        internal bool IsStartElement(string localName)
        {
            return reader.LocalName == localName && reader is
            { NamespaceUri: OpenXmlConst.Main2006SsNs, IsStartElement: true };
        }

        internal void MoveAhead()
        {
            if (!reader.Read())
                throw new InvalidOperationException("Unexpected end of stream.");
        }
    }

    extension(ReadOnlyCollection<OpenXmlAttribute> attributes)
    {
        internal string? GetAttribute(string name)
        {
            // Don't use foreach, performance critical
            var length = attributes.Count;
            for (var i = 0; i < length; ++i)
            {
                var attr = attributes[i];
                if (attr.LocalName == name && string.IsNullOrEmpty(attr.NamespaceUri))
                    return attr.Value;
            }

            return null;
        }

        internal string? GetAttribute(string name, string namespaceUri)
        {
            // Don't use foreach, performance critical
            var length = attributes.Count;
            for (var i = 0; i < length; ++i)
            {
                var attr = attributes[i];
                if (attr.LocalName == name && attr.NamespaceUri == namespaceUri)
                    return attr.Value;
            }

            return null;
        }

        internal bool GetBoolAttribute(string name, bool defaultValue)
        {
            var attribute = attributes.GetAttribute(name);
            return ParseBool(attribute, defaultValue);
        }

        internal int? GetIntAttribute(string name)
        {
            var attribute = attributes.GetAttribute(name);
            if (!string.IsNullOrEmpty(attribute))
                return int.Parse(attribute);

            return null;
        }

        internal uint? GetUintAttribute(string name)
        {
            var attribute = attributes.GetAttribute(name);
            if (!string.IsNullOrEmpty(attribute))
                return uint.Parse(attribute);

            return null;
        }

        internal double? GetDoubleAttribute(string name, string namespaceUri)
        {
            var attribute = attributes.GetAttribute(name, namespaceUri);
            if (!string.IsNullOrEmpty(attribute))
                return double.Parse(attribute, NumberStyles.Float, XLHelper.ParseCulture);

            return null;
        }

        internal double? GetDoubleAttribute(string name)
        {
            var attribute = attributes.GetAttribute(name);
            if (!string.IsNullOrEmpty(attribute))
                return double.Parse(attribute, NumberStyles.Float, XLHelper.ParseCulture);

            return null;
        }

        /// <summary>
        /// Get value of attribute with type <c>ST_CellRef</c>.
        /// </summary>
        internal XLSheetPoint? GetCellRefAttribute(string name)
        {
            var attribute = attributes.GetAttribute(name);
            if (!string.IsNullOrEmpty(attribute))
                return XLSheetPoint.Parse(attribute);

            return null;
        }

        /// <summary>
        /// Get value of attribute with type <c>ST_Ref</c>.
        /// </summary>
        internal XLSheetRange? GetRefAttribute(string name)
        {
            var attribute = attributes.GetAttribute(name);
            if (!string.IsNullOrEmpty(attribute))
                return XLSheetRange.Parse(attribute);

            return null;
        }
    }

    private static bool ParseBool(string? input, bool defaultValue)
    {
        if (string.IsNullOrEmpty(input))
            return defaultValue;

        var isTrue = input == "1" || string.Equals("true", input, StringComparison.OrdinalIgnoreCase);
        if (isTrue)
            return true;

        var isFalse = input == "0" || string.Equals("false", input, StringComparison.OrdinalIgnoreCase);
        return isFalse ? false : throw new FormatException($"Unable to parse '{input}' to bool.");
    }
}
