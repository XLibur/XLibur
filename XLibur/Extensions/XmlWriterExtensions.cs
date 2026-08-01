using System;
using System.Globalization;
using System.Xml;
using XLibur.Excel;
using XLibur.Excel.IO;

#pragma warning disable S1244 // Intentional exact float comparison for Excel formula compatibility

namespace XLibur.Extensions;

internal static class XmlWriterExtensions
{
    /// <summary>
    /// Enough for any <c>double</c> in "G15" form (longest is 22 chars, e.g.
    /// <c>-1.23456789012345E-308</c>) and for any <c>int</c>/<c>uint</c>.
    /// </summary>
    private const int NumberBufferLength = 32;

    [ThreadStatic] private static char[]? _tNumberBuffer;

    extension(XmlWriter w)
    {
        public void WriteAttribute(string attrName, string value)
        {
            w.WriteStartAttribute(attrName);
            w.WriteValue(value);
            w.WriteEndAttribute();
        }

        public void WriteAttribute(string attrName, int value)
        {
            w.WriteStartAttribute(attrName);
            w.WriteNumberValue(value);
            w.WriteEndAttribute();
        }

        public void WriteAttribute(string attrName, uint value)
        {
            w.WriteStartAttribute(attrName);
            w.WriteNumberValue(value);
            w.WriteEndAttribute();
        }

        public void WriteAttribute(string attrName, double value)
        {
            w.WriteStartAttribute(attrName);
            w.WriteNumberValue(value);
            w.WriteEndAttribute();
        }

        public void WriteAttribute(string attrName, bool value)
        {
            w.WriteStartAttribute(attrName);
            w.WriteValue(value ? "1" : "0");
            w.WriteEndAttribute();
        }

        /// <summary>
        /// Write date in a format <c>2015-01-01T00:00:00</c> (ignore kind).
        /// </summary>
        public void WriteAttribute(string attrName, DateTime value)
        {
            w.WriteStartAttribute(attrName);
            w.WriteValue(value.ToString("s"));
            w.WriteEndAttribute();
        }

        public void WriteAttribute(string attrName, string ns, double value)
        {
            w.WriteStartAttribute(attrName, ns);
            w.WriteNumberValue(value);
            w.WriteEndAttribute();
        }

        public void WriteAttributeOptional(string attrName, string? value)
        {
            if (!string.IsNullOrEmpty(value))
                w.WriteAttribute(attrName, value);
        }

        public void WriteAttributeOptional(string attrName, uint? value)
        {
            if (value is not null)
                w.WriteAttribute(attrName, value.Value);
        }

        public void WriteAttributeOptional(string attrName, int? value)
        {
            if (value is not null)
                w.WriteAttribute(attrName, value.Value);
        }

        public void WriteAttributeOptional(string attrName, bool? value)
        {
            if (value is not null)
                w.WriteAttribute(attrName, value.Value);
        }

        public void WriteAttributeDefault(string attrName, bool value, bool defaultValue)
        {
            if (value != defaultValue)
                w.WriteAttribute(attrName, value);
        }

        public void WriteAttributeDefault(string attrName, int value, int defaultValue)
        {
            if (value != defaultValue)
                w.WriteAttribute(attrName, value);
        }

        public void WriteAttributeDefault(string attrName, uint value, uint defaultValue)
        {
            if (value != defaultValue)
                w.WriteAttribute(attrName, value);
        }

        /// <summary>
        /// Write a double using the same "G15" invariant representation as
        /// <see cref="ObjectExtensions.ToInvariantString{T}"/>, but formatted into a reusable
        /// buffer instead of allocating a string per value.
        /// </summary>
        public void WriteNumberValue(double value)
        {
            var buffer = _tNumberBuffer ??= new char[NumberBufferLength];
            if (!value.TryFormat(buffer, out var charsWritten, "G15", CultureInfo.InvariantCulture))
            {
                // Unreachable for double with "G15" (longest output is ~22 chars), but a silent
                // empty attribute would corrupt the file, so fall back rather than trust it.
                w.WriteString(value.ToString("G15", CultureInfo.InvariantCulture));
                return;
            }

            w.WriteRaw(buffer, 0, charsWritten);
        }

        /// <summary>
        /// Write an <see cref="int"/> without going through <see cref="XmlWriter.WriteValue(int)"/>,
        /// which allocates a string for every value.
        /// </summary>
        public void WriteNumberValue(int value)
        {
            var buffer = _tNumberBuffer ??= new char[NumberBufferLength];
            if (!value.TryFormat(buffer, out var charsWritten, default, CultureInfo.InvariantCulture))
            {
                w.WriteString(value.ToString(CultureInfo.InvariantCulture));
                return;
            }

            w.WriteRaw(buffer, 0, charsWritten);
        }

        /// <summary>
        /// Write a <see cref="uint"/> without going through <c>XmlWriter.WriteValue</c>,
        /// which allocates a string for every value.
        /// </summary>
        public void WriteNumberValue(uint value)
        {
            var buffer = _tNumberBuffer ??= new char[NumberBufferLength];
            if (!value.TryFormat(buffer, out var charsWritten, default, CultureInfo.InvariantCulture))
            {
                w.WriteString(value.ToString(CultureInfo.InvariantCulture));
                return;
            }

            w.WriteRaw(buffer, 0, charsWritten);
        }

        public void WritePreserveSpaceAttr()
        {
            w.WriteAttributeString("xml", "space", OpenXmlConst.Xml1998Ns, "preserve");
        }

        public void WriteEmptyElement(string elName)
        {
            w.WriteStartElement(elName, OpenXmlConst.Main2006SsNs);
            w.WriteEndElement();
        }

        public void WriteColor(string elName, XLColor xlColor, bool isDifferential = false)
        {
            w.WriteStartElement(elName, OpenXmlConst.Main2006SsNs);
            switch (xlColor.ColorType)
            {
                case XLColorType.Automatic:
                    // Only reached where the element itself is required. Callers that may omit the
                    // element - a font colour, say - should skip writing it at all for an automatic
                    // colour rather than emit auto="1".
                    w.WriteAttribute("auto", 1);
                    break;

                case XLColorType.Color:
                    w.WriteAttributeString("rgb", xlColor.Color.ToHex());
                    break;

                case XLColorType.Indexed:
                    // 64 is 'transparent' and should be ignored for differential formats
                    if (!isDifferential || xlColor.Indexed != 64)
                        w.WriteAttribute("indexed", xlColor.Indexed);
                    break;

                case XLColorType.Theme:
                    w.WriteAttribute("theme", (int)xlColor.ThemeColor);

                    if (xlColor.ThemeTint != 0)
                        w.WriteAttribute("tint", xlColor.ThemeTint);
                    break;
            }

            w.WriteEndElement();
        }
    }
}
