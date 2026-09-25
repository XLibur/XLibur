using System.IO;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;

namespace XLibur.Excel.IO;

/// <summary>
/// Creates the <see cref="XmlWriter"/>s XLibur writes package parts with, so every hand-written
/// part gets the same encoding. The counterpart of <see cref="PartXmlReader"/>.
/// </summary>
/// <remarks>
/// Parts are written as UTF-8 without a byte order mark, which is what Excel writes. A settings
/// object built by hand with <c>Encoding.UTF8</c> silently adds a BOM, which is how the rich data
/// parts came to be the one exception.
/// </remarks>
internal static class PartXmlWriter
{
    /// <summary>
    /// A writer over <paramref name="part"/>, replacing any content it had. Disposing the writer
    /// closes the part stream.
    /// </summary>
    internal static XmlWriter Create(OpenXmlPart part) =>
        Create(part.GetStream(FileMode.Create), closeOutput: true);

    /// <summary>
    /// A writer over an arbitrary stream, such as a zip entry or an in-memory buffer.
    /// </summary>
    /// <param name="stream">The stream to write to.</param>
    /// <param name="closeOutput">Whether disposing the writer also closes <paramref name="stream"/>.</param>
    internal static XmlWriter Create(Stream stream, bool closeOutput) =>
        XmlWriter.Create(stream, new XmlWriterSettings
        {
            CloseOutput = closeOutput,
            Encoding = XLHelper.NoBomUTF8,
        });
}
