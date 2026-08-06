using System.IO;
using System.Xml;

namespace XLibur.Excel.IO;

/// <summary>
/// Creates the <see cref="XmlReader"/>s XLibur opens over package parts, with one consistent
/// hardening policy.
/// </summary>
/// <remarks>
/// A package part is untrusted input: it is whatever the producer of the file put there, and
/// XLibur is routinely pointed at documents from outside the caller's control. Every reader
/// therefore refuses DTDs outright and carries no resolver, so no part can pull in an external
/// entity or expand one internally.
/// <para>
/// .NET already defaults <see cref="XmlReaderSettings.DtdProcessing"/> to
/// <see cref="DtdProcessing.Prohibit"/> and <see cref="XmlReaderSettings.XmlResolver"/> to null,
/// so this is not fixing a live hole. It is stated explicitly and in one place so the guarantee
/// is auditable and cannot be lost by a settings object that someone later builds by hand.
/// </para>
/// </remarks>
internal static class PartXmlReader
{
    /// <summary>
    /// A reader for XLibur's own streaming parse paths, which look only at elements, attributes
    /// and text.
    /// </summary>
    /// <param name="stream">The part stream to read. Left open when the reader is disposed.</param>
    /// <param name="ignoreWhitespace">
    /// False where whitespace is content — a shared-string <c>&lt;t&gt;</c> holding only spaces is
    /// legitimate text and must survive.
    /// </param>
    internal static XmlReader Create(Stream stream, bool ignoreWhitespace = true) =>
        XmlReader.Create(stream, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreWhitespace = ignoreWhitespace,
            IgnoreComments = true,
            IgnoreProcessingInstructions = true,
            CloseInput = false,
        });

    /// <summary>
    /// A reader that preserves comments, processing instructions and whitespace, for callers that
    /// build a document model from the part rather than scanning it.
    /// </summary>
    internal static XmlReader CreateVerbatim(Stream stream) =>
        XmlReader.Create(stream, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            CloseInput = false,
        });
}
