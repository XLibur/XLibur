using System;
using System.Globalization;
using System.IO;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;

namespace XLibur.Excel.IO;

/// <summary>
/// Writes a sheet's <c>xl/threadedComments/threadedComment{N}.xml</c>.
/// </summary>
internal static class ThreadedCommentPartWriter
{
    /// <summary>
    /// The timestamp format Excel writes: no time zone designator, hundredths of a second, and UTC
    /// by convention. Writing more precision than this makes Excel rewrite the part on open.
    /// </summary>
    private const string DateFormat = "yyyy-MM-ddTHH:mm:ss.ff";

    internal static void GenerateContent(WorksheetThreadedCommentsPart part, XLWorksheet worksheet)
    {
        var settings = new XmlWriterSettings
        {
            CloseOutput = true,
            Encoding = XLHelper.NoBomUTF8
        };

        var partStream = part.GetStream(FileMode.Create);
        using var xml = XmlWriter.Create(partStream, settings);

        xml.WriteStartElement("ThreadedComments", PersonPartWriter.ThreadedCommentsNs);
        xml.WriteAttributeString("xmlns", "x", null, OpenXmlConst.Main2006SsNs);

        var refBuffer = new char[10];
        foreach (var (point, root) in worksheet.Internals.CellsCollection.GetThreadedComments())
        {
            var reference = new string(refBuffer, 0, point.Format(refBuffer));

            WriteComment(xml, root, reference, root);
            if (root.RepliesInternal is { } replies)
            {
                foreach (var reply in replies)
                    WriteComment(xml, reply, reference, root);
            }
        }

        xml.WriteEndElement(); // ThreadedComments
        xml.Close();
    }

    private static void WriteComment(XmlWriter xml, XLThreadedComment comment, string reference,
        XLThreadedComment root)
    {
        var isRoot = ReferenceEquals(comment, root);

        xml.WriteStartElement("threadedComment", PersonPartWriter.ThreadedCommentsNs);
        xml.WriteAttributeString("ref", reference);
        xml.WriteAttributeString("dT", comment.CreatedUtc.ToString(DateFormat, CultureInfo.InvariantCulture));
        xml.WriteAttributeString("personId", CommentWriteSource.FormatId(comment.Author.Id));
        xml.WriteAttributeString("id", CommentWriteSource.FormatId(comment.Id));

        if (!isRoot)
            xml.WriteAttributeString("parentId", CommentWriteSource.FormatId(root.Id));

        // 'done' is a property of the thread and only meaningful on the root. Excel omits it when
        // the thread is unresolved.
        if (isRoot && comment.Resolved)
            xml.WriteAttributeString("done", OpenXmlConst.TrueValue);

        xml.WriteElementString("text", PersonPartWriter.ThreadedCommentsNs, comment.Text);

        // Mentions are preserved verbatim from the file they were loaded from. There is no API to
        // build or inspect them, and editing the text drops them because their offsets index into
        // the text they were written against.
        if (comment.MentionsXml is { } mentions)
            xml.WriteRaw(mentions);

        xml.WriteEndElement(); // threadedComment
    }

    /// <summary>
    /// Truncates a timestamp to the precision <see cref="DateFormat"/> persists, so that a thread
    /// created through the API holds exactly the value a later load will read back.
    /// </summary>
    internal static DateTime TruncateToFilePrecision(DateTime value)
    {
        const long centisecond = TimeSpan.TicksPerMillisecond * 10;
        return new DateTime(value.Ticks - value.Ticks % centisecond, value.Kind);
    }
}
